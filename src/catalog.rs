use std::collections::HashSet;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use windows::Win32::Graphics::DirectWrite::{
    DWRITE_FACTORY_TYPE_SHARED, DWRITE_FONT_AXIS_RANGE, DWRITE_FONT_PROPERTY_ID_FAMILY_NAME,
    DWRITE_FONT_SIMULATIONS_NONE, DWRITE_FONT_STRETCH_NORMAL, DWRITE_FONT_STYLE_NORMAL,
    DWRITE_FONT_WEIGHT_NORMAL, DWriteCreateFactory, IDWriteFactory, IDWriteFactory3,
    IDWriteFactory7, IDWriteFontFace, IDWriteFontFace5, IDWriteFontFile, IDWriteFontResource,
    IDWriteFontSet, IDWriteFontSetBuilder, IDWriteFontSetBuilder1, IDWriteLocalizedStrings,
    IDWriteStringList,
};
use windows::core::{Interface, PCWSTR};

pub(crate) type FontId = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
pub(crate) enum FontSource {
    System,
    FirstTeam,
    Library,
}

impl FontSource {
    pub fn label(self) -> &'static str {
        match self {
            Self::System => "システム",
            Self::FirstTeam => "ローカルフォント",
            Self::Library => "プラグイン専用",
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct FontAxis {
    tag: [u8; 4],
    min: f32,
    max: f32,
}

impl FontAxis {
    fn new(tag: [u8; 4], min: f32, max: f32) -> Option<Self> {
        (min.is_finite() && max.is_finite() && min < max).then_some(Self { tag, min, max })
    }

    pub fn label(self) -> String {
        format!(
            "{}: {:.1} - {:.1}",
            String::from_utf8_lossy(&self.tag),
            self.min,
            self.max
        )
    }

    #[cfg(test)]
    pub fn test_axis() -> Self {
        Self {
            tag: *b"wght",
            min: 100.0,
            max: 900.0,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct FontItem {
    pub id: FontId,
    pub display_name: String,
    pub family_name: String,
    pub path: Option<PathBuf>,
    pub source: FontSource,
    pub axes: Vec<FontAxis>,
    pub favorite: bool,
}

impl FontItem {
    pub fn new_system(family_name: String, axes: Vec<FontAxis>) -> Self {
        Self {
            id: format!("system:{}", family_name.to_lowercase()),
            display_name: family_name.clone(),
            family_name,
            path: None,
            source: FontSource::System,
            axes,
            favorite: false,
        }
    }

    pub fn new_local(
        family_name: String,
        path: PathBuf,
        source: FontSource,
        axes: Vec<FontAxis>,
    ) -> Self {
        let file_name = path
            .file_name()
            .map(|name| name.to_string_lossy())
            .unwrap_or_default();
        Self {
            id: format!(
                "local:{}:{}",
                source.label(),
                path.to_string_lossy().to_lowercase()
            ),
            display_name: format!("{family_name} [{file_name}]"),
            family_name,
            path: Some(path),
            source,
            axes,
            favorite: source == FontSource::FirstTeam,
        }
    }

    pub fn is_system(&self) -> bool {
        self.source == FontSource::System
    }

    pub fn local_path(&self) -> &str {
        self.path
            .as_ref()
            .and_then(|path| path.to_str())
            .unwrap_or("")
    }
}

pub(crate) fn enumerate(first_team_dir: &Path, library_dir: &Path) -> Result<Vec<FontItem>> {
    unsafe {
        let factory: IDWriteFactory7 = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)
            .context("DirectWrite factory creation failed")?;
        let mut fonts = system_fonts(&factory)?;
        let mut paths = HashSet::new();
        fonts.extend(local_fonts(
            &factory,
            first_team_dir,
            FontSource::FirstTeam,
            &mut paths,
        ));
        fonts.extend(local_fonts(
            &factory,
            library_dir,
            FontSource::Library,
            &mut paths,
        ));
        Ok(fonts)
    }
}

unsafe fn system_fonts(factory: &IDWriteFactory7) -> Result<Vec<FontItem>> {
    let base: IDWriteFactory = factory.cast()?;
    let mut collection = None;
    unsafe { base.GetSystemFontCollection(&mut collection, false)? };
    let collection = collection.context("system font collection was null")?;
    let mut fonts = Vec::with_capacity(collection.GetFontFamilyCount() as usize);
    for index in 0..collection.GetFontFamilyCount() {
        let Ok(family) = (unsafe { collection.GetFontFamily(index) }) else {
            continue;
        };
        let Ok(names) = (unsafe { family.GetFamilyNames() }) else {
            continue;
        };
        let Some(name) = (unsafe { localized_name(&names) }) else {
            continue;
        };
        let axes = unsafe {
            family
                .GetFirstMatchingFont(
                    DWRITE_FONT_WEIGHT_NORMAL,
                    DWRITE_FONT_STRETCH_NORMAL,
                    DWRITE_FONT_STYLE_NORMAL,
                )
                .and_then(|font| font.CreateFontFace())
                .and_then(|face: IDWriteFontFace| face.cast::<IDWriteFontFace5>())
                .map(|face| collect_axes(&face))
                .unwrap_or_default()
        };
        fonts.push(FontItem::new_system(name, axes));
    }
    Ok(fonts)
}

unsafe fn local_fonts(
    factory: &IDWriteFactory7,
    directory: &Path,
    source: FontSource,
    seen_paths: &mut HashSet<PathBuf>,
) -> Vec<FontItem> {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return Vec::new();
    };
    entries
        .flatten()
        .filter_map(|entry| {
            let path = entry.path();
            let supported = path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    matches!(
                        extension.to_ascii_lowercase().as_str(),
                        "ttf" | "otf" | "ttc"
                    )
                });
            (supported && seen_paths.insert(path.clone())).then_some(path)
        })
        .filter_map(|path| {
            let (family, axes) = unsafe { local_font_info(factory, &path) }
                .inspect_err(|error| {
                    tracing::warn!("{}を読み込めませんでした: {error}", path.display())
                })
                .ok()?;
            Some(FontItem::new_local(family, path, source, axes))
        })
        .collect()
}

unsafe fn local_font_info(
    factory: &IDWriteFactory7,
    path: &Path,
) -> windows::core::Result<(String, Vec<FontAxis>)> {
    let path_wide = wide(&path.to_string_lossy());
    let file: IDWriteFontFile =
        unsafe { factory.CreateFontFileReference(PCWSTR(path_wide.as_ptr()), None)? };
    let factory3: IDWriteFactory3 = factory.cast()?;
    let builder: IDWriteFontSetBuilder = unsafe { factory3.CreateFontSetBuilder()? };
    let builder: IDWriteFontSetBuilder1 = builder.cast()?;
    unsafe { builder.AddFontFile(&file)? };
    let set: IDWriteFontSet = unsafe { builder.CreateFontSet()? };
    let family = unsafe { font_set_family_name(&set) }.unwrap_or_else(|| {
        path.file_stem()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| "Unknown Font".to_string())
    });
    let files = [Some(file)];
    let axes = unsafe {
        factory
            .CreateFontFace(
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_FACE_TYPE_TRUETYPE,
                &files,
                0,
                DWRITE_FONT_SIMULATIONS_NONE,
            )
            .and_then(|face: IDWriteFontFace| face.cast::<IDWriteFontFace5>())
            .map(|face| collect_axes(&face))
            .unwrap_or_default()
    };
    Ok((family, axes))
}

unsafe fn localized_name(strings: &IDWriteLocalizedStrings) -> Option<String> {
    let mut index = 0;
    let mut exists = windows::core::BOOL::from(false);
    for locale in ["ja-jp", "en-us"] {
        let locale = wide(locale);
        if unsafe {
            strings
                .FindLocaleName(PCWSTR(locale.as_ptr()), &mut index, &mut exists)
                .is_ok()
        } && exists.as_bool()
        {
            break;
        }
    }
    let length = unsafe { strings.GetStringLength(index).ok()? } as usize;
    let mut buffer = vec![0u16; length + 1];
    unsafe { strings.GetString(index, &mut buffer).ok()? };
    Some(String::from_utf16_lossy(&buffer[..length]))
}

unsafe fn font_set_family_name(set: &IDWriteFontSet) -> Option<String> {
    let strings: IDWriteStringList = unsafe {
        set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_FAMILY_NAME)
            .ok()?
    };
    let length = unsafe { strings.GetStringLength(0).ok()? } as usize;
    let mut buffer = vec![0u16; length + 1];
    unsafe { strings.GetString(0, &mut buffer).ok()? };
    Some(String::from_utf16_lossy(&buffer[..length]))
}

unsafe fn collect_axes(face: &IDWriteFontFace5) -> Vec<FontAxis> {
    let Ok(resource): windows::core::Result<IDWriteFontResource> =
        (unsafe { face.GetFontResource() })
    else {
        return Vec::new();
    };
    let count = resource.GetFontAxisCount();
    let mut ranges = vec![DWRITE_FONT_AXIS_RANGE::default(); count as usize];
    if count > 0 && unsafe { resource.GetFontAxisRanges(&mut ranges) }.is_err() {
        return Vec::new();
    }
    ranges
        .into_iter()
        .filter_map(|range| {
            FontAxis::new(
                range.axisTag.0.to_le_bytes(),
                range.minValue,
                range.maxValue,
            )
        })
        .collect()
}

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_labels_are_stable() {
        assert_eq!(FontSource::System.label(), "システム");
        assert_eq!(FontSource::FirstTeam.label(), "ローカルフォント");
        assert_eq!(FontSource::Library.label(), "プラグイン専用");
    }

    #[test]
    fn fixed_axis_is_not_variable() {
        assert!(FontAxis::new(*b"wght", 400.0, 400.0).is_none());
    }

    #[test]
    fn variable_axis_requires_a_finite_range() {
        assert!(FontAxis::new(*b"wght", 100.0, 900.0).is_some());
        assert!(FontAxis::new(*b"wght", f32::NAN, 900.0).is_none());
        assert!(FontAxis::new(*b"wght", 900.0, 100.0).is_none());
    }
}
