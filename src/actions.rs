use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
use aviutl2::generic::EditHandle;

use crate::alias::{self, ObjectKind};
use crate::catalog::{FontItem, FontSource};

const DEFAULT_OBJECT_SECONDS: f64 = 1.1;
const TEXT_EFFECTS: [&str; 3] = ["テキスト", "Variable Font Text", "Variable Font Object"];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum FontImport {
    Copied(PathBuf),
    AlreadyPresent(PathBuf),
}

pub(crate) fn import_font_file(source: &Path, font_dir: &Path) -> Result<FontImport> {
    if !source.is_file() {
        bail!(
            "ドロップされたフォントファイルが見つかりません: {}",
            source.display()
        );
    }
    let extension = source
        .extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase)
        .context("フォントファイルの拡張子を取得できません")?;
    if !matches!(extension.as_str(), "ttf" | "otf" | "ttc") {
        bail!("対応していないフォント形式です: .{extension}");
    }

    let file_name = source
        .file_name()
        .context("フォントファイル名を取得できません")?;
    std::fs::create_dir_all(font_dir)
        .with_context(|| format!("Fontフォルダを作成できません: {}", font_dir.display()))?;
    let target = font_dir.join(file_name);
    if target.exists() {
        let source_path = source.canonicalize().ok();
        let target_path = target.canonicalize().ok();
        if source_path.is_some() && source_path == target_path {
            return Ok(FontImport::AlreadyPresent(target));
        }
        bail!("Fontフォルダに同名ファイルがあります: {}", target.display());
    }

    std::fs::copy(source, &target).with_context(|| {
        format!(
            "フォントファイルをコピーできませんでした: {} -> {}",
            source.display(),
            target.display()
        )
    })?;
    Ok(FontImport::Copied(target))
}

pub(crate) fn create_object(
    handle: &EditHandle,
    font: &FontItem,
    text: &str,
    kind: ObjectKind,
) -> Result<()> {
    if !handle.is_ready() {
        bail!("AviUtl2の編集機能を初期化中です");
    }
    let started = std::time::Instant::now();
    tracing::debug!(
        font = %font.display_name,
        text_len = text.len(),
        "FontPreview create object start"
    );
    let info = handle.get_edit_info();
    let alias = alias::build(
        font,
        text,
        alias::frame_length(info.fps, DEFAULT_OBJECT_SECONDS),
        kind,
    );
    handle
        .call_edit_section(move |edit| {
            edit.create_object_from_alias(&alias, edit.info.layer, edit.info.frame, 0)?;
            Ok::<_, aviutl2::generic::EditSectionError>(())
        })
        .context("編集セクションを開けませんでした")??;
    tracing::debug!(
        elapsed_ms = started.elapsed().as_millis(),
        "FontPreview create object finish"
    );
    Ok(())
}

pub(crate) fn apply_to_selection(handle: &EditHandle, font: &FontItem) -> Result<usize> {
    if !handle.is_ready() {
        bail!("AviUtl2の編集機能を初期化中です");
    }
    let started = std::time::Instant::now();
    tracing::debug!(font = %font.display_name, "FontPreview apply to selection start");
    let family = font.family_name.clone();
    let file = font.local_path().to_string();
    let system = font.is_system();
    let count = handle
        .call_edit_section(move |edit| {
            let mut objects = edit.get_selected_objects()?;
            if objects.is_empty() {
                objects.extend(edit.get_focused_object()?);
            }
            let mut updated = 0;
            for object in objects {
                let mut changed = false;
                if system {
                    changed |= edit
                        .set_object_effect_item(object, "テキスト", 0, "フォント", &family)
                        .is_ok();
                }
                for effect in ["Variable Font Text", "Variable Font Object"] {
                    changed |= edit
                        .set_object_effect_item(
                            object,
                            effect,
                            0,
                            "フォント",
                            if system { &family } else { "" },
                        )
                        .is_ok();
                    changed |= edit
                        .set_object_effect_item(
                            object,
                            effect,
                            0,
                            "フォントファイル",
                            if system { "" } else { &file },
                        )
                        .is_ok();
                }
                updated += usize::from(changed);
            }
            Ok::<_, aviutl2::generic::EditSectionError>(updated)
        })
        .context("編集セクションを開けませんでした")??;
    if count == 0 {
        bail!("更新できるオブジェクトが選択されていません");
    }
    tracing::debug!(
        count,
        elapsed_ms = started.elapsed().as_millis(),
        "FontPreview apply to selection finish"
    );
    Ok(count)
}

pub(crate) fn selected_text(handle: &EditHandle) -> Result<Option<String>> {
    if !handle.is_ready() {
        return Ok(None);
    }
    let started = std::time::Instant::now();
    tracing::debug!("FontPreview selected_text edit section start");
    let result = handle
        .call_read_section(|edit| {
            let focused = edit.get_focused_object()?;
            let (object, source, selected_count) = if let Some(object) = focused {
                (Some(object), "focused", 0)
            } else {
                let selected = edit.get_selected_objects()?;
                let count = selected.len();
                (selected.into_iter().next(), "selection", count)
            };
            tracing::debug!(
                source,
                selected_count,
                has_object = object.is_some(),
                "FontPreview selected_text object resolved"
            );
            let Some(object) = object else {
                return Ok::<_, aviutl2::generic::EditSectionError>(None);
            };
            for effect in TEXT_EFFECTS {
                match edit.get_object_effect_item(object, effect, 0, "テキスト") {
                    Ok(text) => {
                        tracing::debug!(
                            effect,
                            text_len = text.len(),
                            "FontPreview selected_text effect item read"
                        );
                        if !text.is_empty() {
                            return Ok(Some(text));
                        }
                    }
                    Err(error) => tracing::debug!(
                        effect,
                        error = %error,
                        "FontPreview selected_text effect item unavailable"
                    ),
                }
            }
            Ok::<_, aviutl2::generic::EditSectionError>(None)
        })
        .context("編集セクションを参照できませんでした")?
        .context("テキスト取得に失敗しました");
    tracing::debug!(
        elapsed_ms = started.elapsed().as_millis(),
        has_text = result
            .as_ref()
            .ok()
            .and_then(|text| text.as_ref())
            .is_some(),
        "FontPreview selected_text edit section finish"
    );
    result
}

pub(crate) fn move_local_font(
    font: &FontItem,
    first_team_dir: &Path,
    library_dir: &Path,
) -> Result<PathBuf> {
    let source = font
        .path
        .as_ref()
        .context("ローカルフォントではありません")?;
    let target_dir = match font.source {
        FontSource::FirstTeam => library_dir,
        FontSource::Library => first_team_dir,
        FontSource::System => bail!("システムフォントは移動できません"),
    };
    std::fs::create_dir_all(target_dir)?;
    let file_name = source.file_name().context("ファイル名を取得できません")?;
    let target = target_dir.join(file_name);
    if target.exists() {
        bail!("移動先に同名ファイルがあります: {}", target.display());
    }
    if !source.exists() {
        bail!("フォントファイルが見つかりません: {}", source.display());
    }
    std::fs::rename(source, &target).with_context(|| {
        format!(
            "フォントを移動できませんでした: {} -> {}",
            source.display(),
            target.display()
        )
    })?;
    Ok(target)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::catalog::FontSource;

    fn unique_temp_dir(name: &str) -> PathBuf {
        std::env::temp_dir().join(format!("font-preview-{name}-{}", std::process::id()))
    }

    #[test]
    fn moves_library_font_to_first_team() {
        let root = unique_temp_dir("move");
        let first = root.join("Font");
        let library = root.join("FontLibrary");
        std::fs::create_dir_all(&library).unwrap();
        let source = library.join("test.ttf");
        std::fs::write(&source, b"font").unwrap();
        let font = FontItem::new_local("Test".to_string(), source, FontSource::Library, Vec::new());
        let target = move_local_font(&font, &first, &library).unwrap();
        assert_eq!(target, first.join("test.ttf"));
        assert!(target.exists());
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn imports_dropped_font_into_font_directory() {
        let root = unique_temp_dir("drop-import");
        let source_dir = root.join("Source");
        let font_dir = root.join("Font");
        std::fs::create_dir_all(&source_dir).unwrap();
        let source = source_dir.join("test.TTF");
        std::fs::write(&source, b"font").unwrap();

        let result = import_font_file(&source, &font_dir).unwrap();
        assert_eq!(result, FontImport::Copied(font_dir.join("test.TTF")));
        assert_eq!(std::fs::read(font_dir.join("test.TTF")).unwrap(), b"font");
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn rejects_unsupported_dropped_file() {
        let root = unique_temp_dir("drop-unsupported");
        std::fs::create_dir_all(&root).unwrap();
        let source = root.join("test.txt");
        std::fs::write(&source, b"not a font").unwrap();
        assert!(import_font_file(&source, &root.join("Font")).is_err());
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn does_not_overwrite_existing_font() {
        let root = unique_temp_dir("drop-collision");
        let source_dir = root.join("Source");
        let font_dir = root.join("Font");
        std::fs::create_dir_all(&source_dir).unwrap();
        std::fs::create_dir_all(&font_dir).unwrap();
        let source = source_dir.join("test.otf");
        let target = font_dir.join("test.otf");
        std::fs::write(&source, b"source").unwrap();
        std::fs::write(&target, b"existing").unwrap();
        assert!(import_font_file(&source, &font_dir).is_err());
        assert_eq!(std::fs::read(target).unwrap(), b"existing");
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn rejects_name_collision() {
        let root = unique_temp_dir("collision");
        let first = root.join("Font");
        let library = root.join("FontLibrary");
        std::fs::create_dir_all(&first).unwrap();
        std::fs::create_dir_all(&library).unwrap();
        let source = library.join("test.ttf");
        std::fs::write(&source, b"source").unwrap();
        std::fs::write(first.join("test.ttf"), b"target").unwrap();
        let font = FontItem::new_local("Test".to_string(), source, FontSource::Library, Vec::new());
        assert!(move_local_font(&font, &first, &library).is_err());
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn text_effect_priority_is_stable() {
        assert_eq!(
            TEXT_EFFECTS,
            ["テキスト", "Variable Font Text", "Variable Font Object"]
        );
    }
}
