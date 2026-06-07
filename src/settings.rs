use std::collections::HashSet;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

use crate::catalog::{FontId, FontItem, FontSource};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub(crate) enum FilterMode {
    All,
    System,
    FirstTeam,
    Library,
    Favorites,
    Variable,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub(crate) enum SortMode {
    FavoriteName,
    NameAsc,
    NameDesc,
    Source,
    VariableFirst,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub(crate) struct Settings {
    pub system_favorites: HashSet<FontId>,
    pub local_favorites: HashSet<FontId>,
    pub move_local_fonts_with_favorites: bool,
    pub filter: FilterMode,
    pub sort: SortMode,
    pub sync_text: bool,
    pub sample: String,
    pub preview_font_size: f32,
    pub preview_text_color: [u8; 3],
    pub preview_background_color: [u8; 3],
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            system_favorites: HashSet::new(),
            local_favorites: HashSet::new(),
            move_local_fonts_with_favorites: false,
            filter: FilterMode::All,
            sort: SortMode::FavoriteName,
            sync_text: true,
            sample: "あいうABC123".to_string(),
            preview_font_size: 48.0,
            preview_text_color: [0, 0, 0],
            preview_background_color: [255, 255, 255],
        }
    }
}

impl Settings {
    pub fn load(path: &Path) -> (Self, Option<String>) {
        match std::fs::read_to_string(path) {
            Ok(content) => match serde_json::from_str(&content) {
                Ok(settings) => (settings, None),
                Err(error) => (
                    Self::default(),
                    Some(format!("設定ファイルを読み込めませんでした: {error}")),
                ),
            },
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => (Self::default(), None),
            Err(error) => (
                Self::default(),
                Some(format!("設定ファイルを読み込めませんでした: {error}")),
            ),
        }
    }

    pub fn save(&self, path: &Path) -> Result<()> {
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = serde_json::to_string_pretty(self)?;
        std::fs::write(path, content)
            .with_context(|| format!("設定を書き込めませんでした: {}", path.display()))
    }

    pub fn apply_favorites(&self, fonts: &mut [FontItem]) {
        for font in fonts {
            font.favorite = match font.source {
                FontSource::System => self.system_favorites.contains(&font.id),
                FontSource::FirstTeam | FontSource::Library => {
                    self.local_favorites.contains(&local_favorite_id(font))
                }
            };
        }
    }

    pub fn toggle_favorite(&mut self, font: &FontItem) -> bool {
        let (favorites, id) = match font.source {
            FontSource::System => (&mut self.system_favorites, font.id.clone()),
            FontSource::FirstTeam | FontSource::Library => {
                (&mut self.local_favorites, local_favorite_id(font))
            }
        };
        if !favorites.remove(&id) {
            favorites.insert(id);
            true
        } else {
            false
        }
    }
}

fn local_favorite_id(font: &FontItem) -> FontId {
    let file_name = font
        .path
        .as_ref()
        .and_then(|path| path.file_name())
        .map(|name| name.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    format!("local:{}:{}", font.family_name.to_lowercase(), file_name)
}

pub(crate) fn settings_path(app_data: &Path) -> PathBuf {
    app_data.join("Plugin/FontPreview/settings.json")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn toggles_system_favorite() {
        let mut settings = Settings::default();
        let font = FontItem::new_system("Test".to_string(), Vec::new());
        assert!(settings.toggle_favorite(&font));
        assert!(!settings.toggle_favorite(&font));
    }

    #[test]
    fn local_favorite_survives_source_move() {
        let mut settings = Settings::default();
        let library = FontItem::new_local(
            "Test".to_string(),
            PathBuf::from("FontLibrary/Test.ttf"),
            FontSource::Library,
            Vec::new(),
        );
        assert!(settings.toggle_favorite(&library));

        let mut fonts = vec![FontItem::new_local(
            "Test".to_string(),
            PathBuf::from("Font/Test.ttf"),
            FontSource::FirstTeam,
            Vec::new(),
        )];
        settings.apply_favorites(&mut fonts);
        assert!(fonts[0].favorite);
    }

    #[test]
    fn default_json_round_trip() {
        let value = serde_json::to_string(&Settings::default()).unwrap();
        let loaded: Settings = serde_json::from_str(&value).unwrap();
        assert_eq!(loaded.sort, SortMode::FavoriteName);
        assert!(loaded.sync_text);
        assert!(!loaded.move_local_fonts_with_favorites);
        assert_eq!(loaded.preview_font_size, 48.0);
        assert_eq!(loaded.preview_text_color, [0, 0, 0]);
    }

    #[test]
    fn unknown_favorite_id_is_ignored() {
        let mut settings = Settings::default();
        settings
            .system_favorites
            .insert("system:missing".to_string());
        let mut fonts = vec![FontItem::new_system("Known".to_string(), Vec::new())];
        settings.apply_favorites(&mut fonts);
        assert!(!fonts[0].favorite);
    }
}
