use std::collections::HashSet;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

use crate::catalog::{FontId, FontItem, FontSource};
use crate::i18n::format_text;

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

/// リストをキーボードで操作するときの、種別ごとの有効/無効。
///
/// 矢印キーだけ、WASDだけ、といった使い分けをするために分けている。
/// `enabled` はマスタースイッチで、これがoffなら他がどうであれ全て無効。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub(crate) struct KeyBindings {
    /// マスタースイッチ。
    pub enabled: bool,
    /// ↑↓←→ / PageUp PageDown / Home End。
    pub arrows: bool,
    /// W A S D。
    pub letters: bool,
    /// Enter で選択中に適用。
    pub enter_applies: bool,
    /// F でお気に入り切替。
    pub favorite: bool,
    /// 1 2 3 でオブジェクトを追加。
    pub create_objects: bool,
}

impl Default for KeyBindings {
    fn default() -> Self {
        Self {
            enabled: true,
            arrows: true,
            letters: true,
            enter_applies: true,
            favorite: true,
            create_objects: true,
        }
    }
}

impl KeyBindings {
    // マスタースイッチの取りこぼしを防ぐため、問い合わせは必ずこれらを通す。
    pub fn arrows_active(self) -> bool {
        self.enabled && self.arrows
    }

    pub fn letters_active(self) -> bool {
        self.enabled && self.letters
    }

    pub fn enter_active(self) -> bool {
        self.enabled && self.enter_applies
    }

    pub fn favorite_active(self) -> bool {
        self.enabled && self.favorite
    }

    pub fn create_active(self) -> bool {
        self.enabled && self.create_objects
    }
}

/// 詳細列に並ぶオブジェクト追加ボタンの表示。
///
/// Variable Font 系は別のプラグインが入っていない環境では押しても失敗するので、
/// 使わないボタンを隔しておけるようにする。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub(crate) struct CreateButtons {
    /// 標準の「テキスト」オブジェクト。
    pub text: bool,
    /// 「Variable Font Text」オブジェクト。
    pub variable_font_text: bool,
    /// 「Variable Font Object」オブジェクト。
    pub variable_font_object: bool,
}

impl Default for CreateButtons {
    fn default() -> Self {
        Self {
            text: true,
            variable_font_text: true,
            variable_font_object: true,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub(crate) struct Settings {
    pub system_favorites: HashSet<FontId>,
    pub local_favorites: HashSet<FontId>,
    pub move_local_fonts_with_favorites: bool,
    /// リストのキーボード操作。
    pub keys: KeyBindings,
    /// オブジェクト追加ボタンの表示。
    pub create_buttons: CreateButtons,
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
            keys: KeyBindings::default(),
            create_buttons: CreateButtons::default(),
            filter: FilterMode::All,
            sort: SortMode::FavoriteName,
            sync_text: false,
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
            Ok(content) => match serde_json::from_str::<Self>(&content) {
                Ok(settings) => (settings, None),
                Err(error) => (
                    Self::default(),
                    Some(format_text(
                        "設定ファイルを読み込めませんでした: {error}",
                        &[("{error}", error.to_string())],
                    )),
                ),
            },
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => (Self::default(), None),
            Err(error) => (
                Self::default(),
                Some(format_text(
                    "設定ファイルを読み込めませんでした: {error}",
                    &[("{error}", error.to_string())],
                )),
            ),
        }
    }

    pub fn save(&self, path: &Path) -> Result<()> {
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = serde_json::to_string_pretty(self)?;
        std::fs::write(path, content).with_context(|| {
            format_text(
                "設定を書き込めませんでした: {path}",
                &[("{path}", path.display().to_string())],
            )
        })
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
    fn key_bindings_default_to_on_for_existing_settings() {
        // コンテナの`#[serde(default)]`は欠落フィールドを`Default`から埋める。
        // 既存のsettings.jsonにこのキーは無いので、全てtrueになることを固定する。
        let loaded: Settings = serde_json::from_str(r#"{"sample": "test"}"#).unwrap();
        assert_eq!(loaded.keys, KeyBindings::default());
        assert!(loaded.keys.arrows_active());
    }

    #[test]
    fn create_buttons_default_to_visible() {
        let loaded: Settings = serde_json::from_str(r#"{"sample": "test"}"#).unwrap();
        assert_eq!(loaded.create_buttons, CreateButtons::default());
        assert!(loaded.create_buttons.text);
    }

    #[test]
    fn hiding_one_create_button_keeps_the_others() {
        let loaded: Settings =
            serde_json::from_str(r#"{"create_buttons": {"variable_font_text": false}}"#).unwrap();
        assert!(!loaded.create_buttons.variable_font_text);
        assert!(loaded.create_buttons.text);
        assert!(loaded.create_buttons.variable_font_object);
    }

    #[test]
    fn partially_written_key_bindings_fill_in_the_rest() {
        let loaded: Settings = serde_json::from_str(r#"{"keys": {"letters": false}}"#).unwrap();
        assert!(!loaded.keys.letters);
        assert!(loaded.keys.enabled);
        assert!(loaded.keys.arrows);
    }

    #[test]
    fn the_master_switch_overrides_every_group() {
        let off = KeyBindings {
            enabled: false,
            ..KeyBindings::default()
        };
        assert!(!off.arrows_active());
        assert!(!off.letters_active());
        assert!(!off.enter_active());
        assert!(!off.favorite_active());
        assert!(!off.create_active());
    }

    #[test]
    fn groups_can_be_used_one_at_a_time() {
        let arrows_only = KeyBindings {
            letters: false,
            ..KeyBindings::default()
        };
        assert!(arrows_only.arrows_active());
        assert!(!arrows_only.letters_active());

        let letters_only = KeyBindings {
            arrows: false,
            ..KeyBindings::default()
        };
        assert!(!letters_only.arrows_active());
        assert!(letters_only.letters_active());
    }

    #[test]
    fn default_json_round_trip() {
        let value = serde_json::to_string(&Settings::default()).unwrap();
        let loaded: Settings = serde_json::from_str(&value).unwrap();
        assert_eq!(loaded.sort, SortMode::FavoriteName);
        assert!(!loaded.sync_text);
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

    #[test]
    fn sync_text_is_loaded_from_settings_file() {
        let dir =
            std::env::temp_dir().join(format!("font-preview-settings-{}", std::process::id()));
        let path = dir.join("settings.json");
        std::fs::create_dir_all(&dir).unwrap();
        std::fs::write(
            &path,
            r#"{
  "sync_text": true,
  "sample": "test"
}"#,
        )
        .unwrap();
        let (settings, error) = Settings::load(&path);
        assert!(error.is_none());
        assert!(settings.sync_text);
        std::fs::remove_dir_all(dir).unwrap();
    }
}
