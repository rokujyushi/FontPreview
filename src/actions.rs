use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
use aviutl2::generic::GlobalEditHandle;

use crate::alias::{self, ObjectKind};
use crate::catalog::{FontItem, FontSource};

const DEFAULT_OBJECT_SECONDS: f64 = 1.1;
const TEXT_EFFECTS: [&str; 3] = ["テキスト", "Variable Font Text", "Variable Font Object"];

pub(crate) fn create_object(
    handle: &GlobalEditHandle,
    font: &FontItem,
    text: &str,
    kind: ObjectKind,
) -> Result<()> {
    if !handle.is_ready() {
        bail!("AviUtl2の編集機能を初期化中です");
    }
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
    Ok(())
}

pub(crate) fn apply_to_selection(handle: &GlobalEditHandle, font: &FontItem) -> Result<usize> {
    if !handle.is_ready() {
        bail!("AviUtl2の編集機能を初期化中です");
    }
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
    Ok(count)
}

pub(crate) fn selected_text(handle: &GlobalEditHandle) -> Result<Option<String>> {
    if !handle.is_ready() {
        return Ok(None);
    }
    handle
        .call_read_section(|edit| {
            let object = edit
                .get_focused_object()?
                .or_else(|| edit.get_selected_objects().ok()?.into_iter().next());
            let Some(object) = object else {
                return Ok::<_, aviutl2::generic::EditSectionError>(None);
            };
            Ok::<_, aviutl2::generic::EditSectionError>(TEXT_EFFECTS.iter().find_map(|effect| {
                edit.get_object_effect_item(object, effect, 0, "テキスト")
                    .ok()
                    .filter(|text| !text.is_empty())
            }))
        })
        .context("編集セクションを参照できませんでした")?
        .context("テキスト取得に失敗しました")
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
