use crate::catalog::FontItem;

const FALLBACK_LENGTH: usize = 182;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ObjectKind {
    Text,
    VariableFontText,
    VariableFontObject,
}

impl ObjectKind {
    /// 表示順。追加ボタンの並びと、キー操作の数字の両方がこの順に従う。
    pub const ALL: [Self; 3] = [Self::Text, Self::VariableFontText, Self::VariableFontObject];

    /// オブジェクト追加に割り当てる数字キー。1始まりで `ALL` の並びと一致する。
    pub fn shortcut_digit(self) -> usize {
        match self {
            Self::Text => 1,
            Self::VariableFontText => 2,
            Self::VariableFontObject => 3,
        }
    }

    /// 追加ボタンのラベル。設定ウィンドウのチェックボックスと共用する。
    pub fn button_label(self) -> &'static str {
        match self {
            Self::Text => "テキスト +",
            Self::VariableFontText => "VF +",
            Self::VariableFontObject => "VFO +",
        }
    }

    /// ボタンが何を作るのか。VF系は別プラグインが必要なのでそれも書く。
    pub fn description(self) -> &'static str {
        match self {
            Self::Text => "標準の「テキスト」オブジェクトを追加します",
            Self::VariableFontText => {
                "「Variable Font Text」オブジェクトを追加します（対応プラグインが必要）"
            }
            Self::VariableFontObject => {
                "「Variable Font Object」オブジェクトを追加します（対応プラグインが必要）"
            }
        }
    }
}

pub(crate) fn frame_length(fps: aviutl2::common::Rational32, seconds: f64) -> usize {
    if seconds <= 0.0 || *fps.numer() <= 0 || *fps.denom() <= 0 {
        return FALLBACK_LENGTH;
    }
    (seconds * f64::from(*fps.numer()) / f64::from(*fps.denom()))
        .ceil()
        .max(1.0) as usize
}

pub(crate) fn build(font: &FontItem, text: &str, length: usize, kind: ObjectKind) -> String {
    let text = text.replace("\r\n", "\n").replace(['\r', '\n'], "\\n");
    let (effect, family, file) = match kind {
        ObjectKind::Text => ("テキスト", font.family_name.as_str(), ""),
        ObjectKind::VariableFontText => (
            "Variable Font Text",
            if font.is_system() {
                font.family_name.as_str()
            } else {
                ""
            },
            font.local_path(),
        ),
        ObjectKind::VariableFontObject => (
            "Variable Font Object",
            if font.is_system() {
                font.family_name.as_str()
            } else {
                ""
            },
            font.local_path(),
        ),
    };
    let mut alias = format!(
        "[Object]\nframe=0,{}\n[Object.0]\neffect.name={effect}\nフォント={family}\n",
        length.max(1)
    );
    if !file.is_empty() {
        alias.push_str(&format!("フォントファイル={file}\n"));
    }
    if !text.is_empty() {
        alias.push_str(&format!("テキスト={text}\n"));
    }
    alias
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    #[test]
    fn builds_local_variable_font_alias() {
        let font = FontItem::new_local(
            "Test".to_string(),
            PathBuf::from(r"C:\Fonts\Test.ttf"),
            crate::catalog::FontSource::Library,
            Vec::new(),
        );
        let alias = build(&font, "sample", 30, ObjectKind::VariableFontText);
        assert!(alias.contains("effect.name=Variable Font Text"));
        assert!(alias.contains(r"フォントファイル=C:\Fonts\Test.ttf"));
        assert!(alias.contains("テキスト=sample"));
    }

    #[test]
    fn every_kind_has_a_label_and_a_description() {
        for kind in ObjectKind::ALL {
            assert!(!kind.button_label().is_empty());
            assert!(!kind.description().is_empty());
        }
    }

    #[test]
    fn shortcut_digits_follow_the_button_order() {
        // 設定ウィンドウの表示とキー割り当てがずれないことを固定する。
        for (index, kind) in ObjectKind::ALL.into_iter().enumerate() {
            assert_eq!(kind.shortcut_digit(), index + 1);
        }
    }

    #[test]
    fn computes_ntsc_length_with_ceiling() {
        assert_eq!(
            frame_length(aviutl2::common::Rational32::new(30_000, 1001), 1.1),
            33
        );
    }
}
