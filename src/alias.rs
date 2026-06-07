use crate::catalog::FontItem;

const FALLBACK_LENGTH: usize = 182;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ObjectKind {
    Text,
    VariableFontText,
    VariableFontObject,
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
    fn computes_ntsc_length_with_ceiling() {
        assert_eq!(
            frame_length(aviutl2::common::Rational32::new(30_000, 1001), 1.1),
            33
        );
    }
}
