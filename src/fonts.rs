use std::sync::Arc;

use aviutl2_eframe::egui::{self, FontDefinitions, FontFamily};

const FALLBACK_FONTS_KEY: &str = "__FontFallbackFonts";
const FALLBACK_FONT_DATA_NAME: &str = "font-preview::SystemFallback";

pub(crate) fn definitions() -> FontDefinitions {
    let mut definitions = aviutl2_eframe::aviutl2_fonts();
    let fallback_spec = aviutl2::config::translate(FALLBACK_FONTS_KEY);
    let candidates = fallback_candidates(&fallback_spec);
    if candidates.is_empty() {
        return definitions;
    }

    let mut database = fontdb::Database::new();
    database.load_system_fonts();
    for family_name in candidates {
        let Some(font_id) = database.query(&fontdb::Query {
            families: &[fontdb::Family::Name(family_name)],
            ..Default::default()
        }) else {
            continue;
        };

        let mut loaded = false;
        database.with_face_data(font_id, |data, face_index| {
            let mut font_data = egui::FontData::from_owned(data.to_vec());
            font_data.index = face_index;
            definitions
                .font_data
                .insert(FALLBACK_FONT_DATA_NAME.to_string(), Arc::new(font_data));
            for family in [FontFamily::Proportional, FontFamily::Monospace] {
                definitions
                    .families
                    .entry(family)
                    .or_default()
                    .push(FALLBACK_FONT_DATA_NAME.to_string());
            }
            loaded = true;
        });
        if loaded {
            tracing::debug!(family_name, "FontPreview UI fallback font loaded");
            return definitions;
        }
    }

    tracing::warn!(
        candidates = fallback_spec,
        "FontPreview UI fallback font was not found"
    );
    definitions
}

fn fallback_candidates(spec: &str) -> Vec<&str> {
    if spec == FALLBACK_FONTS_KEY {
        return Vec::new();
    }
    spec.split(',')
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_ordered_fallback_candidates() {
        assert_eq!(
            fallback_candidates("Microsoft YaHei UI, Microsoft YaHei, SimSun"),
            ["Microsoft YaHei UI", "Microsoft YaHei", "SimSun"]
        );
    }

    #[test]
    fn untranslated_or_empty_setting_disables_fallback() {
        assert!(fallback_candidates(FALLBACK_FONTS_KEY).is_empty());
        assert!(fallback_candidates("").is_empty());
    }
}
