use std::sync::atomic::{AtomicBool, Ordering};

static READY: AtomicBool = AtomicBool::new(false);

pub(crate) fn initialize() {
    READY.store(true, Ordering::Release);
}

pub(crate) fn text(source: &str) -> String {
    if READY.load(Ordering::Acquire) {
        aviutl2::config::translate(source)
    } else {
        source.to_string()
    }
}

pub(crate) fn format_text(source: &str, values: &[(&str, String)]) -> String {
    let mut translated = text(source);
    for (name, value) in values {
        translated = translated.replace(name, value);
    }
    translated
}
