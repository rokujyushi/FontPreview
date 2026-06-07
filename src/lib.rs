#![cfg(windows)]
#![allow(unsafe_op_in_unsafe_fn)]
#![allow(non_snake_case)]

use aviutl2::AnyResult;

mod actions;
mod alias;
mod catalog;
mod preview;
mod settings;
mod ui;

pub(crate) static EDIT_HANDLE: aviutl2::generic::GlobalEditHandle =
    aviutl2::generic::GlobalEditHandle::new();

#[aviutl2::plugin(GenericPlugin)]
struct FontPreviewPlugin {
    window: aviutl2_eframe::EframeWindow,
}

impl aviutl2::generic::GenericPlugin for FontPreviewPlugin {
    fn new(_info: aviutl2::AviUtl2Info) -> AnyResult<Self> {
        let _ = aviutl2::tracing_subscriber::fmt()
            .with_max_level(if cfg!(debug_assertions) {
                tracing::Level::DEBUG
            } else {
                tracing::Level::INFO
            })
            .event_format(aviutl2::logger::AviUtl2Formatter)
            .with_writer(aviutl2::logger::AviUtl2LogWriter)
            .try_init();

        let app_data = aviutl2::config::app_data_path();
        let window = aviutl2_eframe::EframeWindow::new("FontPreviewClient", move |cc, handle| {
            Ok(Box::new(ui::FontPreviewApp::new(
                cc,
                handle,
                app_data.clone(),
            )))
        })?;
        Ok(Self { window })
    }

    fn plugin_info(&self) -> aviutl2::generic::GenericPluginTable {
        aviutl2::generic::GenericPluginTable {
            name: "Font Preview".to_string(),
            information: format!("Font Preview {} (Rust)", env!("CARGO_PKG_VERSION")),
        }
    }

    fn register(&mut self, host: &mut aviutl2::generic::HostAppHandle) {
        EDIT_HANDLE.init(host.create_edit_handle());
        match self.window.handle() {
            Ok(handle) => {
                if let Err(error) = host.register_window_client("Font Preview", &handle) {
                    tracing::error!("Font Previewウィンドウの登録に失敗しました: {error}");
                }
            }
            Err(error) => tracing::error!("Font Previewウィンドウを取得できませんでした: {error}"),
        }
    }
}

aviutl2::register_generic_plugin!(FontPreviewPlugin);
