#![cfg(windows)]
#![allow(unsafe_op_in_unsafe_fn)]
#![allow(non_snake_case)]

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex, OnceLock};
use std::thread::JoinHandle;
use std::time::{Duration, Instant};

use aviutl2::AnyResult;

mod actions;
mod alias;
mod catalog;
mod preview;
mod settings;
mod ui;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum SelectedTextResult {
    NotReady,
    Text(Option<String>),
    Error(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SelectedTextSnapshot {
    pub revision: u64,
    pub result: SelectedTextResult,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum FontDropOutcome {
    Imported(String),
    AlreadyPresent(String),
    Error(String),
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct FontDropSnapshot {
    pub revision: u64,
    pub outcome: Option<FontDropOutcome>,
}

impl Default for SelectedTextSnapshot {
    fn default() -> Self {
        Self {
            revision: 0,
            result: SelectedTextResult::NotReady,
        }
    }
}

#[derive(Debug, Default)]
pub(crate) struct SharedEditState {
    edit_handle: OnceLock<aviutl2::generic::EditHandle>,
    egui_ctx: OnceLock<aviutl2_eframe::egui::Context>,
    selected_text: Mutex<SelectedTextSnapshot>,
    font_drop: Mutex<FontDropSnapshot>,
    refresh_signal: Mutex<RefreshSignal>,
    refresh_wake: Condvar,
    next_refresh_id: AtomicU64,
}

#[derive(Debug, Default)]
struct RefreshSignal {
    pending: bool,
    shutdown: bool,
    reason: Option<&'static str>,
    request_id: u64,
    coalesced: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RefreshRequest {
    id: u64,
    reason: &'static str,
    coalesced: u64,
}

impl SharedEditState {
    pub(crate) fn init_edit_handle(&self, handle: aviutl2::generic::EditHandle) {
        if self.edit_handle.set(handle).is_err() {
            tracing::warn!("FontPreview edit handle was already initialized");
        }
    }

    pub(crate) fn edit_handle(&self) -> Option<&aviutl2::generic::EditHandle> {
        self.edit_handle.get().filter(|handle| handle.is_ready())
    }

    pub(crate) fn init_egui_ctx(&self, ctx: aviutl2_eframe::egui::Context) {
        if self.egui_ctx.set(ctx).is_err() {
            tracing::warn!("FontPreview egui context was already initialized");
        }
    }

    pub(crate) fn selected_text_snapshot(&self) -> SelectedTextSnapshot {
        self.selected_text
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }

    pub(crate) fn font_drop_snapshot(&self) -> FontDropSnapshot {
        self.font_drop
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }

    fn publish_font_drop(&self, outcome: FontDropOutcome) -> u64 {
        let mut snapshot = self
            .font_drop
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        snapshot.revision = snapshot.revision.wrapping_add(1);
        snapshot.outcome = Some(outcome);
        let revision = snapshot.revision;
        drop(snapshot);
        if let Some(ctx) = self.egui_ctx.get() {
            ctx.request_repaint();
        }
        revision
    }

    fn publish_selected_text(&self, result: SelectedTextResult) -> u64 {
        let mut snapshot = self
            .selected_text
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        snapshot.revision = snapshot.revision.wrapping_add(1);
        snapshot.result = result;
        snapshot.revision
    }

    fn request_selected_text_refresh(&self, reason: &'static str) {
        let request_id = self.next_refresh_id.fetch_add(1, Ordering::Relaxed) + 1;
        let mut signal = self
            .refresh_signal
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if signal.shutdown {
            return;
        }
        if signal.pending {
            signal.coalesced = signal.coalesced.saturating_add(1);
        }
        signal.pending = true;
        signal.reason = Some(reason);
        signal.request_id = request_id;
        self.refresh_wake.notify_one();
        let coalesced = signal.coalesced;
        drop(signal);
        tracing::debug!(
            request_id,
            reason,
            coalesced,
            thread_id = ?std::thread::current().id(),
            "FontPreview selected text refresh queued"
        );
    }

    fn wait_for_selected_text_refresh(&self) -> Option<RefreshRequest> {
        const DEBOUNCE: Duration = Duration::from_millis(150);

        let mut signal = self
            .refresh_signal
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        while !signal.pending && !signal.shutdown {
            signal = self
                .refresh_wake
                .wait(signal)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
        }
        if signal.shutdown {
            return None;
        }

        signal.pending = false;
        let mut reason = signal.reason.take().unwrap_or("edit_event");
        let mut request_id = signal.request_id;
        let mut coalesced = signal.coalesced;
        signal.coalesced = 0;
        loop {
            let (next_signal, timeout) = self
                .refresh_wake
                .wait_timeout(signal, DEBOUNCE)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            signal = next_signal;
            if signal.shutdown {
                return None;
            }
            if signal.pending {
                signal.pending = false;
                reason = signal.reason.take().unwrap_or(reason);
                request_id = signal.request_id;
                coalesced = coalesced.saturating_add(signal.coalesced).saturating_add(1);
                signal.coalesced = 0;
                continue;
            }
            if timeout.timed_out() {
                return Some(RefreshRequest {
                    id: request_id,
                    reason,
                    coalesced,
                });
            }
        }
    }

    fn shutdown_selected_text_worker(&self) {
        let mut signal = self
            .refresh_signal
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        signal.shutdown = true;
        signal.pending = false;
        self.refresh_wake.notify_one();
    }

    fn request_repaint(&self, request: RefreshRequest, revision: u64) {
        if let Some(ctx) = self.egui_ctx.get() {
            tracing::debug!(
                request_id = request.id,
                reason = request.reason,
                revision,
                thread_id = ?std::thread::current().id(),
                "FontPreview repaint request start"
            );
            ctx.request_repaint();
            tracing::debug!(
                request_id = request.id,
                reason = request.reason,
                revision,
                "FontPreview repaint request finish"
            );
        } else {
            tracing::debug!(
                request_id = request.id,
                reason = request.reason,
                "FontPreview repaint skipped before UI initialization"
            );
        }
    }
}

fn selected_text_worker(shared_edit: Arc<SharedEditState>) {
    tracing::debug!(
        thread_id = ?std::thread::current().id(),
        "FontPreview selected text worker started"
    );
    while let Some(request) = shared_edit.wait_for_selected_text_refresh() {
        tracing::debug!(
            request_id = request.id,
            reason = request.reason,
            coalesced = request.coalesced,
            thread_id = ?std::thread::current().id(),
            "FontPreview selected text worker refresh start"
        );
        let Some(handle) = shared_edit.edit_handle() else {
            tracing::debug!(
                request_id = request.id,
                reason = request.reason,
                "FontPreview selected text refresh skipped before ready"
            );
            continue;
        };
        let started = Instant::now();
        let result = match actions::selected_text(handle) {
            Ok(text) => SelectedTextResult::Text(text),
            Err(error) => SelectedTextResult::Error(format!("{error:#}")),
        };
        let revision = shared_edit.publish_selected_text(result);
        tracing::debug!(
            request_id = request.id,
            reason = request.reason,
            revision,
            elapsed_ms = started.elapsed().as_millis(),
            "FontPreview selected text snapshot updated"
        );
        shared_edit.request_repaint(request, revision);
        tracing::debug!(
            request_id = request.id,
            reason = request.reason,
            revision,
            "FontPreview selected text worker refresh finish"
        );
    }
    tracing::debug!("FontPreview selected text worker stopped");
}

fn import_dropped_font(
    shared_edit: &SharedEditState,
    font_dir: &std::path::Path,
    source: &std::path::Path,
) {
    tracing::debug!(source = %source.display(), "FontPreview font drop start");
    let outcome = match actions::import_font_file(source, font_dir) {
        Ok(actions::FontImport::Copied(target)) => {
            tracing::debug!(target = %target.display(), "FontPreview dropped font copied");
            FontDropOutcome::Imported(target.display().to_string())
        }
        Ok(actions::FontImport::AlreadyPresent(target)) => {
            tracing::debug!(target = %target.display(), "FontPreview dropped font already present");
            FontDropOutcome::AlreadyPresent(target.display().to_string())
        }
        Err(error) => {
            tracing::warn!(
                source = %source.display(),
                error = %format!("{error:#}"),
                "FontPreview font drop failed"
            );
            FontDropOutcome::Error(format!("{error:#}"))
        }
    };
    let revision = shared_edit.publish_font_drop(outcome);
    tracing::debug!(revision, "FontPreview font drop finish");
}

#[aviutl2::plugin(GenericPlugin)]
struct FontPreviewPlugin {
    window: aviutl2_eframe::EframeWindow,
    shared_edit: Arc<SharedEditState>,
    selected_text_worker: Option<JoinHandle<()>>,
}

impl aviutl2::generic::GenericPlugin for FontPreviewPlugin {
    fn new(_info: aviutl2::AviUtl2Info) -> AnyResult<Self> {
        let _ = aviutl2::tracing_subscriber::fmt()
            .with_max_level(tracing::Level::DEBUG)
            .with_ansi(false)
            .event_format(aviutl2::logger::AviUtl2Formatter)
            .with_writer(aviutl2::logger::AviUtl2LogWriter)
            .try_init();

        let app_data = aviutl2::config::app_data_path();
        let shared_edit = Arc::new(SharedEditState::default());
        let ui_shared_edit = Arc::clone(&shared_edit);
        let window = aviutl2_eframe::EframeWindow::new("FontPreviewClient", move |cc, handle| {
            Ok(Box::new(ui::FontPreviewApp::new(
                cc,
                handle,
                app_data.clone(),
                Arc::clone(&ui_shared_edit),
            )))
        })?;
        Ok(Self {
            window,
            shared_edit,
            selected_text_worker: None,
        })
    }

    fn plugin_info(&self) -> aviutl2::generic::GenericPluginTable {
        aviutl2::generic::GenericPluginTable {
            name: "Font Preview".to_string(),
            information: format!(
                "Font Preview {} (Rust) by 黒猫大福",
                env!("CARGO_PKG_VERSION")
            ),
        }
    }

    fn register(&mut self, host: &mut aviutl2::generic::HostAppHandle) {
        tracing::debug!("FontPreview register start");
        self.shared_edit.init_edit_handle(host.create_edit_handle());
        tracing::debug!("FontPreview edit handle initialized");
        let worker_state = Arc::clone(&self.shared_edit);
        match std::thread::Builder::new()
            .name("FontPreview text sync".to_string())
            .spawn(move || selected_text_worker(worker_state))
        {
            Ok(worker) => self.selected_text_worker = Some(worker),
            Err(error) => tracing::error!(
                error = %error,
                "FontPreview could not start selected text worker"
            ),
        }
        match self.window.handle() {
            Ok(handle) => {
                tracing::debug!("FontPreview eframe window handle acquired");
                if let Err(error) = host.register_window_client("Font Preview", &handle) {
                    tracing::error!("Font Previewウィンドウの登録に失敗しました: {error}");
                } else {
                    tracing::debug!("FontPreview window client registered");
                }
            }
            Err(error) => tracing::error!("Font Previewウィンドウを取得できませんでした: {error}"),
        }
        let drop_state = Arc::clone(&self.shared_edit);
        let font_dir = aviutl2::config::app_data_path().join("Font");
        let filters = aviutl2::file_filters! {
            "フォントファイル" => ["ttf", "otf", "ttc"]
        };
        host.register_file_drop_handler(
            "Font Previewへフォントを追加",
            &filters,
            move |source| import_dropped_font(&drop_state, &font_dir, &source),
        );
        tracing::debug!("FontPreview font drop handler registered");
        tracing::debug!("FontPreview register finish");
    }

    fn on_project_load(&mut self, project: &mut aviutl2::generic::ProjectFile) {
        let _ = project;
        tracing::debug!("FontPreview on_project_load noop");
    }

    fn on_project_save(&mut self, project: &mut aviutl2::generic::ProjectFile) {
        let _ = project;
        tracing::debug!("FontPreview on_project_save noop");
    }

    fn on_clear_cache(&mut self, edit_section: &aviutl2::generic::EditSection) {
        let _ = edit_section;
        tracing::debug!("FontPreview on_clear_cache noop");
    }

    fn event_update_object_info(&mut self) {
        self.shared_edit
            .request_selected_text_refresh("update_object");
    }

    fn event_change_scene_info(&mut self) {
        self.shared_edit
            .request_selected_text_refresh("change_scene");
    }

    fn event_change_focus_object(&mut self) {
        self.shared_edit
            .request_selected_text_refresh("change_focus_object");
    }
}

impl Drop for FontPreviewPlugin {
    fn drop(&mut self) {
        self.shared_edit.shutdown_selected_text_worker();
        if let Some(worker) = self.selected_text_worker.take()
            && worker.thread().id() != std::thread::current().id()
            && let Err(error) = worker.join()
        {
            tracing::error!(?error, "FontPreview selected text worker panicked");
        }
    }
}

aviutl2::register_generic_plugin!(FontPreviewPlugin);

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selected_text_snapshot_starts_not_ready() {
        let state = SharedEditState::default();
        assert_eq!(
            state.selected_text_snapshot(),
            SelectedTextSnapshot {
                revision: 0,
                result: SelectedTextResult::NotReady,
            }
        );
    }

    #[test]
    fn selected_text_snapshot_tracks_empty_and_errors() {
        let state = SharedEditState::default();
        assert_eq!(
            state.publish_selected_text(SelectedTextResult::Text(None)),
            1
        );
        assert_eq!(
            state.selected_text_snapshot().result,
            SelectedTextResult::Text(None)
        );
        assert_eq!(
            state.publish_selected_text(SelectedTextResult::Error("failed".to_string())),
            2
        );
        assert_eq!(
            state.selected_text_snapshot().result,
            SelectedTextResult::Error("failed".to_string())
        );
    }

    #[test]
    fn identical_selected_text_updates_advance_revision() {
        let state = SharedEditState::default();
        let value = SelectedTextResult::Text(Some("sample".to_string()));
        state.publish_selected_text(value.clone());
        state.publish_selected_text(value.clone());
        assert_eq!(
            state.selected_text_snapshot(),
            SelectedTextSnapshot {
                revision: 2,
                result: value,
            }
        );
    }

    #[test]
    fn refresh_requests_are_coalesced_and_keep_latest_reason() {
        let state = SharedEditState::default();
        state.request_selected_text_refresh("first");
        state.request_selected_text_refresh("latest");
        let signal = state
            .refresh_signal
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        assert!(signal.pending);
        assert_eq!(signal.reason, Some("latest"));
        assert_eq!(signal.request_id, 2);
        assert_eq!(signal.coalesced, 1);
    }

    #[test]
    fn worker_wait_stops_immediately_after_shutdown() {
        let state = SharedEditState::default();
        state.shutdown_selected_text_worker();
        assert_eq!(state.wait_for_selected_text_refresh(), None);
    }

    #[test]
    fn font_drop_snapshot_advances_revision() {
        let state = SharedEditState::default();
        assert_eq!(state.font_drop_snapshot(), FontDropSnapshot::default());
        assert_eq!(
            state.publish_font_drop(FontDropOutcome::Imported("test.ttf".to_string())),
            1
        );
        assert_eq!(
            state.font_drop_snapshot(),
            FontDropSnapshot {
                revision: 1,
                outcome: Some(FontDropOutcome::Imported("test.ttf".to_string())),
            }
        );
    }
}
