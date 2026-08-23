use std::cmp::Ordering;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use aviutl2_eframe::{AviUtl2EframeHandle, eframe, egui};

use crate::actions;
use crate::alias::ObjectKind;
use crate::catalog::{FontId, FontItem, FontSource, enumerate};
use crate::i18n::{format_text, text};
use crate::settings::{FilterMode, Settings, SortMode, settings_path};
use crate::{FontDropOutcome, SelectedTextResult, SelectedTextSnapshot, SharedEditState};

/// リスト1行の高さ。`show_rows`の仮想化とキーボードスクロールの両方がこの値を前提にする。
const ROW_HEIGHT: f32 = 26.0;
/// PageUp/PageDown や ←→ / A D で飛ぶ行数。
const PAGE_STEP: isize = 10;
/// 検索欄の固定 ID。キー操作で「今どの入力欄にいるか」を見分けるのに使う。
const SEARCH_FIELD_ID: &str = "font-preview-search";
const MIN_PREVIEW_WIDTH: f32 = 240.0;
const MIN_PREVIEW_HEIGHT: f32 = 120.0;
/// 詳細列のうち、プレビュー以外（見出し・各種操作・ボタン・ステータス）が使う高さの見積もり。
const PREVIEW_CHROME_HEIGHT: f32 = 320.0;
/// 一覧と詳細の間にあるセパレータと余白の見積もり。
const DETAIL_GUTTER: f32 = 24.0;
/// 描画バッファが際限なく大きくならないようにする上限。
const MAX_PREVIEW_PIXELS: u32 = 4096;

struct PendingMove {
    font_id: FontId,
    destination: FontSource,
    favorite: bool,
}

pub(crate) struct FontPreviewApp {
    window_handle: AviUtl2EframeHandle,
    shared_edit: Arc<SharedEditState>,
    fonts: Vec<FontItem>,
    filtered: Vec<usize>,
    selected: Option<usize>,
    search: String,
    settings: Settings,
    settings_path: PathBuf,
    first_team_dir: PathBuf,
    library_dir: PathBuf,
    preview: Option<egui::TextureHandle>,
    preview_dirty: bool,
    status: Option<String>,
    pending_move: Option<PendingMove>,
    selected_text_revision: u64,
    font_drop_revision: u64,
    /// 次の描画で見える位置までスクロールすべき行。
    scroll_to_row: Option<usize>,
    /// 直前に描画したプレビュー画像の実ピクセルサイズ。
    preview_size: (u32, u32),
    /// 詳細列に割り当てられた領域。パネル分割直後に決まる。
    detail_area: egui::Vec2,
    /// リストがキー操作の対象になっているか。リスト内を押すと立ち、外を押すと降りる。
    list_active: bool,
    /// 前の描画でリストが占めていた矩形。クリックがリスト内かの判定に使う。
    list_rect: egui::Rect,
    /// 設定ウィンドウを開いているか。
    show_options: bool,
    /// 設定ウィンドウが占めていた矩形。閉じているときは `NOTHING`。
    options_rect: egui::Rect,
}

impl FontPreviewApp {
    pub(crate) fn new(
        cc: &eframe::CreationContext<'_>,
        window_handle: AviUtl2EframeHandle,
        app_data: PathBuf,
        shared_edit: Arc<SharedEditState>,
    ) -> Self {
        let started = Instant::now();
        tracing::debug!(app_data = %app_data.display(), "FontPreview app init start");
        cc.egui_ctx.all_styles_mut(|style| {
            style.visuals = aviutl2_eframe::aviutl2_visuals();
        });
        cc.egui_ctx.set_fonts(crate::fonts::definitions());
        shared_edit.init_egui_ctx(cc.egui_ctx.clone());

        let first_team_dir = app_data.join("Font");
        let library_dir = app_data.join("FontLibrary");
        let settings_path = settings_path(&app_data);
        let (settings, settings_status) = Settings::load(&settings_path);
        tracing::debug!(path = %settings_path.display(), "FontPreview settings loaded");
        let (mut fonts, catalog_status) = match enumerate(&first_team_dir, &library_dir) {
            Ok(fonts) => (fonts, None),
            Err(error) => {
                tracing::error!(error = %format!("{error:#}"), "FontPreview catalog enumerate failed");
                (
                    Vec::new(),
                    Some(format!("{}: {error:#}", text("フォント列挙エラー"))),
                )
            }
        };
        settings.apply_favorites(&mut fonts);
        let mut app = Self {
            window_handle,
            shared_edit,
            fonts,
            filtered: Vec::new(),
            selected: None,
            search: String::new(),
            settings,
            settings_path,
            first_team_dir,
            library_dir,
            preview: None,
            preview_dirty: true,
            status: catalog_status.or(settings_status),
            pending_move: None,
            selected_text_revision: 0,
            font_drop_revision: 0,
            scroll_to_row: None,
            preview_size: (0, 0),
            detail_area: egui::vec2(MIN_PREVIEW_WIDTH, MIN_PREVIEW_HEIGHT),
            // 意図せず選択が変わるのを避けるため、初期状態ではキーを拾わない。
            list_active: false,
            list_rect: egui::Rect::NOTHING,
            show_options: false,
            options_rect: egui::Rect::NOTHING,
        };
        app.rebuild_filter();
        tracing::debug!(
            fonts = app.fonts.len(),
            elapsed_ms = started.elapsed().as_millis(),
            "FontPreview app init finish"
        );
        app
    }

    fn rebuild_filter(&mut self) {
        let selected_id = self.selected_font().map(|font| font.id.clone());
        let query = self.search.to_lowercase();
        self.filtered = self
            .fonts
            .iter()
            .enumerate()
            .filter(|(_, font)| filter_matches(font, self.settings.filter))
            .filter(|(_, font)| {
                query.is_empty()
                    || font.display_name.to_lowercase().contains(&query)
                    || font.family_name.to_lowercase().contains(&query)
            })
            .map(|(index, _)| index)
            .collect();
        let sort = self.settings.sort;
        self.filtered
            .sort_by(|left, right| compare_fonts(&self.fonts[*left], &self.fonts[*right], sort));
        self.selected = selected_id
            .and_then(|id| self.fonts.iter().position(|font| font.id == id))
            .filter(|index| self.filtered.contains(index))
            .or_else(|| self.filtered.first().copied());
        // 検索や並び替えで行位置が変わるので、選択を見える位置へ戻す。
        self.scroll_to_row = self.selected_row();
        self.preview_dirty = true;
    }

    /// ポインタの押下位置で、リストがキー操作の対象かどうかを切り替える。
    fn track_list_activation(&mut self, ctx: &egui::Context) {
        if !ctx.input(|input| input.pointer.any_pressed()) {
            return;
        }
        let Some(position) = ctx.pointer_interact_pos() else {
            return;
        };
        // 設定ウィンドウはリストに重なり得るので、そちらへのクリックではアクティブにしない。
        self.list_active =
            self.list_rect.contains(position) && !self.options_rect.contains(position);
    }

    /// 選択中フォントの、表示リストの中での行番号。
    fn selected_row(&self) -> Option<usize> {
        let selected = self.selected?;
        self.filtered.iter().position(|index| *index == selected)
    }

    fn select_row(&mut self, row: usize) {
        let Some(index) = self.filtered.get(row).copied() else {
            return;
        };
        if self.selected != Some(index) {
            self.selected = Some(index);
            self.preview_dirty = true;
        }
        self.scroll_to_row = Some(row);
    }

    /// キー操作を拾うか、拾うならどこまで拾うか。
    ///
    /// キーを広く取り上げると、スライダーを矢印キーで微調整できなくなる、
    /// ボタン操作中に選択フォントが勝手に変わる、といった事故が起きる。
    /// 「今リストを触っている」と見なせるときだけに限定する。
    fn key_scope(&self, ctx: &egui::Context) -> KeyScope {
        key_scope_for(
            self.settings.keys.enabled && self.pending_move.is_none(),
            ctx.memory(|memory| memory.focused()),
            self.list_active,
        )
    }

    /// リストのキーボード操作。
    ///
    /// 拾うキーは`consume_key`で先に取り上げ、`TextEdit`へ届かないようにする。
    fn handle_list_keys(&mut self, ctx: &egui::Context) {
        let scope = self.key_scope(ctx);
        if scope == KeyScope::None {
            return;
        }
        let typing = scope == KeyScope::Search;
        let keys = self.settings.keys;
        let buttons = self.settings.create_buttons;
        let mut step = 0isize;
        let mut edge: Option<usize> = None;
        let mut apply = false;
        let mut favorite = false;
        let mut create: Option<ObjectKind> = None;
        ctx.input_mut(|input| {
            use egui::{Key, Modifiers};
            const NONE: Modifiers = Modifiers::NONE;
            if keys.arrows_active() {
                if input.consume_key(NONE, Key::ArrowDown) {
                    step += 1;
                }
                if input.consume_key(NONE, Key::ArrowUp) {
                    step -= 1;
                }
            }
            if keys.enter_active() && input.consume_key(NONE, Key::Enter) {
                apply = true;
            }
            if typing {
                // 検索欄で拾うのはここまで。以降は文字入力と衝突する。
                return;
            }
            if keys.letters_active() {
                if input.consume_key(NONE, Key::S) {
                    step += 1;
                }
                if input.consume_key(NONE, Key::W) {
                    step -= 1;
                }
            }
            // `|=`は短絡しないので、同じ意味のキーを全て取り上げられる。
            let mut page_down = false;
            let mut page_up = false;
            if keys.arrows_active() {
                page_down |= input.consume_key(NONE, Key::PageDown);
                page_down |= input.consume_key(NONE, Key::ArrowRight);
                page_up |= input.consume_key(NONE, Key::PageUp);
                page_up |= input.consume_key(NONE, Key::ArrowLeft);
                if input.consume_key(NONE, Key::Home) {
                    edge = Some(0);
                }
                if input.consume_key(NONE, Key::End) {
                    edge = Some(usize::MAX);
                }
            }
            if keys.letters_active() {
                page_down |= input.consume_key(NONE, Key::D);
                page_up |= input.consume_key(NONE, Key::A);
            }
            if page_down {
                step += PAGE_STEP;
            }
            if page_up {
                step -= PAGE_STEP;
            }
            if keys.favorite_active() && input.consume_key(NONE, Key::F) {
                favorite = true;
            }
            if keys.create_active() {
                for kind in ObjectKind::ALL {
                    // 隠しているボタンのキーは受け付けない。
                    if !create_button_visible(buttons, kind) {
                        continue;
                    }
                    let Some(key) = digit_key(kind.shortcut_digit()) else {
                        continue;
                    };
                    if input.consume_key(NONE, key) {
                        create = Some(kind);
                    }
                }
            }
        });

        if let Some(row) = edge {
            self.select_row(row.min(self.filtered.len().saturating_sub(1)));
        } else if step != 0
            && let Some(row) = moved_row(self.selected_row(), self.filtered.len(), step)
        {
            self.select_row(row);
        }
        if favorite && let Some(index) = self.selected {
            self.toggle_favorite(index);
        }
        if apply {
            self.apply_selected_font();
        }
        if let Some(kind) = create
            && let Some(font) = self.selected_font().cloned()
        {
            self.create_object_with(&font, kind);
        }
    }

    /// 選択中のフォントで新しいオブジェクトを作る。
    fn create_object_with(&mut self, font: &FontItem, kind: ObjectKind) {
        let result = self
            .shared_edit
            .edit_handle()
            .ok_or_else(|| anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です")))
            .and_then(|handle| actions::create_object(handle, font, &self.settings.sample, kind));
        self.action_result(result);
    }

    /// 選択中のフォントを選択オブジェクトへ適用する。
    fn apply_selected_font(&mut self) {
        let Some(font) = self.selected_font().cloned() else {
            return;
        };
        self.status = Some(
            match self
                .shared_edit
                .edit_handle()
                .ok_or_else(|| anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です")))
                .and_then(|handle| actions::apply_to_selection(handle, &font))
            {
                Ok(count) => format_text(
                    "{count}個のオブジェクトを更新しました",
                    &[("{count}", count.to_string())],
                ),
                Err(error) => format!("{}: {error:#}", text("更新に失敗しました")),
            },
        );
    }

    fn refresh_catalog(&mut self) {
        let started = Instant::now();
        tracing::debug!("FontPreview refresh catalog start");
        let selected_id = self.selected_font().map(|font| font.id.clone());
        match enumerate(&self.first_team_dir, &self.library_dir) {
            Ok(mut fonts) => {
                self.settings.apply_favorites(&mut fonts);
                self.fonts = fonts;
                self.selected =
                    selected_id.and_then(|id| self.fonts.iter().position(|font| font.id == id));
                self.rebuild_filter();
                self.status = Some(format_text(
                    "{count}件のフォントを読み込みました",
                    &[("{count}", self.fonts.len().to_string())],
                ));
                tracing::debug!(
                    fonts = self.fonts.len(),
                    elapsed_ms = started.elapsed().as_millis(),
                    "FontPreview refresh catalog finish"
                );
            }
            Err(error) => {
                tracing::error!(
                    error = %format!("{error:#}"),
                    elapsed_ms = started.elapsed().as_millis(),
                    "FontPreview refresh catalog failed"
                );
                self.status = Some(format!("{}: {error:#}", text("再読み込みエラー")));
            }
        }
    }

    fn selected_font(&self) -> Option<&FontItem> {
        self.selected.and_then(|index| self.fonts.get(index))
    }

    fn save_settings(&mut self) {
        if let Err(error) = self.settings.save(&self.settings_path) {
            self.status = Some(format!("{}: {error:#}", text("設定保存エラー")));
        }
    }

    fn set_sample(&mut self, sample: String) {
        if self.settings.sample == sample {
            return;
        }
        self.settings.sample = sample;
        self.preview_dirty = true;
        self.save_settings();
    }

    fn sync_selected_text(&mut self) {
        let snapshot = self.shared_edit.selected_text_snapshot();
        let previous_revision = self.selected_text_revision;
        let Some(text) = synced_sample(
            self.settings.sync_text,
            &mut self.selected_text_revision,
            &snapshot,
        ) else {
            if self.selected_text_revision != previous_revision
                && let SelectedTextResult::Error(error) = &snapshot.result
            {
                tracing::warn!(
                    revision = snapshot.revision,
                    error,
                    "FontPreview selected text snapshot contains an error"
                );
            }
            return;
        };
        tracing::debug!(
            revision = snapshot.revision,
            len = text.len(),
            "FontPreview selected text snapshot applied"
        );
        self.set_sample(text);
    }

    fn process_font_drop(&mut self) {
        let snapshot = self.shared_edit.font_drop_snapshot();
        if snapshot.revision == 0 || snapshot.revision == self.font_drop_revision {
            return;
        }
        self.font_drop_revision = snapshot.revision;
        match snapshot.outcome {
            Some(FontDropOutcome::Imported(path)) => {
                self.refresh_catalog();
                self.status = Some(format!("{}: {path}", text("フォントを追加しました")));
            }
            Some(FontDropOutcome::AlreadyPresent(path)) => {
                self.refresh_catalog();
                self.status = Some(format!(
                    "{}: {path}",
                    text("フォントは既に追加されています")
                ));
            }
            Some(FontDropOutcome::Error(error)) => {
                self.status = Some(format!("{}: {error}", text("フォント追加エラー")));
            }
            None => {}
        }
    }

    /// プレビューの描画サイズを決める。
    ///
    /// 幅は詳細列に追従し、高さはフォントサイズに応じて伸びる。戻り値は
    /// （実ピクセル幅, 実ピクセル高, 表示サイズ）。実ピクセルで描いて等倍で貼るので、
    /// 以前のように縮小表示でぬれることがなくなる。
    ///
    /// 寸法の元にするのは、パネル分割直後に決まる `detail_area` だけです。
    /// 描画直前の `ui.available_*` を見ると、画像サイズ→レイアウト→画像サイズの
    /// 帰還路ができ、何も操作していないのに再描画が繰り返されます。
    fn preview_canvas(&self, ctx: &egui::Context) -> (u32, u32, egui::Vec2) {
        let points_per_pixel = ctx.pixels_per_point().max(0.5);
        // ウィンドウリサイズ中の再描画を抑えるために16pt刻みへ丸める。
        let width = (self.detail_area.x.max(MIN_PREVIEW_WIDTH) / 16.0).floor() * 16.0;
        let height_limit = (self.detail_area.y - PREVIEW_CHROME_HEIGHT).max(MIN_PREVIEW_HEIGHT);
        let height =
            (self.settings.preview_font_size * 1.7 + 48.0).clamp(MIN_PREVIEW_HEIGHT, height_limit);
        let to_pixels = |points: f32| {
            ((points * points_per_pixel).round() as u32).clamp(16, MAX_PREVIEW_PIXELS)
        };
        (
            to_pixels(width),
            to_pixels(height),
            egui::vec2(width, height),
        )
    }

    fn update_preview(&mut self, ctx: &egui::Context) -> egui::Vec2 {
        let (width, height, size) = self.preview_canvas(ctx);
        if self.preview_size != (width, height) {
            self.preview_size = (width, height);
            self.preview_dirty = true;
        }
        if !self.preview_dirty {
            return size;
        }
        self.preview_dirty = false;
        let Some(font) = self.selected_font() else {
            self.preview = None;
            return size;
        };
        let preview_text = if self.settings.sample.is_empty() {
            "あいうABC123"
        } else {
            &self.settings.sample
        };
        let started = Instant::now();
        tracing::debug!(
            font = %font.display_name,
            text_len = preview_text.len(),
            font_size = self.settings.preview_font_size,
            canvas = %format!("{width}x{height}"),
            "FontPreview detail preview render start"
        );
        // キャンバスを実ピクセルで取るので、文字サイズも同じ倍率で拡大する。
        let font_size = self.settings.preview_font_size * ctx.pixels_per_point().max(0.5);
        match crate::preview::render(
            font,
            preview_text,
            self.settings.preview_background_color,
            self.settings.preview_text_color,
            width,
            height,
            font_size,
        ) {
            Ok(image) => {
                let image = egui::ColorImage::from_rgba_unmultiplied(
                    [image.width as usize, image.height as usize],
                    &image.rgba,
                );
                self.preview =
                    Some(ctx.load_texture("font-preview", image, egui::TextureOptions::LINEAR));
                tracing::debug!(
                    elapsed_ms = started.elapsed().as_millis(),
                    "FontPreview detail preview render finish"
                );
            }
            Err(error) => {
                self.preview = None;
                self.status = Some(format!("{}: {error:#}", text("プレビューエラー")));
                tracing::error!(
                    error = %format!("{error:#}"),
                    elapsed_ms = started.elapsed().as_millis(),
                    "FontPreview detail preview render failed"
                );
            }
        }
        size
    }

    fn toggle_favorite(&mut self, index: usize) {
        let Some(font) = self.fonts.get(index).cloned() else {
            return;
        };
        let favorite = !font.favorite;
        let destination = if favorite {
            FontSource::FirstTeam
        } else {
            FontSource::Library
        };
        if !font.is_system()
            && self.settings.move_local_fonts_with_favorites
            && font.source != destination
        {
            self.pending_move = Some(PendingMove {
                font_id: font.id,
                destination,
                favorite,
            });
            return;
        }

        let favorite = self.settings.toggle_favorite(&font);
        if let Some(current) = self.fonts.get_mut(index) {
            current.favorite = favorite;
        }
        self.save_settings();
        self.rebuild_filter();
    }

    fn apply_pending_move(&mut self) {
        let Some(pending) = self.pending_move.take() else {
            return;
        };
        let Some(font) = self
            .fonts
            .iter()
            .find(|font| font.id == pending.font_id)
            .cloned()
        else {
            self.status = Some(text("移動対象のフォントが見つかりません"));
            return;
        };
        match actions::move_local_font(&font, &self.first_team_dir, &self.library_dir) {
            Ok(_) => {
                let favorite = self.settings.toggle_favorite(&font);
                debug_assert_eq!(favorite, pending.favorite);
                self.save_settings();
                self.refresh_catalog();
                self.status = Some(text(
                    "フォントを移動しました。AviUtl2のメニュー反映には再起動が必要です",
                ));
            }
            Err(error) => self.status = Some(format!("{}: {error:#}", text("フォント移動エラー"))),
        }
    }

    fn show_move_dialog(&mut self, ctx: &egui::Context) {
        let Some(pending) = &self.pending_move else {
            return;
        };
        let destination = text(pending.destination.label());
        let font_name = self
            .fonts
            .iter()
            .find(|font| font.id == pending.font_id)
            .map(|font| font.display_name.clone())
            .unwrap_or_default();
        let mut confirm = false;
        let mut cancel = false;
        egui::Window::new(text("フォントを移動"))
            .collapsible(false)
            .resizable(false)
            .anchor(egui::Align2::CENTER_CENTER, egui::Vec2::ZERO)
            .show(ctx, |ui| {
                ui.label(format_text(
                    "「{font}」を「{destination}」へ移動しますか？",
                    &[
                        ("{font}", font_name.clone()),
                        ("{destination}", destination.clone()),
                    ],
                ));
                ui.label(text("AviUtl2のフォントメニュー反映には再起動が必要です。"));
                ui.horizontal(|ui| {
                    confirm = ui.button(text("移動")).clicked();
                    cancel = ui.button(text("キャンセル")).clicked();
                });
            });
        if confirm {
            self.apply_pending_move();
        } else if cancel {
            self.pending_move = None;
        }
    }

    /// 設定ウィンドウ。
    ///
    /// ツールバーは検索と絞り込みだけに残し、持続的な設定は全てここへ集める。
    /// 項目を増やすときはこの関数にセクションを追加する。
    fn show_options_window(&mut self, ctx: &egui::Context) {
        let mut open = self.show_options;
        let window = egui::Window::new(text("設定"))
            .open(&mut open)
            .collapsible(false)
            .resizable(false)
            .default_width(380.0)
            .show(ctx, |ui| {
                ui.heading(text("動作"));
                if ui
                    .checkbox(
                        &mut self.settings.move_local_fonts_with_favorites,
                        text("星操作でローカルフォント / プラグイン専用を連動"),
                    )
                    .changed()
                {
                    self.save_settings();
                }
                ui.small(text(
                    "有効にすると、お気に入り登録時はFontへ、解除時はFontLibraryへ移動します",
                ));

                ui.add_space(12.0);
                ui.separator();
                ui.heading(text("追加ボタン"));
                ui.small(text(
                    "詳細側に並ぶオブジェクト追加ボタンを選べます。使わないものは隠せます",
                ));
                ui.add_space(4.0);
                let mut buttons_changed = false;
                egui::Grid::new("font-preview-create-buttons")
                    .num_columns(3)
                    .spacing([12.0, 6.0])
                    .show(ui, |ui| {
                        for (flag, kind) in [
                            (&mut self.settings.create_buttons.text, ObjectKind::Text),
                            (
                                &mut self.settings.create_buttons.variable_font_text,
                                ObjectKind::VariableFontText,
                            ),
                            (
                                &mut self.settings.create_buttons.variable_font_object,
                                ObjectKind::VariableFontObject,
                            ),
                        ] {
                            buttons_changed |=
                                ui.checkbox(flag, text(kind.button_label())).changed();
                            ui.monospace(kind.shortcut_digit().to_string());
                            ui.small(text(kind.description()));
                            ui.end_row();
                        }
                    });
                if buttons_changed {
                    self.save_settings();
                }

                ui.add_space(12.0);
                ui.separator();
                ui.heading(text("キー操作"));
                let mut changed = ui
                    .checkbox(
                        &mut self.settings.keys.enabled,
                        text("キーボードでリストを操作"),
                    )
                    .changed();
                if changed {
                    self.list_active = false;
                }
                ui.small(text(
                    "リストをクリックして選択中の間だけキーを受け付けます。切ると全てのキー操作を無効にします",
                ));
                ui.add_space(4.0);
                // マスターがoffの間は種別ごとの設定を薄くして、効いていないことを見せる。
                ui.add_enabled_ui(self.settings.keys.enabled, |ui| {
                    egui::Grid::new("font-preview-key-groups")
                        .num_columns(2)
                        .spacing([16.0, 6.0])
                        .show(ui, |ui| {
                            for (flag, label, description) in [
                                (
                                    &mut self.settings.keys.arrows,
                                    "矢印キー",
                                    "↑↓で1件、←→ / PageUp PageDownで10件、Home Endで先頭・末尾",
                                ),
                                (
                                    &mut self.settings.keys.letters,
                                    "WASDキー",
                                    "W Sで1件、A Dで10件",
                                ),
                                (
                                    &mut self.settings.keys.enter_applies,
                                    "Enter",
                                    "選択中のオブジェクトへ適用",
                                ),
                                (
                                    &mut self.settings.keys.favorite,
                                    "F",
                                    "お気に入りを切り替え",
                                ),
                                (
                                    &mut self.settings.keys.create_objects,
                                    "1 2 3",
                                    "上の追加ボタンを番号で実行（隠したものは無効）",
                                ),
                            ] {
                                changed |= ui.checkbox(flag, text(label)).changed();
                                ui.small(text(description));
                                ui.end_row();
                            }
                        });
                });
                if changed {
                    self.save_settings();
                }
            });
        self.options_rect = window
            .map(|window| window.response.rect)
            .unwrap_or(egui::Rect::NOTHING);
        self.show_options = open;
    }

    fn action_result(&mut self, result: anyhow::Result<()>) {
        self.status = Some(match result {
            Ok(()) => text("操作を完了しました"),
            Err(error) => format!("{}: {error:#}", text("操作に失敗しました")),
        });
    }
}

impl eframe::App for FontPreviewApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        // クリック位置の判定は前回のリスト矩形を使うので、キー処理より先に更新する。
        self.track_list_activation(ui.ctx());
        // 検索欄の TextEdit より先に拾う必要があるので先頭で呼ぶ。
        self.handle_list_keys(ui.ctx());
        self.process_font_drop();
        self.sync_selected_text();

        egui::Panel::top("toolbar").show(ui, |ui| {
            ui.horizontal(|ui| {
                let title = ui.heading("Font Preview");
                if title.interact(egui::Sense::click()).secondary_clicked() {
                    let _ = self.window_handle.show_context_menu();
                }
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if ui.button(text("再読み込み")).clicked() {
                        self.refresh_catalog();
                    }
                    if ui.button(text("設定")).clicked() {
                        self.show_options = !self.show_options;
                    }
                });
            });
            let search_changed = ui
                .add(
                    egui::TextEdit::singleline(&mut self.search)
                        .id(egui::Id::new(SEARCH_FIELD_ID))
                        .hint_text(text("フォント名を検索..."))
                        .desired_width(f32::INFINITY),
                )
                .changed();
            let old_filter = self.settings.filter;
            let old_sort = self.settings.sort;
            ui.horizontal(|ui| {
                ui.label(text("絞り込み:"));
                egui::ComboBox::from_id_salt("font-filter")
                    .selected_text(filter_label(self.settings.filter))
                    .show_ui(ui, |ui| {
                        for mode in [
                            FilterMode::All,
                            FilterMode::System,
                            FilterMode::FirstTeam,
                            FilterMode::Library,
                            FilterMode::Favorites,
                            FilterMode::Variable,
                        ] {
                            ui.selectable_value(
                                &mut self.settings.filter,
                                mode,
                                filter_label(mode),
                            );
                        }
                    });
                ui.separator();
                ui.label(text("並び替え:"));
                egui::ComboBox::from_id_salt("font-sort")
                    .selected_text(sort_label(self.settings.sort))
                    .show_ui(ui, |ui| {
                        for mode in [
                            SortMode::FavoriteName,
                            SortMode::NameAsc,
                            SortMode::NameDesc,
                            SortMode::Source,
                            SortMode::VariableFirst,
                        ] {
                            ui.selectable_value(&mut self.settings.sort, mode, sort_label(mode));
                        }
                    });
            });
            if search_changed
                || old_filter != self.settings.filter
                || old_sort != self.settings.sort
            {
                if old_filter != self.settings.filter || old_sort != self.settings.sort {
                    self.save_settings();
                }
                self.rebuild_filter();
            }
        });

        egui::CentralPanel::default().show(ui, |ui| {
            let list_width = (ui.available_width() * 0.48).clamp(260.0, 430.0);
            let content_height = ui.available_height();
            // 分割直後に確定させる。プレビューの寸法はこれだけから決まる。
            self.detail_area = egui::vec2(
                (ui.available_width() - list_width - DETAIL_GUTTER).max(MIN_PREVIEW_WIDTH),
                content_height,
            );
            ui.horizontal(|ui| {
                let list_area = ui.allocate_ui_with_layout(
                    egui::vec2(list_width, content_height),
                    egui::Layout::top_down(egui::Align::Min),
                    |ui| {
                        ui.label(format_text(
                            "{shown} / {total}件",
                            &[
                                ("{shown}", self.filtered.len().to_string()),
                                ("{total}", self.fonts.len().to_string()),
                            ],
                        ))
                        .on_hover_text(text("リストをクリックするとキー操作が有効になります（一覧は設定ウィンドウ）"));
                        egui::ScrollArea::vertical()
                            .auto_shrink([false, false])
                            .max_height(ui.available_height())
                            .show_rows(ui, ROW_HEIGHT, self.filtered.len(), |ui, rows| {
                                // 仮想化されているので、画面外の行は widget が存在しない。
                                // 行高が一定であることを利用して目標矩形を自分で組み立てる。
                                if let Some(target) = self.scroll_to_row.take() {
                                    let origin = ui.next_widget_position();
                                    // `show_rows`は「行高 + item_spacing.y」刻みで行を置くので、
                                    // 同じピッチで数えないと離れた行ほど位置がずれる。
                                    let pitch = ROW_HEIGHT + ui.spacing().item_spacing.y;
                                    let top =
                                        origin.y + (target as f32 - rows.start as f32) * pitch;
                                    ui.scroll_to_rect(
                                        egui::Rect::from_min_size(
                                            egui::pos2(origin.x, top),
                                            egui::vec2(1.0, ROW_HEIGHT),
                                        ),
                                        None,
                                    );
                                }
                                for row in rows {
                                    let index = self.filtered[row];
                                    let font = self.fonts[index].clone();
                                    // `show_rows`の`skip_ahead_auto_ids`は1行1ウィジェットしか想定していない。
                                    // 1行に複数あるこの実装では、スクロールで先頭行が変わるたびに
                                    // 自動IDがずれる。フォント固有のIDを積んで行位置から切り離す。
                                    ui.push_id(&font.id, |ui| {
                                    ui.horizontal(|ui| {
                                        if ui
                                            .button(if font.favorite { "★" } else { "☆" })
                                            .on_hover_text(
                                                if font.is_system()
                                                    || !self
                                                        .settings
                                                        .move_local_fonts_with_favorites
                                                {
                                                    text("お気に入りを切り替え")
                                                } else if font.source == FontSource::FirstTeam {
                                                    text("プラグイン専用へ移動")
                                                } else {
                                                    text("ローカルフォントへ移動")
                                                },
                                            )
                                            .clicked()
                                        {
                                            self.toggle_favorite(index);
                                        }
                                        ui.vertical(|ui| {
                                            let badges = format!(
                                                "{}{}",
                                                text(font.source.label()),
                                                if font.axes.is_empty() { "" } else { " / VF" }
                                            );
                                            let response = ui.selectable_label(
                                                self.selected == Some(index),
                                                format!("{}  [{}]", font.display_name, badges),
                                            );
                                            response.context_menu(|ui| {
                                                if ui.button(text("フォント名をコピー")).clicked()
                                                {
                                                    ui.copy_text(font.family_name.clone());
                                                    ui.close();
                                                }
                                            });
                                            if response.clicked() {
                                                self.selected = Some(index);
                                                self.preview_dirty = true;
                                            }
                                            if response.double_clicked() {
                                                self.selected = Some(index);
                                                self.apply_selected_font();
                                            }
                                        });
                                    });
                                    });
                                }
                            });
                    },
                );
                self.list_rect = list_area.response.rect;
                if self.list_active {
                    // キーが効く状態かどうかを見て分かるようにする。
                    ui.painter().rect_stroke(
                        self.list_rect,
                        4.0,
                        ui.visuals().selection.stroke,
                        egui::StrokeKind::Inside,
                    );
                }
                ui.separator();
                ui.vertical(|ui| {
                    let selected = self.selected_font().cloned();
                    if let Some(font) = selected {
                        ui.horizontal(|ui| {
                            ui.heading(&font.family_name);
                            if ui.button(text("名前をコピー")).clicked() {
                                ui.copy_text(font.family_name.clone());
                                self.status = Some(text("フォント名をコピーしました"));
                            }
                        });
                        ui.label(format!(
                            "{}{}",
                            text(font.source.label()),
                            if font.favorite {
                                format!(" / {}", text("お気に入り"))
                            } else {
                                String::new()
                            }
                        ));
                        if let Some(path) = &font.path {
                            ui.small(path.display().to_string());
                        }
                        ui.label(if font.axes.is_empty() {
                            text("可変軸: なし")
                        } else {
                            font.axes
                                .iter()
                                .map(|axis| axis.label())
                                .collect::<Vec<_>>()
                                .join(" / ")
                        });
                        ui.add_space(8.0);
                        ui.horizontal(|ui| {
                            let sync_changed = ui
                                .selectable_label(self.settings.sync_text, text("選択に同期"))
                                .clicked();
                            let fixed_changed = ui
                                .selectable_label(!self.settings.sync_text, text("固定"))
                                .clicked();
                            if sync_changed {
                                self.settings.sync_text = true;
                                self.sync_selected_text();
                                self.save_settings();
                            } else if fixed_changed {
                                self.settings.sync_text = false;
                                self.save_settings();
                            }
                        });
                        let mut sample = self.settings.sample.clone();
                        if ui
                            .add(
                                egui::TextEdit::singleline(&mut sample)
                                    .hint_text(text("プレビューテキスト")),
                            )
                            .changed()
                        {
                            self.settings.sync_text = false;
                            self.set_sample(sample);
                        }
                        let mut preview_settings_changed = false;
                        ui.horizontal(|ui| {
                            ui.label(text("サイズ"));
                            preview_settings_changed |= ui
                                .add(
                                    egui::Slider::new(
                                        &mut self.settings.preview_font_size,
                                        8.0..=400.0,
                                    )
                                    .logarithmic(true)
                                    .suffix(" px"),
                                )
                                .changed();
                        });
                        ui.horizontal(|ui| {
                            ui.label(text("文字色"));
                            let mut text_color = egui::Color32::from_rgb(
                                self.settings.preview_text_color[0],
                                self.settings.preview_text_color[1],
                                self.settings.preview_text_color[2],
                            );
                            if egui::color_picker::color_edit_button_srgba(
                                ui,
                                &mut text_color,
                                egui::color_picker::Alpha::Opaque,
                            )
                            .changed()
                            {
                                self.settings.preview_text_color =
                                    text_color.to_array()[..3].try_into().unwrap();
                                preview_settings_changed = true;
                            }
                            ui.label(text("背景"));
                            let mut background_color = egui::Color32::from_rgb(
                                self.settings.preview_background_color[0],
                                self.settings.preview_background_color[1],
                                self.settings.preview_background_color[2],
                            );
                            if egui::color_picker::color_edit_button_srgba(
                                ui,
                                &mut background_color,
                                egui::color_picker::Alpha::Opaque,
                            )
                            .changed()
                            {
                                self.settings.preview_background_color =
                                    background_color.to_array()[..3].try_into().unwrap();
                                preview_settings_changed = true;
                            }
                        });
                        if preview_settings_changed {
                            self.preview_dirty = true;
                            self.save_settings();
                        }
                        let preview_size = self.update_preview(ui.ctx());
                        if let Some(texture) = &self.preview {
                            ui.add(egui::Image::new(texture).fit_to_exact_size(preview_size));
                        }
                        ui.add_space(8.0);
                        ui.horizontal_wrapped(|ui| {
                            let buttons = self.settings.create_buttons;
                            let shortcuts = self.settings.keys.create_active();
                            for kind in ObjectKind::ALL {
                                if !create_button_visible(buttons, kind) {
                                    continue;
                                }
                                let mut hint = text(kind.description());
                                if shortcuts {
                                    hint = format!("{hint}\n[{}]", kind.shortcut_digit());
                                }
                                if ui.button(text(kind.button_label())).on_hover_text(hint).clicked()
                                {
                                    self.create_object_with(&font, kind);
                                }
                            }
                            if ui.button(text("選択中に適用")).clicked() {
                                self.apply_selected_font();
                            }
                        });
                    } else {
                        ui.label(text("一致するフォントがありません"));
                    }
                    if let Some(status) = &self.status {
                        ui.separator();
                        ui.small(status);
                    }
                });
            });
        });
        self.show_options_window(ui.ctx());
        self.show_move_dialog(ui.ctx());
    }

    fn clear_color(&self, visuals: &egui::Visuals) -> [f32; 4] {
        visuals.window_fill.to_normalized_gamma_f32()
    }
}

fn synced_sample(
    sync_enabled: bool,
    consumed_revision: &mut u64,
    snapshot: &SelectedTextSnapshot,
) -> Option<String> {
    if !sync_enabled || snapshot.revision == 0 || snapshot.revision == *consumed_revision {
        return None;
    }
    *consumed_revision = snapshot.revision;
    match &snapshot.result {
        SelectedTextResult::Text(Some(text)) if !text.is_empty() => Some(text.clone()),
        SelectedTextResult::NotReady
        | SelectedTextResult::Text(None)
        | SelectedTextResult::Text(Some(_))
        | SelectedTextResult::Error(_) => None,
    }
}

/// その種別の追加ボタンを表示しているか。
///
/// キー操作もこれを見る。隠しているボタンのキーが生きていると、
/// 使えないオブジェクトをキーだけで作れてしまう。
fn create_button_visible(buttons: crate::settings::CreateButtons, kind: ObjectKind) -> bool {
    match kind {
        ObjectKind::Text => buttons.text,
        ObjectKind::VariableFontText => buttons.variable_font_text,
        ObjectKind::VariableFontObject => buttons.variable_font_object,
    }
}

/// 1始まりの番号を数字キーへ。割り当てのない番号は `None`。
fn digit_key(digit: usize) -> Option<egui::Key> {
    match digit {
        1 => Some(egui::Key::Num1),
        2 => Some(egui::Key::Num2),
        3 => Some(egui::Key::Num3),
        4 => Some(egui::Key::Num4),
        5 => Some(egui::Key::Num5),
        _ => None,
    }
}

/// キーを拾う範囲を決める。
///
/// - `enabled`: 設定で有効で、かつモーダルなダイアログが出ていない
/// - `focused`: 今フォーカスを持っているウィジェット
/// - `list_active`: 直近の押下がリスト内だったか
fn key_scope_for(enabled: bool, focused: Option<egui::Id>, list_active: bool) -> KeyScope {
    if !enabled {
        return KeyScope::None;
    }
    match focused {
        // 検索欄だけは「検索→選択→適用」を途切れさせないために特別扱いする。
        Some(id) if id == egui::Id::new(SEARCH_FIELD_ID) => KeyScope::Search,
        // 他のウィジェット（スライダー、ボタン、プレビューテキスト等）には譲る。
        Some(_) => KeyScope::None,
        None if list_active => KeyScope::List,
        None => KeyScope::None,
    }
}

/// この描画でリストが拾うキーの範囲。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum KeyScope {
    /// 何も拾わない。
    None,
    /// 検索欄にいる。↑↓ と Enter だけ。
    Search,
    /// リストを触っている。全てのキー。
    List,
}

/// 現在行から`delta`行動いた先。端では止まる（回り込まない）。
///
/// 未選択のときは、下方向なら先頭、上方向なら末尾から始める。
fn moved_row(current: Option<usize>, len: usize, delta: isize) -> Option<usize> {
    if len == 0 || delta == 0 {
        return None;
    }
    let last = len - 1;
    let Some(current) = current else {
        return Some(if delta > 0 { 0 } else { last });
    };
    let moved = (current as isize)
        .saturating_add(delta)
        .clamp(0, last as isize) as usize;
    (moved != current).then_some(moved)
}

fn filter_matches(font: &FontItem, filter: FilterMode) -> bool {
    match filter {
        FilterMode::All => true,
        FilterMode::System => font.source == FontSource::System,
        FilterMode::FirstTeam => font.source == FontSource::FirstTeam,
        FilterMode::Library => font.source == FontSource::Library,
        FilterMode::Favorites => font.favorite,
        FilterMode::Variable => !font.axes.is_empty(),
    }
}

fn compare_fonts(left: &FontItem, right: &FontItem, sort: SortMode) -> Ordering {
    let names = || {
        left.display_name
            .to_lowercase()
            .cmp(&right.display_name.to_lowercase())
    };
    match sort {
        SortMode::FavoriteName => right.favorite.cmp(&left.favorite).then_with(names),
        SortMode::NameAsc => names(),
        SortMode::NameDesc => names().reverse(),
        SortMode::Source => left.source.cmp(&right.source).then_with(names),
        SortMode::VariableFirst => right
            .axes
            .is_empty()
            .cmp(&left.axes.is_empty())
            .then_with(names),
    }
}

fn filter_label(filter: FilterMode) -> String {
    text(match filter {
        FilterMode::All => "すべて",
        FilterMode::System => "システム",
        FilterMode::FirstTeam => "ローカルフォント",
        FilterMode::Library => "プラグイン専用",
        FilterMode::Favorites => "お気に入り",
        FilterMode::Variable => "可変フォント",
    })
}

fn sort_label(sort: SortMode) -> String {
    text(match sort {
        SortMode::FavoriteName => "お気に入り優先",
        SortMode::NameAsc => "名前昇順",
        SortMode::NameDesc => "名前降順",
        SortMode::Source => "種類順",
        SortMode::VariableFirst => "可変フォント優先",
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hidden_create_buttons_have_no_shortcut() {
        let buttons = crate::settings::CreateButtons {
            variable_font_text: false,
            ..crate::settings::CreateButtons::default()
        };
        assert!(create_button_visible(buttons, ObjectKind::Text));
        assert!(!create_button_visible(
            buttons,
            ObjectKind::VariableFontText
        ));
        assert!(create_button_visible(
            buttons,
            ObjectKind::VariableFontObject
        ));
    }

    #[test]
    fn every_create_button_has_a_digit_key() {
        for kind in ObjectKind::ALL {
            assert!(digit_key(kind.shortcut_digit()).is_some());
        }
        assert!(digit_key(0).is_none());
    }

    #[test]
    fn keys_are_ignored_until_the_list_is_touched() {
        // 起動直後など、まだリストを押していない状態。
        assert_eq!(key_scope_for(true, None, false), KeyScope::None);
        assert_eq!(key_scope_for(true, None, true), KeyScope::List);
    }

    #[test]
    fn other_widgets_keep_their_own_keys() {
        // スライダーやボタンにフォーカスがある間は、リストがアクティブでも譲る。
        let slider = egui::Id::new("some-other-widget");
        assert_eq!(key_scope_for(true, Some(slider), true), KeyScope::None);
    }

    #[test]
    fn search_field_keeps_only_vertical_keys() {
        let search = egui::Id::new(SEARCH_FIELD_ID);
        assert_eq!(key_scope_for(true, Some(search), false), KeyScope::Search);
        assert_eq!(key_scope_for(true, Some(search), true), KeyScope::Search);
    }

    #[test]
    fn disabling_the_option_turns_every_key_off() {
        let search = egui::Id::new(SEARCH_FIELD_ID);
        assert_eq!(key_scope_for(false, None, true), KeyScope::None);
        assert_eq!(key_scope_for(false, Some(search), true), KeyScope::None);
    }

    #[test]
    fn arrow_movement_stops_at_both_ends() {
        assert_eq!(moved_row(Some(0), 3, -1), None);
        assert_eq!(moved_row(Some(0), 3, 1), Some(1));
        assert_eq!(moved_row(Some(2), 3, 1), None);
        assert_eq!(moved_row(Some(2), 3, -1), Some(1));
    }

    #[test]
    fn page_movement_clamps_instead_of_wrapping() {
        assert_eq!(moved_row(Some(1), 100, PAGE_STEP), Some(11));
        assert_eq!(moved_row(Some(95), 100, PAGE_STEP), Some(99));
        assert_eq!(moved_row(Some(4), 100, -PAGE_STEP), Some(0));
    }

    #[test]
    fn movement_without_a_selection_enters_from_the_matching_end() {
        assert_eq!(moved_row(None, 5, 1), Some(0));
        assert_eq!(moved_row(None, 5, -1), Some(4));
        assert_eq!(moved_row(None, 5, PAGE_STEP), Some(0));
    }

    #[test]
    fn movement_on_an_empty_list_is_a_no_op() {
        assert_eq!(moved_row(None, 0, 1), None);
        assert_eq!(moved_row(Some(0), 0, 1), None);
        assert_eq!(moved_row(Some(0), 3, 0), None);
    }

    fn font(name: &str, source: FontSource, favorite: bool, variable: bool) -> FontItem {
        let mut font = if source == FontSource::System {
            FontItem::new_system(name.to_string(), Vec::new())
        } else {
            FontItem::new_local(
                name.to_string(),
                PathBuf::from(format!("{name}.ttf")),
                source,
                Vec::new(),
            )
        };
        font.favorite = favorite;
        if variable {
            font.axes.push(crate::catalog::FontAxis::test_axis());
        }
        font
    }

    #[test]
    fn filters_first_team_and_favorites() {
        let first = font("First", FontSource::FirstTeam, true, false);
        let system = font("System", FontSource::System, true, false);
        assert!(filter_matches(&first, FilterMode::FirstTeam));
        assert!(!filter_matches(&system, FilterMode::FirstTeam));
        assert!(filter_matches(&system, FilterMode::Favorites));
    }

    #[test]
    fn favorite_sort_places_favorite_first() {
        let normal = font("A", FontSource::System, false, false);
        let favorite = font("Z", FontSource::System, true, false);
        assert_eq!(
            compare_fonts(&normal, &favorite, SortMode::FavoriteName),
            Ordering::Greater
        );
    }

    #[test]
    fn disabled_sync_does_not_consume_snapshot() {
        let snapshot = SelectedTextSnapshot {
            revision: 4,
            result: SelectedTextResult::Text(Some("event text".to_string())),
        };
        let mut consumed = 0;
        assert_eq!(synced_sample(false, &mut consumed, &snapshot), None);
        assert_eq!(consumed, 0);
    }

    #[test]
    fn enabling_sync_consumes_latest_snapshot_once() {
        let snapshot = SelectedTextSnapshot {
            revision: 4,
            result: SelectedTextResult::Text(Some("event text".to_string())),
        };
        let mut consumed = 0;
        assert_eq!(
            synced_sample(true, &mut consumed, &snapshot),
            Some("event text".to_string())
        );
        assert_eq!(consumed, 4);
        assert_eq!(synced_sample(true, &mut consumed, &snapshot), None);
    }

    #[test]
    fn empty_and_error_snapshots_are_consumed_without_text() {
        let mut consumed = 0;
        let empty = SelectedTextSnapshot {
            revision: 1,
            result: SelectedTextResult::Text(None),
        };
        assert_eq!(synced_sample(true, &mut consumed, &empty), None);
        assert_eq!(consumed, 1);

        let error = SelectedTextSnapshot {
            revision: 2,
            result: SelectedTextResult::Error("failed".to_string()),
        };
        assert_eq!(synced_sample(true, &mut consumed, &error), None);
        assert_eq!(consumed, 2);
    }
}
