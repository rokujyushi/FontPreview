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
        self.preview_dirty = true;
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

    fn update_preview(&mut self, ctx: &egui::Context) {
        if !self.preview_dirty {
            return;
        }
        self.preview_dirty = false;
        let Some(font) = self.selected_font() else {
            self.preview = None;
            return;
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
            "FontPreview detail preview render start"
        );
        match crate::preview::render(
            font,
            preview_text,
            self.settings.preview_background_color,
            self.settings.preview_text_color,
            640,
            220,
            self.settings.preview_font_size,
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

    fn action_result(&mut self, result: anyhow::Result<()>) {
        self.status = Some(match result {
            Ok(()) => text("操作を完了しました"),
            Err(error) => format!("{}: {error:#}", text("操作に失敗しました")),
        });
    }
}

impl eframe::App for FontPreviewApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
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
                });
            });
            let search_changed = ui
                .add(
                    egui::TextEdit::singleline(&mut self.search)
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
            if ui
                .checkbox(
                    &mut self.settings.move_local_fonts_with_favorites,
                    text("星操作でローカルフォント / プラグイン専用を連動"),
                )
                .on_hover_text(text(
                    "有効にすると、お気に入り登録時はFontへ、解除時はFontLibraryへ移動します",
                ))
                .changed()
            {
                self.save_settings();
            }
        });

        egui::CentralPanel::default().show(ui, |ui| {
            let list_width = (ui.available_width() * 0.48).clamp(260.0, 430.0);
            let content_height = ui.available_height();
            ui.horizontal(|ui| {
                ui.allocate_ui_with_layout(
                    egui::vec2(list_width, content_height),
                    egui::Layout::top_down(egui::Align::Min),
                    |ui| {
                        ui.label(format_text(
                            "{shown} / {total}件",
                            &[
                                ("{shown}", self.filtered.len().to_string()),
                                ("{total}", self.fonts.len().to_string()),
                            ],
                        ));
                        egui::ScrollArea::vertical()
                            .auto_shrink([false, false])
                            .max_height(ui.available_height())
                            .show_rows(ui, 26.0, self.filtered.len(), |ui, rows| {
                                for row in rows {
                                    let index = self.filtered[row];
                                    let font = self.fonts[index].clone();
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
                                                self.status = Some(
                                                    match self
                                                        .shared_edit
                                                        .edit_handle()
                                                        .ok_or_else(|| {
                                                            anyhow::anyhow!(text(
                                                                "AviUtl2の編集機能を初期化中です"
                                                            ))
                                                        })
                                                        .and_then(|handle| {
                                                            actions::apply_to_selection(
                                                                handle, &font,
                                                            )
                                                        }) {
                                                        Ok(count) => format_text(
                                                            "{count}個のオブジェクトを更新しました",
                                                            &[("{count}", count.to_string())],
                                                        ),
                                                        Err(error) => {
                                                            format!(
                                                                "{}: {error:#}",
                                                                text("更新に失敗しました")
                                                            )
                                                        }
                                                    },
                                                );
                                            }
                                        });
                                    });
                                }
                            });
                    },
                );
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
                                        8.0..=160.0,
                                    )
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
                        self.update_preview(ui.ctx());
                        if let Some(texture) = &self.preview {
                            let size = texture.size_vec2();
                            let scale = (ui.available_width() / size.x).min(1.0);
                            ui.add(egui::Image::new(texture).fit_to_exact_size(size * scale));
                        }
                        ui.add_space(8.0);
                        ui.horizontal_wrapped(|ui| {
                            if ui.button(text("テキスト +")).clicked() {
                                self.action_result(
                                    self.shared_edit
                                        .edit_handle()
                                        .ok_or_else(|| {
                                            anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です"))
                                        })
                                        .and_then(|handle| {
                                            actions::create_object(
                                                handle,
                                                &font,
                                                &self.settings.sample,
                                                ObjectKind::Text,
                                            )
                                        }),
                                );
                            }
                            if ui.button(text("VF +")).clicked() {
                                self.action_result(
                                    self.shared_edit
                                        .edit_handle()
                                        .ok_or_else(|| {
                                            anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です"))
                                        })
                                        .and_then(|handle| {
                                            actions::create_object(
                                                handle,
                                                &font,
                                                &self.settings.sample,
                                                ObjectKind::VariableFontText,
                                            )
                                        }),
                                );
                            }
                            if ui.button(text("VFO +")).clicked() {
                                self.action_result(
                                    self.shared_edit
                                        .edit_handle()
                                        .ok_or_else(|| {
                                            anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です"))
                                        })
                                        .and_then(|handle| {
                                            actions::create_object(
                                                handle,
                                                &font,
                                                &self.settings.sample,
                                                ObjectKind::VariableFontObject,
                                            )
                                        }),
                                );
                            }
                            if ui.button(text("選択中に適用")).clicked() {
                                self.status = Some(
                                    match self
                                        .shared_edit
                                        .edit_handle()
                                        .ok_or_else(|| {
                                            anyhow::anyhow!(text("AviUtl2の編集機能を初期化中です"))
                                        })
                                        .and_then(|handle| {
                                            actions::apply_to_selection(handle, &font)
                                        }) {
                                        Ok(count) => format_text(
                                            "{count}個のオブジェクトを更新しました",
                                            &[("{count}", count.to_string())],
                                        ),
                                        Err(error) => {
                                            format!("{}: {error:#}", text("更新に失敗しました"))
                                        }
                                    },
                                );
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
