# 変更履歴

このファイルは Font Preview のリリースノートです。日付は JST です。

## v5.0.0 — 2026-08-23

多言語対応と、操作性まわりの大きな作り直しを含むメジャー更新です。

### 新機能

#### 多言語対応

- 英語・簡体中文の差分言語ファイル（`English.FontPreview.aul2` / `简体中文.FontPreview.aul2`）を同梱しました。AviUtl2 の「設定」→「言語の設定」で切り替えると、画面表示・操作結果・ドラッグ＆ドロップ項目まで追従します。日本語は既定表示です。
- UI フォントは AviUtl2 のものを優先し、足りないグリフは言語ファイルの `__FontFallbackFonts` に書かれたシステムフォントで補います。簡体中文では `Microsoft YaHei UI` → `Microsoft YaHei` → `SimSun` の順に探します。
- エラーメッセージも翻訳対象になりました。
- インストール説明（`package.txt`）を日本語・英語・簡体中文の併記にしました。

#### 設定ウィンドウ

- ツールバーにあった設定項目を、独立した「設定」ウィンドウへ移しました。ツールバーには検索・絞り込み・並び替えだけが残ります。
- 「動作」「追加ボタン」「キー操作」の 3 セクション構成です。

#### キーボードでのリスト操作

| キー | 動作 |
|---|---|
| ↑ ↓ / W S | 選択を 1 件移動 |
| PageUp PageDown / ← → / A D | 10 件移動 |
| Home / End | 先頭 / 末尾へ移動 |
| Enter | 選択中のオブジェクトへ適用 |
| F | お気に入りを切り替え |
| 1 2 3 | 追加ボタンと同じ順にオブジェクトを追加 |

- **リストをクリックしてから有効**になります。プレビュー側を操作しているだけ、ウィンドウを開いただけでは選択フォントは変わりません。有効な間はリストが枠線で囲まれます。
- 検索欄に入力中は ↑ ↓ と Enter だけを受け付けるので、「検索 → 選択 → 適用」がキーだけで通ります。W A S D などは文字として入力されます。
- スライダーやボタンにフォーカスがあるときは、そちらにキーを譲ります。
- 種別ごとに個別で無効にできます（矢印キーだけ、WASD だけ、といった使い分けが可能です）。まとめて切ることもできます。

#### 追加ボタンの表示切り替え

- 「テキスト +」「VF +」「VFO +」を個別に隠せるようになりました。
- Variable Font 系は対応プラグインが入っていない環境では押しても失敗するため、使わないものを隠しておけます。隠したボタンは数字キーでも実行されません。

### 改善

- **プレビューの拡大**: 描画キャンバスが 640×220 固定でしたが、パネル幅に追従するようになりました。実ピクセルで描いて等倍で表示するため、以前の縮小表示によるぼやけがなくなります。
- フォントサイズの上限を 160px → **400px** に引き上げ、スライダーを対数目盛りにしました。小さい側も刻みやすくなっています。
- プレビューが選択フォントの実際のウェイトで描かれるようになりました。以前は常に Regular 相当で描いていたため、同じファミリの W4 と W8 が同じ絵になっていました。
- キーボードで選択を動かしたとき、選択行が見える位置まで自動でスクロールします。検索や並び替えで行位置が変わったときも同様です。

### 修正

- リストをスクロールすると egui が `changed id between passes` の警告を大量に出し、行のホバー状態やクリック判定が別の行のものと入れ替わり得た問題を修正しました。
- ウィンドウのレイアウトとプレビューのサイズが相互に影響し合い、操作していないのにプレビューの再描画が繰り返される問題を修正しました。

### 動作が変わる点

- 「星操作でローカルフォント / プラグイン専用を連動」のチェックボックスがツールバーから設定ウィンドウへ移動しました。
- プレビューのサイズ指定が実ピクセル基準になったため、同じ数値でも以前と見た目の大きさが変わります。必要に応じて設定し直してください。
- キー操作は既定で有効ですが、リストを一度クリックするまでは効きません。

### 開発者向け

- `DEVELOPMENT_GUIDE.md` を追加しました。責務の境界、主要フロー、壊してはいけない前提をまとめています。

---

## 既知の制限

- モリサワ・フォントワークスなどのサブスクリプション書体で、一覧から適用してもフォントが変わらない場合があります。多ウェイト和文で DirectWrite のファミリ名と AviUtl2 が持つフォント名が食い違うことが原因で、**このリリースには修正が含まれていません**。

## v4.2.0 以前

`git log` および GitHub のリリースページを参照してください。

---

## English summary (v5.0.0)

- **Localization**: bundled English and Simplified Chinese language files; the UI, action results, and drag-and-drop entries follow AviUtl2's language setting. UI fallback fonts fill in missing glyphs.
- **Settings window**: options moved out of the toolbar into a dedicated window with Behavior / Add Buttons / Keyboard sections.
- **Keyboard control**: arrow keys, W A S D, PageUp/PageDown, Home/End, Enter to apply, F to toggle favorite, and 1 2 3 to add objects. Active only after clicking the list; each group can be turned off individually.
- **Add-button visibility**: the Text / VF / VFO buttons can be hidden individually. Hidden buttons ignore their number key.
- **Larger preview**: the canvas now follows the panel width and renders at real pixels, so it is no longer blurred by downscaling. The size limit rose from 160px to 400px, and the preview uses the font's actual weight.
- **Fixes**: widget-id churn while scrolling the list, and a layout feedback loop that redrew the preview repeatedly.
- **Known limitation**: subscription fonts (Morisawa, Fontworks) may still fail to apply. Not fixed in this release.

## 简体中文摘要 (v5.0.0)

- **多语言**：内置英语与简体中文语言文件，界面、操作结果与拖放项目会跟随 AviUtl2 的语言设置。
- **设置窗口**：选项从工具栏移至独立窗口，分为「行为」「添加按钮」「快捷键」三节。
- **键盘操作**：方向键、W A S D、PageUp/PageDown、Home/End、Enter 应用、F 切换收藏、1 2 3 添加对象。需先点击列表才生效，各类可单独关闭。
- **添加按钮可隐藏**：可单独隐藏「文本 +」「VF +」「VFO +」。已隐藏的按钮其数字键也无效。
- **预览放大**：画布随面板宽度变化并按实际像素绘制，不再因缩小而模糊。字号上限由 160px 提升至 400px，并按字体实际字重绘制。
- **修复**：滚动列表时的控件 ID 错乱，以及导致预览反复重绘的布局回环。
- **已知限制**：森泽、Fontworks 等订阅字体可能仍无法应用，本次未修复。
