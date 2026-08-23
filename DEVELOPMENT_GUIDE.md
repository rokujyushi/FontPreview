# Font Preview 開発ガイド

> 対象: `I:\FontPreview` の Rust 実装（`Cargo.toml` の version 4.2.0）  
> コード確認日: 2026-07-26

この文書は利用者向け README ではなく、実装を理解しながら開発するための設計図です。迷ったときは、まず「責務の境界」「主要フロー」「壊してはいけない前提」を見直してください。

## 1. このプラグインがしていること

Font Preview は AviUtl2 の汎用プラグインです。次の3領域をつないでいます。

1. Windows からシステムフォントを、AviUtl2 のデータ領域からローカルフォントを列挙する
2. 選んだフォントを DirectWrite / Direct2D で画像化し、egui に表示する
3. AviUtl2 の編集APIを介し、テキストオブジェクトの作成・更新・選択テキスト同期を行う

本体は `src/` の Rust 実装です。ルートの `FontPreviewNew.cpp`、`AxisMapping.h`、`KeyMapping.h`、`PresetIO.h`、`FontPreview.vcxproj` は旧 C++ 実装です。現在の `aviutl2.toml` は Cargo が作る DLL を成果物にしているため、通常の機能追加は Rust 側だけを変更します。

## 2. 最初に読む順番

実装全体を追うなら、次の順が最短です。

1. `src/lib.rs` — プラグインの生存期間、AviUtl2 イベント、スレッド間共有
2. `src/ui.rs` — 画面状態と、各モジュールを呼ぶタイミング
3. `src/catalog.rs` — フォントモデルと列挙
4. `src/preview.rs` — 選択フォントをプレビュー画像へ変換
5. `src/actions.rs` と `src/alias.rs` — AviUtl2 プロジェクトを変更する処理
6. `src/settings.rs` — 永続化とお気に入りの同一性
7. `src/fonts.rs` と `src/i18n.rs` — UIフォントと言語切り替え

## 3. モジュール構成

| ファイル | 責務 | 主な入口 |
|---|---|---|
| `lib.rs` | DLL/プラグイン登録、共有状態、イベント受信、同期ワーカー | `FontPreviewPlugin::new`, `register` |
| `ui.rs` | egui の状態、フィルタ・ソート、操作の組み立て | `FontPreviewApp::new`, `App::ui` |
| `catalog.rs` | フォント情報のモデル化、システム/ローカル列挙、VF軸取得 | `enumerate` |
| `preview.rs` | DirectWrite + Direct2D + WIC によるオフスクリーン描画 | `render` |
| `actions.rs` | ファイル取込/移動、オブジェクト作成/更新、テキスト取得 | 各 `pub(crate) fn` |
| `alias.rs` | AviUtl2 オブジェクト生成用エイリアス文字列 | `build`, `frame_length` |
| `settings.rs` | `settings.json` のロード/保存、お気に入り管理 | `Settings` |
| `fonts.rs` | egui の多言語フォールバックフォント | `definitions` |
| `i18n.rs` | AviUtl2 の翻訳機構を薄く包む | `text`, `format_text` |

依存方向は概ね次のとおりです。

```text
lib.rs
  ├─ ui.rs ─┬─ catalog.rs
  │         ├─ preview.rs
  │         ├─ actions.rs ── alias.rs
  │         └─ settings.rs
  ├─ actions.rs
  ├─ fonts.rs
  └─ i18n.rs（ほぼ全体から利用）
```

`ui.rs` は調停役であり、Windows API や AviUtl2 API の詳細は各モジュールへ押し出されています。この境界は維持した方が安全です。

## 4. 起動から終了まで

### 4.1 DLLの入口

`#[aviutl2::plugin(GenericPlugin)]` と `register_generic_plugin!` が DLL を AviUtl2 プラグインとして公開します。

`FontPreviewPlugin::new` は次を行います。

- i18n を利用可能にする
- AviUtl2 ログへ `tracing` を接続する
- `SharedEditState` を生成する
- `EframeWindow` を作り、画面生成時に `FontPreviewApp::new` を呼ぶ

この段階では AviUtl2 の `EditHandle` はまだ共有状態へ入りません。

### 4.2 `register`

AviUtl2 が `register` を呼ぶと、プラグインは次を登録します。

- `EditHandle`
- 選択テキスト取得用ワーカースレッド
- `Font Preview` ウィンドウ
- `.ttf/.otf/.ttc` のファイルドロップハンドラ

### 4.3 終了

`Drop` でワーカーへ shutdown を通知し、現在スレッド自身でない限り `join` します。ワーカーを増やす場合も、同様に明示的な停止経路を用意してください。

## 5. 画面状態

`FontPreviewApp` は画面の単一状態です。重要なフィールドは次のとおりです。

- `fonts`: 全フォント。カタログ再読込で置換される
- `filtered`: `fonts` の添字だけを持つ表示順
- `selected`: `fonts` に対する添字
- `settings`: 永続化対象
- `preview`: egui に載せたプレビュー用テクスチャ
- `preview_dirty`: 再描画が必要か
- `status`: 利用者へ見せる直近結果
- `selected_text_revision`, `font_drop_revision`: 共有スナップショットの二重消費防止
- `pending_move`: お気に入り連動のファイル移動を確認する一時状態
- `scroll_to_row`: 次の描画で見える位置までスクロールすべき行
- `preview_size`: 直前に描画したプレビューの実ピクセルサイズ
- `list_active` / `list_rect`: リストがキー操作の対象かと、その判定に使う矩形
- `show_options` / `options_rect`: 設定ウィンドウの開閉と、それが占めていた矩形

### 添字を扱うときの注意

`filtered` と `selected` はどちらも `fonts` の添字です。表示行番号ではありません。カタログ再読込や並べ替えの前後では、添字ではなく `FontId` で選択を退避・復元しています。

一方で `scroll_to_row` と `moved_row` が扱うのは **表示行番号**（`filtered` への添字）です。
この2種類の番号を混ぜないよう、`selected_row` / `select_row` を介して変換します。

### キーボード操作

`handle_list_keys` は `ui()` の**先頭**で呼ばれます。検索欄の `TextEdit` より先に
`consume_key` でキーを取り上げる必要があるためで、この順番は変えられません。

拾う範囲は `key_scope_for` が決めます。**キーを広く取り上げてはいけません**。
広く取ると、スライダーを矢印キーで微調整できなくなる、ボタンを押している間に
選択フォントが勝手に変わる、といった事故が起きます（実際に一度これを踏みました）。

拾うかどうかは **2段階**で決まります。

**1段階目、フォーカスによる範囲（`key_scope_for`）**

| 状態 | 範囲 |
|---|---|
| `keys.enabled` が off / 移動確認ダイアログ中 | `None` |
| 検索欄（`SEARCH_FIELD_ID`）にフォーカス | `Search`（↑↓ と Enter だけ） |
| 他のウィジェットにフォーカス | `None`（スライダーやボタンに譲る） |
| フォーカスなし かつ `list_active` | `List`（全キー） |
| フォーカスなし かつ `!list_active` | `None` |

**2段階目、種別ごとの設定（`settings::KeyBindings`）**

| フラグ | 対象のキー |
|---|---|
| `arrows` | ↑↓←→ / PageUp PageDown / Home End |
| `letters` | W S / A D |
| `enter_applies` | Enter |
| `favorite` | F |
| `create_objects` | 1 2 3（追加ボタンと同じ順） |

矢印キーだけ、WASDだけ、といった使い分けのために分けています。
問い合わせは必ず `arrows_active()` のようなメソッドを通してください。
フィールドを直接見ると、マスタースイッチ `enabled` を取りこぼします。
種別を増やすときは、フラグ・`*_active()`・設定ウィンドウの行・テストを揃えて追加します。

`list_active` は `track_list_activation` が更新します。直近のポインタ押下が
`list_rect` の中なら true、外なら false です。初期値は **false** で、一度もリストを
押していないうちはキーを拾いません。有効な間はリストを枠線で囲って知らせます。

`key_scope_for` は Context を取らない純関数なので、分岐は全てテストされています。
**拾う範囲を変えるときはこの関数とテストを一緒に更新してください。**

検索欄だけを特別扱いできるよう、`TextEdit::id` で固定 ID を与えています。
文字キーを増やすときは、検索中に拾ってしまわないか必ず確認してください。

リストは `ScrollArea::show_rows` で仮想化されており、**画面外の行には widget が存在しません**。
そのため `response.scroll_to_me` は使えず、行高が一定であることを利用して目標矩形を自分で組み立て、
`ui.scroll_to_rect` へ渡します。行のピッチは `ROW_HEIGHT` ではなく **`ROW_HEIGHT + item_spacing.y`** です。

### 行のウィジェットIDは `push_id` で固定する

`show_rows` は ID の整合を取るために `skip_ahead_auto_ids(min_row)` を呼びますが、これは
**1行にウィジェットが1つ**しか想定していません。本実装の1行は星ボタンとラベルの2つを使うため、
スクロールで先頭行が変わるたびに自動IDがずれ、egui が
`changed id between passes` の警告を出します（実際にログへ 72 件出ました）。

各行を `ui.push_id(&font.id, ...)` で包んで、ID を行位置から切り離しています。
**1行にウィジェットを追加するときは、この `push_id` の中に入れてください。**

### 画面の分け方

| 場所 | 置くもの |
|---|---|
| ツールバー | 常に使うものだけ（検索、絞り込み、並び替え、再読み込み、設定ボタン） |
| 詳細列 | 選択フォントに対する操作と、プレビューの見た目（サイズ・色・サンプル文字） |
| **設定ウィンドウ** | **持続的な振る舞いの設定と、参照用のキー一覧** |

設定ウィンドウは `show_options_window` が `egui::Window` で出します。
**オプションを増やすときはここにセクションを追加してください。**ツールバーは広げません。

ウィンドウはリストに重なり得るので、`track_list_activation` は `options_rect` に入った
クリックを除外します。ウィンドウを増やすときは、同じように矩形を除外するかを検討してください。

## 6. フォントのモデルと列挙

### 6.1 `FontItem`

1フォントは次を持ちます。

- `id`: 画面内での識別子
- `display_name`: ファイル名も含む表示名
- `family_name`: DirectWrite/AviUtl2 に渡すファミリー名
- `path`: ローカルフォントだけが持つ
- `source`: `System` / `FirstTeam` / `Library`
- `axes`: 値に幅のある可変軸
- `favorite`: 設定を反映した一時状態

ここで `FirstTeam` は AviUtl2 の `Font`、`Library` はプラグイン用の `FontLibrary` を意味します。

```text
<AviUtl2 app_data>\
  Font\                         ローカルフォント（優先側）
  FontLibrary\                  プラグイン専用保管側
  Plugin\FontPreview\
    settings.json
```

### 6.2 列挙フロー

`catalog::enumerate` は共有 DirectWrite factory を作り、次の順で結合します。

1. Windows のシステムフォントコレクション
2. `Font` 直下の `.ttf/.otf/.ttc`
3. `FontLibrary` 直下の `.ttf/.otf/.ttc`

ローカルディレクトリは再帰探索しません。読めないファイルは警告ログを出してスキップします。ファミリー名は DirectWrite の font set property を優先し、取得できなければファイル stem を使います。

### 6.3 可変フォント判定

`IDWriteFontFace5` → `IDWriteFontResource` から軸範囲を取得します。`minValue == maxValue` の固定軸は捨て、1つでも可変軸が残れば UI 上で VF として扱います。

軸タグは4バイトの OpenType tag です。画面表示用の文字列へ変換されますが、現状は軸値を編集したりプレビューへ適用したりはしていません。

## 7. プレビュー描画

`ui::update_preview` は、選択・サンプル文字・色・サイズが変わって `preview_dirty` になったときだけ `preview::render` を呼びます。

出力サイズは `preview_canvas` が決めます。

- 幅は詳細列に追従する。ウィンドウリサイズ中の再描画を抑えるため 16pt 刻みへ丸める
- 高さはフォントサイズに応じて伸び、`PREVIEW_CHROME_HEIGHT` を引いた値で頭打ち
- **実ピクセル**で描いて等倍で貼る。よってフォントサイズも `pixels_per_point` を掛ける

サイズが変わったことは `preview_size` との比較で検知し、`preview_dirty` を立てます。
この比較を外すと、ウィンドウ幅を変えてもプレビューが古い解像度のまま引き伸ばされます。

#### 寸法の元は `detail_area` だけ

`preview_canvas` は `ui.available_width()` / `ui.available_height()` を**見てはいけません**。
描画直前の残り領域はプレビュー画像自体の大きさに影響されるため、
「画像サイズ → レイアウト → 画像サイズ」の帰還路ができ、何も操作していないのに
再描画が繰り返されます（実際に一度これを踏みました）。

代わりに、CentralPanel で左右に分けた**直後**に `detail_area` を確定させ、それだけを元にします。
ログの `detail preview render start` には `canvas=WxH` が出るので、寸法が振動していないかはそこで見られます。

`preview::render` の流れは次のとおりです。

```text
FontItem
  ↓
システムフォント: system collection を使う
ローカルフォント: 対象ファイルだけの custom collection を作る
  ↓
DirectWrite TextFormat / TextLayout
  ↓
Direct2D で WIC bitmap に描画
  ↓
BGRA premultiplied → egui 用 RGBA
  ↓
TextureHandle
```

文字列は中央揃え、折り返しなしです。カラー絵文字等を考慮し `D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT` を使います。

### COMの扱い

描画ごとに `ComApartment::initialize` が `COINIT_MULTITHREADED` を試し、成功した場合だけ Drop で `CoUninitialize` します。既に別方式で初期化されているスレッドもあり得るため、「必ず自分が初期化した」と仮定しない実装です。

### プレビュー変更時のチェック

- ローカルフォントは OS インストール済みとは限らない
- DirectWrite collection の寿命が TextLayout/描画まで保たれること
- WIC の pixel format と RGBA 変換をセットで確認する
- 高頻度設定を増やす場合、毎フレーム `render` しない

## 8. AviUtl2 編集APIとの境界

編集処理は `actions.rs` に集約されています。

### オブジェクト作成

`create_object` は現在位置・現在レイヤーへ約1.1秒のオブジェクトを作ります。

1. `EditHandle::get_edit_info` から FPS を取得
2. `alias::frame_length` でフレーム数へ切り上げ変換
3. `alias::build` でエイリアス文字列を生成
4. `call_edit_section` 内で `create_object_from_alias`

生成対象は標準 `テキスト`、`Variable Font Text`、`Variable Font Object` の3種です。

ボタンのラベルと説明は `ObjectKind::button_label` / `description` が持ち、詳細列のボタンと
設定ウィンドウのチェックボックスで共用します。**種別を増やすときはこの2つの `match` を
埋めれば、両方の表示が揃います。**

表示するボタンは `settings::CreateButtons` で個別に切れます。Variable Font 系は別の
プラグインが入っていない環境では押しても失敗する（AviUtl2 のログに
`not found effect. ... [Variable Font Text]` が出る）ためです。

キーボードからも `1` `2` `3` で同じものを作れます。番号は `ObjectKind::shortcut_digit`、
並びは `ObjectKind::ALL` で、この2つが一致することはテストで固定しています。

**隠しているボタンのキーは受け付けません。** ボタンを隠すのは「その環境では使えない」からであり、
キーだけ生きていると同じ失敗をキーで踏めてしまいます。判定は `ui::create_button_visible` に
集めてあり、ボタン描画とキー処理の両方がこれを通ります。標準テキストにはファミリー名を、可変フォント系にはシステムならファミリー名、ローカルならファイルパスを入れます。

エイリアスの項目名と effect 名は外部プラグインとの契約です。表記変更は単なるリファクタリングではありません。

### 選択オブジェクトへ適用

`apply_to_selection` は選択オブジェクトを使い、空ならフォーカス中オブジェクトへフォールバックします。対象 effect に対する item 更新が1つでも成功したオブジェクトを更新件数に数えます。

- システムフォント: 標準テキストとVF系の `フォント` を設定
- ローカルフォント: VF系の `フォント` を空にし、`フォントファイル` を設定

どの対応 effect もなければエラーになります。

### 選択テキストの取得

`selected_text` はフォーカス中オブジェクトを優先し、それがなければ最初の選択オブジェクトを使います。次の effect をこの順に調べ、最初の空でない `テキスト` を返します。

1. `テキスト`
2. `Variable Font Text`
3. `Variable Font Object`

この優先順はテストで固定されています。

## 9. 選択テキスト同期の並行処理

この部分は保守時に最も注意が必要です。

AviUtl2 から届く次のイベントは `request_selected_text_refresh` を呼ぶだけです。

- object 情報更新
- scene 変更
- focus object 変更

実際の取得は専用ワーカーが行います。

```text
AviUtl2 event
  ↓ request（複数なら集約）
Condvar
  ↓ 150ms debounce
selected_text_worker
  ↓ actions::selected_text
SelectedTextSnapshot { revision, result }
  ↓ request_repaint
egui UI
  ↓ 未消費 revision のみ反映
sample / preview_dirty / settings.json
```

### なぜこの構造か

- 編集イベント内で重い読取処理を行わない
- 短時間に連続するイベントをまとめる
- UIスレッドとワーカー間で生の可変参照を共有しない
- 同じ値でもイベント単位で revision を進め、更新の事実を失わない

`SharedEditState` は `OnceLock` で `EditHandle` と egui context を1回だけ受け取り、結果は `Mutex` で保護されたスナップショットとして公開します。Mutex が poison しても内部値を回収する方針です。

### 変更時の禁止事項に近い注意

- `App::ui` から毎フレーム `selected_text` を同期呼出ししない
- ワーカーから egui の画面状態を直接変更しない
- revision を値の同一性だけで省略しない
- shutdown フラグと Condvar wake を片方だけ変更しない

## 10. ファイルドロップとカタログ更新

ドロップハンドラは `.ttf/.otf/.ttc` を `<app_data>\Font` へコピーします。

- 同じ実体を同じ場所へドロップ: `AlreadyPresent`
- 同名の別ファイルが既にある: 上書きせずエラー
- コピー成功: shared snapshot を更新し repaint
- UI: 未消費 revision を見つけ、カタログを再列挙

ここでもワーカー/コールバック側は UI 状態を直接触らず、`FontDropSnapshot` を経由します。

## 11. お気に入りとファイル移動

お気に入りには2種類のID戦略があります。

- システムフォント: `FontItem.id`
- ローカルフォント: family 名 + 小文字化したファイル名

ローカル用IDに `FirstTeam` / `Library` を含めないのは、ファイルを両フォルダ間で移動してもお気に入りを維持するためです。

`move_local_fonts_with_favorites` が有効なら、星を付ける操作で `FontLibrary → Font`、外す操作で `Font → FontLibrary` の移動確認を出します。移動成功後にお気に入りを切り替え、保存し、カタログを再読込します。

移動は `rename` であり、移動先の同名ファイルを上書きしません。AviUtl2 の別メニューへの反映には再起動が必要です。

## 12. 設定と互換性

保存先:

```text
<AviUtl2 app_data>\Plugin\FontPreview\settings.json
```

現在の保存項目:

- システム/ローカルのお気に入り
- お気に入りとローカルファイル移動の連動
- フィルタ、ソート
- 選択テキスト同期
- サンプル文字列
- プレビューの文字サイズ、文字色、背景色
- キーボード操作（`keys`：マスターと種別ごとのフラグ。全て既定で有効）
- 追加ボタンの表示（`create_buttons`：種別ごとのフラグ。全て既定で表示）

`Settings` には `#[serde(default)]` があるため、古い JSON に新しいフィールドがなくても既定値で補えます。設定項目を追加するときは `Settings` と `Default` の両方を更新し、旧形式を読むテストを追加してください。

**bool を追加するときは特に注意してください。**コンテナの `#[serde(default)]` は
欠落フィールドを構造体の `Default` から埋めるので、`Default` で `true` にしていれば
既存の settings.json でも有効になります。`bool::default()`（= `false`）にはなりませんが、
逆になると困る値なので `KeyBindings` についてはテストで固定しています。

ロード失敗時は既定値へ戻り、エラー文字列を UI status に渡します。保存は一時ファイル+置換ではなく直接 `write` なので、異常終了時の厳密な耐障害性が必要になったら改善候補です。

## 13. 多言語対応

UI文字列は `i18n::text` または `format_text` を通し、AviUtl2 の `config::translate` に委譲します。翻訳ファイルはルートの `.aul2` ファイルで、リリース時に `Language/` へ配置されます。

`fonts::definitions` は AviUtl2 標準の egui フォント定義に、翻訳キー `__FontFallbackFonts` で指定されたシステムフォントを追加します。カンマ区切りの候補を順に試し、最初に見つかった1書体を proportional / monospace の両方へ末尾追加します。

新しい表示文言を追加するときは、Rust 側だけで終わらせず各言語ファイルも確認します。文字列はAPIの識別子にも使われる箇所があるため、effect/item 名まで翻訳してよいわけではありません。

## 14. ビルド、プレビュー、リリース

通常の確認:

```powershell
cargo test
cargo clippy --all-targets -- -D warnings
cargo build --release
```

成果物は `target/release/FontPreview.dll` です。`aviutl2.toml` は release profile でこれを `Plugin/FontPreview/FontPreview.aux2` としてパッケージ化します。

`aviutl2.toml` の役割:

- `build_group`: Cargo の debug/release コマンド
- `preview`: `.aviutl2-cli/preview` へ検証環境を作る
- `release`: `release/` へ配布物を作る
- `artifacts`: DLL、README、LICENSE、言語ファイル、`package.txt` の配置

ルートの `build_release.ps1` は旧 Visual Studio / C++ ビルドを呼びます。Rust版の日常ビルド手順とは別物です。

## 15. テストの考え方

現状の単体テストは次の境界を守っています。

- フィルタ・ソート順
- 同期 snapshot の一度だけ消費する規則
- お気に入りIDとフォルダ移動後の継続性
- 設定 JSON の既定値互換
- フォント取込/移動時の衝突と形式チェック
- effect 優先順
- エイリアスのローカルフォント表現とFPS変換
- refresh request の集約、shutdown、revision
- 軸判定とUIフォールバック候補

Windows API と AviUtl2 実環境を使う部分は単体テストだけでは保証できません。次の変更では実機確認も必要です。

- DirectWrite/Direct2D/WIC の変更
- alias の effect/item 名変更
- `EditHandle` の呼出方法変更
- 言語ファイル、パッケージ配置
- Font/FontLibrary の実ファイル操作

## 16. よくある変更の進め方

### フィルタを追加する

1. `settings::FilterMode` に variant
2. `ui::filter_label`
3. UI の選択肢
4. `ui::filter_matches`
5. 旧 settings JSON 互換とフィルタテスト
6. 各言語ファイル

### ソートを追加する

1. `settings::SortMode`
2. `ui::sort_label`
3. UI の選択肢
4. `ui::compare_fonts`
5. tie-breaker が安定しているかテスト

### 設定項目を追加する

1. `Settings` のフィールドと `Default`
2. UI変更時の `save_settings`
3. 描画に関係するなら `preview_dirty = true`
4. 欠落フィールドを含む旧 JSON のロードテスト

### 対応するテキスト系 effect を増やす

1. `actions::TEXT_EFFECTS` の読取優先順
2. `apply_to_selection` の更新項目
3. 必要なら `alias::ObjectKind` と `build`
4. UI の操作ボタン
5. 外部 effect の正確な名前・item 名を実機確認

### フォント形式を増やす

拡張子チェックが複数箇所にあります。

- `catalog::local_fonts`
- `actions::import_font_file`
- `lib.rs` の file drop filter

DirectWrite が実際に読み込めるか、パッケージ/利用説明も含めて確認します。

### 非同期処理を追加する

`SharedEditState` の snapshot + revision + repaint パターンを再利用します。UIの所有権を別スレッドへ渡さず、終了時に必ず停止できる構造にします。

## 17. 現在の設計上の割り切りと改善候補

これは即修正リストではなく、変更時に判断材料とするメモです。

- `ui.rs` が大きい: 画面領域や操作単位へ分割する余地がある
- プレビューは毎回 DirectWrite/Direct2D/WIC の factory と collection を組み立てる: キャッシュは可能だが COM/thread/lifetime 設計が必要
- カタログ再読込は全列挙: フォント数が多い環境ではバックグラウンド化の余地がある
- ローカルフォント探索は直下だけ: 再帰対応するとID・重複・移動仕様も変わる
- `settings.json` は直接上書き: atomic save ではない
- ローカルお気に入りIDは family+ファイル名: 別ディレクトリの同名・同familyを区別しない
- `FontItem.id` は画面内ID、永続IDではない: 用途を混ぜない
- 可変軸は検出/表示だけ: 軸値をプレビューへ渡す機能は未実装
- 旧 C++ ソースがルートに同居: Rust版との正本が曖昧になりやすい
- README や一部ファイルが環境によって文字化け表示される場合がある: 編集時は UTF-8 とツールの decoding を確認する

## 18. 壊してはいけない前提

最後に、レビュー時の短いチェックリストです。

- UIは `fonts` の添字と表示行番号を混同していないか
- カタログ置換後に選択を `FontId` で復元しているか
- ローカルフォントをシステムインストール済みと仮定していないか
- ファイル取込/移動で既存ファイルを上書きしていないか
- AviUtl2 の編集処理は `call_edit_section` / `call_read_section` 内か
- event callback や UI frame を重い同期処理で塞いでいないか
- snapshot の revision を一度だけ消費しているか
- preview に影響する変更で `preview_dirty` を立てたか
- 設定追加に serde の旧形式互換があるか
- effect/item 名と翻訳対象の表示文字列を混同していないか
- ワーカーに shutdown と join があるか
- `cargo test`、`cargo clippy --all-targets -- -D warnings`、実機確認のうち必要なものを行ったか

## 19. 調査するときの入口

症状から読む場所を逆引きできます。

| 症状 | 最初に見る場所 |
|---|---|
| フォントが一覧に出ない | `catalog::enumerate`, `local_fonts`, DirectWrite の警告ログ |
| VF と判定されない | `catalog::collect_axes` |
| 一覧の検索/順序がおかしい | `ui::rebuild_filter`, `filter_matches`, `compare_fonts` |
| キー操作が効かない/余計なキーまで拾う | `ui::key_scope_for`, `list_active`, `SEARCH_FIELD_ID` |
| 設定を増やしたい | `ui::show_options_window` と `settings::Settings` |
| 追加ボタンが出ない/押しても失敗する | `settings::CreateButtons` と、AviUtl2 ログの `not found effect.` |
| 1 2 3 で追加できない | `keys.create_objects`, `create_button_visible`, `digit_key` |
| 選択行までスクロールしない/行がずれる | `ui::scroll_to_row` と `ROW_HEIGHT + item_spacing.y` のピッチ |
| プレビューがぼやける/大きくならない | `ui::preview_canvas`, `preview_size` |
| プレビューだけ失敗する | `ui::update_preview`, `preview::render` |
| オブジェクトを追加できない | `actions::create_object`, `alias::build` |
| 選択へ適用できない | `actions::apply_to_selection` と effect/item 名 |
| 選択テキスト同期が遅い/更新されない | `lib.rs` の refresh signal/worker と `ui::synced_sample` |
| お気に入りが消える | `settings::local_favorite_id`, `apply_favorites` |
| 星操作でファイル移動に失敗 | `ui::PendingMove`, `actions::move_local_font` |
| ドロップ後に反映されない | `import_dropped_font`, `FontDropSnapshot`, `process_font_drop` |
| 翻訳されない/文字が欠ける | `i18n.rs`, `fonts.rs`, `.aul2` 言語ファイル |
| 配布物に入らない | `aviutl2.toml` の `artifacts` |

このガイドとコードが食い違う場合はコードが現状の事実です。ただし、意図が変わったのか偶発的にずれたのかを判断してから、コードとこの文書を同じ変更で更新してください。
