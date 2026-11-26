# LED室内現況調査システム（Survey123 ＋ PC側MCP）

現地入力は **ArcGIS Survey123 Field App**、帳票作成は **PC のローカル MCP（Claude デスクトップアプリから利用）** で行います。
図面がない施設でも、部屋リスト・室名札写真・部屋寸法・天井写真で調査できる構成です。

```
Survey123 Field App（スマホ・タブレット／オフライン可）
   │  写真・音声・位置・器具・部屋寸法（送信時に同期）
   ▼
ArcGIS Online（翔テクノロジーの組織。フォームは組織内共有のまま）
   │  クエリと添付ダウンロードのみ（クレジット消費なし）
   ▼
PC のローカル MCP「led_survey123」  ← Claude デスクトップアプリから指示
   └ 器具台帳.xlsx／写真帳.pdf／簡易配置図（DXF・PDF）／報告地点.geojson
```

## フォルダ構成

| パス | 内容 |
| --- | --- |
| `survey123_form/led_genkyo_chosa.xlsx` | Survey123 Connect で公開するフォーム（XLSForm） |
| `survey123_form/media/*.csv` | 器具タイプ・LED置換・用途別点灯時間・対象施設（facilities.csv）のマスタ（フォームとMCPで共用） |
| `mcp_server/led_survey_mcp/` | MCP サーバー本体 |
| `mcp_server/tests/make_sample.py` | 動作確認用の合成データ作成 |

## 公園用_現場写真管理フォームの機能との対応

| 公園用フォームでできること | このシステムでの扱い |
| --- | --- |
| 報告者氏名 | `reporter`（お気に入り回答で固定可） |
| 対象地名（選択式） | 建物名・階（選択式）・通し番号・部屋名、現場写真報告では「対象地名・場所」 |
| 報告事項の繰り返し入力 | `reports`（どの記録の種類でも何件でも追加可） |
| 該当写真 | `rp_photo`（複数枚可）＋部屋・器具・銘板・天井・案内図の写真 |
| 現場メモ（1000字） | `rp_memo`（1000字制限つき） |
| 音声 | `rp_audio`（最大180秒） |
| 撮影場所（地図上の位置） | `rp_loc`（報告ごと）＋`location`（記録ごと） |
| コメント | `rp_comment` |
| オフライン入力・地図表示・Excel出力（Survey123標準） | そのまま利用＋MCPで台帳・写真帳・GeoJSON を作成 |

## 導入手順

### 1. フォームを公開する（Survey123 Connect）
1. Survey123 Connect で「新しいサーベイ」→「ファイルから」で `led_genkyo_chosa.xlsx` を選ぶ
2. 作成されたサーベイフォルダの `media` に `survey123_form/media/*.csv` の4ファイルをコピー（`facilities.csv` は案件ごとに差し替え。現在は南大東村15施設）
3. 「公開」→ 共有先は組織内のグループ
4. Connect の「オプション」で画像サイズを 1920px にしておく（フォーム側でも `max-pixels=1920` を指定済み）

> 写真への印付け（annotate）や描画（draw）は **Survey123 Field App** の機能です。新しい Survey123 Mobile では対応予定の段階なので、調査員の端末には Field App を入れてください。

### 2. マスタを自社の値にする
`fixture_types.csv` の消費電力（W/台）と `led_rules.csv` の品番・消費電力は**仮値**です。自社の採用品番・メーカーカタログ値に置き換えてから、
フォームを再公開し、同じファイルを `mcp_server/led_survey_mcp/masters/` にもコピーしてください（または環境変数 `LED_SURVEY_MASTER_DIR` で `survey123_form/media` を指定）。
`usage_hours.csv`（用途別の点灯時間の初期値）も同様に調整してください。

### 3. MCP を入れる（Windows）
```bat
cd %USERPROFILE%\led_survey_system\mcp_server
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```
ArcGIS Pro が入っている PC では、Pro の Python（arcgis 導入済み）を使っても構いません。

ArcGIS へのログインは、パスワードを設定ファイルに書かずにプロファイルで保存します（1回だけ実行）：
```bat
.venv\Scripts\python -c "from arcgis.gis import GIS; GIS('https://www.arcgis.com', '<ユーザー名>', profile='sho_tech')"
```

`claude_desktop_config_例.json` を参考に、Claude デスクトップアプリの設定（`%APPDATA%\Claude\claude_desktop_config.json`）へ追記し、アプリを再起動します。

### 4. 使い方（Claude への指示例）
- 「LED調査MCPで、案件『○○市役所』をアイテムID xxxxx から取得して」 → `survey_fetch`
- 「概要と部屋の漏れを確認して」 → `survey_summary` / `coverage_check`
- 「1F-03 の銘板写真から型番を読み取って台帳に反映して」 → `list_records` → `view_photo` → `record_correction`
- 「台帳・写真帳・配置図をまとめて作って」 → `build_all`
- 「報告地点を hazard_map_mcp で地図に重ねて」 → `export_map_points`

ArcGIS API が使えない場合は、Survey123 Web サイトの「データ」→「エクスポート」→「Excel」で書き出したファイルを `survey_load_export` で読み込めます（属性のみ。写真は `survey_fetch` が必要）。

出力先：`LED_SURVEY_WORKDIR\<案件名>\出力\<日付>\`

## ツール一覧

| ツール | 内容 | ArcGIS への影響 |
| --- | --- | --- |
| survey_fetch | 回答・写真・音声を取得してキャッシュ | 読み取りのみ |
| survey_load_export | Excel書き出しを読込 | なし |
| app_load_export | スマホアプリ「LED現調」の書き出しzipを取り込み（複数zipの統合可） | なし |
| survey_summary | 部屋数・器具台数・削減電力量・要確認・漏れ警告 | なし |
| coverage_check | 階ごとの登録部屋数と扉の数・申告部屋数の突合 | なし |
| list_records | 部屋・器具・報告事項と写真パスの一覧 | なし |
| view_photo | 写真を Claude に見せる（型番・室名札の読み取り） | なし |
| record_correction | 読み取り結果などの補正をPC側に記録 | なし（ArcGISは変更しない） |
| build_ledger | 器具台帳.xlsx（条件・施設一覧・部屋一覧・器具台帳・数量集計・省エネ算定・網羅チェック・報告事項） | なし |
| build_photo_album | 写真帳.pdf | なし |
| build_layout | 簡易配置図（DXF・PDF） | なし |
| export_map_points | 報告地点.geojson | なし |
| export_photos | 写真を施設ごとのフォルダに「施設名_階_場所_撮影箇所_撮影日_連番」で整理し、写真一覧.xlsx を作成 | なし |
| build_all | 上記の帳票をまとめて作成 | なし |

## 注意
- 電力量単価（初期値30円/kWh）と CO2 排出係数（初期値0.434kg-CO2/kWh）は仮値です。台帳の「条件」シートの黄色セルで変更できます。
- 簡易配置図は図面ではなく、部屋の縦横寸法と器具の配列（または台数）から概略位置を割り付けたものです。
- 報告者氏名や写真に写り込んだ人・車両の扱いは、社内ルールに従ってください。

## スマホアプリ「LED現調」（ArcGISなし・v1.3.0〜）
`smartphone_app/` は、ArcGIS を使わずにスマホだけで現地調査するための専用アプリ（PWA）です。
- 施設→階→部屋→器具の順に入力。写真はアプリ内カメラで**連続撮影**（器具・銘板・天井などを撮りながら切り替え）
- 電波がなくても起動・入力でき、データはスマホ内に保存（入力のたびに自動保存）
- 音声メモ・撮影場所（GPS）・部屋の複製・扉の数による漏れチェック
- 「書出」で zip（data.json＋写真・音声）を作り、パソコンの MCP の `app_load_export` で取り込む。複数のスマホ・複数日の zip はまとめて取り込める
- 公開方法は `smartphone_app/公開手順.md`。画面は `smartphone_app/画面イメージ.png`

## ブラウザ（Survey123 Webアプリ）での利用
フォームは Field App と Webアプリの両方で動くように作っています（v1.2.0）。Webアプリで使えない入力欄の見た目（spinner・predictivetext）は外し、報告事項の写真は1欄1枚×3欄にしました。
Webアプリはインターネット接続が必要で、電波のない場所では入力できません。圏外になりうる現地調査は Field App、事務所での入力や追記は Webアプリ、と使い分けてください。

## 変更履歴
- 1.3.0（2026-11-26）スマホアプリ「LED現調」（PWA）を追加：連続撮影カメラ、オフライン起動・保存、音声・GPS、zip書き出し。MCPに app_load_export（複数zipの統合）を追加
- 1.2.0（2026-11-25）Webアプリ対応：Webアプリ非対応の appearance（spinner・predictivetext）を削除、報告事項の写真を3欄に分割（rp_photo・rp_photo2・rp_photo3。MCPはまとめて扱う）
- 1.1.1（2026-11-25）文字化け対策：zip内の日本語ファイル名にUTF-8の印を付与（Windowsの「すべて展開」で化けない）、マスタCSVをBOM付きUTF-8に変更（Excelで化けない）、MCPはBOMの有無どちらのCSVも読めるように変更
- 1.1.0（2026-11-25）南大東村15施設の調査仕様に対応：施設名を選択式（facilities.csv）、器具に形状・ランプ区分／仕様・銘板の消費電力・設置高さ・高所作業・撤去条件・老朽度／劣化を追加、部屋に照明状況・課題を追加、取付方法にポール・架台、投光器などの屋外器具をマスタに追加、施設一覧シートと写真整理（export_photos）を追加
- 1.0.0（2026-11-24）初版
