# SSBOM Manager 使用手順書 (v1.0.7)

本書は、機種コード **110C3M3AK0**（モデル Virgo）について実際に操作した記録に基づき、ソフトウェアの機能を説明します。スクリーンショットはその操作中にソフトウェアから直接取得したものです。v1.0.6 で撮影した画像はステータスバーに v1.0.6 と表示されていますが、該当手順の操作と画面は v1.0.7 でも変わりません。

画面上の表示（ボタン名など）はベトナム語です。本書ではボタン名を画面のとおりに引用し、説明を付けます。

HTML 版では、ページ上部の言語切替ボタンと**左側のメニュー**で各項目へ移動できます。

## 1. インストールとアップデート

### 1.1 初回インストール

次のインストーラーを実行します。

```
\\fstvn01.kdtvn.local\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\SSBOM_Manager_Setup_1.0.8.exe
```

インストール後にソフトを起動します。バージョンは画面右下に表示されます（v1.0.8）。

### 1.2 LAN 経由のアップデート

LAN 上の `release_update` フォルダに新しいバージョンがあると、"Có Phiên Bản Mới Sẵn Sàng!"（新バージョンあり）ダイアログが表示され、現在のバージョン、新バージョン、パッケージ容量、変更内容が確認できます。"Tải & Cập Nhật Ngay"（今すぐ更新）で更新、"Để sau"（後で）で見送ります。

![アップデートダイアログ](images/g27_software_update.png)

この画像は、実際の 1.0.7 パッケージ情報でダイアログを開いて撮影したものです（説明のため現在バージョンを v1.0.6 としています）。自動ダウンロードとインストールは、本書作成時には再実行していません。

## 2. 例で使ったデータ

サンプルファイル `C3MAKO - Copy.zip` を使用しました。

| ファイル | 内容 |
|---|---|
| `PLM_110C3M3AK0.xlsx` | PLM (Teamcenter) から出力した BOM、1890行 |
| `BOM_110C3M3AK0.xlsm` | 作成済みのサンプル BOM。CTTT シート（2198行）と R3 シートを入力データとして使用 |

注意事項:

- 本書作成時は Teamcenter と SAP に接続していないため、R3 ファイルはサンプル BOM の R3 シートから書き出したもので、SAP から直接出力したファイルではありません。
- 担当者9名の CTTT データは、サンプル BOM の CTTT シートから取得しました。サンプル内の旧アカウント名は、現在の人員リストのアカウントに対応付けています（例: `AnQuyDuong` → `Duong_mecha1.2`）。この対応付けは作成者の仮定であり、試験用です。
- 作業フォルダは `sample_data\run` と `sample_data\run_feat`（試験用コピーで、実データではありません）。
- 設定ダイアログの画像は、実際のアカウントとパスワードをプレースホルダーに置き換えています。

## 3. 画面全体

メインウィンドウには上部のツールバーと2つの画面「Trưởng Nhóm」（リーダー）と「Thành Viên Công Đoạn」（工程担当）があります。

ツールバーのボタン: "Tải BOM (PLM/SAP)"、"Quản lý Bộ lọc BOM PLM"、"Áp dụng bộ lọc cho file có sẵn"、"Cập nhật BOM mới (PLM & R3)"、"Quản lý Trưởng nhóm"、"Nhập liệu Thành viên"、"Giao diện: Sáng [Ctrl+T]"。

### 3.1 ライト／ダークテーマ

ツールバーの "Giao diện" ボタンまたは `Ctrl+T` で切り替えます。"Cấu hình hệ thống" の "Giao diện (UI Theme)" タブでも選べます。再起動は不要です。

![ダークテーマ: プロジェクト作成画面](images/g25_dark_step1.png)

![ダークテーマ: 統合 BOM 比較画面](images/g26_dark_step4.png)

### 3.2 システム設定

メニュー "Hệ thống" → "Cấu hình hệ thống"（ショートカット `Ctrl+,`）で開きます。5つのタブがあります。

**Teamcenter TC24:** ポータル URL、アカウント、パスワード、ブラウザ、ヘッドレス実行、待機時間。"Kiểm tra kết nối TC24"（接続確認）ボタン付き。

![TC24 設定](images/g20_settings_tc24.png)

**SAP R3:** System Description、SAP Client、SAP アカウントとパスワード、言語、Plant、BOM Usage、Alternative、`saplogon.exe` のパス。"Kiểm tra kết nối SAP GUI" ボタン付き。

![SAP R3 設定](images/g21_settings_sap.png)

**Đường dẫn Dữ liệu（パス）:** プロジェクトのルートフォルダ、Master Fix Serial Tool ファイル、レポート出力先、共有人員 DB (SQLite)、LAN アップデートフォルダ。

![パス設定](images/g22_settings_paths.png)

**Giao diện (UI Theme):** ライト／ダークの選択。

![テーマ設定](images/g23_settings_theme.png)

**Email Mẫu Form:** "Chế độ Thử nghiệm"（テストモード）、テスト受信者、既定の送信者、LAN 上のテンプレートファイル（機種辞書、`formnguoidung.xlsm`、`form_ssbom.xlsm`）のパス。

![メール設定](images/g24_settings_email.png)

テストモードが有効な間は、すべてのメールが実際の宛先ではなくテスト担当者に転送されます（4.5、4.17、6 を参照）。本番運用では無効にしてください。

## 4. リーダーの手順

「Trưởng Nhóm」画面には4つのタブがあります: "1. Lập Dự Án & Phân Công"（プロジェクト作成と割当）、"2. Tải & Xử Lý Nguồn BOM"（BOM ソースの取得と処理）、"3. Theo dõi tổng hợp file của phụ trách"（担当者ファイルの進捗管理）、"4. So Sánh BOM Tổng & Báo Cáo"（統合 BOM の比較とレポート）。

### 4.1 ステップ1. プロジェクト作成と割当

1. "Model máy" で `Virgo` を選びます。機種コード `110C3M3AK0` は "1.2 Danh sách Mã Máy" の表にあります。ルートフォルダは `sample_data\run` にしました。
2. "1.3 Phân Công Nhân Sự Phụ Trách" の表で担当者9名を選び、工程（DRUM, LSU, DLP, IMAGE, FUSER, DP, ISU, HONTAI, FRAME UNIT）を割り当てます。
3. "Khởi Tạo Thư Mục && Phân Công" を押します。

![ステップ1: プロジェクト設定](images/e01_step1_setup.png)

次のメッセージが表示されます。

> Đã tạo xong danh sách BOM và các gói file CTTT của phụ trách được phân lập theo từng dòng máy: Model 'Virgo': 1 mã máy
> Cấu trúc: <Thư mục gốc>\<Tên Model>\<Mã máy>\<Tên phụ trách>.xlsm

（意味: 機種ごとに担当者の CTTT ファイル一式を作成しました。構成は <ルート>\<モデル名>\<機種コード>\<担当者名>.xlsm）

![ステップ1: 作成後](images/e02_step1_created.png)

`Virgo\110C3M3AK0\` に `<アカウント>.xlsm` が9個作成されます（Chien_mecha1.2, Duong_mecha1.2, Hung_mecha1.1, Ly_mecha1.2, Nhung_mecha1.2, Thai_mecha1.2, Thuong_mecha1.2, Tu_mecha1.2, VietE_mecha1.2）。これらは `formnguoidung.xlsm` のコピーです。

### 4.2 モデルと機種コードの管理

1.1 の "Quản lý Model & Mã máy..." を押します。ダイアログは3つの部分で構成されます: モデルと生産フェーズの選択、モデルに登録済みの4文字機種コード一覧（例: Virgo は10コード）、部署ごとの担当メールグループ（To、CC1、CC2…）。ボタン: "Tạo Model Mới"（新規モデル）、"Tải lại Excel"、"+ Bổ sung mã vào Model"、"Lưu Nhóm Mail Vào Từ Điển"、"Đưa Danh Sách Mã Này Vào Bảng 1.2"、"Chọn Model Này"。

![モデル・機種コード管理](images/g16_model_manager.png)

### 4.3 人員管理

1.3 の "Quản lý nhân sự..." を押します。"Danh mục Nhân sự Kỹ thuật Sản xuất"（生産技術人員名簿）に、メンバー一覧（例では47名）が表示されます。名前またはアカウントでの検索、部署フィルター、"Chỉ đang hoạt động"（稼働中のみ）フィルターがあります。右側の詳細フォームには "+ Thêm thành viên mới"、"Cập nhật thông tin"、"Xóa thành viên"、"Làm mới ô nhập" のボタンがあります。下部には "Đồng bộ & Nạp lại từ CSDL"、"Nạp lại từ Excel (Refresh)"、"Ghi sang Danhsachthanhvien.xlsm"、"Mở File Excel" があります。共有データはネットワーク上の DB `ssbom_master.db` にあります。

![人員管理](images/g15_roster.png)

### 4.4 PCD 計画のスキャン

1.1 の "Quét Kế Hoạch PCD" を押します。ダイアログは3つの部分で構成されます: 月次生産計画の情報（年、月、計画 Excel ファイル。"Tải Tự Động Từ SharePoint"、"Mở SharePoint PCD..."、"Quét Dữ Liệu" ボタン）、見つかった新規機種コード一覧、フォルダ作成とメール通知の設定。最下部の "Khởi Tạo Thư Mục & Sinh Gói File Làm Việc Cho Mã Đã Chọn" で、選択したコードのフォルダを作成します。

![PCD 計画スキャン](images/g17_pcd_scan.png)

本書作成時は実際の計画ファイルがなかったため、この機能は実行していません。画像はデータなしの状態です。

### 4.5 担当者への割当メール

1.1 の "Gửi Mail Phụ Trách"を押します。メールのプレビュー（宛先、CC、件名、HTML 本文）が開き、"Mở trong Microsoft Outlook" と "Gửi trực tiếp qua Outlook" のボタンがあります。

![割当メール](images/g18_mail_assign.png)

画像ではテストモードが有効で、メールはテスト担当者に転送されます。本文冒頭に、本番時の実際の宛先が表示されます。本書作成時に実際のメール送信はしていません。

### 4.6 ステップ2. BOM ソースの取得と処理

"2.2 Bảng Trạng Thái Nguồn Dữ Liệu" の表は、最初は PLM TC24 列と SAP R3 CS12 列の両方が "Thiếu file"（ファイルなし）です。表の下にボタンがあります: "Tải BOM Tự Động (PLM TC24 / SAP R3)"、"Nạp Tệp Sẵn Có Từ Ổ Đĩa"、"Lọc BOM TC24 Level 1..6 & Sao Lưu"、"Cấu Hình Bộ Lọc BOM"、"Tạo Lọc Trực Quan"、"Lọc Lại File Có Sẵn..."、"Làm mới trạng thái"。

![ステップ2: ファイル未読込](images/e03_step2_before.png)

ソースファイルの取り込み方法は2つあります: PLM/SAP から自動ダウンロード（4.7）、またはディスク上のファイルを読み込み（4.8）。

### 4.7 PLM (Teamcenter) と SAP R3 からの自動ダウンロード

ステップ2の "Tải BOM Tự Động (PLM TC24 / SAP R3)"、またはツールバーの "Tải BOM (PLM/SAP)"（`Ctrl+D`）を押します。ダイアログ上部に3ステップの案内があります:

1. 機種コード／BOM コードの一覧を下の欄に貼り付ける。
2. SAP R3 から取得する場合は有効日を選ぶ（Teamcenter TC24 は常に最新の BOM を取得するため、日付は不要）。
3. 最下部のバーのボタンでダウンロードする。

![ダウンロードダイアログ: Teamcenter タブ](images/g01_plm_download_tc24.png)

構成要素:

- "Dòng máy / Model": 既定は "Tự động nhận diện theo mã BOM"（BOM コードから自動判定）。
- チェックボックス "Tự động áp dụng Bộ lọc BOM (BolocBom) theo từng loại máy ngay sau khi tải"（ダウンロード直後に機種ごとの BOM フィルターを自動適用）。
- "Quét Nạp mã..."、"Dán từ Clipboard"、"Xóa hết": コード一覧の入力。機種フォルダにコードが既にある場合は "Tìm thấy 1 mã BOM trong thư mục Virgo" と表示され、"Nạp tất cả (1)" と "Chi tiết..." ボタンが出ます。
- タブ "Siemens Teamcenter (TC24)": TC24 のアカウントと、PLM の "Tự động chuẩn hóa 14 cột tiêu chuẩn"（14標準列への自動正規化）。
- タブ "SAP R3": "Mã nhà máy (Plant)"、"BOM Usage"、"Alternative (Alt)"、"Ngày hiệu lực"（有効日）。
- 下部バー: "Tải đồng thời cả PLM & SAP R3"（両方同時）、"Chỉ tải BOM PLM (TC24)"、"Chỉ tải BOM SAP R3"、"Đóng"。

![ダウンロードダイアログ: SAP R3 タブ](images/g02_plm_download_sap.png)

アカウント、パスワード、接続パラメータは "Cấu hình hệ thống"（3.2）から取得されます。事前に入力し、"Kiểm tra kết nối" を押してください。

**未実施:** Teamcenter と SAP からの実際のダウンロードにはアカウントとネットワーク接続が必要なため、本書作成時はダイアログを開いて画像を撮影しただけです。上記の説明は、ダウンロード成功の記録ではなく、実際の画面構成に基づいています。

**フォルダまたはプロジェクトからの機種コードスキャン:** "Quét Nạp mã..." でスキャンダイアログが開きます。フォルダまたは現在のプロジェクトから機種コードを検出し、コードごとに PLM ファイルと R3 ファイルの有無を表示します。クイックフィルター "Tất cả"、"Chỉ thiếu R3"、"Chỉ thiếu PLM"、"Đã đủ cả 2" と、"Lọc lại file có sẵn..."、"Nạp vào danh sách tải" ボタンがあります。

![機種コード／BOM スキャン](images/g03_bom_scan.png)

### 4.8 ディスク上の既存 BOM ファイルの読み込み

1. "Nạp Tệp Sẵn Có Từ Ổ Đĩa" を押し、PLM ファイルと R3 ファイルを選びます。
2. "Đã nạp và phân tuyến 2 tệp vào các thư mục máy tương ứng."（2ファイルを各機種フォルダに読み込みました）と表示され、両列が "Sẵn sàng"（準備完了）になります（PLM 123 KB、R3 175 KB）。

![ステップ2: 読込後](images/e04_step2_imported.png)

3. "Lọc BOM TC24 Level 1..6 & Sao Lưu" を押します。"Đã hoàn thành lọc BOM Level 1..6 và sao lưu 1 tệp vào backupTC14full/."（Level 1..6 のフィルタが完了し、1ファイルを backupTC14full/ にバックアップしました）と表示されます。

今回、PLM ファイルはフィルタ前が1890行、フィルタ後が1878行でした。

![ステップ2: フィルタ後](images/e05_step2_filtered.png)

### 4.9 PLM BOM フィルターの設定（BolocBom）

フィルターは、機種ごとに PLM BOM を展開する際に除外する子アセンブリ／部品のリストです。ツールバーの "Quản lý Bộ lọc BOM PLM"、またはステップ2の "Cấu Hình Bộ Lọc BOM" を押します。

ダイアログの構成: 左側に機種一覧（ルール数付き。例: Virgo 77 ルール）、右側にルール表（部品／アセンブリ名、照合モード "Chính xác (Full)"（完全一致）または "Chứa từ (Part)"（部分一致）、部品コード、備考、編集・削除ボタン）、"Thêm Mới Quy Tắc Lọc"（ルール追加）フォーム。データは LAN 上の共有 DB に集中保存されます（右上の "LAN Server (UNC)" 表示）。

下部のボタン: "Áp Dụng Cho File Có Sẵn..."、"Tạo Lọc Trực Quan Từ Cây BOM..."、"Nhập từ Excel (Sheet BolocBom)"、"Xuất ra Excel"、"Khôi Phục Mặc Định"、"Đóng"。

![BOM フィルター設定](images/g04_filter_config.png)

### 4.10 BOM ツリーからのビジュアルフィルター作成

ステップ2の "Tạo Lọc Trực Quan" を押します。PLM Full ファイルを読み込み（"Nạp Tệp BOM PLM Full"）、階層（Level）ごとに BOM ツリーを表示し、"Cấp 1-2"、"Mở rộng tất cả"、"Thu gọn tất cả" で展開・折りたたみ、検索し、除外したい Phantom やアセンブリを直接選べます。"Quy Tắc Lọc Áp Dụng" 列に、各行に適用されているルールが表示されます。最下部のバーに部品総数、フィルター後の部品数、ルールが適用された集合数が表示され、"Lưu Vào Bộ Lọc Của Model" で選択をルールとして保存します。

![ビジュアルフィルター](images/g05_visual_builder.png)

例では、`110C3M3AK0` の PLM ツリーは 1,889 部品で、Virgo の 77 ルールにより 21 集合にルールが適用され、ダイアログ表示上は 1,879 部品が残ります。

### 4.11 既存ファイルへのフィルター適用

ディスク上に PLM ファイルがあり、Teamcenter から再取得せずにフィルターだけをやり直したい場合に使います。ツールバーの "Áp dụng bộ lọc cho file có sẵn"、またはステップ2の "Lọc Lại File Có Sẵn..." を押します。

1. ソースを選びます: フォルダ全体をスキャン、または Excel ファイルを1つ以上選択。サブフォルダのスキャンと、"PLM_" で始まるファイルのみの認識のオプションがあります。
2. 機種を選びます（ファイル名／フォルダ名からの自動判定も可）。適用中のルール数が表示されます（例: "Có 77 quy tắc đang áp dụng cho [Virgo]"）。
3. 書き込み方法を選びます: "Tự động tạo bản sao lưu (.bak) trước khi cắt tỉa"（実行前にバックアップ）、"Ghi đè trực tiếp lên tệp gốc"（元ファイルを上書き）または "Lưu thành tệp mới (*_filtered.xlsx)"（新規ファイルとして保存）。
4. "Bắt Đầu Lọc BOM" を押します。

![既存ファイルへのフィルター: 実行前](images/g06_standalone_before.png)

ダイアログには注意書きがあります: 追加で削除できるのはファイルに残っている行のみで、既にファイルから削除された行は自動では復元されません。復元にはバックアップを使うか、Teamcenter から再取得します。

例の PLM ファイルのコピーでの実行結果:

> Hoàn tất (77 quy tắc): Ban đầu 1,890 dòng -> Còn lại 1,878 dòng (Đã cắt 12 dòng linh kiện con).
> Kết thúc: Đã xử lý 1/1 tệp (1 thành công, 0 lỗi). Tổng số dòng đã cắt tỉa: 12 dòng.

（意味: 完了（77ルール）: 元 1,890 行 → 残り 1,878 行（子部品 12 行を削除）。1/1 ファイル処理（成功1、エラー0）。削除行数合計 12 行。）

上書き前にバックアップが作成されます（`PLM_110C3M3AK0_<日時>.xlsx`、`backup_before_filter` フォルダ内）。

![既存ファイルへのフィルター: 実行後](images/g07_standalone_after.png)

### 4.12 ステップ3. 担当者の入力（「Thành Viên Công Đoạn」画面）

この作業は各エンジニアが自分の PC で行います。例では1台の PC で各アカウントを順に操作しました。

1. 自分の割当ファイルを開きます（"Mở file phân công..."）。担当者名、機種コード、部署、状態 "CHƯA HOÀN THÀNH (Q2 TRỐNG)"（未完了、Q2 が空）が表示されます。
2. タブ "1. Linh kiện CTTT (A:N)" に部品リストを入力します（"Thêm dòng" または "Nhập từ Excel/CSV..."）。

![ステップ3: 部品入力](images/e07_member_entered.png)

3. "Kiểm tra Sơ bộ (Self-Check PLM & R3)" を押し、PLM・R3 と数量を照合します。

![ステップ3: Self-Check 後](images/e08_member_selfcheck.png)

4. "Xác nhận Hoàn Thành (đóng dấu Q2 = OK)" を押します。差異があって "Giải thích"（説明）が未入力の部品がある場合は、続行するか確認されます。続行すると次のメッセージが表示されます。

> Đã đóng dấu Q2 = OK và lưu file thành công!

（Q2 = OK を記録し、ファイルを保存しました）

![ステップ3: 完了確認後](images/e09_member_submitted.png)

例での担当者9名の部品数: Duong 312、Ly 481、Nhung 356、Hung 87、VietE 109、Tu 210、Thai 281、Chien 307、Thuong 55（合計 2198）。

### 4.13 進捗確認と統合（タブ3）

タブ3に担当者ごとのファイル状態が一覧表示されます。"Quét trạng thái ngay" で更新します。

9ファイルすべてが Q2 = OK になると、ステータスバーに "Hoàn tất 100% (9/9 file đạt OK). Sẵn sàng tổng hợp dữ liệu." と表示され、"TỔNG HỢP DỮ LIỆU THÀNH VIÊN" ボタンが押せるようになります。

![ステップ3（リーダー）: 9/9 OK](images/e11_step3_all_ok.png)

"TỔNG HỢP DỮ LIỆU THÀNH VIÊN" を押します。例での結果は2198行です。v1.0.7 以降は、Q2 = OK のファイルだけが提出済みと判定されます。テンプレートから作成しただけで部品が未入力のファイルは、"Chưa hoàn thành" のままです。"Số LK" 列は実際の部品数（部品コードの数）です。

### 4.14 ステップ4. 統合 BOM ファイルの作成（タブ4）

タブ4の "4.1 Khởi Tạo File So Sánh BOM Tổng"、機種コード `110C3M3AK0`。作成前のステータスは "Chưa tạo file BOM tổng hợp."（統合 BOM 未作成）です。

![ステップ4: 作成前](images/e13_step4_before.png)

"Tạo File BOM Tổng && Refresh Pivot Tables" を押すと、ステータスが "Đã tạo xong file: BOM_110C3M3AK0.xlsm (168 KB)" に変わります。

![ステップ4: 作成後](images/e14_step4_generated.png)

ファイルは `Virgo\110C3M3AK0\BOM_110C3M3AK0.xlsm` に保存されます。"Mở File BOM Tổng" と "Mở Thư Mục Máy" で、このファイルとフォルダを開けます。

この 168 KB のファイルは Excel (COM) 連携を無効にして作成したもので、Excel で開いて内容を確認していません。4.15 の BOM 更新では、マクロ付きの完全なサンプル BOM（3.2 MB）を使い、実際の Excel で実行しました。

### 4.15 BOM への新しい PLM・R3 の更新

新しい PLM または R3 ファイルがあり、説明列を維持したまま既存の BOM ファイルに反映したい場合に使います。ボタンは2か所にあります:

- ツールバーの "Cập nhật BOM mới (PLM & R3)": 専用ダイアログを開きます（下記）。
- タブ4の "Cập nhật BOM mới" 欄の3つのボタン "Cập nhật PLM vào BOM"、"Cập nhật R3 vào BOM"、"Cập nhật cả PLM & R3": ダイアログと同じ処理を使います。本書作成時、これらのボタンは個別には押していません。

**ダイアログ "Cập Nhật Dữ Liệu BOM Mới (PLM & SAP R3)" の手順:**

1. "Thư mục máy"（機種フォルダ）を選びます（"Duyệt thư mục..."）。フォルダ内の比較 BOM ファイル、新しい `PLM*` ファイル、新しい `R3*` ファイルが検出されます。ファイルを追加したら "Quét lại" を押します。"Chọn file..." で個別にファイルを指定することもできます。
2. 元ファイルを保管したい場合は "Tự động chuyển file PLM/R3 cũ vào thư mục 'capnhat\old\'" にチェックを入れたままにします。
3. "7. Cập nhật PLM mới"、"8. Cập nhật R3 mới"、または "Cập nhật cả PLM & R3" を押します。

![BOM 更新ダイアログ: フォルダスキャン後](images/g08_update_dialog_scanned.png)

ダイアログには、従来の VBA 版のボタン 7 とボタン 8 に相当すること、および O 列（Giải thích＝説明）、P 列（Người phụ trách＝担当者）、Q 列（Quản lý check＝管理者確認）の3列を引き継ぎ、ピボットテーブルを再同期することが記載されています。

サンプルファイルでの実行（実際の Excel を使用、PLM と R3 の両方）には約 52〜65 秒かかりました:

![BOM 更新ダイアログ: 完了後](images/g09_update_dialog_done.png)

ソフトウェアの完了メッセージ:

![更新完了メッセージ](images/g10_msg_update_done.png)

実行後に確認した内容:

| 項目 | 更新前 | 更新後 |
|---|---|---|
| BOM ファイルサイズ | 3,245,301 バイト | 3,974,239 バイト |
| シート数 | 9 | 11（更新前データの控えとして `PLM_old_2` と `R3_old` が追加） |
| VBA マクロ (`vbaProject.bin`) | あり | あり |

- `capnhat\old\` に BOM のバックアップ2つ（`BOM_110C3M3AK0_backup_<日時>.xlsm`）と、移動された元の PLM・R3 ファイルがあります。
- 更新後（および 4.16 のフィルター後）の BOM ファイルを実際の Excel で開き直して確認しました: 開けること、11 シート、VBA プロジェクトあり。
- "48 linh kiện kế thừa giải trình"（説明を引き継いだ部品 48 件）は、この実行でソフトウェアが表示した数値です。

例についての注意: CTTT シートの R3 比較列は、元のサンプル BOM でも更新後でも `#N/A` と表示されます。ここで使った R3 ファイルは SAP から直接出力したものではないため、この列については結論を出せません。SAP から実際に出力した R3 ファイルでの試験は行っていません。

### 4.16 不整合データの抽出（NG / #N/A）

更新後、タブ4の "Lọc Bất Đồng Bộ (NG/#N/A)" を押すと、NG または #N/A の行だけを残し、OK の行を非表示にします。選択中の機種の BOM ファイルに対して動作します。

例での結果（4.15 の更新後）:

![不整合抽出メッセージ](images/g11_msg_filter_done.png)

フィルター後の BOM ファイルサイズ: 3,398,897 バイト。

更新とフィルター後の BOM ファイルの各シート（実際の Excel で読み取り専用で開いて撮影）:

CTTT シート（"Compare" 列は PLM との比較で OK／NG を示します。R3 と REV の列は 4.15 のとおり `#N/A`／`#REF!` です）:

![CTTT シート](images/g12_excel_cttt.png)

PLM シート（更新後の PLM BOM。"Giải thích"（説明）と "Phụ trách"（担当者）列は保持されています）:

![PLM シート](images/g13_excel_plm.png)

CTTT_Total シート（部品コードごとの数量合計）:

![CTTT_Total シート](images/g14_excel_cttt_total.png)

### 4.17 管理者への BOM 確認依頼メール

タブ4の "Gửi Mail Nhờ Quản Lý Check BOM..." を押します。管理者宛メールのプレビューが開き、統合 BOM が添付ファイルとして付き、"Mở trong Microsoft Outlook" と "Gửi trực tiếp qua Outlook" のボタンがあります。"Xem trước Email" ボタンでは、"4.3 Hệ Thống Gửi Email Outlook 2 Tầng"（2段階 Outlook メール）の2つのフロー（フロー1: メンバーへの18項目確認リマインド、フロー2: 管理者への確認報告）の内容を確認できます。

![管理者宛メール](images/g19_mail_manager.png)

画像ではテストモードが有効です。本書作成時に実際のメール送信はしていません。

### 4.18 JIG と 4M 評価（4.2 欄）

タブ4の "4.2 Quản lý Danh Mục Master List JIG & Đánh Giá 4M (KTSX)" には、JIG 機種の選択、"Nạp JIG từ Master" ボタン、"Có phát sinh thay đổi 4M (Man, Machine, Material, Method)" チェック、4M 評価（Đạt OK／Không đạt NG／Chờ xác nhận）、KTSX 確認技術者欄と評価日があります（4.14 の画像を参照）。本書作成時、この機能は実行していません。

## 5. 作成時に実施した項目と未実施の項目

| 機能 | 状況 |
|---|---|
| プロジェクト作成、フォルダ作成、割当（ステップ1） | 実際に実行 |
| ディスクからのファイル読込、BOM Level 1..6 フィルター（ステップ2） | 実際に実行 |
| メンバー入力、Self-Check、Q2 = OK の確定（ステップ3） | 9アカウントで実際に実行 |
| メンバーデータの統合（ステップ3） | 実際に実行（2198行） |
| 統合 BOM ファイルの作成（ステップ4） | 実行（Excel COM 無効）。ファイルは開いて未確認 |
| BOM への PLM・R3 更新（実際の Excel） | 実際に実行し、Excel で開き直して確認 |
| 不整合データの抽出（NG/#N/A） | 実際に実行 |
| 既存ファイルへのフィルター適用 | コピーに対して実際に実行 |
| フィルター設定、ビジュアルフィルター、機種コードスキャン、モデル／人員管理、システム設定、テーマ切替 | 開いて撮影のみ。データ変更なし |
| Teamcenter・SAP からの自動ダウンロード | 未実施（アカウントとネットワークが必要） |
| 実際の計画ファイルでの PCD スキャン | 未実施 |
| Outlook での実際のメール送信 | 未実施 |
| 4.2 欄 JIG／4M | 未実施 |
| ソフトウェアの自動アップデート（ダウンロードとインストール） | 再実行せず。ダイアログは撮影用に構成 |

## 6. 備考

- メンバー画面: Self-Check 後、PLM と R3 の数量が一致している行でも "Kết Quả So Sánh" 列が NG になることがあります。原因は未確認です。発生した場合はソフトの担当者へ連絡してください。
- ダークテーマ: リーダー画面で未選択のステップのタブが、明るい背景に灰色の文字で表示されます。読めますが、ダークテーマの他の部分とは見た目が異なります。
- テストモードが有効な間は、送信メールがテスト担当者に転送されます。本番運用前に "Cấu hình hệ thống" → "Email Mẫu Form" タブで確認してください。
