# **Mail Merge & PDF Generator Tool (Google Apps Script)**

A robust, web-based Mail Merge tool built with Google Apps Script (GAS). It allows you to generate multiple documents or PDFs by merging data from a Google Spreadsheet into a Google Doc template.

## **✨ Features**

* **User-Friendly Web UI:** No coding required after deployment. Operate everything from a clean browser interface.  
* **Two Output Modes:**  
  * **Individual:** Generates a separate file for each row in the spreadsheet.  
  * **Merged:** Combines all data into a single multi-page document/PDF (Ideal for mass printing).  
* **Flexible Formats:** Output as **PDF** or editable **Google Docs**.  
* **Page Settings:** Supports various paper sizes (A4, A3, B5, etc.) and orientation (Portrait/Landscape).  
* **Smart Validation:** Automatically checks for missing variables or invalid URLs before processing.

## **🚀 Setup & Installation**

1. **Create a Project:**  
   * Go to [Google Apps Script](https://script.google.com/) and create a "New Project".  
2. **Copy Code:**  
   * Create a file named Code.gs (default) and paste the content of Code.js.  
   * Create a file named index.html and paste the content of index.html.  
3. **Deploy as Web App:**  
   * Click **Deploy** \> **New deployment**.  
   * Select type: **Web app**.  
     * **Description:** (Any name, e.g., "Mail Merge Tool")  
     * **Execute as:** User accessing the web app  
     * **Who has access:** Anyone with Google account (or Only myself for private use).  
   * Click **Deploy** and open the generated **Web App URL**.

## **📖 Usage Guide**

### **1\. Prepare the Template (Google Doc)**

Create a Google Doc and use double curly braces for variables.

* **Example:** Dear {{ Name }},  
* **Note:** Variable names must exactly match the Spreadsheet headers.

### **2\. Prepare the Data (Google Sheet)**

Create a Spreadsheet with the following structure:

* **Row 1:** Headers (Variable names).  
* **Column A (Required):** ID/Index. **Processing stops if this cell is empty.**  
* **Column B (Required for Individual Mode):** Output Filename.  
* **Other Columns:** Your data variables.

**Spreadsheet Example:**

| A (ID) | B (Filename) | C (Name) | D (Date) |
| ----- | ----- | ----- | ----- |
| 1 | Invoice\_John | John Doe | 2023-01-01 |
| 2 | Invoice\_Jane | Jane Doe | 2023-01-02 |

### **3\. Run the Tool**

1. Open the Web App URL.  
2. Paste the URLs for the **Template**, **Spreadsheet**, and **Destination Folder**.  
3. Select output settings (PDF/Doc, Paper Size, etc.).  
4. Click **Start**.

# **🇯🇵 Google Apps Script 差し込み印刷・PDF生成ツール**

Googleドキュメントをテンプレート、Googleスプレッドシートをデータソースとして使用する、高機能な差し込み印刷（Mail Merge）ツールです。 直感的なWeb UIを備えており、PDFの一括生成や、1つのファイルへの結合出力に対応しています。

## **✨ 特徴**

* **使いやすいWeb UI:** ブラウザ上で操作できるため、コードを編集する必要はありません。  
* **2つの出力モード:**  
  * **個別出力:** データ1件につき1つのファイルを作成します（ファイル名指定可能）。  
  * **一括出力:** 全データを1つのファイル（PDF/ドキュメント）に結合して出力します（印刷に便利）。  
* **柔軟な出力形式:** **PDF形式**、または編集可能な**Googleドキュメント形式**を選べます。  
* **用紙設定:** A4, A3, B5などのサイズや、縦・横の向きを指定可能です。  
* **事前検証:** 実行前にURLの有効性や、変数の不一致（ヘッダー漏れ）を自動チェックします。

## **🚀 セットアップ方法**

1. **GASプロジェクトの作成:**  
   * [Google Apps Script](https://script.google.com/) にアクセスし、「新しいプロジェクト」を作成します。  
2. **コードの貼り付け:**  
   * プロジェクト内の コード.gs (デフォルト) に、配布された Code.js の中身を上書きして貼り付けます。  
   * ファイルリストの「＋」から「HTML」を選択し、index.html という名前でファイルを作成します。そこに配布された index.html の中身を貼り付けます。  
3. **ウェブアプリとしてデプロイ:**  
   * 右上の**デプロイ** \> **新しいデプロイ**をクリック。  
   * 種類の選択で「**ウェブアプリ**」を選択。  
     * **次のユーザーとして実行:** ウェブアプリケーションにアクセスしているユーザー  
     * **アクセスできるユーザー:** Googleアカウントを持つ全員（自分だけで使う場合は「自分のみ」）  
   * **デプロイ** を押し、表示されたURLにアクセスします。

## **📖 データの作り方**

### **1\. テンプレート (Googleドキュメント)**

差し込みたい箇所を {{ }} で囲んでください。

* **例:** {{ 氏名 }} 様、いつもありがとうございます。  
* **注意:** {{ }} の中の文字は、スプレッドシートの1行目と完全に一致させてください。

### **2\. データソース (Googleスプレッドシート)**

以下のルールで作成してください。

* **1行目:** ヘッダー行（テンプレートの変数名）。  
* **A列 (必須):** 「管理番号」などを入力してください。**このセルが空欄の行で読み込みが停止します。**  
* **B列 (個別出力モードで必須):** 生成されるファイルの「ファイル名」になります（例: 請求書\_{{氏名}} のように変数も使用可）。

**スプレッドシート構成例:**

| A (管理用) | B (ファイル名用) | C (変数1) | D (変数2) |
| ----- | ----- | ----- | ----- |
| 1 | 請求書\_佐藤様 | 佐藤 太郎 | 10,000 |
| 2 | 請求書\_鈴木様 | 鈴木 花子 | 25,000 |

### **3\. ツールの実行**

1. デプロイされたWebアプリのURLを開きます。  
2. テンプレート、スプレッドシート、保存先フォルダのURLをそれぞれの欄に入力します。  
3. 出力形式（PDF/Doc）、用紙サイズ、出力モードを選択します。  
4. **作成を開始する** ボタンをクリックします。

## **⚠️ 注意事項**

* 初回実行時はGoogleアカウントへのアクセス権限（ドライブ、ドキュメント、スプレッドシート）の承認が必要です。  
* 一括出力モードでデータ件数が非常に多い場合、Google Apps Scriptの実行時間制限（6分）により処理が完了しない場合があります。その場合はデータを分割して実行してください。

