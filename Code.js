/**
 * 差し込み印刷/PDF生成ツール
 * Server Side Code
 */

// ウェブアプリとしてアクセスした際にHTMLを表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('差し込み印刷ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Googleサイト埋め込み用
}

/**
 * 入力されたURL/IDから有効なIDを抽出するヘルパー関数
 */
function extractIdFromUrl(input) {
  if (!input) return null;
  // すでにIDらしき文字列（英数字とハイフン、アンダースコア）のみの場合はそのまま返す
  if (/^[a-zA-Z0-9-_]+$/.test(input)) return input;
  
  try {
    // URL形式の場合
    if (input.indexOf('drive.google.com') !== -1 || input.indexOf('docs.google.com') !== -1) {
      // フォルダの場合 folders/ID
      var folderMatch = input.match(/folders\/([a-zA-Z0-9-_]+)/);
      if (folderMatch) return folderMatch[1];
      
      // ファイルの場合 /d/ID
      var fileMatch = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (fileMatch) return fileMatch[1];
      
      // ID=パラメータの場合
      var idMatch = input.match(/id=([a-zA-Z0-9-_]+)/);
      if (idMatch) return idMatch[1];
    }
    return null; 
  } catch (e) {
    return null;
  }
}

/**
 * URLからgid（シートID）を抽出するヘルパー関数
 */
function extractGidFromUrl(url) {
  if (!url) return null;
  var match = url.match(/[#&]gid=([0-9]+)/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * ステップ1: 事前検証とデータ取得
 * URLの有効性チェック、データの取得、ヘッダーと変数の照合を行う
 * @param {string} outputMode - 'individual' (個別) または 'merged' (一括)
 */
function validateAndFetchData(templateUrl, sheetUrl, folderUrl, outputMode) {
  var result = {
    success: false,
    message: '',
    data: [],
    headers: [],
    validationErrors: []
  };

  try {
    var templateId = extractIdFromUrl(templateUrl);
    var sheetId = extractIdFromUrl(sheetUrl);
    var folderId = extractIdFromUrl(folderUrl);
    var sheetGid = extractGidFromUrl(sheetUrl); // gidを取得

    // 1. ID抽出チェック
    if (!templateId) result.validationErrors.push('テンプレートドキュメントのURLが無効です。');
    if (!sheetId) result.validationErrors.push('スプレッドシートのURLが無効です。');
    if (!folderId) result.validationErrors.push('保存先フォルダのURLが無効です。');

    if (result.validationErrors.length > 0) return result;

    // 2. アクセス権限と存在チェック
    var templateFile, folder;
    try {
      templateFile = DriveApp.getFileById(templateId);
      if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
        result.validationErrors.push('テンプレートはGoogleドキュメント形式である必要があります。');
      }
    } catch (e) {
      result.validationErrors.push('テンプレートドキュメントにアクセスできません。権限を確認してください。');
    }

    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      result.validationErrors.push('保存先フォルダにアクセスできません。権限を確認してください。');
    }

    var sheet;
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      
      // gid指定がある場合はそのシートを探す
      if (sheetGid !== null) {
        var sheets = ss.getSheets();
        for (var i = 0; i < sheets.length; i++) {
          if (sheets[i].getSheetId() === sheetGid) {
            sheet = sheets[i];
            break;
          }
        }
      }
      
      // 見つからない場合やgid指定がない場合は1シート目を対象とする
      if (!sheet) {
        sheet = ss.getSheets()[0];
      }
      
      if (sheet.getLastRow() < 2) {
        result.validationErrors.push('スプレッドシートにデータがありません（ヘッダーとデータが必要です）。');
      }
    } catch (e) {
      result.validationErrors.push('スプレッドシートにアクセスできません。権限を確認してください。');
    }

    if (result.validationErrors.length > 0) return result;

    // 3. データ取得と整合性チェック
    var values = sheet.getDataRange().getDisplayValues(); // 表示形式のまま取得（日付などを文字列として扱う）
    var headers = values[0];
    var rawRows = values.slice(1);

    // 有効データ判定と必須項目チェック
    var validRows = [];
    var emptyFileNameRows = [];

    for (var i = 0; i < rawRows.length; i++) {
      var row = rawRows[i];
      // A列(番号)が空なら、そこでデータの読み込みを終了する
      if (!row[0] || row[0] === '') {
        break;
      }
      
      // 「個別出力」モードの場合のみ、B列(ファイル名)の必須チェックを行う
      if (outputMode !== 'merged') {
        if (!row[1] || row[1] === '') {
          emptyFileNameRows.push((i + 2) + '行目');
        }
      }
      
      validRows.push(row);
    }

    // エラーがあれば通知
    if (emptyFileNameRows.length > 0) {
      result.validationErrors.push('個別出力モードでは B列(ファイル名) が必須です。以下の行を確認してください: ' + emptyFileNameRows.join(', '));
      return result;
    }
    
    if (validRows.length === 0) {
      result.validationErrors.push('有効なデータ行が見つかりません。A列(番号)を入力してください。');
      return result;
    }

    // テンプレート内の変数 {{ ... }} を抽出してチェック
    var doc = DocumentApp.openById(templateId);
    var bodyText = doc.getBody().getText();
    var templateVars = bodyText.match(/\{\{\s*.*?\s*\}\}/g) || [];
    
    // 変数の重複除外と整形
    var uniqueVars = [];
    templateVars.forEach(function(v) {
      var varName = v.replace(/\{\{|\}\}/g, '').trim();
      if (uniqueVars.indexOf(varName) === -1) uniqueVars.push(varName);
    });

    // テンプレートにある変数がスプレッドシートのヘッダーに存在するか確認
    var missingVars = [];
    uniqueVars.forEach(function(v) {
      if (headers.indexOf(v) === -1) {
        missingVars.push(v);
      }
    });

    if (missingVars.length > 0) {
      result.validationErrors.push('以下の変数がスプレッドシートのヘッダーに見つかりません: ' + missingVars.join(', '));
      return result; // ここで返すと処理を中断させる
    }

    // 成功
    result.success = true;
    result.data = validRows; // validRows を返す
    result.headers = headers;
    result.ids = { template: templateId, folder: folderId }; // IDも返却してクライアントで保持

  } catch (e) {
    result.success = false;
    result.message = '予期せぬエラーが発生しました: ' + e.toString();
  }

  return result;
}

/**
 * 共通処理: 用紙サイズを適用する
 */
function applyPageSize(body, config) {
  // 用紙サイズの設定 (Points単位: 1 inch = 72 points)
  var sizes = {
    'A4': [595.276, 841.89],
    'A3': [841.89, 1190.55],
    'A5': [419.528, 595.276],
    'B4': [728.5, 1031.8],
    'B5': [515.9, 728.5]
  };
  
  var size = sizes[config.paperSize] || sizes['A4'];
  var width = size[0];
  var height = size[1];

  if (config.orientation === 'landscape') {
    // 横向きの場合入れ替え
    var temp = width;
    width = height;
    height = temp;
  }

  // ページサイズ適用
  body.setPageWidth(width);
  body.setPageHeight(height);
}

/**
 * 変数置換処理（共通）
 */
function replaceVariables(body, rowData, headers) {
  headers.forEach(function(header, index) {
    var value = rowData[index] || '';
    // {{ ヘッダー名 }} のパターン（スペース許容）を作成
    var escapedHeader = header.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // 正規表現: {{ (任意の空白) ヘッダー (任意の空白) }}
    body.replaceText('\\{\\{\\s*' + escapedHeader + '\\s*\\}\\}', value);
  });
}

/**
 * PDFとしてエクスポートするヘルパー関数
 * DriveApp.getAs()ではなく、GoogleドキュメントのエクスポートURLを叩く
 * 改修: タブIDを指定して、特定のタブのみをPDF化する
 */
function exportAsPdf(fileId) {
  // ドキュメントを開く
  var doc = DocumentApp.openById(fileId);
  
  // 最初のタブ（または特定のタブ）を取得
  // これにより、余計な表紙などを除外し、メインのタブだけをPDF化できる
  var tabs = doc.getTabs();
  var targetTab = tabs[0]; // 1つ目のタブを指定
  var tabId = targetTab.getId();

  // Googleドキュメントの標準エクスポートURL（タブ指定を追加）
  var url = "https://docs.google.com/document/d/" + fileId + "/export?format=pdf&tab=" + tabId;
  
  var options = {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error("PDF変換に失敗しました: " + response.getContentText());
  }
  
  return response.getBlob();
}


/**
 * ステップ2: 個別ファイルの生成処理 (1件ずつ呼び出される想定)
 */
function generateSingleDocument(templateId, folderId, rowData, headers, config) {
  var result = {
    success: false,
    fileName: '',
    url: '',
    error: ''
  };

  try {
    var templateFile = DriveApp.getFileById(templateId);
    var folder = DriveApp.getFolderById(folderId);
    
    // ファイル名の決定ロジック
    var fileNameTemplate = rowData[1] || 'NoName'; 
    headers.forEach(function(header, index) {
      var value = rowData[index] || '';
      var escapedHeader = header.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var regex = new RegExp('\\{\\{\\s*' + escapedHeader + '\\s*\\}\\}', 'g');
      fileNameTemplate = fileNameTemplate.replace(regex, value);
    });
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
    var newFileName = fileNameTemplate + '_' + timestamp;
    
    var newFile = templateFile.makeCopy(newFileName, folder);
    var doc = DocumentApp.openById(newFile.getId());
    var body = doc.getBody();

    applyPageSize(body, config); // サイズ適用
    replaceVariables(body, rowData, headers); // 置換実行

    doc.saveAndClose();

    if (config.outputType === 'pdf') {
      // 変更点: exportAsPdf関数を使用（内部でタブ指定ロジックを実行）
      var pdfBlob = exportAsPdf(newFile.getId());
      pdfBlob.setName(newFileName + '.pdf');
      var pdfFile = folder.createFile(pdfBlob);
      newFile.setTrashed(true);
      result.fileName = pdfFile.getName();
      result.url = pdfFile.getUrl();
    } else {
      result.fileName = newFile.getName();
      result.url = newFile.getUrl();
    }
    result.success = true;

  } catch (e) {
    result.success = false;
    result.error = e.toString();
  }

  return result;
}

/**
 * 一括結合ファイルの生成処理
 */
function generateMergedDocument(templateId, folderId, allRowData, headers, config) {
  var result = {
    success: false,
    fileName: '',
    url: '',
    error: ''
  };

  try {
    var templateFile = DriveApp.getFileById(templateId);
    var folder = DriveApp.getFolderById(folderId);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
    
    // 結合ファイル名決定
    var mergedFileName = "一括出力_" + timestamp;
    
    // 1. ベースとなるファイル作成（1件目の処理としてテンプレートをコピー）
    var mergedFile = templateFile.makeCopy(mergedFileName, folder);
    var mergedDoc = DocumentApp.openById(mergedFile.getId());
    var mergedBody = mergedDoc.getBody();
    
    // 用紙サイズ設定
    applyPageSize(mergedBody, config);
    
    // 2. 1件目のデータ置換
    if (allRowData.length > 0) {
      replaceVariables(mergedBody, allRowData[0], headers);
    }
    
    // 3. 2件目以降をループ処理で追記
    for (var i = 1; i < allRowData.length; i++) {
      var rowData = allRowData[i];
      
      // 改ページを挿入
      mergedBody.appendPageBreak();
      
      // 一時ファイルを作成して置換（置換スコープを限定するため）
      var tempFile = templateFile.makeCopy('temp_' + i);
      var tempDoc = DocumentApp.openById(tempFile.getId());
      var tempBody = tempDoc.getBody();
      
      // 一時ファイル内で置換
      replaceVariables(tempBody, rowData, headers);
      tempDoc.saveAndClose();
      
      // 再度開いて中身をコピー
      tempDoc = DocumentApp.openById(tempFile.getId());
      tempBody = tempDoc.getBody();
      
      // 要素を順番にコピーして結合ファイルに追加
      var numChildren = tempBody.getNumChildren();
      for (var j = 0; j < numChildren; j++) {
        var element = tempBody.getChild(j).copy();
        var type = element.getType();
        
        if (type == DocumentApp.ElementType.PARAGRAPH) {
          mergedBody.appendParagraph(element);
        } else if (type == DocumentApp.ElementType.TABLE) {
          mergedBody.appendTable(element);
        } else if (type == DocumentApp.ElementType.LIST_ITEM) {
          mergedBody.appendListItem(element);
        } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
          mergedBody.appendImage(element);
        }
      }
      
      // 一時ファイルは削除
      tempFile.setTrashed(true);
    }
    
    mergedDoc.saveAndClose();

    // 出力形式分岐
    if (config.outputType === 'pdf') {
      // 変更点: exportAsPdf関数を使用（内部でタブ指定ロジックを実行）
      var pdfBlob = exportAsPdf(mergedFile.getId());
      pdfBlob.setName(mergedFileName + '.pdf');
      var pdfFile = folder.createFile(pdfBlob);
      
      mergedFile.setTrashed(true);
      
      result.fileName = pdfFile.getName();
      result.url = pdfFile.getUrl();
    } else {
      result.fileName = mergedFile.getName();
      result.url = mergedFile.getUrl();
    }
    
    result.success = true;

  } catch (e) {
    result.success = false;
    result.error = e.toString();
  }

  return result;
}
