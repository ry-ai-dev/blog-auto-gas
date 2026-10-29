/**
 * blog-auto-gas / src/Code.gs
 *
 * シート「記事」を使ったブログ記事自動生成・投稿の土台。
 * この段階では Claude API / WordPress への実際の呼び出しは行わない(スタブ)。
 */

// ===== シート・列の定義 =====
const SHEET_NAME = '記事';

const COL = {
  KEYWORD: 1,       // A: キーワード
  SUB_KEYWORD: 2,   // B: サブキーワード
  STATUS: 3,        // C: ステータス
  TITLE: 4,         // D: タイトル
  WP_URL: 5,        // E: WordPress URL
  META_DESCRIPTION: 6, // F: メタディスクリプション
  BODY: 7,          // G: 本文
  CHAR_COUNT: 8,    // H: 文字数
  ERROR: 9          // I: エラー内容
};

const HEADER_ROW = 1;

// ===== ステータス値 =====
const STATUS = {
  NOT_GENERATED: '未生成',
  GENERATING: '生成中',
  GENERATED: '生成済み',
  POSTED: '投稿済み',
  ERROR: 'エラー'
};

// スクリプトプロパティで必須のキー名(値はここに書かない)
const REQUIRED_PROPERTY_KEYS = [
  'ANTHROPIC_API_KEY',
  'WP_URL',
  'WP_USER',
  'WP_APP_PASSWORD',
  'DRY_RUN'
];

/**
 * スプレッドシートを開いたときにメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('記事生成')
    .addItem('選択行の記事を生成', 'generateForSelectedRow')
    .addItem('設定を確認', 'checkSettings')
    .addToUi();
}

/**
 * 選択行の記事を生成する(現段階ではスタブ)。
 * ヘッダー行・空行は除外し、ステータスが「生成中」「投稿済み」の行は
 * 重複実行防止のため何もしない。
 */
function generateForSelectedRow() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      ui.alert('エラー: シート「' + SHEET_NAME + '」が見つかりません。');
      return;
    }

    const row = sheet.getActiveCell().getRow();

    if (row <= HEADER_ROW) {
      ui.alert('ヘッダー行が選択されています。記事の行を選択してください。');
      return;
    }

    const keyword = sheet.getRange(row, COL.KEYWORD).getValue();
    if (!keyword) {
      ui.alert('キーワードが空のため処理を中止しました。');
      return;
    }

    const status = sheet.getRange(row, COL.STATUS).getValue();
    if (status === STATUS.GENERATING || status === STATUS.POSTED) {
      ui.alert('ステータスが「' + status + '」のため、重複実行を防止して処理を中止しました。');
      return;
    }

    try {
      // この段階ではスタブ: ステータスを「生成中」にしてから「未生成」に戻すだけ。
      // Claude API を呼び出して本文を生成する処理は次の段階で実装する。
      updateStatus(sheet, row, STATUS.GENERATING);

      // TODO(次の段階): Claude API でキーワードから記事を生成し、
      // タイトル・メタディスクリプション・本文・文字数を各列に書き込む。

      updateStatus(sheet, row, STATUS.NOT_GENERATED);

      ui.alert('スタブ処理が完了しました(キーワード: ' + keyword + ')。本文生成は未実装です。');
    } catch (innerError) {
      updateStatus(sheet, row, STATUS.ERROR, innerError.message);
      ui.alert('エラーが発生しました: ' + innerError.message);
    }
  } catch (error) {
    ui.alert('想定外のエラーが発生しました: ' + error.message);
  }
}

/**
 * スクリプトプロパティに必要なキーが設定されているかをダイアログで表示する。
 * 値そのものは一切表示しない(存在するかどうかの○/×のみ)。
 */
function checkSettings() {
  const ui = SpreadsheetApp.getUi();

  try {
    const properties = PropertiesService.getScriptProperties();
    const lines = REQUIRED_PROPERTY_KEYS.map(function (key) {
      const value = properties.getProperty(key);
      const mark = value ? '○' : '×';
      return mark + ' ' + key;
    });

    ui.alert('スクリプトプロパティの設定状況\n\n' + lines.join('\n'));
  } catch (error) {
    ui.alert('設定確認中にエラーが発生しました: ' + error.message);
  }
}

/**
 * 指定行のステータス(・任意でエラー内容)を更新する。
 * @param {Sheet} sheet 対象シート
 * @param {number} row 更新対象の行番号
 * @param {string} status 書き込むステータス
 * @param {string=} errorMessage エラー内容(省略可)
 */
function updateStatus(sheet, row, status, errorMessage) {
  sheet.getRange(row, COL.STATUS).setValue(status);

  if (status === STATUS.ERROR) {
    sheet.getRange(row, COL.ERROR).setValue(errorMessage || '');
  } else {
    sheet.getRange(row, COL.ERROR).setValue('');
  }
}
