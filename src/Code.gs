/**
 * @OnlyCurrentDoc
 *
 * blog-auto-gas / src/Code.gs
 *
 * シート「記事」を使ったブログ記事自動生成・投稿の土台。
 * 段階2: Claude API (Messages API) を呼び出して本文を生成する処理を実装。
 * WordPress への投稿はまだ実装していない。
 */

// ===== シート・列の定義 =====
const SHEET_NAME = '記事';
const SAMPLE_SHEET_NAME = 'サンプル記事';

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

// ===== Claude API 関連の設定(値はスクリプトプロパティから読む) =====
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL = 'claude-sonnet-5-5';
const MAX_TOKENS = 6000;
const ANTHROPIC_VERSION = '2023-06-01';

// ===== 生成結果のバリデーション条件 =====
const META_DESCRIPTION_MAX_LENGTH = 120;
const BODY_MIN_LENGTH = 2000;
const BODY_MAX_LENGTH = 3000;
const ERROR_MESSAGE_MAX_LENGTH = 200;

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
 * 選択行の記事を Claude API で生成する。
 * ヘッダー行・空行は除外し、ステータスが「生成中」「生成済み」「投稿済み」の行は
 * 重複実行防止のため何もしない。1回の実行で処理するのは選択中の1行のみ。
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
    const subKeyword = sheet.getRange(row, COL.SUB_KEYWORD).getValue();

    const status = sheet.getRange(row, COL.STATUS).getValue();
    if (status === STATUS.GENERATING || status === STATUS.GENERATED || status === STATUS.POSTED) {
      ui.alert('ステータスが「' + status + '」のため、重複実行を防止して処理を中止しました。');
      return;
    }

    const sampleArticle = getSampleArticleText();
    if (!sampleArticle) {
      ui.alert(
        'エラー: サンプル記事が読み込めません。\n' +
        'シート「' + SAMPLE_SHEET_NAME + '」を作成し、セルA1にサンプル記事の本文を貼り付けてください。'
      );
      return;
    }

    updateStatus(sheet, row, STATUS.GENERATING);

    try {
      const prompt = buildPrompt(keyword, subKeyword, sampleArticle);
      const apiResult = callClaudeApi(prompt);
      const article = extractJsonFromText(apiResult.text, apiResult.stopReason);

      const bodyText = article && typeof article.content_html === 'string'
        ? stripHtmlTags(article.content_html)
        : '';
      const bodyLength = bodyText.length;

      // バリデーション結果にかかわらず、取得できた項目はシートに残す。
      if (article && typeof article.title === 'string' && article.title) {
        sheet.getRange(row, COL.TITLE).setValue(article.title);
      }
      if (article && typeof article.meta_description === 'string') {
        sheet.getRange(row, COL.META_DESCRIPTION).setValue(article.meta_description);
      }
      if (article && typeof article.content_html === 'string' && article.content_html) {
        sheet.getRange(row, COL.BODY).setValue(article.content_html);
        sheet.getRange(row, COL.CHAR_COUNT).setValue(bodyLength);
      }

      const validationError = validateArticle(article, bodyLength);
      if (validationError) {
        throw new Error(validationError);
      }

      updateStatus(sheet, row, STATUS.GENERATED);

      ui.alert(
        '記事の生成が完了しました。\n\n' +
        'キーワード: ' + keyword + '\n' +
        '入力トークン: ' + apiResult.usage.input_tokens + '\n' +
        '出力トークン: ' + apiResult.usage.output_tokens + '\n' +
        '所要時間: ' + apiResult.elapsedSeconds.toFixed(1) + ' 秒'
      );
    } catch (innerError) {
      const safeMessage = truncateErrorMessage(innerError.message);
      updateStatus(sheet, row, STATUS.ERROR, safeMessage);
      ui.alert('エラーが発生しました: ' + safeMessage);
    }
  } catch (error) {
    ui.alert('想定外のエラーが発生しました: ' + error.message);
  }
}

/**
 * シート「サンプル記事」のセルA1からサンプル記事の本文を読み込む。
 * シートが無い、またはA1が空の場合は空文字列を返す。
 * @return {string} サンプル記事の本文(前後の空白は除去)
 */
function getSampleArticleText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SAMPLE_SHEET_NAME);
  if (!sheet) {
    return '';
  }

  const value = sheet.getRange('A1').getValue();
  const text = typeof value === 'string' ? value : String(value || '');
  return text.trim();
}

/**
 * Claude API に渡すプロンプトを組み立てる。
 * 役割 / サンプル記事(トーンのみ参考) / 構成 / 制約 の4要素で構成する。
 * @param {string} keyword メインキーワード
 * @param {string} subKeyword サブキーワード(空の場合あり)
 * @param {string} sampleArticle シート「サンプル記事」A1の本文
 * @return {string} 組み立てたプロンプト全文
 */
function buildPrompt(keyword, subKeyword, sampleArticle) {
  const roleSection =
    'あなたはBtoB SaaS企業のマーケティングライターです。読者である企業の意思決定者・担当者に向けて、' +
    '分かりやすく信頼できる記事を書いてください。';

  const sampleSection =
    '【トーン・文体のサンプル】\n' +
    '以下は文体とトーンの見本として提示するサンプル記事です。このサンプルの内容・具体的な表現は' +
    'そのまま転用・引用しないでください。あくまで「書き方の雰囲気」の参考としてのみ使ってください。\n' +
    '---\n' + sampleArticle + '\n---';

  const structureSection =
    '【構成】\n' +
    '1. タイトル\n' +
    '2. 導入文(約300字)\n' +
    '3. 見出し(H2)を3つ、それぞれ本文を付ける\n' +
    '4. まとめ(約200字)';

  const constraintsSection =
    '【制約】\n' +
    '- 本文全体(HTMLタグを除く)は2,000〜3,000字にすること\n' +
    '- メインキーワード「' + keyword + '」' +
    (subKeyword ? '、サブキーワード「' + subKeyword + '」' : '') +
    'を不自然にならない範囲で本文に含めること\n' +
    '- サンプル記事の表現をそのまま使わず、オリジナルの表現で書くこと\n' +
    '- 法令・統計・数値などは断定的に書きすぎないこと(例:「〜とされています」のような柔らかい表現を使う)\n\n' +
    '以下のJSON形式のみを出力してください。JSON以外の説明文やコードフェンスは付けないでください。\n' +
    '{"title": "...", "meta_description": "...(120字以内)", ' +
    '"content_html": "...(h2, p などのHTMLタグを使った本文)"}';

  return roleSection + '\n\n' + sampleSection + '\n\n' + structureSection + '\n\n' + constraintsSection;
}

/**
 * Claude API (Messages API) を呼び出し、生成テキストとトークン使用量、所要時間を返す。
 * content配列のうちtypeが'text'のブロックはすべて連結して本文とする(先頭要素のみに限定しない)。
 * @param {string} prompt 送信するプロンプト
 * @return {{text: string, usage: {input_tokens: number, output_tokens: number}, elapsedSeconds: number, stopReason: string}}
 */
function callClaudeApi(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error('スクリプトプロパティ ANTHROPIC_API_KEY が設定されていません。');
  }

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: MAX_TOKENS,
    messages: [
      { role: 'user', content: prompt }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const startTime = new Date().getTime();
  const httpResponse = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  const elapsedSeconds = (new Date().getTime() - startTime) / 1000;

  const responseCode = httpResponse.getResponseCode();
  if (responseCode !== 200) {
    // レスポンス本文はエラーメッセージに含めない(認証情報が混入する可能性を避けるため)。
    throw new Error('Claude APIの呼び出しに失敗しました(HTTP ' + responseCode + ')。');
  }

  let responseJson;
  try {
    responseJson = JSON.parse(httpResponse.getContentText());
  } catch (e) {
    throw new Error('Claude APIの応答を解析できませんでした(HTTP ' + responseCode + ')。');
  }

  const contentBlocks = Array.isArray(responseJson.content) ? responseJson.content : [];
  const text = contentBlocks
    .filter(function (block) { return block && block.type === 'text' && typeof block.text === 'string'; })
    .map(function (block) { return block.text; })
    .join('');

  const stopReason = responseJson.stop_reason || '不明';
  const usage = responseJson.usage || { input_tokens: 0, output_tokens: 0 };

  if (!text) {
    throw new Error(buildEmptyResponseErrorMessage(responseCode, stopReason, contentBlocks, usage));
  }

  return { text: text, usage: usage, elapsedSeconds: elapsedSeconds, stopReason: stopReason };
}

/**
 * 本文が空だった場合のエラーメッセージを組み立てる。
 * APIキーやヘッダーなど秘密情報は含めず、原因調査に使える情報(HTTPステータス・stop_reason・
 * contentのtype一覧・usage)のみを含める。
 * @param {number} responseCode HTTPステータスコード
 * @param {string} stopReason レスポンスのstop_reason
 * @param {Array} contentBlocks レスポンスのcontent配列
 * @param {{input_tokens: number, output_tokens: number}} usage トークン使用量
 * @return {string} エラーメッセージ
 */
function buildEmptyResponseErrorMessage(responseCode, stopReason, contentBlocks, usage) {
  const blockTypes = contentBlocks
    .map(function (block) { return block && block.type ? block.type : '不明'; })
    .join(', ') || 'なし';

  let message = 'Claude APIの応答が空のため、JSONを取得できませんでした。' +
    '(HTTP ' + responseCode + '、stop_reason: ' + stopReason + '、' +
    'contentのtype一覧: ' + blockTypes + '、' +
    'usage: input=' + usage.input_tokens + ' output=' + usage.output_tokens + ')';

  if (stopReason === 'max_tokens') {
    message += ' ※出力が上限(max_tokens)で途切れたと考えられます。';
  }

  return message;
}

/**
 * Claude APIの応答テキストからJSONオブジェクトを取り出す。
 * コードフェンス(```json ... ```)付きの場合や、前後に説明文が付いている場合にも対応する。
 * @param {string} text 応答テキスト
 * @param {string} [stopReason] レスポンスのstop_reason(エラー文言の補足に使う)
 * @return {Object} 解析したJSONオブジェクト
 */
function extractJsonFromText(text, stopReason) {
  if (!text) {
    throw new Error('Claude APIの応答が空のため、JSONを取得できませんでした。');
  }

  let candidate = text.trim();

  const fenceMatch = candidate.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (fenceMatch) {
    candidate = fenceMatch[1].trim();
  }

  const firstBrace = candidate.indexOf('{');
  const lastBrace = candidate.lastIndexOf('}');
  if (firstBrace === -1 || lastBrace === -1 || lastBrace < firstBrace) {
    throw new Error(buildJsonExtractErrorMessage(text, stopReason));
  }
  candidate = candidate.substring(firstBrace, lastBrace + 1);

  try {
    return JSON.parse(candidate);
  } catch (e) {
    throw new Error(buildJsonExtractErrorMessage(text, stopReason));
  }
}

/**
 * JSON取り出し失敗時のエラーメッセージを組み立てる。
 * 本文全体は出さず、先頭30字・末尾30字のみを含める。
 * @param {string} text 応答テキスト
 * @param {string} [stopReason] レスポンスのstop_reason
 * @return {string} エラーメッセージ
 */
function buildJsonExtractErrorMessage(text, stopReason) {
  const head = text.substring(0, 30);
  const tail = text.length > 30 ? text.substring(text.length - 30) : '';

  let message = 'Claude APIの応答のJSON解析に失敗しました。' +
    '(先頭30字: "' + head + '" / 末尾30字: "' + tail + '")';

  if (stopReason === 'max_tokens') {
    message += ' ※出力が上限(max_tokens)で途切れたと考えられます。';
  }

  return message;
}

/**
 * 生成結果の必須項目・文字数制約を検証する。
 * @param {Object} article 生成されたJSONオブジェクト({title, meta_description, content_html})
 * @param {number} bodyLength content_htmlからHTMLタグを除いた文字数
 * @return {?string} 問題があればエラー理由の文字列、問題なければ null
 */
function validateArticle(article, bodyLength) {
  if (!article || typeof article.title !== 'string' || !article.title) {
    return '生成結果に title が見つかりません。';
  }
  if (!article || typeof article.meta_description !== 'string' || !article.meta_description) {
    return '生成結果に meta_description が見つかりません。';
  }
  if (!article || typeof article.content_html !== 'string' || !article.content_html) {
    return '生成結果に content_html が見つかりません。';
  }
  if (article.meta_description.length > META_DESCRIPTION_MAX_LENGTH) {
    return 'meta_description が ' + META_DESCRIPTION_MAX_LENGTH + ' 字を超えています(' +
      article.meta_description.length + ' 字)。';
  }
  if (bodyLength < BODY_MIN_LENGTH || bodyLength > BODY_MAX_LENGTH) {
    return '本文の文字数が範囲外です(' + bodyLength + ' 字、想定 ' +
      BODY_MIN_LENGTH + '〜' + BODY_MAX_LENGTH + ' 字)。';
  }
  return null;
}

/**
 * HTMLタグを除去したプレーンテキストを返す(文字数カウント用)。
 * @param {string} html HTML文字列
 * @return {string} タグを除去した文字列
 */
function stripHtmlTags(html) {
  return html.replace(/<[^>]*>/g, '');
}

/**
 * I列に書き込むエラーメッセージを200字以内に切り詰める。
 * @param {string} message 元のエラーメッセージ
 * @return {string} 200字以内のメッセージ
 */
function truncateErrorMessage(message) {
  const safe = String(message || '');
  return safe.length > ERROR_MESSAGE_MAX_LENGTH ? safe.substring(0, ERROR_MESSAGE_MAX_LENGTH) : safe;
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
