/**
 * @OnlyCurrentDoc
 *
 * blog-auto-gas / src/Code.gs
 *
 * シート「記事」を使ったブログ記事自動生成・投稿の土台。
 * 段階2: Claude API (Messages API) を呼び出して本文を生成する処理を実装。
 * 段階3: WordPress REST API への下書き保存(DRY_RUN対応)を実装。
 */

// ===== シート・列の定義 =====
const SHEET_NAME = '記事';
const SAMPLE_SHEET_NAME = 'サンプル記事';

const COL = {
  KEYWORD: 1,       // A: キーワード
  SUB_KEYWORD: 2,   // B: サブキーワード
  STATUS: 3,        // C: ステータス
  TITLE: 4,         // D: タイトル
  WP_URL: 5,        // E: WordPress URL(投稿後のURLを書き込む)
  META_DESCRIPTION: 6, // F: メタディスクリプション
  BODY: 7,          // G: 本文
  CHAR_COUNT: 8,    // H: 文字数
  ERROR: 9,         // I: エラー内容
  POST_ID: 10       // J: WordPress投稿ID(段階3で追加。二重投稿防止に使う)
};

const HEADER_ROW = 1;

// ===== ステータス値 =====
const STATUS = {
  NOT_GENERATED: '未生成',
  GENERATING: '生成中',
  GENERATED: '生成済み',
  DRAFT_SAVED: '下書き保存済み',
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
const MAX_TOKENS = 8000;
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
    .addItem('WordPressに下書き保存(選択行)', 'postDraftToWordPress')
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
 * 選択行の記事をWordPressに下書きとして保存する。記事生成とは別操作とし、
 * 「生成→人が確認→投稿」の流れを保つ。
 * DRY_RUN が明示的に 'false' でない限り試運転とし、実際の通信は行わない。
 */
function postDraftToWordPress() {
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

    const status = sheet.getRange(row, COL.STATUS).getValue();
    if (status !== STATUS.GENERATED) {
      ui.alert(
        'ステータスが「' + STATUS.GENERATED + '」ではないため処理を中止しました(現在: 「' + status + '」)。'
      );
      return;
    }

    const title = sheet.getRange(row, COL.TITLE).getValue();
    if (!title || typeof title !== 'string') {
      ui.alert('タイトルが空のため処理を中止しました。');
      return;
    }

    const bodyHtml = sheet.getRange(row, COL.BODY).getValue();
    if (!bodyHtml || typeof bodyHtml !== 'string') {
      ui.alert('本文が空のため処理を中止しました。');
      return;
    }

    const existingPostId = sheet.getRange(row, COL.POST_ID).getValue();
    if (existingPostId) {
      ui.alert(
        'この行はすでにWordPressの投稿ID(' + existingPostId + ')が設定されているため、' +
        '二重投稿防止のため処理を中止しました。'
      );
      return;
    }

    let config;
    try {
      config = getWordPressConfig();
    } catch (configError) {
      ui.alert('設定エラー: ' + configError.message);
      return;
    }

    const metaDescriptionValue = sheet.getRange(row, COL.META_DESCRIPTION).getValue();
    const excerpt = typeof metaDescriptionValue === 'string' ? metaDescriptionValue : String(metaDescriptionValue || '');
    const sanitizedHtml = sanitizeContentHtml(bodyHtml);

    if (config.dryRun) {
      const hostname = extractHostname(config.baseUrl);
      const bodyPlainText = stripHtmlTags(sanitizedHtml);

      // 認証ヘッダー・ユーザー名・パスワードは出さない。送信内容の確認用のみ。
      Logger.log(
        '[WordPress試運転] title: ' + title + '\n' +
        'status: draft\n' +
        'excerpt: ' + excerpt + '\n' +
        '本文の文字数: ' + bodyPlainText.length + ' 字\n' +
        '本文の先頭100字: ' + bodyPlainText.substring(0, 100)
      );

      ui.alert(
        '【試運転】WordPress下書き投稿の内容確認\n\n' +
        '送信先ホスト: ' + hostname + '\n' +
        'タイトル: ' + title + '\n' +
        '本文の文字数: ' + bodyPlainText.length + ' 字\n' +
        'excerptの文字数: ' + excerpt.length + ' 字\n' +
        'status: draft\n\n' +
        '※試運転のため投稿していません。(再実行できます)'
      );
      return;
    }

    try {
      const result = callWordPressCreateDraft(config, title, sanitizedHtml, excerpt);
      sheet.getRange(row, COL.POST_ID).setValue(result.id);
      if (result.link) {
        sheet.getRange(row, COL.WP_URL).setValue(result.link);
      }
      updateStatus(sheet, row, STATUS.DRAFT_SAVED);

      ui.alert(
        'WordPressへの下書き保存が完了しました。\n\n' +
        '投稿ID: ' + result.id + (result.link ? '\nURL: ' + result.link : '')
      );
    } catch (postError) {
      const safeMessage = truncateErrorMessage(postError.message);
      updateStatus(sheet, row, STATUS.ERROR, safeMessage);
      ui.alert('エラーが発生しました: ' + safeMessage);
    }
  } catch (error) {
    ui.alert('想定外のエラーが発生しました: ' + error.message);
  }
}

/**
 * WordPress接続設定をスクリプトプロパティから読み込み、検証する。
 * DRY_RUN は 'false' が明示されたときのみ本番投稿とし、未設定・不明な値は試運転として扱う。
 * @return {{baseUrl: string, user: string, password: string, dryRun: boolean}}
 */
function getWordPressConfig() {
  const properties = PropertiesService.getScriptProperties();
  const rawUrl = properties.getProperty('WP_URL');
  const user = properties.getProperty('WP_USER');
  const password = properties.getProperty('WP_APP_PASSWORD');
  const dryRunValue = properties.getProperty('DRY_RUN');

  if (!rawUrl) {
    throw new Error('スクリプトプロパティ WP_URL が設定されていません。');
  }
  if (rawUrl.indexOf('https://') !== 0) {
    throw new Error('WP_URL は https:// で始まる必要があります。');
  }
  if (!user) {
    throw new Error('スクリプトプロパティ WP_USER が設定されていません。');
  }
  if (!password) {
    throw new Error('スクリプトプロパティ WP_APP_PASSWORD が設定されていません。');
  }

  const baseUrl = rawUrl.replace(/\/+$/, '');
  const dryRun = dryRunValue !== 'false';

  return { baseUrl: baseUrl, user: user, password: password, dryRun: dryRun };
}

/**
 * URLからホスト名のみを取り出す(ダイアログ表示用。パス以下は含めない)。
 * @param {string} url https:// で始まるURL
 * @return {string} ホスト名
 */
function extractHostname(url) {
  return url.replace(/^https:\/\//, '').split('/')[0];
}

/**
 * WordPressに送信する前に、HTML中の危険な要素を取り除く(念のためのサニタイズ)。
 * <script>・<iframe>・on○○= イベントハンドラ属性を除去する。
 * @param {string} html 元のHTML
 * @return {string} サニタイズ後のHTML
 */
function sanitizeContentHtml(html) {
  let sanitized = String(html || '');
  sanitized = sanitized.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  sanitized = sanitized.replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, '');
  sanitized = sanitized.replace(/\son\w+\s*=\s*"[^"]*"/gi, '');
  sanitized = sanitized.replace(/\son\w+\s*=\s*'[^']*'/gi, '');
  return sanitized;
}

/**
 * WordPress REST API (wp/v2/posts) に下書きを作成する。
 * status は常に 'draft' 固定とし、他の値を渡せないようにする。
 * @param {{baseUrl: string, user: string, password: string}} config getWordPressConfig() の戻り値
 * @param {string} title タイトル
 * @param {string} contentHtml サニタイズ済み本文HTML
 * @param {string} excerpt 抜粋(メタディスクリプション)
 * @return {{id: number, link: string}} 作成された投稿のIDとURL
 */
function callWordPressCreateDraft(config, title, contentHtml, excerpt) {
  const endpoint = config.baseUrl + '/wp-json/wp/v2/posts';
  const credentials = Utilities.base64Encode(config.user + ':' + config.password);

  const payload = {
    title: title,
    content: contentHtml,
    excerpt: excerpt,
    status: 'draft'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Basic ' + credentials
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const httpResponse = UrlFetchApp.fetch(endpoint, options);
  const responseCode = httpResponse.getResponseCode();

  if (responseCode !== 200 && responseCode !== 201) {
    throw new Error(buildWordPressErrorMessage(responseCode, httpResponse.getContentText()));
  }

  let responseJson;
  try {
    responseJson = JSON.parse(httpResponse.getContentText());
  } catch (e) {
    throw new Error('WordPressの応答を解析できませんでした(HTTP ' + responseCode + ')。');
  }

  return { id: responseJson.id, link: responseJson.link };
}

/**
 * WordPress投稿失敗時のエラーメッセージを組み立てる。
 * レスポンス本文全体やヘッダーは含めず、HTTPステータスと code フィールドのみを含める。
 * @param {number} responseCode HTTPステータスコード
 * @param {string} responseText レスポンス本文(codeフィールド抽出のためだけに使う)
 * @return {string} エラーメッセージ
 */
function buildWordPressErrorMessage(responseCode, responseText) {
  let code = '不明';
  try {
    const parsed = JSON.parse(responseText);
    if (parsed && typeof parsed.code === 'string') {
      code = parsed.code;
    }
  } catch (e) {
    // 解析できない場合は code を「不明」のままにする(レスポンス本文は出さない)。
  }
  return 'WordPressへの投稿に失敗しました(HTTP ' + responseCode + '、code: ' + code + ')。';
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
    '3. 見出し(H2)を3つ、それぞれ本文を付ける(H3は必要な場合のみ少なく使う)\n' +
    '4. まとめ(約200字、必ず最後まで書き切る)';

  const constraintsSection =
    '【制約】\n' +
    '- 本文全体(HTMLタグを除く)は2,000〜3,000字にすること\n' +
    '- 見出し(H2)は3つとし、H3は必要な場合のみ少なく使うこと\n' +
    '- 表(table)は使わず、段落(p)とリスト(ul/ol)で書くこと\n' +
    '- 出力が途中で切れないよう、必ず「まとめ」まで書き切ること\n' +
    '- メインキーワード「' + keyword + '」' +
    (subKeyword ? '、サブキーワード「' + subKeyword + '」' : '') +
    'を不自然にならない範囲で本文に含めること\n' +
    '- サンプル記事の表現をそのまま使わず、オリジナルの表現で書くこと\n' +
    '- 法令・統計・数値などは断定的に書きすぎないこと(例:「〜とされています」のような柔らかい表現を使う)\n\n' +
    '以下のJSON形式のみを出力してください。JSON以外の説明文やコードフェンスは付けないでください。\n' +
    '{"title": "...", "meta_description": "...(120字以内)", ' +
    '"content_html": "...(h2, p, ul などのHTMLタグを使った本文。tableは使わない)"}';

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
