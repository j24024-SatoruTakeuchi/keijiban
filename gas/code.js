const SHEET_NAME = 'posts';
const SETTINGS_SHEET_NAME = 'settings';
const HEADERS = ['id', 'name', 'message', 'created_at'];

const doGet = (e) => {
  const params = e && e.parameter ? e.parameter : {};
  const mode = params.mode || 'list_view';

  try {
    if (mode === 'list_view' || mode === 'list_posts') {
      return jsonResponse({ ok: true, posts: listPosts() });
    }

    if (mode === 'add_item' || mode === 'add_post') {
      return jsonResponse({ ok: true, post: addPost(params) });
    }

    if (mode === 'convert_message') {
      return jsonResponse({ ok: true, convertedMessage: convertMessage(params) });
    }

    if (mode === 'health') {
      assertPostSheet();
      return jsonResponse({ ok: true, message: 'ok' });
    }

    return jsonResponse({
      ok: false,
      error: 'UNKNOWN_MODE',
      message: '対応していない mode です。'
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      error: error.code || 'SERVER_ERROR',
      message: error.message || '処理中にエラーが発生しました。'
    });
  }
};

const convertMessage = (params) => {
  const id = normalizeText(params.id, 128);

  if (!id) {
    const error = new Error('投稿IDが指定されていません。');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const post = findPostById(id);

  if (!post) {
    const error = new Error('指定された投稿が見つかりません。');
    error.code = 'POST_NOT_FOUND';
    throw error;
  }

  const messages = [
    {
      role: 'system',
      content: 'あなたは掲示板の投稿を、意味を変えずにやさしい言葉へ変換します。相手を責める表現を避け、短く自然な日本語で返してください。'
    },
    {
      role: 'user',
      content: '次の投稿をやさしい言葉に変換してください。'
    },
    {
      role: 'assistant',
      content: 'わかりました。原文の意味を保って、丁寧でやわらかい表現にします。'
    },
    {
      role: 'user',
      content: post.message
    }
  ];

  return callOpenRouter(messages);
};

const findPostById = (id) => {
  const sheet = assertPostSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const row = rows.find((item) => String(item[0]) === id);

  if (!row) {
    return null;
  }

  return {
    id: String(row[0]),
    name: String(row[1] || '匿名'),
    message: String(row[2] || ''),
    createdAt: formatDate(row[3])
  };
};

const callOpenRouter = (messages) => {
  const apiKey = getOpenRouterApiKey();
  const model = 'openai/gpt-oss-120b:free';
  const url = 'https://openrouter.ai/api/v1/chat/completions';
  const payload = {
    model,
    messages,
    temperature: 0.7,
    max_tokens: 512
  };
  const headers = {
    Authorization: 'Bearer ' + apiKey,
    'Content-Type': 'application/json',
    'HTTP-Referer': 'https://script.google.com/',
    'X-Title': 'GAS OpenRouter Sample'
  };
  const options = {
    method: 'post',
    headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const status = res.getResponseCode();
  const text = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error('OpenRouter API Error: HTTP ' + status + '\n' + text);
  }

  const json = JSON.parse(text);
  const content = json && json.choices && json.choices[0] && json.choices[0].message
    ? json.choices[0].message.content
    : '';

  if (!content) {
    throw new Error('OpenRouter APIから変換結果を取得できませんでした。');
  }

  Logger.log(content);
  return content;
};

const getOpenRouterApiKey = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    const error = new Error('スプレッドシートに settings シートを作成し、B2 にOpenRouterのAPIキーを入れてください。');
    error.code = 'SETTINGS_SHEET_NOT_FOUND';
    throw error;
  }

  const apiKey = String(sheet.getRange('B2').getValue() || '').trim();

  if (!apiKey) {
    const error = new Error('settings シートの B2 にOpenRouterのAPIキーを入れてください。');
    error.code = 'API_KEY_NOT_FOUND';
    throw error;
  }

  return apiKey;
};

const listPosts = () => {
  const sheet = assertPostSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return rows
    .filter((row) => row[0] && row[2])
    .map((row) => ({
      id: String(row[0]),
      name: String(row[1] || '匿名'),
      message: String(row[2] || ''),
      createdAt: formatDate(row[3])
    }))
    .reverse();
};

const addPost = (params) => {
  const name = normalizeText(params.name, 24) || '匿名';
  const message = normalizeText(params.message, 500);

  if (!message) {
    const error = new Error('本文を入力してください。');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = assertPostSheet();
    const now = new Date();
    const post = {
      id: Utilities.getUuid(),
      name,
      message,
      createdAt: formatDate(now)
    };

    sheet.appendRow([post.id, post.name, post.message, now]);
    return post;
  } finally {
    lock.releaseLock();
  }
};

const assertPostSheet = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    const error = new Error('スプレッドシートに posts シートを作成し、1行目に id, name, message, created_at を入れてください。');
    error.code = 'SHEET_NOT_FOUND';
    throw error;
  }

  return sheet;
};

const normalizeText = (value, maxLength) => {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim()
    .slice(0, maxLength);
};

const formatDate = (value) => {
  const date = value instanceof Date ? value : new Date(value);

  if (Number.isNaN(date.getTime())) {
    return '';
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
};

const jsonResponse = (payload) => {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify(payload));
  return output;
};
