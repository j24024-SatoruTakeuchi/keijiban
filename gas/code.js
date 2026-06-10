const SHEET_NAME = 'posts';
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
