/**
 * 6年 国語「おすすめパンフレットを作ろう」学習支援アプリ（GASウェブアプリ）
 *
 * ■ セットアップ
 *  - このスクリプトはスプレッドシートに紐づけて（コンテナバインドで）使うことを想定。
 *    スタンドアロンで使う場合は SPREADSHEET_ID にスプレッドシートIDを設定する。
 *  - 必要なシート（設定・名簿・進捗・振り返り）は初回アクセス時に自動作成される。
 *
 * ■ デプロイ設定
 *  - 実行するユーザー：自分
 *  - アクセスできるユーザー：全員
 */

// スタンドアロンスクリプトの場合のみIDを設定する（コンテナバインドなら空のまま）
const SPREADSHEET_ID = '';

const SHEET_SETTINGS = '設定';
const SHEET_ROSTER = '名簿';
const SHEET_PROGRESS = '進捗';
const SHEET_JOURNAL = '振り返り';

const ROSTER_HEADERS = ['年度', 'コード', '氏名'];
const PROGRESS_HEADERS = [
  '年度', 'コード', '氏名',
  'mission1_data', 'mission1_done',
  'mission2_data', 'mission2_done',
  'mission3_data', 'mission3_done',
  'mission4_data', 'mission4_done',
  'mission5_data', 'mission5_done',
  '更新日時'
];
const JOURNAL_HEADERS = ['年度', 'コード', '氏名', '時間目', '振り返り内容', '記入日時', '教員コメント', 'コメント日時'];

/* ============================================================
 * エントリポイント
 * ============================================================ */

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('おすすめパンフレットを作ろう')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
 * 共通ヘルパー
 * ============================================================ */

function ss_() {
  // 1) IDが直接指定されていればそれを使う
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  // 2) スプレッドシートに紐づいたスクリプトなら、そのシートを使う
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  // 3) 単体スクリプトの場合：初回にデータ用スプレッドシートを自動作成し、以後使い回す
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('DATA_SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      // 保存済みのシートが削除されていた場合は作り直す
    }
  }
  // 同時アクセスで2つ作ってしまわないよう、作成はロックの中で行う
  const lock = LockService.getScriptLock();
  let locked = false;
  try { locked = lock.tryLock(20000); } catch (e) {}
  try {
    const recheckId = props.getProperty('DATA_SPREADSHEET_ID');
    if (recheckId) {
      try { return SpreadsheetApp.openById(recheckId); } catch (e) {}
    }
    const created = SpreadsheetApp.create('おすすめパンフレットを作ろう（データ）');
    props.setProperty('DATA_SPREADSHEET_ID', created.getId());
    return created;
  } finally {
    if (locked) lock.releaseLock();
  }
}

function ensureSheets_() {
  const ss = ss_();

  // 4シートすべて揃っていれば何もしない（通常アクセスはここで即リターン）
  if (ss.getSheetByName(SHEET_SETTINGS) && ss.getSheetByName(SHEET_ROSTER) &&
      ss.getSheetByName(SHEET_PROGRESS) && ss.getSheetByName(SHEET_JOURNAL)) {
    return;
  }

  // 初回だけ：同時アクセスで同じシートを二重に作ろうとしないようロックする
  const lock = LockService.getScriptLock();
  let locked = false;
  try { locked = lock.tryLock(20000); } catch (e) {}
  try {
    initSheet_(ss, SHEET_SETTINGS, function (sh) {
      // 日本の年度（4月始まり）で初期値を決める
      const d = new Date();
      const nendo = (d.getMonth() + 1) >= 4 ? d.getFullYear() : d.getFullYear() - 1;
      sh.getRange(1, 1, 2, 2).setValues([
        ['現在の年度', nendo],
        ['教員パスワード', '1258']
      ]);
      sh.getRange('B2').setNumberFormat('@'); // パスワードを文字列として保持
    });
    initSheet_(ss, SHEET_ROSTER, function (sh) {
      sh.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    });
    initSheet_(ss, SHEET_PROGRESS, function (sh) {
      sh.getRange(1, 1, 1, PROGRESS_HEADERS.length).setValues([PROGRESS_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    });
    initSheet_(ss, SHEET_JOURNAL, function (sh) {
      sh.getRange(1, 1, 1, JOURNAL_HEADERS.length).setValues([JOURNAL_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/** シートがなければ作って初期化する。同時実行で先に作られていた場合は静かにスキップする。 */
function initSheet_(ss, name, initFn) {
  if (ss.getSheetByName(name)) return;
  let sh;
  try {
    sh = ss.insertSheet(name);
  } catch (e) {
    // 別の実行が先に作成済みならOK。そうでなければ本当のエラー
    if (ss.getSheetByName(name)) return;
    throw e;
  }
  initFn(sh);
}

function getSettings_() {
  const sh = ss_().getSheetByName(SHEET_SETTINGS);
  const year = Number(sh.getRange('B1').getValue());
  const password = String(sh.getRange('B2').getValue()).trim();
  return {
    year: (year >= 2000 && year <= 2100) ? year : new Date().getFullYear(),
    password: password || '1258'
  };
}

function normCode_(code) {
  const s = String(code || '').trim();
  if (!/^[1-9]\d{3}$/.test(s)) {
    throw new Error('コードは4桁の数字で入力してください。');
  }
  return s;
}

function normName_(name) {
  let s = String(name || '').trim();
  // 制御文字と、数式として解釈されうる先頭文字を除去
  s = s.replace(/[\u0000-\u001f\u007f]/g, '');
  s = s.replace(/^[=+@\-]+/, '');
  if (!s || s.length > 20) {
    throw new Error('名前は1〜20文字で入力してください。');
  }
  return s;
}

function fmtDate_(d) {
  if (d instanceof Date && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d HH:mm');
  }
  return '';
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('こんでいます。少し待ってからもう一度ためしてください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * 「年度」列がA列・「コード」列がB列のシートから、年度＋コードに一致する行番号を返す。
 * 見つからなければ -1。TextFinderでコード列だけを検索して高速化している。
 */
function findRowByYearCode_(sheet, year, code) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const matches = sheet.getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(String(code)).matchEntireCell(true).findAll();
  for (let i = 0; i < matches.length; i++) {
    const row = matches[i].getRow();
    if (String(sheet.getRange(row, 1).getValue()) === String(year)) return row;
  }
  return -1;
}

/** 振り返りシートから 年度＋コード＋時間目 に一致する行番号を返す（なければ -1）。 */
function findJournalRow_(sheet, year, code, hour) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const matches = sheet.getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(String(code)).matchEntireCell(true).findAll();
  for (let i = 0; i < matches.length; i++) {
    const row = matches[i].getRow();
    const vals = sheet.getRange(row, 1, 1, 4).getValues()[0];
    if (String(vals[0]) === String(year) && Number(vals[3]) === Number(hour)) return row;
  }
  return -1;
}

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

/* ============================================================
 * 児童用API
 * ============================================================ */

/**
 * コードでログイン。現在の年度の名簿に登録済みかどうかを返す。
 */
function apiLogin(code) {
  code = normCode_(code);
  ensureSheets_();
  const year = getSettings_().year;
  const sh = ss_().getSheetByName(SHEET_ROSTER);
  const row = findRowByYearCode_(sh, year, code);
  if (row === -1) {
    return { registered: false, year: year };
  }
  return { registered: true, year: year, name: String(sh.getRange(row, 3).getValue()) };
}

/**
 * 初回ログイン時の氏名登録。
 */
function apiRegister(code, name) {
  code = normCode_(code);
  name = normName_(name);
  ensureSheets_();
  return withLock_(function () {
    const year = getSettings_().year;
    const sh = ss_().getSheetByName(SHEET_ROSTER);
    const row = findRowByYearCode_(sh, year, code);
    if (row === -1) {
      sh.appendRow([year, code, name]);
    } else {
      // 別端末で先に登録されていた場合は既存の名前を優先する
      name = String(sh.getRange(row, 3).getValue());
    }
    return { ok: true, name: name, year: year };
  });
}

/**
 * 児童の全データ（ミッション進捗＋振り返り）をまとめて取得。
 */
function apiLoadStudent(code) {
  code = normCode_(code);
  ensureSheets_();
  const ss = ss_();
  const year = getSettings_().year;

  const roster = ss.getSheetByName(SHEET_ROSTER);
  const rrow = findRowByYearCode_(roster, year, code);
  if (rrow === -1) {
    throw new Error('登録が見つかりません。最初の画面からやり直してください。');
  }
  const name = String(roster.getRange(rrow, 3).getValue());

  const prog = ss.getSheetByName(SHEET_PROGRESS);
  const prow = findRowByYearCode_(prog, year, code);
  const missions = {};
  if (prow !== -1) {
    const vals = prog.getRange(prow, 4, 1, 10).getValues()[0];
    for (let i = 0; i < 5; i++) {
      missions[i + 1] = {
        data: vals[i * 2] ? String(vals[i * 2]) : '',
        done: isTrue_(vals[i * 2 + 1])
      };
    }
  } else {
    for (let i = 1; i <= 5; i++) missions[i] = { data: '', done: false };
  }

  const jsh = ss.getSheetByName(SHEET_JOURNAL);
  const journals = [];
  const jlast = jsh.getLastRow();
  if (jlast > 1) {
    const data = jsh.getRange(2, 1, jlast - 1, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (String(r[0]) === String(year) && String(r[1]) === String(code)) {
        journals.push({
          hour: Number(r[3]),
          content: String(r[4] || ''),
          updatedAt: fmtDate_(r[5]),
          comment: String(r[6] || ''),
          commentAt: fmtDate_(r[7])
        });
      }
    }
  }

  return { year: year, name: name, missions: missions, journals: journals };
}

/**
 * ミッションのデータと完了フラグを保存。
 */
function apiSaveMission(code, missionNo, dataJson, done) {
  code = normCode_(code);
  missionNo = Number(missionNo);
  if (!(missionNo >= 1 && missionNo <= 5)) throw new Error('不正なミッション番号です。');
  if (typeof dataJson !== 'string') throw new Error('データ形式が不正です。');
  if (dataJson.length > 45000) throw new Error('データが大きすぎて保存できません。文章を少し短くしてください。');
  ensureSheets_();
  return withLock_(function () {
    const ss = ss_();
    const year = getSettings_().year;
    const roster = ss.getSheetByName(SHEET_ROSTER);
    const rrow = findRowByYearCode_(roster, year, code);
    if (rrow === -1) throw new Error('名前の登録が見つかりません。最初の画面からやり直してください。');
    const name = String(roster.getRange(rrow, 3).getValue());

    const prog = ss.getSheetByName(SHEET_PROGRESS);
    let prow = findRowByYearCode_(prog, year, code);
    if (prow === -1) {
      prog.appendRow([year, code, name, '', false, '', false, '', false, '', false, '', false, new Date()]);
      prow = prog.getLastRow();
    }
    const col = 4 + (missionNo - 1) * 2;
    prog.getRange(prow, col).setValue(dataJson);
    prog.getRange(prow, col + 1).setValue(done === true);
    prog.getRange(prow, 3).setValue(name);
    prog.getRange(prow, 14).setValue(new Date());
    return { ok: true, done: done === true };
  });
}

/**
 * 振り返り（学習ジャーナル）を保存。同じ時間目は上書き。
 */
function apiSaveJournal(code, hour, content) {
  code = normCode_(code);
  hour = Number(hour);
  if (!(hour >= 1 && hour <= 10)) throw new Error('不正な時間目です。');
  content = String(content || '');
  if (content.length > 3000) throw new Error('文章が長すぎて保存できません。');
  ensureSheets_();
  return withLock_(function () {
    const ss = ss_();
    const year = getSettings_().year;
    const roster = ss.getSheetByName(SHEET_ROSTER);
    const rrow = findRowByYearCode_(roster, year, code);
    if (rrow === -1) throw new Error('名前の登録が見つかりません。最初の画面からやり直してください。');
    const name = String(roster.getRange(rrow, 3).getValue());

    const jsh = ss.getSheetByName(SHEET_JOURNAL);
    const jrow = findJournalRow_(jsh, year, code, hour);
    if (jrow === -1) {
      jsh.appendRow([year, code, name, hour, content, new Date(), '', '']);
    } else {
      jsh.getRange(jrow, 3).setValue(name);
      jsh.getRange(jrow, 5).setValue(content);
      jsh.getRange(jrow, 6).setValue(new Date());
    }
    return { ok: true, updatedAt: fmtDate_(new Date()) };
  });
}

/* ============================================================
 * 教員用API（すべてパスワードをサーバー側で検証する）
 * ============================================================ */

function checkTeacher_(password) {
  const settings = getSettings_();
  if (String(password || '').trim() !== settings.password) {
    throw new Error('パスワードがちがいます。');
  }
  return settings;
}

/**
 * 教員ログイン。成功したら現在の年度と選択可能な年度一覧を返す。
 */
function apiTeacherAuth(password) {
  ensureSheets_();
  const settings = checkTeacher_(password);
  return { ok: true, currentYear: settings.year, years: listYears_(settings.year) };
}

function listYears_(currentYear) {
  const set = {};
  set[currentYear] = true;
  const sheets = [SHEET_ROSTER, SHEET_PROGRESS, SHEET_JOURNAL];
  const ss = ss_();
  sheets.forEach(function (nm) {
    const sh = ss.getSheetByName(nm);
    const last = sh.getLastRow();
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
        const y = Number(r[0]);
        if (y >= 2000 && y <= 2100) set[y] = true;
      });
    }
  });
  return Object.keys(set).map(Number).sort(function (a, b) { return b - a; });
}

/**
 * タブ1：指定年度の児童×ミッション達成状況一覧。
 */
function apiTeacherGetProgress(password, year) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  const ss = ss_();

  const roster = ss.getSheetByName(SHEET_ROSTER);
  const students = [];
  const rlast = roster.getLastRow();
  if (rlast > 1) {
    roster.getRange(2, 1, rlast - 1, 3).getValues().forEach(function (r) {
      if (String(r[0]) === String(year)) {
        students.push({ code: String(r[1]), name: String(r[2]) });
      }
    });
  }
  students.sort(function (a, b) { return a.code.localeCompare(b.code); });

  const prog = ss.getSheetByName(SHEET_PROGRESS);
  const progMap = {};
  const plast = prog.getLastRow();
  if (plast > 1) {
    prog.getRange(2, 1, plast - 1, 14).getValues().forEach(function (r) {
      if (String(r[0]) === String(year)) {
        progMap[String(r[1])] = {
          done: [isTrue_(r[4]), isTrue_(r[6]), isTrue_(r[8]), isTrue_(r[10]), isTrue_(r[12])],
          updatedAt: fmtDate_(r[13])
        };
      }
    });
  }

  return students.map(function (s) {
    const p = progMap[s.code] || { done: [false, false, false, false, false], updatedAt: '' };
    return { code: s.code, name: s.name, done: p.done, updatedAt: p.updatedAt };
  });
}

/**
 * タブ1：児童1人分のミッション入力内容の詳細。
 */
function apiTeacherGetStudentDetail(password, year, code) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  code = normCode_(code);
  const ss = ss_();

  const roster = ss.getSheetByName(SHEET_ROSTER);
  const rrow = findRowByYearCode_(roster, year, code);
  const name = rrow === -1 ? '' : String(roster.getRange(rrow, 3).getValue());

  const prog = ss.getSheetByName(SHEET_PROGRESS);
  const prow = findRowByYearCode_(prog, year, code);
  const missions = {};
  if (prow !== -1) {
    const vals = prog.getRange(prow, 4, 1, 10).getValues()[0];
    for (let i = 0; i < 5; i++) {
      missions[i + 1] = {
        data: vals[i * 2] ? String(vals[i * 2]) : '',
        done: isTrue_(vals[i * 2 + 1])
      };
    }
  } else {
    for (let i = 1; i <= 5; i++) missions[i] = { data: '', done: false };
  }
  return { code: code, name: name, missions: missions };
}

/**
 * タブ2：指定年度・時間目の全児童の振り返り一覧（未記入者も含む）。
 */
function apiTeacherGetJournals(password, year, hour) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  hour = Number(hour);
  const ss = ss_();

  const roster = ss.getSheetByName(SHEET_ROSTER);
  const students = [];
  const rlast = roster.getLastRow();
  if (rlast > 1) {
    roster.getRange(2, 1, rlast - 1, 3).getValues().forEach(function (r) {
      if (String(r[0]) === String(year)) {
        students.push({ code: String(r[1]), name: String(r[2]) });
      }
    });
  }
  students.sort(function (a, b) { return a.code.localeCompare(b.code); });

  const jsh = ss.getSheetByName(SHEET_JOURNAL);
  const jmap = {};
  const jlast = jsh.getLastRow();
  if (jlast > 1) {
    jsh.getRange(2, 1, jlast - 1, 8).getValues().forEach(function (r) {
      if (String(r[0]) === String(year) && Number(r[3]) === hour) {
        jmap[String(r[1])] = {
          content: String(r[4] || ''),
          updatedAt: fmtDate_(r[5]),
          comment: String(r[6] || ''),
          commentAt: fmtDate_(r[7])
        };
      }
    });
  }

  return students.map(function (s) {
    const j = jmap[s.code] || { content: '', updatedAt: '', comment: '', commentAt: '' };
    return {
      code: s.code, name: s.name,
      written: !!j.content,
      content: j.content, updatedAt: j.updatedAt,
      comment: j.comment, commentAt: j.commentAt
    };
  });
}

/**
 * タブ2：振り返りへの教員コメント保存。
 */
function apiTeacherSaveComment(password, year, code, hour, comment) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  code = normCode_(code);
  hour = Number(hour);
  if (!(hour >= 1 && hour <= 10)) throw new Error('不正な時間目です。');
  comment = String(comment || '');
  if (comment.length > 1000) throw new Error('コメントが長すぎます。');

  return withLock_(function () {
    const ss = ss_();
    const jsh = ss.getSheetByName(SHEET_JOURNAL);
    let jrow = findJournalRow_(jsh, year, code, hour);
    if (jrow === -1) {
      // 児童が未記入でもコメントを残せるよう、空の行を作成する
      const roster = ss.getSheetByName(SHEET_ROSTER);
      const rrow = findRowByYearCode_(roster, year, code);
      const name = rrow === -1 ? '' : String(roster.getRange(rrow, 3).getValue());
      jsh.appendRow([year, code, name, hour, '', '', comment, new Date()]);
    } else {
      jsh.getRange(jrow, 7).setValue(comment);
      jsh.getRange(jrow, 8).setValue(new Date());
    }
    return { ok: true, commentAt: fmtDate_(new Date()) };
  });
}

/**
 * タブ3：現在の年度を変更。過去データは削除しない。
 */
function apiTeacherSetYear(password, year) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  if (!(year >= 2000 && year <= 2100)) throw new Error('年度は2000〜2100の範囲で入力してください。');
  return withLock_(function () {
    ss_().getSheetByName(SHEET_SETTINGS).getRange('B1').setValue(year);
    return { ok: true, currentYear: year, years: listYears_(year) };
  });
}

/**
 * 児童の氏名を修正（名簿・進捗・振り返りの氏名列をまとめて更新）。
 */
function apiTeacherUpdateName(password, year, code, newName) {
  ensureSheets_();
  checkTeacher_(password);
  year = Number(year);
  code = normCode_(code);
  newName = normName_(newName);

  return withLock_(function () {
    const ss = ss_();
    const roster = ss.getSheetByName(SHEET_ROSTER);
    const rrow = findRowByYearCode_(roster, year, code);
    if (rrow === -1) throw new Error('その児童は名簿に見つかりません。');
    roster.getRange(rrow, 3).setValue(newName);

    const prog = ss.getSheetByName(SHEET_PROGRESS);
    const prow = findRowByYearCode_(prog, year, code);
    if (prow !== -1) prog.getRange(prow, 3).setValue(newName);

    const jsh = ss.getSheetByName(SHEET_JOURNAL);
    const jlast = jsh.getLastRow();
    if (jlast > 1) {
      const matches = jsh.getRange(2, 2, jlast - 1, 1)
        .createTextFinder(String(code)).matchEntireCell(true).findAll();
      matches.forEach(function (m) {
        const row = m.getRow();
        if (String(jsh.getRange(row, 1).getValue()) === String(year)) {
          jsh.getRange(row, 3).setValue(newName);
        }
      });
    }
    return { ok: true, name: newName };
  });
}
