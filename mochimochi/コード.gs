/**
 * 「モチモチの木」自由進度学習支援ソフト
 * Google Apps Script Web アプリ（バックエンド）
 *
 * ■ 著作権について
 *   本文・原文は一切保存・表示しません。児童が教科書を見て自分で入力した文のみを扱います。
 *
 * ■ 使い方
 *   1. このスクリプトを含む GAS プロジェクトにひもづくスプレッドシートを用意（新規でよい）。
 *   2. [デプロイ] → [新しいデプロイ] → 種類「ウェブアプリ」。
 *      「次のユーザーとして実行」= 自分、「アクセスできるユーザー」= 全員（校内ドメイン等）。
 *   3. 初回は setup() を一度実行すると、必要なシートと見本の名簿が自動作成されます。
 *      （setup を実行しなくても、doGet 初回アクセス時に自動でシートを用意します。）
 */

/* =========================================================================
 * 定数
 * ======================================================================= */
var SHEETS = {
  ROSTER:   '名簿',
  SCENE:    '場面記録',
  PARALLEL: '並行読書記録',
  SUMMARY:  'まとめ'
};

var HEADERS = {
  ROSTER:   ['出席番号', '氏名'],
  SCENE:    ['タイムスタンプ', '氏名', '場面番号', '登場人物名', '性格語', '根拠文', '理由コメント', '変化コメント'],
  PARALLEL: ['タイムスタンプ', '氏名', '作品名', '登場人物名', '性格語', '根拠文', '感想'],
  SUMMARY:  ['氏名', 'イチオシ登場人物', '出典作品', 'あらすじ', 'プロフィール', '性格まとめ文', '完成フラグ', '更新日時']
};

// 「モチモチの木」の場面情報（本文は含めない。案内文のみ）
var SCENE_META = [
  { no: 1, title: 'おくびょう豆太',   guide: '教科書を読んでから、豆太の様子を書こう。' },
  { no: 2, title: 'やい、木ぃ',       guide: '教科書を読んでから、豆太の気もちを書こう。' },
  { no: 3, title: '霜月二十日のばん', guide: '教科書を読んでから、この場面の豆太を書こう。' },
  { no: 4, title: '豆太は見た',       guide: '教科書を読んでから、豆太の行動を書こう。' },
  { no: 5, title: '弱虫でも、やさしけりゃ', guide: '教科書を読んでから、さいごの豆太を書こう。' }
];

var SCENE_PEOPLE = ['豆太', 'じさま', '医者様'];

var PARALLEL_WORKS = ['花さき山', '八郎', '三コ', 'ソメコとオニ', '岩じさま'];

var PROP_NAME_VISIBLE = 'FRIENDS_NAME_VISIBLE'; // 'true' / 'false'

/* =========================================================================
 * エントリポイント
 * ======================================================================= */
function doGet(e) {
  ensureSheets_();
  var view = (e && e.parameter && e.parameter.view) || 'student';
  var t = HtmlService.createTemplateFromFile('index');
  t.initialView = view; // 'student' or 'teacher'
  return t.evaluate()
    .setTitle('モチモチの木 じゆうしんど学しゅう')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** HTML 内で include() できるように（分割しないので通常は未使用） */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================================================================
 * セットアップ / シート管理
 * ======================================================================= */
function setup() {
  ensureSheets_();
  seedRosterIfEmpty_();
  return '準備ができました。名簿シートに児童の氏名を入力してください。';
}

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheets_() {
  var ss = getSS_();
  Object.keys(SHEETS).forEach(function (key) {
    var name = SHEETS[key];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    var headers = HEADERS[key];
    // 1 行目にヘッダーが無ければ書き込む
    var firstCell = sh.getRange(1, 1).getValue();
    if (firstCell === '' || firstCell === null) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // 既定シート「シート1」が空なら削除
  var blank = ss.getSheetByName('シート1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) {
    ss.deleteSheet(blank);
  }
  seedRosterIfEmpty_();
}

function seedRosterIfEmpty_() {
  var sh = getSS_().getSheetByName(SHEETS.ROSTER);
  if (sh.getLastRow() <= 1) {
    var sample = [];
    for (var i = 1; i <= 6; i++) {
      sample.push([i, '（見本）児童' + i]);
    }
    sh.getRange(2, 1, sample.length, 2).setValues(sample);
  }
}

/** ヘッダー名 → 列 index（0 始まり）のマップ */
function headerMap_(sh) {
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  head.forEach(function (h, i) { map[h] = i; });
  return map;
}

/* =========================================================================
 * 共通データ取得
 * ======================================================================= */
function getBootstrap() {
  return {
    roster: getRoster(),
    scenes: SCENE_META,
    scenePeople: SCENE_PEOPLE,
    parallelWorks: PARALLEL_WORKS,
    nameVisible: isFriendsNameVisible_()
  };
}

function getRoster() {
  var sh = getSS_().getSheetByName(SHEETS.ROSTER);
  if (sh.getLastRow() <= 1) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return values
    .filter(function (r) { return String(r[1]).trim() !== ''; })
    .map(function (r) { return { no: r[0], name: String(r[1]).trim() }; });
}

/* =========================================================================
 * 進捗（ホーム画面用）
 * ======================================================================= */
function getProgress(name) {
  name = String(name || '').trim();
  var sceneCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  var sceneRows = readRows_(SHEETS.SCENE);
  sceneRows.forEach(function (r) {
    if (r['氏名'] === name) {
      var no = Number(r['場面番号']);
      if (sceneCounts[no] !== undefined) sceneCounts[no]++;
    }
  });

  var parallelCount = 0;
  var parallelWorks = {};
  readRows_(SHEETS.PARALLEL).forEach(function (r) {
    if (r['氏名'] === name) {
      parallelCount++;
      if (r['作品名']) parallelWorks[r['作品名']] = true;
    }
  });

  var summary = getSummary(name);

  return {
    scenes: sceneCounts,
    parallelCount: parallelCount,
    parallelWorkCount: Object.keys(parallelWorks).length,
    summaryDone: summary ? summary.done : false,
    summaryStarted: !!summary
  };
}

/* =========================================================================
 * 場面記録（画面③）
 * ======================================================================= */
function saveSceneRecord(data) {
  var name = String(data.name || '').trim();
  if (!name) throw new Error('氏名がありません。');
  if (!data.person) throw new Error('登場人物をえらんでください。');
  if (!String(data.trait || '').trim()) throw new Error('性格・気もちを書いてください。');

  var sh = getSS_().getSheetByName(SHEETS.SCENE);
  sh.appendRow([
    new Date(),
    name,
    Number(data.scene),
    data.person,
    String(data.trait || '').trim(),
    String(data.evidence || '').trim(),
    String(data.reason || '').trim(),
    String(data.change || '').trim()
  ]);
  return getSceneRecords(name);
}

/** 自分の場面記録一覧 */
function getSceneRecords(name) {
  name = String(name || '').trim();
  return readRows_(SHEETS.SCENE)
    .filter(function (r) { return r['氏名'] === name; })
    .map(function (r) {
      return {
        scene: Number(r['場面番号']),
        person: r['登場人物名'],
        trait: r['性格語'],
        evidence: r['根拠文'],
        reason: r['理由コメント'],
        change: r['変化コメント'],
        timestamp: r['タイムスタンプ']
      };
    });
}

/** 友達の意見（同じ場面・同じ人物、自分以外）をランダムに数件 */
function getFriendsOpinions(scene, person, myName, limit) {
  scene = Number(scene);
  limit = limit || 5;
  myName = String(myName || '').trim();
  var visible = isFriendsNameVisible_();

  var pool = readRows_(SHEETS.SCENE).filter(function (r) {
    return Number(r['場面番号']) === scene &&
      r['登場人物名'] === person &&
      r['氏名'] !== myName &&
      String(r['性格語'] || '').trim() !== '';
  });

  shuffle_(pool);
  return pool.slice(0, limit).map(function (r) {
    return {
      name: visible ? r['氏名'] : 'ともだち',
      trait: r['性格語'],
      evidence: r['根拠文']
    };
  });
}

/* =========================================================================
 * 並行読書記録（画面④）
 * ======================================================================= */
function saveParallelRecord(data) {
  var name = String(data.name || '').trim();
  if (!name) throw new Error('氏名がありません。');
  if (!String(data.work || '').trim()) throw new Error('作品名を書いてください。');
  if (!String(data.person || '').trim()) throw new Error('登場人物名を書いてください。');

  var sh = getSS_().getSheetByName(SHEETS.PARALLEL);
  sh.appendRow([
    new Date(),
    name,
    String(data.work || '').trim(),
    String(data.person || '').trim(),
    String(data.trait || '').trim(),
    String(data.evidence || '').trim(),
    String(data.impression || '').trim()
  ]);
  return getParallelRecords(name);
}

function getParallelRecords(name) {
  name = String(name || '').trim();
  return readRows_(SHEETS.PARALLEL)
    .filter(function (r) { return r['氏名'] === name; })
    .map(function (r) {
      return {
        work: r['作品名'],
        person: r['登場人物名'],
        trait: r['性格語'],
        evidence: r['根拠文'],
        impression: r['感想'],
        timestamp: r['タイムスタンプ']
      };
    });
}

/* =========================================================================
 * まとめ（画面⑤）
 * ======================================================================= */
function getSummary(name) {
  name = String(name || '').trim();
  var sh = getSS_().getSheetByName(SHEETS.SUMMARY);
  if (sh.getLastRow() <= 1) return null;
  var map = headerMap_(sh);
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][map['氏名']]).trim() === name) {
      return rowToSummary_(values[i], map);
    }
  }
  return null;
}

function rowToSummary_(row, map) {
  return {
    name: row[map['氏名']],
    person: row[map['イチオシ登場人物']],
    source: row[map['出典作品']],
    story: row[map['あらすじ']],
    profile: row[map['プロフィール']],
    traitSummary: row[map['性格まとめ文']],
    done: row[map['完成フラグ']] === true || row[map['完成フラグ']] === '完成' || row[map['完成フラグ']] === 'TRUE',
    updated: row[map['更新日時']]
  };
}

/** 児童データ（人物選択時に呼び出す）: 場面記録＋並行読書をまとめて返す */
function getSummarySources(name) {
  return {
    scenes: getSceneRecords(name),
    parallel: getParallelRecords(name)
  };
}

/** まとめの保存（同一氏名は上書き = upsert） */
function saveSummary(data) {
  var name = String(data.name || '').trim();
  if (!name) throw new Error('氏名がありません。');
  var sh = getSS_().getSheetByName(SHEETS.SUMMARY);
  var map = headerMap_(sh);
  var done = data.done === true;
  var row = [];
  row[map['氏名']] = name;
  row[map['イチオシ登場人物']] = String(data.person || '').trim();
  row[map['出典作品']] = String(data.source || '').trim();
  row[map['あらすじ']] = String(data.story || '').trim();
  row[map['プロフィール']] = String(data.profile || '').trim();
  row[map['性格まとめ文']] = String(data.traitSummary || '').trim();
  row[map['完成フラグ']] = done ? '完成' : '';
  row[map['更新日時']] = new Date();

  // 既存行を探す
  var targetRow = -1;
  if (sh.getLastRow() > 1) {
    var names = sh.getRange(2, map['氏名'] + 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === name) { targetRow = i + 2; break; }
    }
  }
  if (targetRow === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  }
  return getSummary(name);
}

/* =========================================================================
 * 教員用ダッシュボード（画面⑥）
 * ======================================================================= */
function getTeacherData() {
  var roster = getRoster();
  var sceneRows = readRows_(SHEETS.SCENE);
  var parallelRows = readRows_(SHEETS.PARALLEL);
  var summarySh = getSS_().getSheetByName(SHEETS.SUMMARY);

  // 進捗一覧
  var progress = roster.map(function (stu) {
    var counts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
    sceneRows.forEach(function (r) {
      if (r['氏名'] === stu.name && counts[Number(r['場面番号'])] !== undefined) {
        counts[Number(r['場面番号'])]++;
      }
    });
    var pCount = 0, works = {};
    parallelRows.forEach(function (r) {
      if (r['氏名'] === stu.name) { pCount++; if (r['作品名']) works[r['作品名']] = true; }
    });
    var sm = getSummary(stu.name);
    return {
      no: stu.no,
      name: stu.name,
      scenes: counts,
      parallelCount: pCount,
      parallelWorks: Object.keys(works),
      summaryStarted: !!sm,
      summaryDone: sm ? sm.done : false,
      summaryPerson: sm ? sm.person : ''
    };
  });

  // 場面記録（検索・フィルタ用に全件）
  var sceneRecords = sceneRows.map(function (r) {
    return {
      timestamp: fmtDate_(r['タイムスタンプ']),
      name: r['氏名'],
      scene: Number(r['場面番号']),
      person: r['登場人物名'],
      trait: r['性格語'],
      evidence: r['根拠文'],
      reason: r['理由コメント'],
      change: r['変化コメント']
    };
  });

  var parallelRecords = parallelRows.map(function (r) {
    return {
      timestamp: fmtDate_(r['タイムスタンプ']),
      name: r['氏名'],
      work: r['作品名'],
      person: r['登場人物名'],
      trait: r['性格語'],
      evidence: r['根拠文'],
      impression: r['感想']
    };
  });

  // まとめ一覧
  var summaries = [];
  if (summarySh.getLastRow() > 1) {
    var map = headerMap_(summarySh);
    summarySh.getRange(2, 1, summarySh.getLastRow() - 1, summarySh.getLastColumn())
      .getValues().forEach(function (row) {
        if (String(row[map['氏名']]).trim() === '') return;
        var s = rowToSummary_(row, map);
        s.updated = fmtDate_(s.updated);
        summaries.push(s);
      });
  }

  return {
    scenes: SCENE_META,
    scenePeople: SCENE_PEOPLE,
    progress: progress,
    sceneRecords: sceneRecords,
    parallelRecords: parallelRecords,
    summaries: summaries,
    nameVisible: isFriendsNameVisible_()
  };
}

/* 友達の意見での氏名表示の切り替え（教員用） */
function isFriendsNameVisible_() {
  var v = PropertiesService.getScriptProperties().getProperty(PROP_NAME_VISIBLE);
  return v === null ? true : (v === 'true');
}

function setFriendsNameVisible(flag) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_NAME_VISIBLE, flag ? 'true' : 'false');
  return isFriendsNameVisible_();
}

/* =========================================================================
 * ユーティリティ
 * ======================================================================= */
function readRows_(sheetName) {
  var sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() <= 1) return [];
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  return values.map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function shuffle_(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) !== '[object Date]') return String(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'MM/dd HH:mm');
}
