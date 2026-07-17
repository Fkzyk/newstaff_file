/**
 * 新店把握シート：新店情報シート貼付 → 新店（情報）フロー 自動入力
 *
 * 使い方（詳細は gas/README.md）:
 *   1. スプレッドシートの [拡張機能] → [Apps Script] を開く
 *   2. このファイルの中身をまるごと貼り付けて保存
 *   3. 一度だけ syncNewStore を手動実行して権限を許可
 *   以降、「新店情報シート貼付」に新店情報シートを貼り付けると
 *   「新店（情報）フロー」の該当店舗の行が自動で埋まります。
 *
 * 動作:
 *   - 店番で照合し、同じ店番の行があれば更新／無ければ最下行に追加（アップサート）
 *   - 手動で入れる列（採用フローの日程 C〜J、面接会場 Q、面接担当者 T、
 *     担当者電話番号 U）は自動では触りません（上書きしません）
 *   - 営業部・営業部長は「営業部」シートから店番／店名で照合して取得
 */

// ===== シート名（実物に合わせる。変更時はここだけ直す） =====
var SHEET_SRC    = '新店情報シート貼付';   // 貼り付け元フォーム
var SHEET_FLOW   = '新店（情報）フロー';   // 自動入力する一覧
var SHEET_MASTER = '営業部';               // 営業部・営業部長のマスター

// ===== 貼付フォーム内の「固定セル」位置 =====
var SRC = {
  storeNoName: 'C4', // 「341  横浜霧ケ丘」= 店番＋店名
  address:     'H4', // 住所（→都道府県・市町村・店舗住所）
  tel:         'C5', // 電話番号（→店舗電話番号）
  handover:    'E6', // 引渡予定日（→引渡日）
  grandOpen:   'K6'  // グランドOP
};

// ===== フロー一覧の列番号（A=1）。ヘッダーは3行目・データは4行目から =====
var FLOW = {
  headerRow: 3,
  firstDataRow: 4,
  col: {
    storeNo:   1,  // A 所属コード
    storeName: 2,  // B 所属名
    dept:      11, // K 営業部        ← 営業部シートから
    pref:      12, // L 都道府県      ← 住所から
    city:      13, // M 市町村        ← 住所から
    manager:   14, // N 営業部長      ← 営業部シートから
    handover:  15, // O 引渡日        ← フォーム E6
    grandOpen: 16, // P グランドOP    ← フォーム K6
    tel:       18, // R 店舗電話番号  ← フォーム C5
    address:   19  // S 店舗住所      ← フォーム H4
  }
};

// ===== 営業部マスターの列番号（A=1） =====
var MASTER = {
  headerRow: 1,
  firstDataRow: 2,
  col: { dept: 1, pref: 2, city: 3, storeNo: 4, storeName: 5, manager: 11 }
};

/**
 * 貼付シートが編集されたら自動実行（シンプルトリガー）。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== SHEET_SRC) return;
    syncNewStore();
  } catch (err) {
    // 貼り付け中の中途半端な状態でのエラーは黙って無視（次の編集で再実行される）
  }
}

/**
 * メニューからの手動実行用。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('新店フロー')
    .addItem('貼付フォームを今すぐ取り込む', 'syncNewStore')
    .addToUi();
}

/**
 * 本体：貼付フォームを読み取り、フロー一覧へアップサートする。
 */
function syncNewStore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src  = ss.getSheetByName(SHEET_SRC);
  var flow = ss.getSheetByName(SHEET_FLOW);
  if (!src || !flow) return;

  // --- フォームから値を取得 ---
  var raw = getStr_(src, SRC.storeNoName);
  var parsed = parseStoreNoName_(raw);
  if (!parsed.storeNo && !parsed.storeName) return; // 店番も店名も無ければ何もしない

  var address = getStr_(src, SRC.address);
  var pc = splitAddress_(address); // {pref, city}

  var master = lookupMaster_(ss, parsed.storeNo, parsed.storeName); // {dept, manager} or null

  // --- 書き込む列だけを定義（手動列には触れない） ---
  var writes = {};
  writes[FLOW.col.storeNo]   = parsed.storeNo;   // A
  writes[FLOW.col.storeName] = parsed.storeName; // B
  if (pc.pref) writes[FLOW.col.pref] = pc.pref;  // L
  if (pc.city) writes[FLOW.col.city] = pc.city;  // M
  writes[FLOW.col.handover]  = getVal_(src, SRC.handover);  // O
  writes[FLOW.col.grandOpen] = getVal_(src, SRC.grandOpen); // P
  var tel = getStr_(src, SRC.tel);
  if (tel) writes[FLOW.col.tel] = tel;           // R
  if (address) writes[FLOW.col.address] = address; // S
  if (master) {
    if (master.dept)    writes[FLOW.col.dept]    = master.dept;    // K
    if (master.manager) writes[FLOW.col.manager] = master.manager; // N
  }

  // --- 対象行を決定（店番で照合／無ければ最下行に追加） ---
  var targetRow = findRowByStoreNo_(flow, parsed.storeNo);
  if (targetRow < 0) targetRow = Math.max(flow.getLastRow() + 1, FLOW.firstDataRow);

  // --- 書き込み ---
  Object.keys(writes).forEach(function (colStr) {
    var col = Number(colStr);
    var v = writes[col];
    if (v === '' || v === null || v === undefined) return;
    flow.getRange(targetRow, col).setValue(v);
  });
}

// ================= ヘルパー =================

function getStr_(sheet, a1) {
  var v = sheet.getRange(a1).getValue();
  return (v === null || v === undefined) ? '' : String(v).trim();
}
function getVal_(sheet, a1) {
  return sheet.getRange(a1).getValue(); // 日付はDateのまま返す
}

/** 「341  横浜霧ケ丘」→ {storeNo:341, storeName:'横浜霧ケ丘'} */
function parseStoreNoName_(raw) {
  var s = String(raw || '').replace(/　/g, ' ').trim();
  var m = s.match(/^\s*(\d+)\s*(.*)$/);
  if (m) return { storeNo: Number(m[1]), storeName: (m[2] || '').trim() };
  return { storeNo: '', storeName: s };
}

/** 住所 → {pref:都道府県, city:市区町村} */
function splitAddress_(address) {
  var s = String(address || '').trim();
  if (!s) return { pref: '', city: '' };
  // 「京都府」を「京都」で切らないよう、都道府県は明示パターンで取る
  var mp = s.match(/^(東京都|北海道|京都府|大阪府|.{2,3}県)/);
  var pref = mp ? mp[1] : '';
  var rest = mp ? s.slice(pref.length) : s;
  // 郡部（例: 邑楽郡大泉町）は「〇〇郡」を市町村とする（営業部シートの表記に合わせる）。
  // 「郡山市」のような地名は後ろに町村が続かないので誤爆しない。
  var mg = rest.match(/^(.+?郡)(?=.+?[町村])/);
  var mc = rest.match(/^(.+?[市区町村])/); // 政令市は「市」で止まる（横浜市／京都市／葛飾区）
  var city = mg ? mg[1] : (mc ? mc[1] : '');
  return { pref: pref, city: city };
}

/** 営業部マスターを店番→店名の順で照合。{dept, manager} を返す */
function lookupMaster_(ss, storeNo, storeName) {
  var m = ss.getSheetByName(SHEET_MASTER);
  if (!m) return null;
  var last = m.getLastRow();
  if (last < MASTER.firstDataRow) return null;
  var vals = m.getRange(1, 1, last, MASTER.col.manager).getValues();

  var targetName = normName_(storeName);
  var byName = null;
  for (var r = MASTER.firstDataRow - 1; r < vals.length; r++) {
    var row = vals[r];
    var no = row[MASTER.col.storeNo - 1];
    if (storeNo !== '' && no !== '' && Number(no) === Number(storeNo)) {
      return { dept: clean_(row[MASTER.col.dept - 1]), manager: stripYear_(row[MASTER.col.manager - 1]) };
    }
    if (!byName && targetName && normName_(row[MASTER.col.storeName - 1]) === targetName) {
      byName = { dept: clean_(row[MASTER.col.dept - 1]), manager: stripYear_(row[MASTER.col.manager - 1]) };
    }
  }
  return byName; // 店番一致が無ければ店名一致（緩め）を返す
}

/** フロー一覧で同じ店番の行を探す（見つからなければ -1） */
function findRowByStoreNo_(flow, storeNo) {
  if (storeNo === '' || storeNo === null || storeNo === undefined) return -1;
  var last = flow.getLastRow();
  if (last < FLOW.firstDataRow) return -1;
  var col = flow.getRange(FLOW.firstDataRow, FLOW.col.storeNo, last - FLOW.firstDataRow + 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var v = col[i][0];
    if (v !== '' && v !== null && Number(v) === Number(storeNo)) {
      return FLOW.firstDataRow + i;
    }
  }
  return -1;
}

/** 店名の表記ゆれ吸収：空白除去＋「が/ヶ/ヵ/ケ」を統一 */
function normName_(x) {
  return String(x || '')
    .replace(/[\s　]/g, '')
    .replace(/[がヶヵケ]/g, 'ケ');
}

/** 「貴島 遼介'20」→「貴島 遼介」 */
function stripYear_(x) {
  return String(x || '').replace(/['’][0-9].*$/, '').trim();
}

function clean_(x) {
  return (x === null || x === undefined) ? '' : String(x).trim();
}
