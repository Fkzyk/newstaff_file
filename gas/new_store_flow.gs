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
 *   - 店番で照合し、同じ店番の行があれば更新。店番が無い新店は店名で照合。
 *     どちらも無ければ最下行に追加（重複行は作らない）
 *   - フォームを貼り直した(C4を含む編集をした)ときは、フォームの内容を
 *     そのまま写す=フォームで空欄の項目はフローも空欄に直す(古い値を残さない)。
 *     単セルの手直しでは空欄は書かない(既存値を消さない)
 *   - 営業部・営業部長と、住所分割に失敗した都道府県・市町村は
 *     手入力を尊重して自動では消さない
 *   - 手動で入れる列(採用フローの日程 C〜J、面接会場 Q、面接担当者 T)には
 *     一切触れない
 *   - フローの T列(面接担当者)を入力・変更すると、U列(担当者電話番号)を
 *     「連絡先」シート(Googleコンタクトのエクスポート)から自動入力する。
 *     「長屋（1週間）→石井」のような複数名・注釈付きにも対応
 *   - 結果は画面右下のお知らせ(トースト)で毎回表示。エラーも表示する
 */

// ===== シート名(実物に合わせる。変更時はここだけ直す) =====
var SHEET_SRC      = '新店情報シート貼付';   // 貼り付け元フォーム
var SHEET_FLOW     = '新店（情報）フロー';   // 自動入力する一覧
var SHEET_MASTER   = '営業部';               // 営業部・営業部長のマスター
var SHEET_CONTACTS = '連絡先';               // 面接担当者の電話番号マスター(Googleコンタクト形式)

// ===== 貼付フォーム内の「固定セル」位置 =====
var SRC = {
  storeNoName: 'C4', // 「341  横浜霧ケ丘」= 店番＋店名
  address:     'H4', // 住所(→都道府県・市町村・店舗住所)
  tel:         'C5', // 電話番号(→店舗電話番号)
  handover:    'E6', // 引渡予定日(→引渡日)
  grandOpen:   'K6'  // グランドOP
};
var SRC_FORM_LAST_ROW = 9; // フォーム部分の最終行(これより下=業者表の編集では同期しない)

// ===== フロー一覧の列番号(A=1)。ヘッダーは3行目・データは4行目から =====
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
    // Q=17 面接会場は手動列。K〜Sの一括書き込みでも現状値を保持する
    tel:       18, // R 店舗電話番号  ← フォーム C5
    address:   19, // S 店舗住所      ← フォーム H4
    interviewer:    20, // T 面接担当者(手動入力)
    interviewerTel: 21  // U 担当者電話番号 ← T列入力時に連絡先シートから自動入力
  }
};

// ===== 営業部マスターの列番号(A=1) =====
var MASTER = {
  headerRow: 1,
  firstDataRow: 2,
  col: { dept: 1, pref: 2, city: 3, storeNo: 4, storeName: 5, manager: 11 }
};

// ===== 連絡先シートの列番号(A=1。Googleコンタクトのエクスポート形式) =====
var CONTACTS = {
  firstDataRow: 2,
  col: { first: 1, last: 3, phone1: 25, phone2: 27 } // A=名, C=姓, Y=電話1, AA=電話2
};

/**
 * 貼付シートが編集されたら自動実行(シンプルトリガー)。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheetName = e.range.getSheet().getName();
    if (sheetName === SHEET_SRC) {
      // フォームより下(業者一覧など)の編集では同期しない
      if (e.range.getRow() > SRC_FORM_LAST_ROW) return;
      // C4(店番・店名)を含む編集=フォームの貼り直しとみなし「そのまま写す」モード。
      // それ以外の単発修正は「空欄では消さない」モード。
      syncNewStore(rangeContains_(e.range, SRC.storeNoName));
    } else if (sheetName === SHEET_FLOW) {
      // T列(面接担当者)が編集されたら、その行のU列に電話番号を自動入力
      var colT = FLOW.col.interviewer;
      if (e.range.getColumn() > colT || e.range.getLastColumn() < colT) return;
      var startRow = Math.max(e.range.getRow(), FLOW.firstDataRow);
      var endRow = e.range.getLastRow();
      if (endRow < startRow) return;
      fillInterviewerPhones_(e.range.getSheet(), startRow, endRow, true);
    }
  } catch (err) {
    toast_('自動反映でエラーが起きました: ' + err);
  }
}

/**
 * メニューからの手動実行用。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('新店フロー')
    .addItem('貼付フォームを今すぐ取り込む', 'syncNewStore')
    .addItem('面接担当者の電話番号を一括入力', 'fillAllInterviewerPhones')
    .addToUi();
}

/**
 * メニュー用: フロー一覧全行のU列(担当者電話番号)を連絡先から入力する。
 * 手で入れた値を消さないよう、U列が空欄の行だけ埋める。
 */
function fillAllInterviewerPhones() {
  var flow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FLOW);
  if (!flow) { toast_('シート「' + SHEET_FLOW + '」が見つかりません。'); return; }
  var last = flow.getLastRow();
  if (last < FLOW.firstDataRow) { toast_('フロー一覧にデータ行がありません。'); return; }
  fillInterviewerPhones_(flow, FLOW.firstDataRow, last, false);
}

/**
 * 本体：貼付フォームを読み取り、フロー一覧へアップサートする。
 * mirror=true(既定): フォームの内容をそのまま写す(空欄はフローも空欄にする)。
 * mirror=false: 空欄の項目は書かない(単セル修正時の安全モード)。
 */
function syncNewStore(mirror) {
  mirror = (mirror !== false);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src  = ss.getSheetByName(SHEET_SRC);
  var flow = ss.getSheetByName(SHEET_FLOW);
  if (!src) { toast_('シート「' + SHEET_SRC + '」が見つかりません。シート名を確認してください。'); return; }
  if (!flow) { toast_('シート「' + SHEET_FLOW + '」が見つかりません。シート名を確認してください。'); return; }

  // --- フォームを一括読取(C4:K6) ---
  var block = src.getRange(4, 3, 3, 9).getValues(); // 行4〜6 × 列C〜K
  var raw       = clean_(block[0][0]); // C4 店番・店名
  var address   = clean_(block[0][5]); // H4 住所
  var tel       = clean_(block[1][0]); // C5 電話番号
  var handover  = block[2][2];         // E6 引渡予定日(Dateのまま)
  var grandOpen = block[2][8];         // K6 グランドOP(Dateのまま)

  var parsed = parseStoreNoName_(raw);
  if (!parsed.storeNo && !parsed.storeName) return; // フォームが空(貼り付け前)なら何もしない

  var pc = splitAddress_(address); // {pref, city}
  var master = lookupMaster_(ss, parsed.storeNo, parsed.storeName); // {dept, manager} or null

  // --- 対象行を決定(店番→店名の順で照合。無ければ最下行に追加) ---
  var targetRow = findTargetRow_(flow, parsed.storeNo, parsed.storeName);
  var isNew = targetRow < 0;
  if (isNew) targetRow = Math.max(flow.getLastRow() + 1, FLOW.firstDataRow);

  // --- A:B(所属コード・所属名)。空では消さない ---
  var ab = flow.getRange(targetRow, 1, 1, 2).getValues()[0];
  if (parsed.storeNo !== '') ab[0] = parsed.storeNo;
  if (parsed.storeName)      ab[1] = parsed.storeName;
  flow.getRange(targetRow, 1, 1, 2).setValues([ab]);

  // --- K:S(11〜19列)を一括で読み→必要な列だけ差し替え→一括で書く。
  //     Q(面接会場・手動列)は読んだ現状値をそのまま書き戻すので変わらない ---
  var ks = flow.getRange(targetRow, FLOW.col.dept, 1, FLOW.col.address - FLOW.col.dept + 1).getValues()[0];
  var put = function (col, v, force) {
    if (force || (v !== '' && v !== null && v !== undefined)) ks[col - FLOW.col.dept] = v;
  };
  // 営業部・営業部長: 照合できたときだけ書く(手入力を消さない)
  if (master) {
    put(FLOW.col.dept, master.dept);
    put(FLOW.col.manager, master.manager);
  }
  // 都道府県・市町村: 住所から取れたときだけ書く。
  // mirrorモードで住所自体が空なら、住所と一緒に空欄へ戻す
  if (address) {
    put(FLOW.col.pref, pc.pref);
    put(FLOW.col.city, pc.city);
  } else if (mirror) {
    put(FLOW.col.pref, '', true);
    put(FLOW.col.city, '', true);
  }
  // 引渡日・グランドOP・電話・住所: mirrorモードではフォームをそのまま写す(空欄含む)
  put(FLOW.col.handover,  handover,  mirror);
  put(FLOW.col.grandOpen, grandOpen, mirror);
  put(FLOW.col.tel,       tel,       mirror);
  put(FLOW.col.address,   address,   mirror);
  flow.getRange(targetRow, FLOW.col.dept, 1, ks.length).setValues([ks]);

  var label = (parsed.storeNo !== '' ? parsed.storeNo + ' ' : '') + parsed.storeName;
  toast_((isNew ? '新しい行を追加しました' : '既存の行を更新しました') + ': ' + label + '(' + targetRow + '行目)' +
         (master ? '' : ' ※営業部が見つからず空欄です'));
}

// ============ 面接担当者の電話番号(連絡先シートから) ============

/**
 * フローのstartRow〜endRowについて、T列(面接担当者)の名前を連絡先で調べ
 * U列(担当者電話番号)に入力する。
 * overwrite=true: U列を上書き(T列を変更した直後用)
 * overwrite=false: U列が空欄の行だけ埋める(一括入力用。手入力を消さない)
 */
function fillInterviewerPhones_(flow, startRow, endRow, overwrite) {
  var contacts = loadContacts_();
  if (!contacts) { toast_('シート「' + SHEET_CONTACTS + '」が見つかりません。'); return; }
  if (!contacts.length) { toast_('シート「' + SHEET_CONTACTS + '」に電話番号入りの連絡先がありません。'); return; }

  var n = endRow - startRow + 1;
  // 電話番号の先頭の0が数値化で消えないよう表示値で読む
  var tu = flow.getRange(startRow, FLOW.col.interviewer, n, 2).getDisplayValues();
  var updated = 0, problems = [];
  for (var i = 0; i < n; i++) {
    var name = clean_(tu[i][0]);
    if (!name) continue;                            // T列が空の行は触らない
    if (!overwrite && clean_(tu[i][1])) continue;   // 一括入力ではU列入力済みを守る
    var res = interviewerPhoneText_(name, contacts);
    problems = problems.concat(res.missing);
    if (res.text && res.text !== clean_(tu[i][1])) {
      flow.getRange(startRow + i, FLOW.col.interviewerTel).setValue(res.text);
      updated++;
    }
  }
  var msg = updated > 0 ? '担当者電話番号を' + updated + '件入力しました' : '';
  if (problems.length) {
    msg += (msg ? '。' : '') + '連絡先で特定できません: ' + problems.join('、');
  }
  if (msg) toast_(msg);
}

/**
 * 「長屋（1週間）→石井」のようなT列の値を電話番号文字列にする。
 * 1名なら電話番号のみ、複数名なら「名前+番号」を「、」でつなぐ。
 * 特定できなかった名前は missing に入れて返す(U列には書かない)。
 */
function interviewerPhoneText_(nameCell, contacts) {
  var s = String(nameCell).replace(/[（(][^）)]*[）)]/g, ''); // （注釈）を除去
  var tokens = s.split(/[→、,，・／/＆&\n]+/).map(clean_).filter(function (x) { return x; });
  var parts = [], missing = [];
  for (var i = 0; i < tokens.length; i++) {
    var r = resolveContact_(tokens[i], contacts);
    if (r.phone) parts.push({ name: tokens[i], phone: r.phone });
    if (r.miss) missing.push(r.miss);
  }
  var text = '';
  if (tokens.length === 1 && parts.length === 1) text = parts[0].phone;
  else text = parts.map(function (p) { return p.name + p.phone; }).join('、');
  return { text: text, missing: missing };
}

/**
 * 名前1つを連絡先で解決する。優先順:
 *   ①姓+名の完全一致 ②1文字違い(姓2文字一致が条件。沙/紗のような字体違い対策)
 *   ③姓だけで1人に決まる ④前方一致で1人に決まる
 * 複数候補・該当なしは入力しない(missで理由を返す)。
 */
function resolveContact_(token, contacts) {
  var t = normName_(token);
  if (!t) return { phone: '', miss: '' };
  var exact = [], near = [], byLast = [], prefix = [];
  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    if (c.full === t) exact.push(c);
    if (c.full.length === t.length && t.length >= 3 && c.full.slice(0, 2) === t.slice(0, 2)) {
      var diff = 0;
      for (var k = 0; k < t.length; k++) if (c.full.charAt(k) !== t.charAt(k)) diff++;
      if (diff <= 1) near.push(c);
    }
    if (c.last === t) byLast.push(c);
    if (c.full.indexOf(t) === 0) prefix.push(c);
  }
  if (exact.length === 1) return { phone: exact[0].phone, miss: '' };
  if (near.length === 1) return { phone: near[0].phone, miss: '' };
  if (byLast.length === 1) return { phone: byLast[0].phone, miss: '' };
  if (byLast.length > 1) return { phone: '', miss: token + '(同姓' + byLast.length + '名)' };
  if (prefix.length === 1) return { phone: prefix[0].phone, miss: '' };
  return { phone: '', miss: token + '(見つからない)' };
}

/**
 * 連絡先シート(Googleコンタクト形式)を読み込む。
 * 姓(C列)と電話番号(Y列、無ければAA列)がある行だけ対象。
 * 戻り値: [{full:'姓名(正規化)', last:'姓(正規化)', phone:'090-...'}]。シートが無ければnull
 */
function loadContacts_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONTACTS);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < CONTACTS.firstDataRow) return [];
  // 電話番号の書式(先頭0・ハイフン)を保つため表示値で読む
  var vals = sh.getRange(CONTACTS.firstDataRow, 1, last - CONTACTS.firstDataRow + 1, CONTACTS.col.phone2).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var lastName = clean_(vals[i][CONTACTS.col.last - 1]);
    var firstName = clean_(vals[i][CONTACTS.col.first - 1]);
    var phone = clean_(vals[i][CONTACTS.col.phone1 - 1]) || clean_(vals[i][CONTACTS.col.phone2 - 1]);
    if (!lastName || !phone) continue;
    out.push({ full: normName_(lastName + firstName), last: normName_(lastName), phone: phone });
  }
  return out;
}

// ================= ヘルパー =================

/** 編集範囲がa1セルを含むか */
function rangeContains_(range, a1) {
  var cell = range.getSheet().getRange(a1);
  var r = cell.getRow(), c = cell.getColumn();
  return range.getRow() <= r && range.getLastRow() >= r &&
         range.getColumn() <= c && range.getLastColumn() >= c;
}

/** 「341  横浜霧ケ丘」→ {storeNo:341, storeName:'横浜霧ケ丘'}。全角数字・全角空白も可 */
function parseStoreNoName_(raw) {
  var s = String(raw || '')
    .replace(/　/g, ' ')
    .replace(/[０-９]/g, function (d) { return String.fromCharCode(d.charCodeAt(0) - 0xFEE0); })
    .trim();
  var m = s.match(/^\s*(\d+)\s*(.*)$/);
  if (m) return { storeNo: Number(m[1]), storeName: (m[2] || '').trim() };
  return { storeNo: '', storeName: s };
}

// 名前の途中に市/町/村/郡を含み、単純な正規表現では途中で切れてしまう市名
var CITY_EXCEPTIONS = [
  '四日市市', '廿日市市', '野々市市', '十日町市', '大和郡山市',
  '武蔵村山市', '東村山市', '田村市', '大町市', '蒲郡市',
  '小郡市', '大村市', '羽村市', '村上市', '郡山市'
];

/** 住所 → {pref:都道府県, city:市区町村} */
function splitAddress_(address) {
  var s = String(address || '').trim();
  if (!s) return { pref: '', city: '' };
  // 「京都府」を「京都」で切らないよう、都道府県は明示パターンで取る
  var mp = s.match(/^(東京都|北海道|京都府|大阪府|.{2,3}県)/);
  var pref = mp ? mp[1] : '';
  var rest = mp ? s.slice(pref.length) : s;
  // 例外市名(十日町市・蒲郡市など)を最優先で確認
  for (var i = 0; i < CITY_EXCEPTIONS.length; i++) {
    if (rest.indexOf(CITY_EXCEPTIONS[i]) === 0) return { pref: pref, city: CITY_EXCEPTIONS[i] };
  }
  // 郡部(例: 邑楽郡大泉町)は「〇〇郡」を市町村とする(営業部シートの表記に合わせる)
  var mg = rest.match(/^(.+?郡)(?=.+?[町村])/);
  var mc = rest.match(/^(.+?[市区町村])/); // 政令市は「市」で止まる(横浜市/京都市/葛飾区)
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
    if (storeNo !== '' && no !== '' && Number(no) === Number(storeNo)) return masterHit_(row);
    if (!byName && targetName && normName_(row[MASTER.col.storeName - 1]) === targetName) {
      byName = masterHit_(row);
    }
  }
  return byName; // 店番一致が無ければ店名一致(表記ゆれ吸収)を返す
}

function masterHit_(row) {
  return {
    dept: clean_(row[MASTER.col.dept - 1]),
    manager: stripYear_(row[MASTER.col.manager - 1])
  };
}

/**
 * フロー一覧で対象行を探す(見つからなければ -1)。
 * 店番一致を最優先、無ければ店名一致(店番未採番の新店でも同じ行を更新できる)。
 */
function findTargetRow_(flow, storeNo, storeName) {
  var last = flow.getLastRow();
  if (last < FLOW.firstDataRow) return -1;
  var vals = flow.getRange(FLOW.firstDataRow, 1, last - FLOW.firstDataRow + 1, 2).getValues();
  var target = normName_(storeName);
  var byName = -1;
  for (var i = 0; i < vals.length; i++) {
    var no = vals[i][0];
    if (storeNo !== '' && no !== '' && Number(no) === Number(storeNo)) return FLOW.firstDataRow + i;
    if (byName < 0 && target && normName_(vals[i][1]) === target) byName = FLOW.firstDataRow + i;
  }
  return byName;
}

/** 店名の表記ゆれ吸収：空白除去＋「が/ヶ/ヵ」を「ケ」に統一 */
function normName_(x) {
  return String(x || '')
    .replace(/[\s　]/g, '')
    .replace(/[がヶヵ]/g, 'ケ');
}

/** 「貴島 遼介'20」→「貴島 遼介」 */
function stripYear_(x) {
  return String(x || '').replace(/['’][0-9].*$/, '').trim();
}

function clean_(x) {
  return (x === null || x === undefined) ? '' : String(x).trim();
}

/** 画面右下のお知らせ表示(失敗しても本体処理には影響させない) */
function toast_(msg) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '新店フロー', 8);
  } catch (e) { /* トーストが出せない環境では無視 */ }
}
