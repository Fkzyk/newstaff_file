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
 *   - 採用日程(C〜I列)は引渡日から逆算して空欄セルだけ自動入力する。
 *     法則(児玉・横浜霧ケ丘の実績と一致):
 *       稟議申請依頼 = 引渡日の60日前以降の直近月曜
 *       媒体に求人依頼/掲載開始/面接開始 = そこから1週間刻み
 *       時給調査依頼 = 見積取得 = 稟議申請依頼の3週間前 / 面接会場 = その1週間後
 *   - 期日チェック: 期限超過=赤い太字+薄赤背景、3日以内=赤い字+薄黄背景。
 *     シートを開いたとき・メニュー・毎朝のリマインドメール(任意)で確認できる。
 *     済んだ項目はセルをグレーにするか取り消し線を引けば対象外になる
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

// ===== 採用日程(C〜I列)の自動入力と期日チェック =====
var SCHEDULE = {
  firstCol: 3, // C 時給調査依頼
  lastCol:  9, // I 面接開始
  names: ['時給調査依頼', '見積取得', '面接会場', '稟議申請依頼', '媒体に求人依頼', '掲載開始', '面接開始'],
  attentionDays: 3,            // この日数以内に迫った期日を「危険」とする
  fontDanger:  '#cc0000',      // 危険な期日は赤い字
  bgOverdue:   '#f4cccc',      // 期限超過の背景(薄赤)
  bgSoon:      '#fff2cc',      // 3日以内の背景(薄黄)
  doneGreys: ['#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#b7b7b7', '#999999', '#666666']
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
      var startRow = Math.max(e.range.getRow(), FLOW.firstDataRow);
      var endRow = e.range.getLastRow();
      if (endRow < startRow) return;
      // T列(面接担当者)が編集されたら、その行のU列に電話番号を自動入力
      if (colIn_(e.range, FLOW.col.interviewer)) {
        fillInterviewerPhones_(e.range.getSheet(), startRow, endRow, true);
      }
      // O列(引渡日)が編集されたら、その行の空欄日程を逆算入力
      if (colIn_(e.range, FLOW.col.handover)) {
        var filled = fillSchedules_(e.range.getSheet(), startRow, endRow);
        if (filled) toast_('引渡日から日程を' + filled + 'セル入力しました');
      }
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
    .addItem('日程を自動入力(引渡日から逆算)', 'fillAllSchedules')
    .addItem('期限チェックを今すぐ実行', 'checkDeadlines')
    .addSeparator()
    .addItem('✓ 選択した日程を完了にする', 'markScheduleDone')
    .addItem('選択した日程の完了を取り消す', 'unmarkScheduleDone')
    .addSeparator()
    .addItem('毎朝のリマインドメールを有効にする', 'enableDailyReminder')
    .addItem('リマインドメールを止める', 'disableDailyReminder')
    .addToUi();
  // シートを開いたときに取り残しを自己修復してから、期日の色を最新化する。
  // (スクリプト更新前に追加された行など、onEditが効かなかった行もここで必ず追いつく)
  try {
    var flow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FLOW);
    if (flow) {
      var healedSched = 0, healedTel = 0;
      var lastRow = flow.getLastRow();
      if (lastRow >= FLOW.firstDataRow) {
        healedSched = fillSchedules_(flow, FLOW.firstDataRow, lastRow);
        healedTel = fillInterviewerPhones_(flow, FLOW.firstDataRow, lastRow, false, true);
      }
      var items = refreshAttention_(flow);
      var msgs = [];
      if (healedSched || healedTel) {
        msgs.push('取り残しを自動補完: ' +
          (healedSched ? '日程' + healedSched + 'セル' : '') +
          (healedSched && healedTel ? '・' : '') +
          (healedTel ? '電話番号' + healedTel + '件' : ''));
      }
      if (items.length) {
        msgs.push('日程の要注意が' + items.length + '件あります(赤い字=期限超過または3日以内)。' + items[0] + (items.length > 1 ? ' ほか' : ''));
      }
      if (msgs.length) toast_(msgs.join(' / '));
    }
  } catch (err) { /* 開くのを妨げない */ }
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

  // 引渡日が分かっていれば、採用日程(C〜I)の空欄セルを逆算で埋める
  if (handover instanceof Date) fillSchedules_(flow, targetRow, targetRow);

  var label = (parsed.storeNo !== '' ? parsed.storeNo + ' ' : '') + parsed.storeName;
  toast_((isNew ? '新しい行を追加しました' : '既存の行を更新しました') + ': ' + label + '(' + targetRow + '行目)' +
         (master ? '' : ' ※営業部が見つからず空欄です'));
}

// ============ 採用日程の逆算入力と期日チェック ============

/**
 * メニュー用: フロー一覧全行の採用日程(C〜I)を引渡日から逆算して入力する。
 * 空欄のセルだけ埋める(手で入れた日付・既存の日付は変更しない)。
 */
function fillAllSchedules() {
  var flow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FLOW);
  if (!flow) { toast_('シート「' + SHEET_FLOW + '」が見つかりません。'); return; }
  var last = flow.getLastRow();
  if (last < FLOW.firstDataRow) { toast_('フロー一覧にデータ行がありません。'); return; }
  var filled = fillSchedules_(flow, FLOW.firstDataRow, last);
  toast_(filled ? '日程を' + filled + 'セル入力しました(入力済みの日付は変更していません)'
                : '入力できる空欄がありません(引渡日が入っていて、面接開始がまだ先の行の空欄だけ埋めます)');
  refreshAttention_(flow);
}

/**
 * startRow〜endRowの各行について、引渡日(O列)から日程を逆算し
 * C〜I列の「空欄セルだけ」埋める。埋めたセル数を返す。
 * すでに面接開始予定日を過ぎている行(対応済みの古い店)は触らない。
 */
function fillSchedules_(flow, startRow, endRow) {
  var n = endRow - startRow + 1;
  var width = SCHEDULE.lastCol - SCHEDULE.firstCol + 1;
  var range = flow.getRange(startRow, SCHEDULE.firstCol, n, width);
  var grid = range.getValues();
  var handovers = flow.getRange(startRow, FLOW.col.handover, n, 1).getValues();
  var today = today_();
  var filled = 0;
  for (var i = 0; i < n; i++) {
    var o = handovers[i][0];
    if (!(o instanceof Date)) continue;
    var plan = scheduleFromHandover_(o);
    if (plan[plan.length - 1].getTime() < today.getTime()) continue; // 面接開始が過去=昔の店は触らない
    for (var j = 0; j < width; j++) {
      if (grid[i][j] === '' || grid[i][j] === null) { grid[i][j] = plan[j]; filled++; }
    }
  }
  if (filled) range.setValues(grid);
  return filled;
}

/**
 * 引渡日から日程を逆算する。児玉・横浜霧ケ丘の実績と一致する法則:
 *   稟議申請依頼 = 引渡日の60日前以降の直近月曜
 *   媒体に求人依頼 = +1週 / 掲載開始 = +2週 / 面接開始 = +3週(≒グランドOPの2か月前)
 *   時給調査依頼 = 見積取得 = 稟議申請依頼の3週間前 / 面接会場 = その1週間後
 * 戻り値: [C,D,E,F,G,H,I] の7つの日付
 */
function scheduleFromHandover_(handover) {
  var f = nextMonday_(addDays_(handover, -60)); // 稟議申請依頼
  var c = addDays_(f, -21);                     // 時給調査依頼=見積取得
  return [c, c, addDays_(c, 7), f, addDays_(f, 7), addDays_(f, 14), addDays_(f, 21)];
}

/**
 * メニュー用: 期日チェックを実行して色を最新化し、結果をお知らせする。
 */
function checkDeadlines() {
  var flow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FLOW);
  if (!flow) { toast_('シート「' + SHEET_FLOW + '」が見つかりません。'); return; }
  var items = refreshAttention_(flow);
  if (!items.length) { toast_('期限が近い・過ぎている日程はありません。'); return; }
  toast_('要注意 ' + items.length + '件: ' + items.slice(0, 3).join(' / ') + (items.length > 3 ? ' ほか' : ''));
}

/**
 * C〜I列の期日を今日と比べて色を付け直す。
 *   期限超過 = 赤い太字 + 薄赤背景 / 3日以内 = 赤い字 + 薄黄背景 / それ以外 = 通常に戻す
 * 対象外: 日付でないセル、グレー背景・グレー字・取り消し線のセル(=対応済み)、
 *         グランドOPが過ぎた行(開店済みの店)。
 * 戻り値: 要注意項目の一覧(リマインドメール・お知らせ用)
 */
function refreshAttention_(flow) {
  var last = flow.getLastRow();
  if (last < FLOW.firstDataRow) return [];
  var n = last - FLOW.firstDataRow + 1;
  var width = SCHEDULE.lastCol - SCHEDULE.firstCol + 1;
  var range = flow.getRange(FLOW.firstDataRow, SCHEDULE.firstCol, n, width);
  var vals = range.getValues();
  var bgs = range.getBackgrounds();
  var fonts = range.getFontColors();
  var weights = range.getFontWeights();
  var lines = range.getFontLines();
  var meta = flow.getRange(FLOW.firstDataRow, 1, n, FLOW.col.grandOpen).getValues();

  var today = today_();
  var soonLimit = addDays_(today, SCHEDULE.attentionDays);
  var items = [];
  for (var i = 0; i < n; i++) {
    var op = meta[i][FLOW.col.grandOpen - 1];
    var rowClosed = (op instanceof Date) && op.getTime() < today.getTime(); // 開店済み
    var label = (clean_(meta[i][0]) ? clean_(meta[i][0]) + ' ' : '') + clean_(meta[i][1]);
    for (var j = 0; j < width; j++) {
      var bg = String(bgs[i][j] || '').toLowerCase();
      var fc = String(fonts[i][j] || '').toLowerCase();
      var oursBg = (bg === SCHEDULE.bgOverdue || bg === SCHEDULE.bgSoon);
      var oursFont = (fc === SCHEDULE.fontDanger);
      var reset = function () { // 通常表示に戻す(自分が付けた色だけ)
        if (oursBg) bgs[i][j] = null;
        if (oursFont) { fonts[i][j] = null; weights[i][j] = 'normal'; }
      };
      var v = vals[i][j];
      if (!(v instanceof Date) || rowClosed) { reset(); continue; }
      var done = lines[i][j] === 'line-through' || SCHEDULE.doneGreys.indexOf(bg) >= 0 ||
                 (fc && SCHEDULE.doneGreys.indexOf(fc) >= 0);
      if (done) continue; // 対応済み(グレー/取り消し線)は触らない
      var t = new Date(v.getTime()); t.setHours(0, 0, 0, 0);
      if (t.getTime() < today.getTime()) {
        bgs[i][j] = SCHEDULE.bgOverdue; fonts[i][j] = SCHEDULE.fontDanger; weights[i][j] = 'bold';
        items.push(label + ': ' + SCHEDULE.names[j] + ' ' + fmtDate_(t) + '【期限超過】');
      } else if (t.getTime() <= soonLimit.getTime()) {
        bgs[i][j] = SCHEDULE.bgSoon; fonts[i][j] = SCHEDULE.fontDanger; weights[i][j] = 'normal';
        var left = Math.round((t.getTime() - today.getTime()) / 86400000);
        items.push(label + ': ' + SCHEDULE.names[j] + ' ' + fmtDate_(t) + '(' + (left === 0 ? '今日' : 'あと' + left + '日') + ')');
      } else {
        reset();
      }
    }
  }
  range.setBackgrounds(bgs);
  range.setFontColors(fonts);
  range.setFontWeights(weights);
  return items;
}

/**
 * メニュー用: 選択中の日程セルを「完了」にする(グレー+取り消し線)。
 * 完了にすると赤字・リマインドの対象から外れる。
 */
function markScheduleDone() { applyDoneMark_(true); }

/** メニュー用: 選択中の日程セルの「完了」を取り消す(必要なら赤字に戻る) */
function unmarkScheduleDone() { applyDoneMark_(false); }

function applyDoneMark_(done) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEET_FLOW) {
    toast_('「' + SHEET_FLOW + '」シートで、完了にしたい日程のセルを選んでから押してください。');
    return;
  }
  var rangeList = ss.getActiveRangeList();
  if (!rangeList) { toast_('完了にしたい日程のセルを選んでから押してください。'); return; }

  var count = 0;
  var ranges = rangeList.getRanges();
  for (var k = 0; k < ranges.length; k++) {
    // 選択範囲のうち、日程エリア(データ行×C〜I列)に重なる部分だけに適用
    var r1 = Math.max(ranges[k].getRow(), FLOW.firstDataRow);
    var r2 = ranges[k].getLastRow();
    var c1 = Math.max(ranges[k].getColumn(), SCHEDULE.firstCol);
    var c2 = Math.min(ranges[k].getLastColumn(), SCHEDULE.lastCol);
    if (r2 < r1 || c2 < c1) continue;
    var target = sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    if (done) {
      target.setFontColor('#999999').setFontLine('line-through')
            .setBackground('#efefef').setFontWeight('normal');
    } else {
      target.setFontColor(null).setFontLine('none')
            .setBackground(null).setFontWeight('normal');
    }
    var vals = target.getValues();
    for (var i = 0; i < vals.length; i++) {
      for (var j = 0; j < vals[i].length; j++) if (vals[i][j] instanceof Date) count++;
    }
  }
  if (!count) {
    toast_('選択の中に日程(時給調査依頼〜面接開始の日付)がありません。日付のセルを選んでから押してください。');
    return;
  }
  if (!done) refreshAttention_(sheet); // 取り消した日程がまだ危険なら赤字に戻す
  toast_(done ? '✓ ' + count + '件の日程を完了にしました(グレー+取り消し線)'
              : count + '件の完了を取り消しました');
}

/**
 * 毎朝の自動リマインド(時間トリガーから実行)。
 * 要注意の日程がある日だけ、自分宛にまとめメールを送る。
 */
function dailyReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var flow = ss.getSheetByName(SHEET_FLOW);
  if (!flow) return;
  // 取り残しの自己修復(onEditが効かなかった行への追いつき)をここでも行う
  if (flow.getLastRow() >= FLOW.firstDataRow) {
    fillSchedules_(flow, FLOW.firstDataRow, flow.getLastRow());
    fillInterviewerPhones_(flow, FLOW.firstDataRow, flow.getLastRow(), false, true);
  }
  var items = refreshAttention_(flow);
  if (!items.length) return;
  var email = Session.getEffectiveUser().getEmail();
  if (!email) return;
  MailApp.sendEmail(email,
    '【新店フロー】日程リマインド(' + items.length + '件)',
    '新店把握シートで、期限が近い・過ぎている日程があります。\n\n・' + items.join('\n・') +
    '\n\n対応が済んだ項目は、セルをグレーにするか取り消し線を引くとリマインド対象から外れます。\n' + ss.getUrl());
}

/** メニュー用: 毎朝8時台のリマインドメールを有効にする */
function enableDailyReminder() {
  deleteReminderTriggers_();
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(8).create();
  toast_('毎朝8時台のリマインドメールを有効にしました(要注意の日程がある日だけ届きます)');
}

/** メニュー用: リマインドメールを止める */
function disableDailyReminder() {
  deleteReminderTriggers_();
  toast_('リマインドメールを止めました');
}

function deleteReminderTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ============ 面接担当者の電話番号(連絡先シートから) ============

/**
 * フローのstartRow〜endRowについて、T列(面接担当者)の名前を連絡先で調べ
 * U列(担当者電話番号)に入力する。入力・削除した件数の合計を返す。
 * overwrite=true: U列を上書きし、T列が空ならU列も消す(T列を編集した直後用)
 * overwrite=false: U列が空欄の行だけ埋める(一括入力・自己修復用。手入力を消さない)
 * quiet=true: お知らせを出さない(シートを開いたときの自己修復用)
 */
function fillInterviewerPhones_(flow, startRow, endRow, overwrite, quiet) {
  var contacts = loadContacts_();
  if (!contacts) { if (!quiet) toast_('シート「' + SHEET_CONTACTS + '」が見つかりません。'); return 0; }
  if (!contacts.length) { if (!quiet) toast_('シート「' + SHEET_CONTACTS + '」に電話番号入りの連絡先がありません。'); return 0; }

  var n = endRow - startRow + 1;
  // 電話番号の先頭の0が数値化で消えないよう表示値で読む
  var tu = flow.getRange(startRow, FLOW.col.interviewer, n, 2).getDisplayValues();
  var updated = 0, cleared = 0, problems = [];
  for (var i = 0; i < n; i++) {
    var name = clean_(tu[i][0]);
    if (!name) {
      // 担当者を消したら電話番号も消す(T列を編集した直後だけ。一括入力では触らない)
      if (overwrite && clean_(tu[i][1])) {
        flow.getRange(startRow + i, FLOW.col.interviewerTel).setValue('');
        cleared++;
      }
      continue;
    }
    if (!overwrite && clean_(tu[i][1])) continue;   // 一括入力ではU列入力済みを守る
    var res = interviewerPhoneText_(name, contacts);
    problems = problems.concat(res.missing);
    if (res.text && res.text !== clean_(tu[i][1])) {
      flow.getRange(startRow + i, FLOW.col.interviewerTel).setValue(res.text);
      updated++;
    }
  }
  var parts = [];
  if (updated) parts.push('担当者電話番号を' + updated + '件入力しました');
  if (cleared) parts.push('担当者が消されたので電話番号も' + cleared + '件消しました');
  if (problems.length) parts.push('連絡先で特定できません: ' + problems.join('、'));
  if (!quiet && parts.length) toast_(parts.join('。'));
  return updated + cleared;
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

/** 編集範囲がcol列(番号)を含むか */
function colIn_(range, col) {
  return range.getColumn() <= col && range.getLastColumn() >= col;
}

/** 日付にn日足す(引く) */
function addDays_(d, n) {
  var x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  x.setHours(0, 0, 0, 0);
  return x;
}

/** その日以降の直近の月曜(その日が月曜ならその日) */
function nextMonday_(d) {
  return addDays_(d, (8 - d.getDay()) % 7);
}

/** 今日の0時 */
function today_() {
  var t = new Date();
  t.setHours(0, 0, 0, 0);
  return t;
}

/** 日付を「7/20」形式にする */
function fmtDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d');
}

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
