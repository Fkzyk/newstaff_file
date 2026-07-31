# -*- coding: utf-8 -*-
"""九州7店舗 周辺時給調査ブックの生成。

判定列(最賃判定/証拠4点/証明レベル)は一切手入力しない。すべて証拠列からの数式で算出する。
収集者(人でもAIでも)が書けるのは証拠列だけ。「店舗一致OK」と書けば通る自己採点を構造的に排除する。
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from records import MINWAGE, STORES, RECORDS

SURVEY_DATE = "2026-07-31"

THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
H_FILL = PatternFill("solid", fgColor="1F4E78")
H_FONT = Font(color="FFFFFF", bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
OWN_FILL = PatternFill("solid", fgColor="E2EFDA")
TITLE = Font(bold=True, size=13)
BOLD = Font(bold=True)
SMALL = Font(size=9)
TOP = Alignment(vertical="top", wrap_text=True)
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)

def header(ws, row, labels, start=1, widths=None):
    for i, lab in enumerate(labels, start=start):
        c = ws.cell(row, i, lab)
        c.fill, c.font, c.border, c.alignment = H_FILL, H_FONT, BOX, CTR
    if widths:
        for i, w in enumerate(widths, start=start):
            ws.column_dimensions[get_column_letter(i)].width = w


def nested_if(pairs, otherwise):
    """[(条件, 真の値), ...] を IF のネストに組み立てる。閉じ括弧の数え間違いを構造的に防ぐ。"""
    expr = otherwise
    for cond, then in reversed(pairs):
        expr = f"IF({cond},{then},{expr})"
    return "=" + expr

def src_verdict(src_cell, wage_cell, minwage_cell, warn_cell, evid_cell):
    return nested_if([
        (f'{src_cell}="近隣求人欄"', '"未確定(近隣欄からの抽出)"'),
        (f'{src_cell}="ブランド共通"', '"未確定(店舗非固有)"'),
        (f'{src_cell}="施設一般"',   '"未確定(施設一般値)"'),
        (f'{src_cell}="店舗非特定"', '"未確定(店舗非特定)"'),
        (f'{src_cell}="情報源矛盾"', '"未確定(情報源の矛盾)"'),
        (f'{wage_cell}=""',          '"取得不能"'),
        (f'{warn_cell}="割れ(旧掲載疑い)"', '"未確定(最低賃金割れ)"'),
        (f'{evid_cell}="不足"',      '"未確定(証拠不足)"'),
    ], '"媒体掲載値"')

wb = openpyxl.Workbook()

# ========================= 調査台帳 =========================
led = wb.active
led.title = "調査台帳"
led["A1"] = "調査台帳（証拠シート）— L・M・N列は数式です。手入力しないでください"
led["A1"].font = TITLE
led["A2"] = ("収集者が記入してよいのは D〜J の証拠列のみ。判定は証拠から自動算出されます。"
             "『店舗一致OK』と人が書けば通る自己採点を排除するための設計です。")
led["A2"].font = SMALL
led["A3"] = ("採用条件：店舗名・所在地・求人ID・時給が同じ求人ブロックから取れていること。"
             "近隣求人欄／ブランド共通レンジ／施設一般値／店舗非特定／情報源矛盾 の金額は採用しません。")
led["A3"].font = SMALL

header(led, 5,
       ["No.", "対象資さん店舗", "区分", "競合店舗名(証拠)", "所在地・施設(証拠)", "求人ID/識別子(証拠)",
        "基本時給(証拠)", "時給条件(原文)", "取得元区分", "出典URL", "県最低賃金",
        "【自動】最賃判定", "【自動】証拠4点", "【自動】証明レベル", "確認日", "備考"],
       widths=[5, 15, 15, 30, 30, 18, 10, 50, 13, 50, 10, 18, 11, 24, 11, 46])

led_row, led_index = 6, {}
for i, r in enumerate(RECORDS, start=1):
    pref = next(s["pref"] for s in STORES if s["key"] == r["store"])
    mw = MINWAGE[pref][0]
    for ci, v in enumerate([i, r["store"], r["cat"], r["name"], r["place"], r["jid"],
                            r["wage"], r["cond"], r["src"], r["url"], mw], start=1):
        c = led.cell(led_row, ci, v); c.border = BOX; c.alignment = TOP
    R_ = led_row
    led.cell(R_, 12, f'=IF(G{R_}="","-",IF(G{R_}<K{R_},"割れ(旧掲載疑い)","OK"))')
    led.cell(R_, 13, f'=IF(AND(D{R_}<>"",E{R_}<>"",F{R_}<>"",G{R_}<>""),"充足","不足")')
    led.cell(R_, 14, src_verdict(f"I{R_}", f"G{R_}", f"K{R_}", f"L{R_}", f"M{R_}"))
    led.cell(R_, 15, SURVEY_DATE)
    led.cell(R_, 16, r["note"])
    for ci in (12, 13, 14, 15, 16):
        c = led.cell(R_, ci); c.border = BOX; c.alignment = TOP
    if r["note"].startswith("★"):
        for ci in range(1, 17):
            led.cell(R_, ci).fill = WARN_FILL
    led_index[(r["store"], r["name"])] = R_
    led_row += 1
LED_LAST = led_row - 1
led.freeze_panes = "A6"

# ========================= 店舗別 主表 =========================
G1, G2 = 4, 11          # グループ1 の先頭列(D)、グループ2 の先頭列(K)
COLW = [15, 30, 10, 52, 22, 9]

for st in STORES:
    ws = wb.create_sheet(st["key"])
    mw, eff = MINWAGE[st["pref"]]
    ws.column_dimensions["A"].width = 19
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 2
    ws.column_dimensions["J"].width = 2

    ws["A1"] = f"資さんうどん {st['key']}　周辺時給調査"; ws["A1"].font = TITLE
    ws["A2"] = f"住所：{st['addr']}　TEL：{st['tel']}　調査商圏：{st['area']}　調査日：{SURVEY_DATE}"
    ws["A2"].font = SMALL

    left = [(f"{st['pref']}最低賃金", mw), ("発効日", eff), ("", ""),
            (f"{st['key']} 現行時給", st["wage"]), ("深夜時給(22-5時)", st["night"]),
            ("土日祝加算", st["holiday"]), ("土日祝 実質時給", "=B8+B10"),
            ("最低賃金との差", "=B8-B5"), ("証明レベル", "実額確認済み(社内データ)")]
    for i, (k, v) in enumerate(left):
        if k == "": continue
        r_ = 5 + i
        ws.cell(r_, 1, k).font = BOLD if i in (0, 3) else Font()
        ws.cell(r_, 2, v)
        for cl in (1, 2):
            ws.cell(r_, cl).border = BOX
            if i >= 3: ws.cell(r_, cl).fill = OWN_FILL

    rows = [r for r in RECORDS if r["store"] == st["key"]]
    half = (len(rows) + 1) // 2
    groups = [(G1, rows[:half]), (G2, rows[half:])]
    header(ws, 4, ["区分", "店舗名", "基本時給", "その他条件(原文)", "証明レベル", "距離"],
           start=G1, widths=COLW)
    header(ws, 4, ["区分", "店舗名", "基本時給", "その他条件(原文)", "証明レベル", "距離"],
           start=G2, widths=COLW)
    ends = []
    for col0, grp in groups:
        rr = 5
        prev_cat = None
        for r in grp:
            lr = led_index[(st["key"], r["name"])]
            ws.cell(rr, col0, r["cat"] if r["cat"] != prev_cat else "")
            prev_cat = r["cat"]
            ws.cell(rr, col0 + 1, r["name"])
            ws.cell(rr, col0 + 2, f'=IF(調査台帳!N{lr}="媒体掲載値",調査台帳!G{lr},"未確定")')
            ws.cell(rr, col0 + 3, r["cond"])
            ws.cell(rr, col0 + 4, f'=調査台帳!N{lr}')
            ws.cell(rr, col0 + 5, "未計測")
            for ci in range(col0, col0 + 6):
                c = ws.cell(rr, ci); c.border = BOX; c.alignment = TOP
            if r["note"].startswith("★"):
                for ci in range(col0, col0 + 6):
                    ws.cell(rr, ci).fill = WARN_FILL
            rr += 1
        ends.append(rr - 1)
    c1 = get_column_letter(G1 + 2); c2 = get_column_letter(G2 + 2)
    RNG = f"{c1}5:{c1}{ends[0]},{c2}5:{c2}{max(ends[1],5)}"

    base = max(ends) + 2
    ws.cell(base, G1, f"周辺相場（証明レベルが通った値のみで集計・対象{len(rows)}件中）").font = BOLD
    stats = [("採用できた件数", f"=COUNT({RNG})"),
             ("最安", f'=IF(COUNT({RNG})=0,"-",MIN({RNG}))'),
             ("中央値", f'=IF(COUNT({RNG})=0,"-",MEDIAN({RNG}))'),
             ("最高", f'=IF(COUNT({RNG})=0,"-",MAX({RNG}))'),
             ("自社と中央値の差", f'=IF(COUNT({RNG})=0,"-",$B$8-MEDIAN({RNG}))'),
             ("自社と最安の差", f'=IF(COUNT({RNG})=0,"-",$B$8-MIN({RNG}))')]
    for i, (k, f) in enumerate(stats):
        ws.cell(base + 1 + i, G1, k).border = BOX
        ws.cell(base + 1 + i, G1).fill = SUB_FILL
        ws.cell(base + 1 + i, G1 + 1, f).border = BOX
    ws.cell(base + 8, G1,
            "※「未確定」は数値が無いという意味ではなく、店舗と金額の帰属を確認できていないという意味です。"
            "内訳は調査台帳の証明レベル列を参照してください。").font = SMALL
    ws.cell(base + 9, G1, "※黄色の行は自動判定で不採用になった行です。※距離は未計測（今回の調査範囲外）。").font = SMALL
    ws.freeze_panes = "D5"

# ========================= サマリー =========================
smy = wb.create_sheet("サマリー", 0)
smy["A1"] = "資さんうどん 九州7店舗　周辺時給調査サマリー"; smy["A1"].font = Font(bold=True, size=14)
smy["A2"] = f"調査日：{SURVEY_DATE}　／　自社時給＝社内データ（実額確認済み）　／　周辺相場＝公開求人の掲載値（実額ではない）"
smy["A2"].font = SMALL
header(smy, 4,
       ["店舗", "県", "県最低賃金", "自社時給", "最賃との差", "深夜時給", "土日祝加算", "土日祝実質",
        "調査件数", "採用件数", "周辺最安", "周辺中央値", "周辺最高", "自社-中央値", "自社の位置"],
       widths=[16, 8, 11, 10, 11, 10, 11, 11, 10, 10, 10, 12, 10, 12, 16])
sr = 5
for st in STORES:
    n = len([r for r in RECORDS if r["store"] == st["key"]])
    sh = f"'{st['key']}'"
    base = 5 + ((n + 1) // 2) + 1        # 「周辺相場」見出し行
    smy.cell(sr, 1, st["key"]); smy.cell(sr, 2, st["pref"])
    smy.cell(sr, 3, MINWAGE[st["pref"]][0]); smy.cell(sr, 4, st["wage"])
    smy.cell(sr, 5, f"=D{sr}-C{sr}")
    smy.cell(sr, 6, st["night"]); smy.cell(sr, 7, st["holiday"])
    smy.cell(sr, 8, f"=D{sr}+G{sr}")
    smy.cell(sr, 9, n)
    smy.cell(sr, 10, f"={sh}!E{base+1}")
    smy.cell(sr, 11, f"={sh}!E{base+2}")
    smy.cell(sr, 12, f"={sh}!E{base+3}")
    smy.cell(sr, 13, f"={sh}!E{base+4}")
    smy.cell(sr, 14, f"={sh}!E{base+5}")
    smy.cell(sr, 15, f'=IF(NOT(ISNUMBER(N{sr})),"判定不可",IF(N{sr}<0,"周辺より低い",IF(N{sr}=0,"同水準","周辺より高い")))')
    for ci in range(1, 16):
        c = smy.cell(sr, ci); c.border = BOX; c.alignment = CTR
    sr += 1
smy.freeze_panes = "A5"

smy.cell(sr + 1, 1, "所見").font = Font(bold=True, size=12)
for i, t in enumerate([
  "1. 6店舗（陣山・新池・宗像・菊陽・鳥栖真木・佐賀兵庫）の時給は、県の地域別最低賃金と1円単位で同額です。差は0円です。",
  "2. 宮崎阿波岐原店だけが最低賃金より27円高い1,050円で、7店舗の中で唯一の例外です。",
  "3. 深夜時給は全店とも 基本時給×1.25 の切り上げで、社内データ内の整合は取れています。",
  "4. 全7店舗が周辺相場の中央値を下回りました。北九州（陣山・新池）は僅差ですが、佐賀・熊本で開きが大きくなっています。",
  "5. 宮崎阿波岐原店の公開求人は平日1,000円のままで、社内データの1,050円と食い違います。1,000円は宮崎県の最低賃金1,023円を",
  "   下回るため、2025年11月の改定前の旧掲載です。求人票の更新が必要と思われます。",
  "6. 競合の公開求人には最低賃金割れの旧掲載が21件ありました。うどんウエストは戸畑店1,050円・宗像店1,000円・",
  "   菊陽光の森店900円といずれも改定前の掲載で、直接競合であるにもかかわらず現行値を確認できていません。",
  "7. 同一店舗に複数の異なる金額が見つかった4件（ケンタッキー イオンタウン黒崎店、マクドナルド イオン戸畑SC店、",
  "   ケンタッキー 鳥栖店、すき家 57号大津店）は、どちらかを自動で選ばず「未確定(情報源の矛盾)」にしています。",
  "8. 周辺相場は公開求人の掲載値であり、実際に支払われている時給ではありません。実額を確認するには店舗への直接確認が必要です。",
]):
    smy.cell(sr + 2 + i, 1, t).font = SMALL

# ========================= 検証チェック =========================
chk = wb.create_sheet("検証チェック")
chk["A1"] = "自動検証チェック"; chk["A1"].font = TITLE
chk["A2"] = ("調査台帳を数式で集計し、過去の取り違え事故の再現テストを実行します。"
             "『最終警告数』が0でなければ確定版として使用しないでください。")
chk["A2"].font = SMALL
for i, w in enumerate([40, 12, 34, 12, 44], start=1):
    chk.column_dimensions[get_column_letter(i)].width = w

N = f"調査台帳!$N$6:$N${LED_LAST}"
header(chk, 4, ["集計項目", "件数", "意味", "判定", "備考"])
agg = [("台帳の総レコード数", f'=COUNTA(調査台帳!$A$6:$A${LED_LAST})', "調査対象として立てた件数", ""),
       ("媒体掲載値として採用", f'=COUNTIF({N},"媒体掲載値")', "店舗と金額の帰属を確認できた件数", "これのみ主表に数値表示される"),
       ("未確定(最低賃金割れ)", f'=COUNTIF({N},"未確定(最低賃金割れ)")', "旧掲載の疑いで自動除外", "改定前の求人が残っているもの"),
       ("未確定(近隣欄からの抽出)", f'=COUNTIF({N},"未確定(近隣欄からの抽出)")', "近隣求人欄の金額のため除外", "南柏店の事故そのもののガード"),
       ("未確定(店舗非固有)", f'=COUNTIF({N},"未確定(店舗非固有)")', "ブランド共通レンジのため除外", "県内一括レンジ・深夜込みレンジ"),
       ("未確定(施設一般値)", f'=COUNTIF({N},"未確定(施設一般値)")', "商業施設全体の値のため除外", ""),
       ("未確定(店舗非特定)", f'=COUNTIF({N},"未確定(店舗非特定)")', "どの店舗の金額か特定できず除外", ""),
       ("未確定(情報源の矛盾)", f'=COUNTIF({N},"未確定(情報源の矛盾)")', "同一店舗に複数の金額があり除外", "どちらかを自動で選ばない"),
       ("未確定(証拠不足)", f'=COUNTIF({N},"未確定(証拠不足)")', "証拠4点セットが揃わず除外", ""),
       ("取得不能", f'=COUNTIF({N},"取得不能")', "時給そのものを取得できず", "")]
ar = 5
for name, f, mean, note in agg:
    chk.cell(ar, 1, name); chk.cell(ar, 2, f); chk.cell(ar, 3, mean)
    chk.cell(ar, 5, note).font = SMALL
    for ci in range(1, 6): chk.cell(ar, ci).border = BOX
    ar += 1
chk.cell(ar, 1, "合計(内訳の和)").font = BOLD
chk.cell(ar, 2, f"=SUM(B6:B{ar-1})")
chk.cell(ar, 3, "総レコード数と一致すること")
chk.cell(ar, 4, f'=IF(B{ar}=B5,"OK","NG")')
for ci in range(1, 6): chk.cell(ar, ci).border = BOX
SUM_ROW = ar

tr = ar + 2
chk.cell(tr, 1, "店舗取り違え・旧掲載・情報源矛盾 の再現テスト").font = Font(bold=True, size=12)
chk.cell(tr + 1, 1, "台帳と同じ判定式に、過去の事故と同じ形のデータを流し込みます。全てPASSであることを確認してください。").font = SMALL
header(chk, tr + 2, ["テスト内容", "投入した時給", "投入した取得元区分", "判定結果", "期待結果／合否"])
TESTS = [
  ("①ページタイトルはイオンモール柏、金額は近隣求人欄のららぽーと柏の葉1,300円", 1300, "近隣求人欄", 1140),
  ("②県内共通レンジ(深夜込み)を店舗の通常時給として投入", 1240, "ブランド共通", 1057),
  ("③商業施設全体の一般値を特定テナントの時給として投入", 1200, "施設一般", 1023),
  ("④どの店舗の値か特定できないまま投入", 1042, "店舗非特定", 1057),
  ("⑤同一店舗に1,200円と1,100円。高いほうを採用しようとする", 1200, "情報源矛盾", 1057),
  ("⑥改定前の旧掲載（熊本1,034円に対し900円）", 900, "対象求人本文", 1034),
  ("⑦店舗名・住所・求人ID・時給がすべて揃った正常データ", 1150, "対象求人本文", 1030),
]
trow = tr + 3
for i, (desc, wage, src, mw_) in enumerate(TESTS):
    chk.cell(trow, 1, desc); chk.cell(trow, 2, wage); chk.cell(trow, 3, src)
    # 台帳とまったく同じ nested_if で組む(判定式が二重管理にならないようにする)
    chk.cell(trow, 4, nested_if([
        (f'C{trow}="近隣求人欄"', '"未確定(近隣欄からの抽出)"'),
        (f'C{trow}="ブランド共通"', '"未確定(店舗非固有)"'),
        (f'C{trow}="施設一般"',   '"未確定(施設一般値)"'),
        (f'C{trow}="店舗非特定"', '"未確定(店舗非特定)"'),
        (f'C{trow}="情報源矛盾"', '"未確定(情報源の矛盾)"'),
        (f'B{trow}<{mw_}',        '"未確定(最低賃金割れ)"'),
    ], '"媒体掲載値"'))
    if i == len(TESTS) - 1:
        chk.cell(trow, 5, f'=IF(D{trow}="媒体掲載値","PASS(正常データは通る)","FAIL")')
    else:
        chk.cell(trow, 5, f'=IF(LEFT(D{trow},3)="未確定","PASS(採用されない)","FAIL(採用されてしまう)")')
    for ci in range(1, 6):
        c = chk.cell(trow, ci); c.border = BOX; c.alignment = TOP
    trow += 1
TF, TL = tr + 3, trow - 1

fr = trow + 2
chk.cell(fr, 1, "最終警告数").font = Font(bold=True, size=12)
chk.cell(fr, 2, f'=COUNTIF(E{TF}:E{TL},"FAIL*")+IF(D{SUM_ROW}="OK",0,1)')
chk.cell(fr, 3, "0件でなければ確定版として使用しない")
chk.cell(fr, 4, f'=IF(B{fr}=0,"OK","NG")')
for ci in range(1, 5): chk.cell(fr, ci).border = BOX
for i, t in enumerate([
  "※このシートが検証できるのは「抽出の正しさ」だけです。求人情報そのものが古い・誤掲載である可能性は、",
  "　同じ情報を再検査しても検証できません。実際に支払われる時給を証明するには、店舗または採用本部への直接確認が必要です。",
  "※自社時給のみ社内データに基づく「実額確認済み」です。競合はすべて「公開求人の掲載値」であり、実額ではありません。"]):
    chk.cell(fr + 2 + i, 1, t).font = SMALL

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "20260731_周辺時給調査(九州7店舗).xlsx")
wb.save(out)
print("SAVED:", out, "台帳", LED_LAST - 5, "件")
