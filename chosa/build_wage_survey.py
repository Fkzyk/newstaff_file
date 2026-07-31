# -*- coding: utf-8 -*-
"""九州7店舗 周辺時給調査ブックの生成。

設計方針(2026-07-31):
  判定列(最低賃金割れ/証拠充足/証明レベル)は一切手入力しない。すべて証拠列からの数式で算出する。
  収集者(AI)が書けるのは「証拠列」だけ。これにより「storeMatch=true と書けば通る」自己採点を構造的に排除する。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SURVEY_DATE = "2026-07-31"

# 県別最低賃金(2025年度改定・発効日)
MINWAGE = {
    "福岡県": (1057, "2025-11-16"),
    "熊本県": (1034, "2026-01-01"),
    "佐賀県": (1030, "2025-11-21"),
    "宮崎県": (1023, "2025-11-16"),
}

# 自社7店舗。wage=None は社内データ未提供
STORES = [
    dict(key="陣山店",       pref="福岡県", addr="福岡県北九州市八幡西区陣山1-3-33",     tel="093-883-6018",
         wage=1057, night=1322, holiday=130),
    dict(key="新池店",       pref="福岡県", addr="福岡県北九州市戸畑区新池3-12-3",       tel="093-616-8024",
         wage=1057, night=1322, holiday=130),
    dict(key="宗像店",       pref="福岡県", addr="福岡県宗像市稲元2-4-8",                tel="0940-34-1177",
         wage=1057, night=1322, holiday=130),
    dict(key="菊陽店",       pref="熊本県", addr="熊本県菊池郡菊陽町津久礼2750-2",       tel="096-234-8668",
         wage=1034, night=1293, holiday=130),
    dict(key="鳥栖真木店",   pref="佐賀県", addr="佐賀県鳥栖市真木町1113-6",             tel="0942-50-5105",
         wage=1030, night=1288, holiday=130),
    dict(key="佐賀兵庫店",   pref="佐賀県", addr="佐賀県佐賀市兵庫町大字藤木1487-4",     tel="0952-97-5501",
         wage=1030, night=1288, holiday=130),
    dict(key="宮崎阿波岐原店", pref="宮崎県", addr="宮崎県宮崎市阿波岐原町請田2420",     tel="0985-65-3393",
         wage=None, night=None, holiday=130),
]

# 調査台帳の証拠レコード。
#   src: 店舗固有求人 / ブランド共通 / 施設一般 / 取得不能
#   wage は「通常時間帯の基本時給」。取れなければ None
R = lambda store, cat, name, place, jid, wage, cond, src, url, note="": dict(
    store=store, cat=cat, name=name, place=place, jid=jid,
    wage=wage, cond=cond, src=src, url=url, note=note)

RECORDS = [
    # ---- 陣山店(黒崎エリア) ----
    R("陣山店", "牛丼", "すき家 3号黒崎店", "北九州市八幡西区(黒崎駅近く)", "3号黒崎店", 1080,
      "求人表示 1,080～1,350円", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/fukuoka/kitakyusyushi/kitakyushushiyahatanishiku/"),
    R("陣山店", "ファミレス", "ジョイフル 北九州黒崎店", "北九州市八幡西区", "job155176519", 1107,
      "9-15時 1,200円／15-22時 1,107円／22-5時 1,564円／5-9時 1,551円／高校生 1,107円", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/fukuoka/kitakyusyushi/kitakyushushiyahatanishiku/job155176519/"),
    R("陣山店", "ファストフード", "マクドナルド フレスポ黒崎店", "北九州市八幡西区", "job147130904", 1060,
      "22-5時 25%UP", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/fukuoka/kitakyusyushi/kitakyushushiyahatanishiku/job147130904/"),
    R("陣山店", "牛丼", "吉野家 3号線黒崎西店", "北九州市八幡西区", "ysn_044844", 1060,
      "求人表示 1,060～1,325円", "店舗固有求人",
      "https://www.yoshinoya.com/baito/op71872/alist/kyushu_fukuoka_kitakyusyushi_kitakyushushiyahatanishiku/"),
    R("陣山店", "うどん(直接競合)", "うどんウエスト(八幡西区)", "", "", 1050,
      "八幡西区として1,050円～と表示", "店舗非特定",
      "https://west-saiyou.net/", "★八幡西区のどの店舗の値か特定できない。数式が自動で弾く"),
    R("陣山店", "ちゃんぽん", "リンガーハット コムシティ黒崎店", "北九州市八幡西区(JR黒崎駅前コムシティ3F)", "job15735", 1050,
      "八幡西区のリンガーハットとして1,050円～と表示", "店舗非特定",
      "https://nishitetsu-store-job.net/jobfind-pc/job/All/15735",
      "★運営は西鉄ストア。1,050円がこの店舗の値か確認できない"),
    R("陣山店", "ファミレス", "ガスト イオンタウン黒崎店", "北九州市八幡西区西曲里町3-3", "", 1240,
      "福岡県内共通レンジ 1,240～1,428円(深夜割増込み)のみ取得", "ブランド共通",
      "https://arbaito.skylark.co.jp/csaiyo/f1dr/pc_job/list/GT/kyuo/fukuoka/all",
      "★県内共通かつ深夜込みレンジ。通常時給として使うと南柏店サイゼリヤと同型の事故になる"),

    # ---- 新池店(戸畑区) ----
    R("新池店", "回転寿司", "スシロー 戸畑鞘ヶ谷店", "北九州市戸畑区鞘ヶ谷", "スシロー戸畑鞘ヶ谷店", 1100,
      "求人表示 1,100～1,425円", "店舗固有求人",
      "https://www.hatarako.net/fukuoka/kitakyushushi/kitakyushushitobataku/ktp30/"),
    R("新池店", "ファストフード", "マクドナルド 戸畑夜宮店", "北九州市戸畑区夜宮", "job156939010", 1060,
      "22-5時 25%UP", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/fukuoka/kitakyusyushi/kitakyushushitobataku/job156939010/"),
    R("新池店", "ファストフード", "マクドナルド イオン戸畑SC店", "北九州市戸畑区汐井町2-2", "40564", None,
      "高校生 1,057円～／22-5時 25%UP。一般時給は未取得", "未取得",
      "https://crewrecruiting.mcdonalds.co.jp/map/40564", "高校生時給しか取れず。通常時給と混同しないため未採用"),
    R("新池店", "うどん(直接競合)", "うどんウエスト 戸畑店", "北九州市戸畑区初音町9-25", "", None,
      "店舗の時給掲載を取得できず", "未取得",
      "https://west-saiyou.net/", "★小倉片野店の1,050円は別店舗。使用しない"),
    R("新池店", "ファミレス", "ジョイフル(戸畑区)", "", "", 1042,
      "検索で並んだのは本城店1,042円・小倉熊本店1,042円。いずれも戸畑区外", "店舗非特定",
      "https://joyful.saiyo-job.jp/dsaiyo/cct7/pc_job/list/all/kyuo/fukuoka/all",
      "★検索で出た本城店・小倉熊本店はいずれも戸畑区外。別店舗の値なので使用しない"),

    # ---- 宗像店 ----
    R("宗像店", "ファミレス", "ジョイフル 宗像徳重店", "福岡県宗像市徳重", "ジョイフル宗像徳重店", 1150,
      "9:00-15:00 の時給", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/fukuoka/kitakyusyushiigai/munakatashi/food/"),
    R("宗像店", "ファストフード", "マクドナルド 宗像ミスターマックス店", "福岡県宗像市", "40031", None,
      "時給の記載を取得できず", "未取得",
      "https://crewrecruiting.mcdonalds.co.jp/map/40031"),
    R("宗像店", "牛丼", "すき家 宗像赤間店", "福岡県宗像市赤間", "593", None,
      "時給の記載を取得できず", "未取得",
      "https://jobs.sukiya.jp/shops/593"),
    R("宗像店", "うどん(直接競合)", "うどんウエスト 宗像店", "福岡県宗像市", "store/62", None,
      "店舗の時給掲載を取得できず", "未取得",
      "https://www.shop-west.jp/store/62.html", "★佐賀市の1,100円は別県の別店舗。使用しない"),

    # ---- 菊陽店(光の森エリア) ----
    R("菊陽店", "うどん(直接競合)", "丸亀製麺 ゆめタウン光の森店", "熊本県菊池郡菊陽町(JR光の森駅北口徒歩7分)",
      "A10701917410", 1200,
      "求人表示 1,200～1,500円／閉店作業 1,100円～／土曜+50円・日祝+100円", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/A10701917410/MDkyujin_d.htm"),
    R("菊陽店", "ファミレス", "ジョイフル 菊陽店", "熊本県菊池郡菊陽町", "job155176308", 1084,
      "9-15時 1,250円／15-18時 1,084円／18-22時 1,200円／22-0時 1,355円／0-9時 1,438円／高校生 1,084円",
      "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/kumamoto/kumamotoshiigai/kikuchigun/job155176308/"),
    R("菊陽店", "うどん(直接競合)", "うどんウエスト 菊陽光の森店", "熊本県菊池郡菊陽町光の森", "job158925669", 900,
      "うどんコーナー 900円～/950円～／深夜早朝 1,400円～／土日祝+50円", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/kumamoto/kumamotoshiigai/kikuchigun/job158925669/",
      "★熊本県最低賃金1,034円を大きく下回る=2026-01-01改定前の旧掲載。実額とみなせない"),
    R("菊陽店", "ファストフード", "マクドナルド 光の森ゆめタウン店", "熊本県菊池郡菊陽町光の森7-33-1", "43542", 1040,
      "熊本県内一般として1,040円～、深夜1,250円と表示", "ブランド共通",
      "https://crewrecruiting.mcdonalds.co.jp/map/43542", "★県内一般の1,040円～は別店舗の値。使用しない"),

    # ---- 鳥栖真木店 ----
    R("鳥栖真木店", "ファミレス", "ジョイフル 鳥栖中央店", "佐賀県鳥栖市", "job156010261", 1200,
      "求人表示 1,200～1,450円", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/saga/tosushi/job156010261/"),
    R("鳥栖真木店", "うどん(直接競合)", "丸亀製麺 鳥栖店", "佐賀県鳥栖市", "A10701916420", 980,
      "開店準備・仕込みスタッフ 980円～(高校生同額)", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/A10701916420/MDkyujin_d.htm",
      "★佐賀県最低賃金1,030円を下回る=旧掲載。かつ職種限定(仕込み)で通常ホールの値ではない"),
    R("鳥栖真木店", "ファストフード", "マクドナルド 鳥栖フレスポ店", "佐賀県鳥栖市", "41505", None,
      "時給の記載を取得できず", "未取得",
      "https://crewrecruiting.mcdonalds.co.jp/map/41505"),
    R("鳥栖真木店", "焼肉", "ウエスト 焼肉 鳥栖店", "佐賀県鳥栖市", "043-06", None,
      "時給の記載を取得できず", "未取得",
      "https://baito.mynavi.jp/cl-002555807358/job-109509242/"),

    # ---- 佐賀兵庫店 ----
    R("佐賀兵庫店", "うどん(直接競合)", "うどんウエスト 佐賀光法店", "佐賀県佐賀市光法", "job84382253", 1150,
      "うどんコーナー 1,150円～／高校生 1,030円～", "店舗固有求人",
      "https://www.baitoru.com/kyushu/jlist/saga/sagashi/job84382253/",
      "兵庫町ではなく佐賀市光法。距離は未計測"),
    R("佐賀兵庫店", "商業施設内飲食", "天麩羅こむぎ モラージュ佐賀", "佐賀県佐賀市巨勢町(モラージュ佐賀)", "", 1050,
      "1,050円以上／高校生可", "店舗固有求人",
      "https://townwork.net/job_search/kw/%E3%83%A2%E3%83%A9%E3%83%BC%E3%82%B8%E3%83%A5+%E3%83%90%E3%82%A4%E3%83%88+%E4%BD%90%E8%B3%80%E7%9C%8C+%E4%BD%90%E8%B3%80%E5%B8%82/"),
    R("佐賀兵庫店", "うどん(直接競合)", "丸亀製麺 佐賀店", "佐賀県佐賀市", "A10701917124", 1000,
      "ホールスタッフ 1,000円～／22時以降25%UP", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/A10701917124/MDkyujin_d.htm",
      "★佐賀県最低賃金1,030円を下回る=旧掲載"),
    R("佐賀兵庫店", "ファストフード", "ロッテリア モラージュ佐賀店", "佐賀県佐賀市(モラージュ佐賀)", "1092", None,
      "時給の記載を取得できず", "未取得",
      "https://lotteria-arbeit.net/jobfind-pc/job/All/1092"),

    # ---- 宮崎阿波岐原店 ----
    R("宮崎阿波岐原店", "うどん(直接競合)", "丸亀製麺 イオンモール宮崎店", "宮崎県宮崎市(イオンモール宮崎)",
      "AC0121911288", 1100,
      "1,100円～／12/31・1/2・1/3は+200円、1/1は+300円", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/AC0121911288/MDkyujin_d.htm"),
    R("宮崎阿波岐原店", "ファミレス", "ジョイフル 宮崎田野店", "宮崎県宮崎市田野町", "job-131030808", 1073,
      "求人表示 1,073～1,467円／日祝+50円", "店舗固有求人",
      "https://baito.mynavi.jp/cl-002216705350/job-131030808/",
      "田野町は市南部。阿波岐原からは遠く同一商圏か要確認"),
    R("宮崎阿波岐原店", "うどん(直接競合)", "丸亀製麺 宮崎住吉店", "宮崎県宮崎市住吉", "A91031990530", 1000,
      "1,000円～／22時以降25%UP", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/A91031990530/MDkyujin_d.htm",
      "★宮崎県最低賃金1,023円を下回る=旧掲載"),
    R("宮崎阿波岐原店", "うどん(直接競合)", "丸亀製麺 宮崎店", "宮崎県宮崎市", "A10701916440", 950,
      "開店準備・仕込みスタッフ 950円～", "店舗固有求人",
      "https://toridoll-job.com/toridoll02/A10701916440/MDkyujin_d.htm",
      "★最低賃金割れ かつ職種限定"),
    R("宮崎阿波岐原店", "牛丼", "すき家(宮崎市内)", "", "", 1080,
      "高校生1,023円／深夜1,350円／早朝+150円／通常1,080円～。新別府店・大島店・郡司分店が混在", "店舗非特定",
      "https://work.sukiya.jp/brand-jobfind/job/All/32871",
      "★複数店舗の値が混在し、どの店舗の1,080円か特定できない。採用しない"),
    R("宮崎阿波岐原店", "商業施設内飲食", "イオンモール宮崎(飲食テナント)", "宮崎県宮崎市新別府町(イオンモール宮崎)", "", 1200,
      "施設全体で飲食1,200円～・カフェ960円以上", "施設一般",
      "https://am-miyazaki-recruit.jp/job/-/info/list",
      "★施設全体の一般値であり特定テナントの時給ではない"),
]

# ============================== 書式 ==============================
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
H_FILL = PatternFill("solid", fgColor="1F4E78")
H_FONT = Font(color="FFFFFF", bold=True, size=10)
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
OWN_FILL = PatternFill("solid", fgColor="E2EFDA")
TITLE = Font(bold=True, size=13)
SMALL = Font(size=9)

def header(ws, row, labels, widths=None, start=1):
    for i, lab in enumerate(labels, start=start):
        c = ws.cell(row, i, lab)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=start):
            ws.column_dimensions[get_column_letter(i)].width = w

wb = openpyxl.Workbook()

# ============================== 調査台帳 ==============================
# 先に作る。判定列はすべて数式。
led = wb.active
led.title = "調査台帳"
led["A1"] = "調査台帳（証拠シート）— 判定列は数式です。手入力しないでください"
led["A1"].font = TITLE
led["A2"] = ("収集者が記入してよいのは D〜J の証拠列のみ。L・M・N の判定は証拠から自動算出されます。"
             "『店舗一致OK』と人が書けば通る自己採点を構造的に排除するための設計です。")
led["A2"].font = SMALL
led["A3"] = ("採用条件：店舗名・所在地・求人ID・時給が同じ求人ブロックから取れていること(証拠4点セット)。"
             "近隣求人欄・ブランド共通レンジ・施設一般値からの金額は採用しません。")
led["A3"].font = SMALL

LED_HDR = ["No.", "対象資さん店舗", "区分", "競合店舗名(証拠)", "所在地・施設(証拠)", "求人ID/識別子(証拠)",
           "基本時給(証拠)", "時給条件(原文)", "取得元区分", "出典URL", "県最低賃金",
           "【自動】最賃判定", "【自動】証拠4点", "【自動】証明レベル", "確認日", "備考"]
LED_W = [5, 14, 14, 26, 30, 16, 10, 46, 14, 52, 10, 18, 12, 24, 11, 46]
header(led, 5, LED_HDR, LED_W)

led_row = 6
led_index = {}   # (store, name) -> row
for i, r in enumerate(RECORDS, start=1):
    pref = next(s["pref"] for s in STORES if s["key"] == r["store"])
    mw = MINWAGE[pref][0]
    vals = [i, r["store"], r["cat"], r["name"], r["place"], r["jid"],
            r["wage"], r["cond"], r["src"], r["url"], mw]
    for ci, v in enumerate(vals, start=1):
        c = led.cell(led_row, ci, v)
        c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=(ci in (5, 8, 10, 16)))
    R_ = led_row
    # L: 最賃判定
    led.cell(R_, 12, f'=IF(G{R_}="","-",IF(G{R_}<K{R_},"割れ(旧掲載疑い)","OK"))')
    # M: 証拠4点セット
    led.cell(R_, 13, f'=IF(AND(D{R_}<>"",E{R_}<>"",F{R_}<>"",G{R_}<>""),"充足","不足")')
    # N: 証明レベル(すべて証拠からの算出)
    led.cell(R_, 14,
        f'=IF(I{R_}="ブランド共通","未確定(店舗非固有)",'
        f'IF(I{R_}="施設一般","未確定(施設一般値)",'
        f'IF(I{R_}="近隣求人欄","未確定(近隣欄からの抽出)",'
        f'IF(I{R_}="店舗非特定","未確定(店舗非特定)",'
        f'IF(G{R_}="","取得不能",'
        f'IF(L{R_}="割れ(旧掲載疑い)","未確定(最低賃金割れ)",'
        f'IF(M{R_}="不足","未確定(証拠不足)","媒体掲載値")))))))')
    led.cell(R_, 15, SURVEY_DATE)
    led.cell(R_, 16, r["note"])
    for ci in (12, 13, 14, 15, 16):
        cc = led.cell(R_, ci); cc.border = BOX
        cc.alignment = Alignment(vertical="top", wrap_text=(ci == 16))
    if r["note"].startswith("★"):
        for ci in range(1, 17):
            led.cell(R_, ci).fill = WARN_FILL
    led_index[(r["store"], r["name"])] = R_
    led_row += 1
LED_LAST = led_row - 1
led.freeze_panes = "A6"

print("台帳行数:", LED_LAST - 5)

# ============================== 店舗別 主表 ==============================
for st in STORES:
    ws = wb.create_sheet(st["key"])
    mw, eff = MINWAGE[st["pref"]]
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 2
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["F"].width = 11
    ws.column_dimensions["G"].width = 48
    ws.column_dimensions["H"].width = 22
    ws.column_dimensions["I"].width = 10

    ws["A1"] = f"資さんうどん {st['key']}　周辺時給調査"
    ws["A1"].font = TITLE
    ws["A2"] = f"住所：{st['addr']}　TEL：{st['tel']}　調査日：{SURVEY_DATE}"
    ws["A2"].font = SMALL

    # 左ブロック: 基準情報
    ws["A4"] = f"{st['pref']}最低賃金"; ws["A4"].font = Font(bold=True)
    ws["B4"] = mw
    ws["A5"] = "発効日"; ws["B5"] = eff
    ws["A7"] = f"{st['key']} 現行時給"; ws["A7"].font = Font(bold=True)
    if st["wage"]:
        ws["B7"] = st["wage"]
        ws["A8"] = "深夜時給(22-5時)"; ws["B8"] = st["night"]
        ws["A9"] = "土日祝加算"; ws["B9"] = st["holiday"]
        ws["A10"] = "土日祝 実質時給"; ws["B10"] = f"=B7+B9"
        ws["A11"] = "最低賃金との差"; ws["B11"] = f"=B7-B4"
        ws["A12"] = "証明レベル"; ws["B12"] = "実額確認済み(社内データ)"
        for r_ in range(7, 13):
            ws[f"A{r_}"].fill = OWN_FILL; ws[f"B{r_}"].fill = OWN_FILL
    else:
        ws["B7"] = "未提供"
        ws["B7"].fill = BAD_FILL
        ws["A8"] = "備考"
        ws["B8"] = "社内データ未提供。公開求人の1,000円は最低賃金割れの旧掲載のため採用不可"
        ws["B8"].font = SMALL
        ws["A9"] = "土日祝加算"; ws["B9"] = st["holiday"]
        ws["A12"] = "証明レベル"; ws["B12"] = "取得不能"
        ws["B12"].fill = BAD_FILL
    for r_ in range(4, 13):
        for cl in ("A", "B"):
            if ws[f"{cl}{r_}"].value is not None:
                ws[f"{cl}{r_}"].border = BOX

    # 右ブロック: 周辺競合
    header(ws, 4, ["区分", "店舗名", "基本時給", "その他条件", "証明レベル", "距離"], start=4)

    rows = [r for r in RECORDS if r["store"] == st["key"]]
    rr = 5
    for r in rows:
        lr = led_index[(st["key"], r["name"])]
        ws.cell(rr, 4, r["cat"])
        ws.cell(rr, 5, r["name"])
        # 時給は台帳の証明レベルが通ったものだけ表示する
        ws.cell(rr, 6, f'=IF(調査台帳!N{lr}="媒体掲載値",調査台帳!G{lr},"未確定")')
        ws.cell(rr, 7, r["cond"])
        ws.cell(rr, 8, f'=調査台帳!N{lr}')
        ws.cell(rr, 9, "未計測")
        for ci in range(4, 10):
            c = ws.cell(rr, ci); c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=(ci == 7))
        if r["note"].startswith("★"):
            for ci in range(4, 10):
                ws.cell(rr, ci).fill = WARN_FILL
        rr += 1

    # 相場サマリー
    first, last = 5, rr - 1
    ws.cell(rr + 1, 4, "周辺相場(採用できた値のみ)").font = Font(bold=True)
    ws.cell(rr + 2, 4, "件数")
    ws.cell(rr + 2, 5, f'=COUNT(F{first}:F{last})')
    ws.cell(rr + 3, 4, "最安")
    ws.cell(rr + 3, 5, f'=IF(COUNT(F{first}:F{last})=0,"-",MIN(F{first}:F{last}))')
    ws.cell(rr + 4, 4, "最高")
    ws.cell(rr + 4, 5, f'=IF(COUNT(F{first}:F{last})=0,"-",MAX(F{first}:F{last}))')
    ws.cell(rr + 5, 4, "中央値")
    ws.cell(rr + 5, 5, f'=IF(COUNT(F{first}:F{last})=0,"-",MEDIAN(F{first}:F{last}))')
    ws.cell(rr + 6, 4, "自社と中央値の差")
    if st["wage"]:
        ws.cell(rr + 6, 5, f'=IF(COUNT(F{first}:F{last})=0,"-",B7-MEDIAN(F{first}:F{last}))')
    else:
        ws.cell(rr + 6, 5, "自社時給が未提供のため算出不可")
    ws.cell(rr + 8, 4, "※「未確定」は数値が無いという意味ではなく、店舗と金額の帰属を確認できていないという意味です。").font = SMALL
    ws.cell(rr + 9, 4, "※距離は未計測です。今回の調査範囲外としています。").font = SMALL
    ws.freeze_panes = "A5"

# ============================== サマリー ==============================
smy = wb.create_sheet("サマリー", 0)
smy["A1"] = "資さんうどん 九州7店舗　周辺時給調査サマリー"
smy["A1"].font = Font(bold=True, size=14)
smy["A2"] = f"調査日：{SURVEY_DATE}　／　自社時給＝社内データ（実額確認済み）　／　周辺相場＝公開求人の掲載値"
smy["A2"].font = SMALL
header(smy, 4,
       ["店舗", "県", "県最低賃金", "自社時給", "最賃との差", "深夜時給", "土日祝加算",
        "土日祝実質", "周辺件数", "周辺最安", "周辺中央値", "周辺最高", "自社-中央値", "自社の位置"],
       [16, 8, 11, 10, 11, 10, 11, 11, 10, 10, 12, 10, 12, 22])
sr = 5
for st in STORES:
    mw = MINWAGE[st["pref"]][0]
    smy.cell(sr, 1, st["key"])
    smy.cell(sr, 2, st["pref"])
    smy.cell(sr, 3, mw)
    smy.cell(sr, 4, st["wage"] if st["wage"] else "未提供")
    smy.cell(sr, 5, f'=IF(ISNUMBER(D{sr}),D{sr}-C{sr},"-")')
    smy.cell(sr, 6, st["night"] if st["night"] else "-")
    smy.cell(sr, 7, st["holiday"])
    smy.cell(sr, 8, f'=IF(ISNUMBER(D{sr}),D{sr}+G{sr},"-")')
    sr += 1

# 周辺統計はシート内の集計セルを直接参照する
sr = 5
for st in STORES:
    n = len([r for r in RECORDS if r["store"] == st["key"]])
    base = 5 + n + 1          # 「周辺相場」見出し行
    smy.cell(sr, 9,  f"='{st['key']}'!E{base+1}")
    smy.cell(sr, 10, f"='{st['key']}'!E{base+2}")
    smy.cell(sr, 11, f"='{st['key']}'!E{base+4}")
    smy.cell(sr, 12, f"='{st['key']}'!E{base+3}")
    smy.cell(sr, 13, f'=IF(AND(ISNUMBER(D{sr}),ISNUMBER(K{sr})),D{sr}-K{sr},"-")')
    smy.cell(sr, 14,
        f'=IF(NOT(ISNUMBER(M{sr})),"判定不可",'
        f'IF(M{sr}<0,"周辺より低い",IF(M{sr}=0,"周辺と同水準","周辺より高い")))')
    for ci in range(1, 15):
        c = smy.cell(sr, ci); c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
    sr += 1

smy.cell(sr + 1, 1, "所見").font = Font(bold=True)
notes = [
    "1. 7店舗中6店舗（社内データ提供分すべて）の時給は、県の地域別最低賃金と1円単位で同額です。差は0円です。",
    "2. 深夜時給は全店とも 基本時給×1.25 の切り上げで、社内データ内の整合は取れています。",
    "3. 宮崎阿波岐原店の社内データが未提供です。公開求人には平日1,000円と出ていますが、",
    "   宮崎県の最低賃金1,023円を下回るため、2025年11月の改定前の旧掲載です。数値として採用していません。",
    "4. 公開求人には最低賃金割れの旧掲載が複数ありました（丸亀製麺 鳥栖店980円・佐賀店1,000円・宮崎住吉店1,000円、",
    "   うどんウエスト 菊陽光の森店900円）。いずれも自動判定で「未確定」に落としています。",
    "5. 周辺相場は公開求人の掲載値であり、実際に支払われている時給ではありません。証明レベルの違いは調査台帳を参照してください。",
]
for i, t in enumerate(notes):
    smy.cell(sr + 2 + i, 1, t).font = SMALL

# ============================== 検証チェック ==============================
chk = wb.create_sheet("検証チェック")
chk["A1"] = "自動検証チェック"
chk["A1"].font = TITLE
chk["A2"] = ("このシートは調査台帳を数式で集計し、取り違え事故の再現テストを実行します。"
             "『最終警告数』が0でなければ、確定版として使用しないでください。")
chk["A2"].font = SMALL
for i, w in enumerate([34, 12, 34, 12, 46], start=1):
    chk.column_dimensions[get_column_letter(i)].width = w

N = f"調査台帳!$N$6:$N${LED_LAST}"
G = f"調査台帳!$G$6:$G${LED_LAST}"
K = f"調査台帳!$K$6:$K${LED_LAST}"
L = f"調査台帳!$L$6:$L${LED_LAST}"

header(chk, 4, ["集計項目", "件数", "意味", "判定", "備考"])
agg = [
    ("台帳の総レコード数", f'=COUNTA(調査台帳!$A$6:$A${LED_LAST})', "調査対象として立てた件数", "", ""),
    ("媒体掲載値として採用", f'=COUNTIF({N},"媒体掲載値")', "店舗と金額の帰属を確認できた件数", "", "これのみ主表に数値表示される"),
    ("未確定(最低賃金割れ)", f'=COUNTIF({N},"未確定(最低賃金割れ)")', "旧掲載の疑いで自動除外", "", "改定前の求人が残っているもの"),
    ("未確定(店舗非固有)", f'=COUNTIF({N},"未確定(店舗非固有)")', "ブランド共通レンジのため除外", "", "南柏店サイゼリヤと同型の事故を防ぐ"),
    ("未確定(施設一般値)", f'=COUNTIF({N},"未確定(施設一般値)")', "商業施設全体の値のため除外", "", ""),
    ("未確定(近隣欄からの抽出)", f'=COUNTIF({N},"未確定(近隣欄からの抽出)")', "近隣求人欄の金額のため除外", "", "南柏店の事故そのもののガード"),
    ("未確定(店舗非特定)", f'=COUNTIF({N},"未確定(店舗非特定)")', "どの店舗の金額か特定できず除外", "", ""),
    ("未確定(証拠不足)", f'=COUNTIF({N},"未確定(証拠不足)")', "証拠4点セットが揃わず除外", "", ""),
    ("取得不能", f'=COUNTIF({N},"取得不能")', "時給そのものを取得できず", "", ""),
]
ar = 5
for name, f, mean, judge, note in agg:
    chk.cell(ar, 1, name); chk.cell(ar, 2, f); chk.cell(ar, 3, mean)
    chk.cell(ar, 5, note).font = SMALL
    for ci in range(1, 6):
        chk.cell(ar, ci).border = BOX
    ar += 1
chk.cell(ar, 1, "合計(内訳の和)").font = Font(bold=True)
chk.cell(ar, 2, f"=SUM(B{6}:B{ar-1})")
chk.cell(ar, 3, "総レコード数と一致すること")
chk.cell(ar, 4, f'=IF(B{ar}=B5,"OK","NG")')
for ci in range(1, 6):
    chk.cell(ar, ci).border = BOX
SUM_ROW = ar

# ---- 取り違え再現テスト ----
tr = ar + 2
chk.cell(tr, 1, "店舗取り違え・旧掲載 再現テスト").font = Font(bold=True, size=12)
chk.cell(tr + 1, 1,
    "台帳と同じ判定式に、過去の事故と同じ形のデータを流し込みます。すべてFAILにならないことを確認してください。").font = SMALL
header(chk, tr + 2, ["テスト内容", "投入した時給", "投入した取得元区分", "判定結果", "期待結果／合否"])

TESTS = [
    ("①ページタイトルはイオンモール柏、金額は近隣求人欄のららぽーと柏の葉1,300円", 1300, "近隣求人欄", 1140),
    ("①-b ブランド共通レンジを店舗の通常時給として投入", 1240, "ブランド共通", 1057),
    ("①-c どの店舗の値か特定できないまま投入", 1042, "店舗非特定", 1057),
    ("②改定前の旧掲載（熊本1,034円に対し900円）", 900, "店舗固有求人", 1034),
    ("③商業施設全体の一般値を特定テナントの時給として投入", 1200, "施設一般", 1023),
    ("④店舗名・住所・求人ID・時給がすべて揃った正常データ", 1150, "店舗固有求人", 1030),
]
trow = tr + 3
for i, (desc, wage, src, mw_) in enumerate(TESTS):
    chk.cell(trow, 1, desc)
    chk.cell(trow, 2, wage)
    chk.cell(trow, 3, src)
    # 台帳とまったく同じ判定ロジック(証拠4点は充足している前提=最も通りやすい条件で試す)
    chk.cell(trow, 4,
        f'=IF(C{trow}="ブランド共通","未確定(店舗非固有)",'
        f'IF(C{trow}="施設一般","未確定(施設一般値)",'
        f'IF(C{trow}="近隣求人欄","未確定(近隣欄からの抽出)",'
        f'IF(C{trow}="店舗非特定","未確定(店舗非特定)",'
        f'IF(B{trow}<{mw_},"未確定(最低賃金割れ)","媒体掲載値")))))')
    last = (i == len(TESTS) - 1)
    if last:
        chk.cell(trow, 5, f'=IF(D{trow}="媒体掲載値","PASS(正常データは通る)","FAIL")')
    else:
        chk.cell(trow, 5, f'=IF(LEFT(D{trow},3)="未確定",\"PASS(採用されない)\",\"FAIL(採用されてしまう)\")')
    for ci in range(1, 6):
        c = chk.cell(trow, ci); c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=(ci == 1))
    trow += 1
TEST_FIRST, TEST_LAST = tr + 3, trow - 1

# ---- 最終警告数 ----
fr = trow + 2
chk.cell(fr, 1, "最終警告数").font = Font(bold=True, size=12)
chk.cell(fr, 2, f'=COUNTIF(E{TEST_FIRST}:E{TEST_LAST},"FAIL*")+IF(D{SUM_ROW}="OK",0,1)')
chk.cell(fr, 3, "0件でなければ確定版として使用しない")
chk.cell(fr, 4, f'=IF(B{fr}=0,"OK","NG")')
for ci in range(1, 5):
    chk.cell(fr, ci).border = BOX
chk.cell(fr + 2, 1,
    "※このシートが検証できるのは「抽出の正しさ」だけです。求人情報そのものが古い・誤掲載である可能性は"
    "この方法では検証できません。実際に支払われる時給を証明するには、店舗または採用本部への直接確認が必要です。").font = SMALL
chk.cell(fr + 3, 1,
    "※自社時給のみ社内データに基づく「実額確認済み」です。競合はすべて「公開求人の掲載値」であり、実額ではありません。").font = SMALL

out = "/home/user/newstaff_file/chosa/20260731_周辺時給調査(九州7店舗).xlsx"
wb.save(out)
print("SAVED:", out)
