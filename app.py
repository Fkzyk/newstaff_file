# -*- coding: utf-8 -*-
"""入社書類作成アプリ(Streamlit)。

デスクトップの「入社書類アプリ」アイコン(start.batが作成)から起動すると、
新入社員一覧のスプレッドシートを自動で読み込み、ボタン一つで
  入社のご案内 / 雇用契約書 / 入社辞令(原本等長置換) (+ 転居者には社宅案内一式)
を Word/Excel と PDF で生成し、添付付きメール下書き(.eml)まで作成する。
"""
from __future__ import annotations

import datetime as dt
import json
import os
import platform
import re
import subprocess
from pathlib import Path

import requests
import streamlit as st

from generator import defaults
from generator.data import load_people
import webbrowser

from generator.pipeline import (build_mail_content, check_grades,
                                find_stale_pdfs, generate_jobkan_mails,
                                generate_person, gmail_compose_url,
                                jobkan_mail_contents, person_dir,
                                refresh_person)

st.set_page_config(page_title="入社書類作成アプリ", page_icon="📄", layout="wide")

CUSTOM_CSS = """
<style>
:root { --brand:#3457d5; --brand-d:#274bb5; --ink:#1f2733; --muted:#5b6472;
        --line:#e6e8ee; --bg:#f4f6fa; --ok:#1a7f4b; --warn:#b26a00; }
.stApp { background: var(--bg); }
.block-container { max-width: 1080px; padding-top: 1.4rem; padding-bottom: 4rem; }
html, body, [class*="css"] { -webkit-font-smoothing: antialiased; }
h1,h2,h3 { color: var(--ink); letter-spacing:.01em; }
/* header banner */
.app-hero { background: linear-gradient(120deg,#3457d5,#5b7bf0); color:#fff;
  border-radius:18px; padding:22px 26px; margin-bottom:18px;
  box-shadow:0 6px 24px rgba(52,87,213,.18); }
.app-hero h1 { color:#fff; margin:0; font-size:1.6rem; }
.app-hero p { margin:.35rem 0 0; opacity:.92; font-size:.95rem; }
/* buttons */
.stButton>button, .stDownloadButton>button {
  border-radius:11px; font-weight:600; border:1px solid var(--line);
  padding:.5rem 1rem; transition:.12s ease; }
.stButton>button:hover { border-color:#c7cede; transform:translateY(-1px); }
.stButton>button[kind="primary"] {
  background:var(--brand); border-color:var(--brand); color:#fff;
  box-shadow:0 4px 14px rgba(52,87,213,.28); }
.stButton>button[kind="primary"]:hover { background:var(--brand-d); }
/* metric cards */
[data-testid="stMetric"] { background:#fff; border:1px solid var(--line);
  border-radius:14px; padding:14px 18px; box-shadow:0 1px 3px rgba(20,30,60,.04); }
[data-testid="stMetricLabel"] p { color:var(--muted); font-weight:600; }
/* tabs */
.stTabs [data-baseweb="tab-list"] { gap:4px; background:#eef1f7;
  padding:5px; border-radius:12px; }
.stTabs [data-baseweb="tab"] { font-weight:600; border-radius:9px;
  padding:9px 18px; color:var(--muted); }
.stTabs [aria-selected="true"] { background:#fff; color:var(--ink);
  box-shadow:0 1px 4px rgba(20,30,60,.08); }
/* alerts a touch softer */
[data-testid="stAlert"] { border-radius:12px; }
/* containers with border act as cards */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius:14px; }
section[data-testid="stSidebar"] { background:#fff; border-right:1px solid var(--line); }
</style>
"""

DEFAULT_OUTPUT = Path.home() / "Desktop" / "準備済_入社社員必要書類"
# 大元シート(新入社員一覧)。CLAUDE.md参照。
DEFAULT_SHEET_URL = ("https://docs.google.com/spreadsheets/d/"
                     "1v4w-kRFmN49vt5dQE5dj2zhoJJOPGvQUPOxm6L7tMdA/edit")


# ---------------------------------------------------------------- 設定の保存
def settings_path(out_dir: Path) -> Path:
    return out_dir / "アプリ設定.json"


def load_settings(out_dir: Path) -> dict:
    p = settings_path(out_dir)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_settings(out_dir: Path, settings: dict):
    out_dir.mkdir(parents=True, exist_ok=True)
    settings_path(out_dir).write_text(
        json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- 読み込み
def fetch_google_sheet(url: str) -> bytes:
    m = re.search(r"/spreadsheets/d/([\w-]+)", url)
    if not m:
        raise ValueError("スプレッドシートのURLではないようです")
    export = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    r = requests.get(export, timeout=30)
    if r.status_code != 200 or b"<html" in r.content[:200].lower():
        raise ValueError("スプレッドシートを自動で読み込めませんでした")
    return r.content


def _set_sheet(data: bytes):
    """シートを読み込む。内容が変わったときは選択状態をリセットする。"""
    if st.session_state.get("sheet_bytes") != data:
        st.session_state["sheet_bytes"] = data
        st.session_state.pop("bulk_select", None)
        st.session_state.pop("people_editor", None)


def open_folder(path: Path):
    """エクスプローラー等で出力フォルダを開く(アプリはPC上で動いている)。"""
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


from generator.pipeline import build_jobkan_reminders

# ---------------------------------------------------------------- 画面
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(
    '<div class="app-hero"><h1>📄 入社書類作成アプリ</h1>'
    '<p>新入社員の名簿から、入社書類とメールをまとめて作成します。</p></div>',
    unsafe_allow_html=True)

today = dt.date.today()

# ===== サイドバー: 設定(普段は触らない) =====
saved = load_settings(DEFAULT_OUTPUT)
with st.sidebar:
    st.markdown("### ⚙️ 設定")
    st.caption("普段は変更不要です。")
    with st.expander("保存先・読み込み"):
        out_dir_str = st.text_input("保存先フォルダ", value=saved.get("out_dir", str(DEFAULT_OUTPUT)))
        out_dir = Path(out_dir_str).expanduser()
        if out_dir != DEFAULT_OUTPUT:
            saved = load_settings(out_dir) or saved
        sheet_url = st.text_input("新入社員一覧のURL", saved.get("sheet_url", DEFAULT_SHEET_URL))
        default_year = st.number_input(
            "入社日の年(シートに年がない場合)",
            value=today.year, min_value=2020, max_value=2100, step=1)
        hakko_date = st.date_input("案内状の発行日(日付欄)", value=today, format="YYYY/MM/DD")
    with st.expander("会社情報・差出人"):
        company = dict(defaults.COMPANY)
        company.update(saved.get("company", {}))
        company["担当者名"] = st.text_input("担当者名(差出人)", company["担当者名"])
        company["問い合わせ"] = st.text_input("問い合わせ先", company["問い合わせ"])
        company["社長名"] = st.text_input("代表者名(辞令)", company["社長名"])
        company["辞令配属先"] = st.text_input("辞令の配属先", company["辞令配属先"])
        gmail_account = st.text_input(
            "Gmail送信アカウント",
            saved.get("gmail_account", "kazuyuki.furukawa@sukesan.co.jp"))
    with st.expander("メール文面"):
        st.caption("{氏名} {入社日} {添付一覧} {社宅段落} などが自動で差し込まれます。")
        subject_tpl = st.text_input(
            "件名", saved.get("mail_subject_v4", defaults.MAIL_SUBJECT_DEFAULT))
        body_tpl = st.text_area(
            "本文", saved.get("mail_body_v4", defaults.MAIL_BODY_DEFAULT), height=300)
        shataku_text = st.text_area(
            "社宅対象の方だけに入る段落",
            saved.get("mail_shataku_v4", defaults.MAIL_SHATAKU_PARAGRAPH_DEFAULT),
            height=110)
    st.divider()
    if st.button("📂 保存先フォルダを開く", width="stretch"):
        open_folder(out_dir)

# ===== シートの自動読み込み =====
if "sheet_bytes" not in st.session_state and "sheet_error" not in st.session_state:
    with st.spinner("新入社員一覧(スプレッドシート)を読み込んでいます…"):
        try:
            _set_sheet(fetch_google_sheet(sheet_url))
        except Exception as e:
            st.session_state["sheet_error"] = str(e)

if "sheet_bytes" not in st.session_state:
    st.error("スプレッドシートを自動で読み込めませんでした。下のどちらかで読み込んでください。")
    up = st.file_uploader(
        "① スプレッドシートを「ファイル → ダウンロード → Microsoft Excel(.xlsx)」で"
        "保存し、ここにドラッグ&ドロップ", type=["xlsx"])
    if up is not None:
        _set_sheet(up.getvalue())
        st.session_state.pop("sheet_error", None)
        st.rerun()
    if st.button("🔄 ② もう一度自動読み込みを試す"):
        st.session_state.pop("sheet_error", None)
        st.rerun()
    st.stop()

try:
    people = load_people(st.session_state["sheet_bytes"], int(default_year))
except Exception as e:
    st.error(f"一覧を読み取れませんでした: {e}")
    st.stop()
if not people:
    st.warning("新入社員の行が見つかりませんでした(2行目ヘッダー・3行目以降データの前提)。")
    st.stop()

grades_error = None
try:
    grades_ok, grade_lines = check_grades(people)
except Exception as e:
    grades_ok, grade_lines, grades_error = True, [], str(e)
future = [p for p in people if p.nyusha_date and p.nyusha_date >= today]

# ===== ステータス表示(ひと目で状態が分かる) =====
m1, m2, m3, m4 = st.columns(4)
m1.metric("名簿(読込済)", f"{len(people)}名")
m2.metric("等級・月給", "確認できず" if grades_error
          else "一致 ✅" if grades_ok else "不一致 ⚠️")
m3.metric("これから入社", f"{len(future)}名")
next_jobkan = None
for d in sorted({p.nyusha_date for p in people if p.nyusha_date and p.email}):
    for cand in (defaults.jobkan_announce_date(d), defaults.jobkan_auto_date(d)):
        if cand >= today and (next_jobkan is None or cand < next_jobkan):
            next_jobkan = cand
if next_jobkan:
    dleft = (next_jobkan - today).days
    m4.metric("次のジョブカン送信", f"{next_jobkan.month}/{next_jobkan.day}",
              "本日" if dleft == 0 else f"あと{dleft}日", delta_color="off")
else:
    m4.metric("次のジョブカン送信", "—")

if grades_error:
    st.warning(f"等級・月給チェックを実行できませんでした(給与テーブルをご確認ください): {grades_error}")
elif not grades_ok:
    st.error("❌ 等級・月給が不一致です。作成前にシートをご確認ください。")
with st.expander("等級・月給チェックの明細を見る"):
    for line in grade_lines:
        st.write(line)

# 反映漏れフラッシュ
if st.session_state.pop("flash", None):
    st.success("✅ 修正をすべて反映しました(PDFとメール下書きを作り直しました)")

# 修正の反映漏れ(全体アラート)
stale_people = []
for p in people:
    if p.nyusha_date and person_dir(out_dir, p).exists():
        stale = find_stale_pdfs(person_dir(out_dir, p))
        if stale:
            stale_people.append((p, stale))
if stale_people:
    with st.container(border=True):
        st.warning("⚠️ Word/Excelの修正がPDFにまだ反映されていません:\n"
                   + "\n".join(f"- {p.name}: {'、'.join(s)}" for p, s in stale_people))
        if st.button("♻️ 反映されていない分をまとめて作り直す", type="primary"):
            for p, _ in stale_people:
                with st.spinner(f"{p.name} さんの分を作り直し中…"):
                    refresh_person(p, out_dir, company, subject_tpl, body_tpl,
                                   shataku_text=shataku_text)
            st.session_state["flash"] = True
            st.rerun()

tab_docs, tab_mail, tab_jobkan = st.tabs(
    ["  ①  書類を作る  ", "  ②  入社案内メールを送る  ", "  ③  ジョブカン案内メール  "])

# =========================================================== ① 書類を作る
with tab_docs:
    st.subheader("作成する方を選ぶ")
    st.caption("チェックが付いた方だけ作成します。入社日が過ぎた方・対応済み(グレー)は"
               "自動で外れます。社宅案内は「転居」列から自動判定(ここで変更可)。")
    bc1, bc2, _sp = st.columns([1, 1, 3])
    if bc1.button("✅ 全員選択"):
        st.session_state["bulk_select"] = True
        st.session_state.pop("people_editor", None)
    if bc2.button("⬜ 全員解除"):
        st.session_state["bulk_select"] = False
        st.session_state.pop("people_editor", None)
    bulk = st.session_state.get("bulk_select")

    rows = []
    for p in people:
        default_make = (not p.warnings and not p.done
                        and p.nyusha_date is not None and p.nyusha_date >= today)
        rows.append({
            "作成": bulk if bulk is not None else default_make,
            "氏名": p.name,
            "入社日": defaults.fmt_full(p.nyusha_date) if p.nyusha_date else "?",
            "所属": p.shozoku_name,
            "等級": p.grade,
            "社宅案内": p.shataku,
            "状態": ("✅対応済み" if p.done else
                     "入社日経過" if p.nyusha_date and p.nyusha_date < today else "作成対象"),
            "注意": " / ".join(p.warnings),
        })
    edited = st.data_editor(
        rows, hide_index=True, width="stretch", key="people_editor",
        disabled=["氏名", "入社日", "所属", "等級", "状態", "注意"],
        column_config={
            "作成": st.column_config.CheckboxColumn(help="書類とメールを作成する"),
            "社宅案内": st.column_config.CheckboxColumn(help="社宅案内とマニュアルを添付する"),
        })

    selected = []
    for p, row in zip(people, edited):
        p.shataku = bool(row["社宅案内"])
        if row["作成"]:
            if p.warnings:
                st.error(f"{p.name}: {' / '.join(p.warnings)} — シートを修正するまで作成できません")
            else:
                selected.append(p)

    cohort_settings = saved.get("cohorts", {})
    cohorts = {}
    for d in sorted({p.nyusha_date for p in selected}):
        key = d.isoformat()
        base = defaults.cohort_defaults(d)
        ends = [p.tresen_end for p in selected if p.nyusha_date == d and p.tresen_end]
        if ends:
            base["チェックアウト日"] = defaults.fmt_md(max(ends))
        base.update(cohort_settings.get(key, {}))
        with st.expander(f"📅 {defaults.fmt_full_youbi(d)}入社の入社式情報(変えたいときだけ開く)"):
            c1, c2 = st.columns(2)
            base["集合時間"] = c1.text_input("集合時間", base["集合時間"], key=f"{key}集合")
            base["入社式場所"] = c2.text_input("入社式場所", base["入社式場所"], key=f"{key}場所")
            base["トレセン住所"] = c1.text_input("トレーニングセンター住所",
                                                 base["トレセン住所"], key=f"{key}トレセン")
            base["ホテル名"] = c2.text_input("ホテル名", base["ホテル名"], key=f"{key}ホテル")
            base["ホテル住所"] = c1.text_input("ホテル住所", base["ホテル住所"], key=f"{key}ホテル住所")
            base["チェックイン日"] = c2.text_input("チェックイン日", base["チェックイン日"],
                                                   key=f"{key}チェックイン")
            base["チェックアウト日"] = c1.text_input("チェックアウト日", base["チェックアウト日"],
                                                     key=f"{key}チェックアウト")
        cohorts[key] = base

    st.divider()
    st.markdown("#### 🚀 これ1つで全部")
    st.caption("押すと、選んだ全員について次を一気に実行します:")
    st.markdown(
        "1. 書類を作成(入社のご案内・雇用契約書・**入社辞令**・転居者は社宅一式)\n"
        "2. **入社案内メール**のGmail作成画面を全員分まとめて開く\n"
        "3. **ジョブカン案内メール**の下書き(PDF添付入り)を用意\n"
        "4. **ジョブカンの送信予定をGoogleカレンダーに登録**(忘れ防止)\n"
        "5. 保存先フォルダを開く")
    run = st.button(f"🚀 {len(selected)}名分をまとめて実行する",
                    type="primary", disabled=not selected, width="stretch")
    st.caption(f"保存先: {out_dir} / 送信元: {gmail_account}")

    if run:
        save_settings(out_dir, {
            "out_dir": str(out_dir), "company": company, "cohorts": cohorts,
            "mail_subject_v4": subject_tpl, "mail_body_v4": body_tpl,
            "mail_shataku_v4": shataku_text, "sheet_url": sheet_url,
            "gmail_account": gmail_account,
        })
        results, errors = [], []
        bar = st.progress(0.0)
        status = st.empty()
        for i, p in enumerate(selected):
            try:
                status.info(f"{p.name} さんの書類を作成中…")
                folder, notes = generate_person(
                    p, out_dir, cohorts[p.nyusha_date.isoformat()], company,
                    subject_tpl, body_tpl, hakko_date=hakko_date,
                    shataku_text=shataku_text)
                results.append((p, folder, notes))
            except Exception as e:
                errors.append((p, str(e)))
            bar.progress((i + 1) / len(selected))
        status.empty()
        # 2. 全員分の入社案内Gmailをまとめて開く(添付だけは各自ドラッグ)
        #    PDF(添付)が1つも無い人はメールを開かない(添付漏れ送信の防止)
        opened, mail_skipped = 0, []
        for p, folder, notes in results:
            try:
                subject, body, atts = build_mail_content(
                    folder, p, company, subject_tpl, body_tpl, shataku_text)
                if not atts:
                    mail_skipped.append(p.name)
                    continue
                webbrowser.open(gmail_compose_url(
                    p.email, subject, body, cc=defaults.ANNAI_CC,
                    account=gmail_account))
                opened += 1
            except Exception:
                mail_skipped.append(p.name)
        # 3. ジョブカン下書きを用意(これから入社する方の入社日ごと)
        from generator.pipeline import (generate_jobkan_mails,
                                        jobkan_calendar_events, gcal_event_url)
        jobkan_dates = sorted({p.nyusha_date for p in people
                               if p.nyusha_date and p.nyusha_date >= today and p.email})
        for d in jobkan_dates:
            try:
                generate_jobkan_mails(people, out_dir, d)
            except Exception:
                pass
        # 4. ジョブカンの送信予定をGoogleカレンダーに登録(忘れ防止)+ .ics保存
        cal_events = jobkan_calendar_events(people)
        cal_opened = 0
        for day, title, detail in cal_events[:8]:
            try:
                webbrowser.open(gcal_event_url(title, day, detail, account=gmail_account))
                cal_opened += 1
            except Exception:
                pass
        try:
            build_jobkan_reminders(people, out_dir)  # .icsも保存(バックアップ)
        except Exception:
            pass
        # 5. フォルダを開く
        if results:
            open_folder(out_dir)
            st.success(f"✅ 完了しました。書類 {len(results)}名分 / 入社案内Gmail {opened}件 / "
                       f"ジョブカン下書き {len(jobkan_dates)}件 / カレンダー予定 {cal_opened}件を開きました。")
            if mail_skipped:
                st.error("⚠️ PDFが作成できず、入社案内メールを開かなかった方: "
                         + "、".join(mail_skipped)
                         + "。Microsoft Officeを閉じてアプリを再起動し、"
                         "「作った書類を直したいとき」でPDFを作ってから送ってください"
                         "(添付漏れ防止のため保留しました)。")
            with st.container(border=True):
                st.markdown(f"**残りの手作業(送信元: {gmail_account})**")
                st.markdown(
                    "1. 開いた各**入社案内Gmail**に、フォルダから本文「■添付書類」のPDFを"
                    "ドラッグ&ドロップ → 送信\n"
                    "2. 開いた各**Googleカレンダー**の予定を「保存」(ジョブカン送信日の通知)\n"
                    "3. ジョブカン案内メールは**その予定日になったら**「③ジョブカン案内メール」タブから送信")
            for p, folder, notes in results:
                st.write(f"- **{p.name}**{'(社宅案内あり)' if p.shataku else ''} "
                         f"→ `{folder.name}`")
                for n in notes:
                    st.warning(f"{p.name}: {n}")
        for p, msg in errors:
            st.error(f"❌ {p.name}: {msg}")

    with st.expander("♻️ 作った書類を直したいとき(PDFとメールの作り直し)"):
        st.caption("フォルダ内の Word / Excel を直接修正して保存 → 下のボタンで、"
                   "PDF変換とメール作り直しだけを行います(修正内容はそのまま残ります)。")
        fixable = [p for p in people if not p.warnings and p.nyusha_date
                   and person_dir(out_dir, p).exists()]
        if fixable:
            names = [f"{p.name}({defaults.fmt_md(p.nyusha_date)}入社)" for p in fixable]
            idx = st.selectbox("対象者", range(len(fixable)), format_func=lambda i: names[i])
            if st.button("♻️ この方のPDFとメール下書きを作り直す"):
                p = fixable[idx]
                try:
                    with st.spinner("変換中…"):
                        folder = refresh_person(p, out_dir, company, subject_tpl,
                                                body_tpl, shataku_text=shataku_text)
                    st.success(f"✅ 作り直しました → `{folder}`")
                except Exception as e:
                    st.error(f"❌ {e}")
        else:
            st.write("(まだ書類を作成した方がいません)")

# =========================================================== ② メールを送る
with tab_mail:
    st.subheader("入社案内メールを送る")
    st.caption(f"To=本人・CC=人事2名で、Gmail({gmail_account})の作成画面が"
               "宛先・件名・本文入りで開きます。添付はGmailの仕様で自動では付かないため、"
               "保存先フォルダから本文の「■添付書類」のPDFをドラッグしてください。"
               "(「①書類を作る」の一括ボタンでも全員分のGmailが開きます)")
    mailable = [p for p in people if not p.warnings and p.nyusha_date
                and person_dir(out_dir, p).exists()]
    if not mailable:
        st.info("まだ書類を作成した方がいません。「① 書類を作る」で作成すると、"
                "ここに送信ボタンが並びます。")
    if mailable:
        if st.button(f"📧 {len(mailable)}名分のGmailをまとめて開く",
                     type="primary", width="stretch"):
            opened, skipped = 0, []
            try:
                for p in mailable:
                    folder = person_dir(out_dir, p)
                    subject, body, atts = build_mail_content(
                        folder, p, company, subject_tpl, body_tpl, shataku_text)
                    if not atts:
                        skipped.append(p.name)
                        continue
                    webbrowser.open(gmail_compose_url(p.email, subject, body,
                                                      cc=defaults.ANNAI_CC,
                                                      account=gmail_account))
                    opened += 1
                open_folder(out_dir)
                st.success(f"Gmailを{opened}件開きました。保存先フォルダから"
                           "PDFをドラッグして添付し、送信してください。")
                if skipped:
                    st.error("⚠️ PDFが無いためメールを開かなかった方: "
                             + "、".join(skipped) + "(先にPDFを作成してください)")
            except (KeyError, ValueError, IndexError) as e:
                st.error("メール文面のテンプレートに誤りがあります(サイドバーの"
                         f"「メール文面」をご確認ください)。詳細: {e}")
        st.caption("個別に開き直したいときは下のボタンをどうぞ。")
    for p in mailable:
        with st.container(border=True):
            c1, c2 = st.columns([3, 2])
            c1.markdown(f"**{p.name}**　<span style='color:#5b6472'>"
                        f"{defaults.fmt_md(p.nyusha_date)}入社"
                        f"{'・社宅あり' if p.shataku else ''}</span>",
                        unsafe_allow_html=True)
            c1.caption(f"To: {p.email} / CC: {'、'.join(defaults.ANNAI_CC)}")
            if c2.button("📧 Gmailで開く", key=f"gmail{p.row}", type="primary",
                         width="stretch"):
                try:
                    folder = person_dir(out_dir, p)
                    subject, body, attachments = build_mail_content(
                        folder, p, company, subject_tpl, body_tpl, shataku_text)
                    if not attachments:
                        c2.error("PDFがまだありません。先にPDFを作成してください。")
                    else:
                        webbrowser.open(gmail_compose_url(p.email, subject, body,
                                                          cc=defaults.ANNAI_CC,
                                                          account=gmail_account))
                        open_folder(folder)
                        c2.success("Gmailとフォルダを開きました。")
                        c2.caption("添付: " + "、".join(a.name for a in attachments))
                except (KeyError, ValueError, IndexError) as e:
                    c2.error(f"メール文面のテンプレートに誤りがあります。詳細: {e}")

# =========================================================== ③ ジョブカン
with tab_jobkan:
    upcoming = []
    for d in sorted({p.nyusha_date for p in people if p.nyusha_date and p.email}):
        ann, auto = defaults.jobkan_announce_date(d), defaults.jobkan_auto_date(d)
        members = [p.name for p in people if p.nyusha_date == d and p.email]
        if ann >= today:
            upcoming.append((ann, "事前案内メール", d, members))
        if auto >= today:
            upcoming.append((auto, "登録日リマインドメール", d, members))
    upcoming.sort()

    st.subheader("⏰ 送信予定(忘れ防止)")
    st.caption("入社月の前月25日(土日は前倒し)にジョブカンが登録メールを自動送信します。"
               "その1週間前に事前案内、当日に登録日リマインドを送る想定です。")
    if upcoming:
        n = upcoming[0]
        days = (n[0] - today).days
        when = "本日" if days == 0 else f"あと{days}日"
        st.info(f"📨 次の送信: **{n[0]:%Y/%m/%d}({defaults.WEEKDAYS[n[0].weekday()]})** "
                f"— {defaults.fmt_md(n[2])}入社の{n[1]}({when})")
        with st.container(border=True):
            for d0, kind, nd, members in upcoming:
                st.write(f"- **{d0:%Y/%m/%d}({defaults.WEEKDAYS[d0.weekday()]})** "
                         f"{kind} — {defaults.fmt_md(nd)}入社({'、'.join(members)})")
            if st.button("📅 送信予定をカレンダー(Outlook/Google)に登録"):
                ics, _lines = build_jobkan_reminders(people, out_dir)
                open_folder(ics.parent)
                st.success(f"✅ 「{ics.name}」を作りました。ダブルクリックすると"
                           "カレンダーに予定と当日通知が登録されます。")
    else:
        st.write("(これから入社する方がいません)")

    st.divider()
    st.subheader("📮 メールを作る")
    st.caption("To=古川さん本人 / CC=人事3名 / BCC=同じ入社日の新入社員全員(自動)。"
               "日付は自動記入、案内PDFも自動添付。PDFの差し替えは "
               "templates/jobkan_workflow.pdf を入れ替え。")
    dates = sorted({p.nyusha_date for p in people
                    if p.nyusha_date and p.nyusha_date >= today and p.email})
    if not dates:
        st.write("(これから入社する方がいません)")
    for d in dates:
        members = [p.name for p in people if p.nyusha_date == d and p.email]
        with st.container(border=True):
            st.markdown(f"**{defaults.fmt_md(d)}入社**（{len(members)}名: {'、'.join(members)}）")
            st.caption(f"事前案内 {defaults.fmt_md(defaults.jobkan_announce_date(d))} ／ "
                       f"登録日 {defaults.fmt_md(defaults.jobkan_auto_date(d))}")
            c1, c2, c3 = st.columns(3)
            try:
                if c1.button("📧 事前案内をGmailで", key=f"jbg1{d}", width="stretch"):
                    mails, bcc, _ = jobkan_mail_contents(people, d)
                    webbrowser.open(gmail_compose_url(
                        ", ".join(defaults.JOBKAN_TO), mails[0][1], mails[0][2],
                        cc=defaults.JOBKAN_CC, bcc=bcc, account=gmail_account))
                    folder, _o, _n = generate_jobkan_mails(people, out_dir, d)
                    open_folder(folder)
                    st.success(f"Gmailとフォルダを開きました(CC人事3名・BCC {len(bcc)}名)。")
                if c2.button("📧 登録日をGmailで", key=f"jbg2{d}", width="stretch"):
                    mails, bcc, _ = jobkan_mail_contents(people, d)
                    webbrowser.open(gmail_compose_url(
                        ", ".join(defaults.JOBKAN_TO), mails[1][1], mails[1][2],
                        cc=defaults.JOBKAN_CC, bcc=bcc, account=gmail_account))
                    folder, _o, _n = generate_jobkan_mails(people, out_dir, d)
                    open_folder(folder)
                    st.success(f"Gmailとフォルダを開きました(CC人事3名・BCC {len(bcc)}名)。")
                if c3.button("📄 .emlで保存", key=f"jobkan{d}", width="stretch"):
                    folder, outs, names = generate_jobkan_mails(people, out_dir, d)
                    st.success(f"✅ 2通の下書き(PDF添付入り)を保存 → `{folder}`")
            except Exception as e:
                st.error(f"❌ {e}")
