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

st.set_page_config(page_title="入社書類作成アプリ", page_icon="📄", layout="centered")

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


# ---------------------------------------------------------------- 画面
st.title("📄 入社書類作成アプリ")

# ===== 詳細設定(普段は触らない) =====
saved = load_settings(DEFAULT_OUTPUT)
with st.expander("⚙️ 詳細設定(普段は変更不要)", expanded=False):
    out_dir_str = st.text_input("出力先フォルダ", value=saved.get("out_dir", str(DEFAULT_OUTPUT)))
    out_dir = Path(out_dir_str).expanduser()
    if out_dir != DEFAULT_OUTPUT:
        saved = load_settings(out_dir) or saved
    hakko_date = st.date_input("案内状の発行日(日付欄)", value=dt.date.today(),
                               format="YYYY/MM/DD")
    default_year = st.number_input(
        "入社日の年(シートに年の記載がない場合に使用)",
        value=dt.date.today().year, min_value=2020, max_value=2100, step=1)
    sheet_url = st.text_input("新入社員一覧のURL", saved.get("sheet_url", DEFAULT_SHEET_URL))
    st.divider()
    company = dict(defaults.COMPANY)
    company.update(saved.get("company", {}))
    company["担当者名"] = st.text_input("担当者名(案内状の差出人)", company["担当者名"])
    company["問い合わせ"] = st.text_input("問い合わせ先", company["問い合わせ"])
    company["社長名"] = st.text_input("代表者名(辞令)", company["社長名"])
    company["辞令配属先"] = st.text_input("辞令の配属先", company["辞令配属先"])
    gmail_account = st.text_input(
        "Gmailの送信アカウント", saved.get("gmail_account", "kazuyuki.furukawa@sukesan.co.jp"))

# ===== 1. シートの自動読み込み =====
if "sheet_bytes" not in st.session_state and "sheet_error" not in st.session_state:
    with st.spinner("新入社員一覧(スプレッドシート)を読み込んでいます…"):
        try:
            _set_sheet(fetch_google_sheet(sheet_url))
        except Exception as e:
            st.session_state["sheet_error"] = str(e)

col_status, col_reload = st.columns([4, 1])
if "sheet_bytes" in st.session_state:
    col_status.success("✅ 新入社員一覧を読み込みました")
if col_reload.button("🔄 再読込"):
    st.session_state.pop("sheet_error", None)
    try:
        with st.spinner("読み込み中…"):
            data = fetch_google_sheet(sheet_url)
        st.session_state.pop("sheet_bytes", None)
        _set_sheet(data)
        st.rerun()
    except Exception as e:
        st.session_state["sheet_error"] = str(e)

if "sheet_bytes" not in st.session_state:
    st.error("スプレッドシートを自動で読み込めませんでした。"
             "下のどちらかの方法で読み込んでください。")
    up = st.file_uploader(
        "方法1: スプレッドシートを「ファイル → ダウンロード → Microsoft Excel(.xlsx)」で"
        "保存し、ここにドラッグ&ドロップ", type=["xlsx"])
    if up is not None:
        _set_sheet(up.getvalue())
        st.session_state.pop("sheet_error", None)
        st.rerun()
    st.caption("方法2: 上の「⚙️ 詳細設定」でURLを確認して「🔄 再読込」を押す")
    st.stop()

try:
    people = load_people(st.session_state["sheet_bytes"], int(default_year))
except Exception as e:
    st.error(f"一覧を読み取れませんでした: {e}")
    st.stop()
if not people:
    st.warning("新入社員の行が見つかりませんでした(2行目がヘッダー、3行目以降がデータの前提です)。")
    st.stop()

# ===== 2. 等級・月給チェック(毎回自動実行) =====
grades_ok, grade_lines = check_grades(people)
if grades_ok:
    with st.expander("✅ 等級・月給チェック: 全員一致(詳細を見る)", expanded=False):
        for line in grade_lines:
            st.write(line)
else:
    st.error("❌ 等級・月給チェックで不一致があります。作成前にシートを確認してください。")
    for line in grade_lines:
        st.write(line)

# ===== 3. 対象者 =====
st.header("作成する方を確認")
st.caption("チェックが付いている方の分だけ作成します。入社日が過ぎた方・シートでグレーアウト"
           "された方(対応済み)は自動でチェックが外れています。「社宅案内」は転居列から自動判定。")

bc1, bc2, _sp = st.columns([1, 1, 4])
if bc1.button("✅ 全員選択"):
    st.session_state["bulk_select"] = True
    st.session_state.pop("people_editor", None)
if bc2.button("⬜ 全員解除"):
    st.session_state["bulk_select"] = False
    st.session_state.pop("people_editor", None)
bulk = st.session_state.get("bulk_select")

today = dt.date.today()
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
                 "入社日経過" if p.nyusha_date and p.nyusha_date < today else ""),
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

# ===== 4. 入社式の情報(自動計算・変更したいときだけ開く) =====
cohort_settings = saved.get("cohorts", {})
cohorts = {}
for d in sorted({p.nyusha_date for p in selected}):
    key = d.isoformat()
    base = defaults.cohort_defaults(d)
    ends = [p.tresen_end for p in selected if p.nyusha_date == d and p.tresen_end]
    if ends:
        base["チェックアウト日"] = defaults.fmt_md(max(ends))
    base.update(cohort_settings.get(key, {}))
    with st.expander(f"📅 {defaults.fmt_full_youbi(d)}入社の入社式情報(ホテル等を変えたいときだけ開く)"):
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

# ===== 5. メール文面(変更したいときだけ開く) =====
with st.expander("✉️ メールの件名・本文を変えたいときだけ開く"):
    st.caption("{氏名} {入社日} {入社日full} {添付一覧} {社宅段落} などが自動で差し込まれます。")
    subject_tpl = st.text_input(
        "件名", saved.get("mail_subject_v4", defaults.MAIL_SUBJECT_DEFAULT))
    body_tpl = st.text_area(
        "本文", saved.get("mail_body_v4", defaults.MAIL_BODY_DEFAULT), height=380)
    shataku_text = st.text_area(
        "社宅対象の方にだけ入る段落(対象外の方では行ごと消えます)",
        saved.get("mail_shataku_v4", defaults.MAIL_SHATAKU_PARAGRAPH_DEFAULT),
        height=120)

# ===== 修正の反映漏れチェック(毎回自動) =====
if st.session_state.pop("flash", None):
    st.success("✅ 修正をすべて反映しました(PDFとメール下書きを作り直しました)")
stale_people = []
for p in people:
    if p.nyusha_date and person_dir(out_dir, p).exists():
        stale = find_stale_pdfs(person_dir(out_dir, p))
        if stale:
            stale_people.append((p, stale))
if stale_people:
    st.warning("⚠️ Word/Excelの修正がPDFにまだ反映されていません:\n"
               + "\n".join(f"- {p.name}: {'、'.join(s)}" for p, s in stale_people))
    if st.button("♻️ 反映されていない分をまとめて作り直す", type="primary"):
        for p, _ in stale_people:
            with st.spinner(f"{p.name} さんの分を作り直し中…"):
                refresh_person(p, out_dir, company, subject_tpl, body_tpl,
                               shataku_text=shataku_text)
        st.session_state["flash"] = True
        st.rerun()

# ===== 6. 作成ボタン =====
st.header("書類とメール下書きを作る")
run = st.button(f"🚀 {len(selected)}名分を作成する", type="primary",
                disabled=not selected, width="stretch")
st.caption(f"保存先: {out_dir}")

if run:
    save_settings(out_dir, {
        "out_dir": str(out_dir),
        "company": company,
        "cohorts": cohorts,
        "mail_subject_v4": subject_tpl,
        "mail_body_v4": body_tpl,
        "mail_shataku_v4": shataku_text,
        "sheet_url": sheet_url,
        "gmail_account": gmail_account,
    })
    results, errors = [], []
    bar = st.progress(0.0)
    status = st.empty()
    for i, p in enumerate(selected):
        try:
            status.info(f"{p.name} さんの書類を作成中…(しばらくお待ちください)")
            folder, notes = generate_person(
                p, out_dir, cohorts[p.nyusha_date.isoformat()], company,
                subject_tpl, body_tpl, hakko_date=hakko_date,
                shataku_text=shataku_text)
            results.append((p, folder, notes))
        except Exception as e:
            errors.append((p, str(e)))
        bar.progress((i + 1) / len(selected))
    status.empty()
    st.session_state["last_results"] = True
    if results:
        st.success(f"✅ {len(results)}名分を作成しました")
        for p, folder, notes in results:
            marker = "(社宅案内あり)" if p.shataku else ""
            st.write(f"- **{p.name}** {marker}")
            for n in notes:
                st.warning(f"{p.name}: {n}")
    for p, msg in errors:
        st.error(f"❌ {p.name}: {msg}")
    if results:
        st.info("次にやること: 下の「📧 Gmailで開く」を押すと、宛先・件名・本文入りの"
                "Gmail作成画面と書類フォルダが開きます。本文の■添付書類にある"
                "PDFをドラッグして添付し、送信してください。")

if st.button("📂 保存先フォルダを開く", width="stretch"):
    open_folder(out_dir)

# ===== メールを送る(Gmail) =====
st.header("メールを送る")
st.caption(f"Gmail({gmail_account})の作成画面が、宛先・件名・本文入りで開きます。"
           "添付だけはGmailの仕様で自動で付かないため、同時に開くフォルダから"
           "本文の「■添付書類」のPDFをドラッグしてください。")
mailable = [p for p in people if not p.warnings and p.nyusha_date
            and person_dir(out_dir, p).exists()]
if not mailable:
    st.caption("(まだ書類を作成した方がいません。作成するとここにボタンが並びます)")
for p in mailable:
    c1, c2 = st.columns([3, 2])
    c1.write(f"**{p.name}**({defaults.fmt_md(p.nyusha_date)}入社"
             f"{'・社宅あり' if p.shataku else ''})")
    if c2.button("📧 Gmailで開く", key=f"gmail{p.row}"):
        folder = person_dir(out_dir, p)
        subject, body, attachments = build_mail_content(
            folder, p, company, subject_tpl, body_tpl, shataku_text)
        webbrowser.open(gmail_compose_url(p.email, subject, body,
                                          account=gmail_account))
        open_folder(folder)
        st.success(f"Gmailとフォルダを開きました。添付するPDF: "
                   f"{'、'.join(a.name for a in attachments)}")

# ===== 7. ジョブカン案内メール =====
with st.expander("📮 ジョブカン案内メールの下書きを作る(入社日ごと・BCC自動設定)"):
    st.caption("宛先は人事の3名、BCCに同じ入社日の新入社員全員が自動で入ります。"
               "案内資料PDFの添付と日付の記入は送信前に手動で行ってください(本文冒頭に"
               "チェックリストが入ります)。")
    dates = sorted({p.nyusha_date for p in people
                    if p.nyusha_date and p.nyusha_date >= today and p.email})
    if not dates:
        st.write("(これから入社する方がいません)")
    for d in dates:
        members = [p.name for p in people if p.nyusha_date == d and p.email]
        st.write(f"**{defaults.fmt_md(d)}入社**({len(members)}名: {'、'.join(members)})")
        c1, c2, c3 = st.columns(3)
        try:
            if c1.button("📧 Gmailで1通目(事前案内)", key=f"jbg1{d}"):
                mails, bcc, _ = jobkan_mail_contents(people, d)
                webbrowser.open(gmail_compose_url(
                    ", ".join(defaults.JOBKAN_TO), mails[0][1], mails[0][2],
                    bcc=bcc, account=gmail_account))
                st.success(f"Gmailを開きました(BCC {len(bcc)}名入り)")
            if c2.button("📧 Gmailで2通目(リマインド)", key=f"jbg2{d}"):
                mails, bcc, _ = jobkan_mail_contents(people, d)
                webbrowser.open(gmail_compose_url(
                    ", ".join(defaults.JOBKAN_TO), mails[1][1], mails[1][2],
                    bcc=bcc, account=gmail_account))
                st.success(f"Gmailを開きました(BCC {len(bcc)}名入り)")
            if c3.button("📄 .emlで保存", key=f"jobkan{d}"):
                folder, outs, names = generate_jobkan_mails(people, out_dir, d)
                st.success(f"✅ 2通の下書きを保存しました → `{folder}`")
        except Exception as e:
            st.error(f"❌ {e}")

# ===== 8. 修正したいとき =====
with st.expander("♻️ 作った書類を直したいとき(PDFとメールの作り直し)"):
    st.caption("フォルダ内の Word / Excel を直接修正して保存 → 下のボタンで、PDFへの変換と"
               "メール下書きの作り直しだけを行います(修正したWord/Excelはそのまま残ります)。")
    fixable = [p for p in people if not p.warnings and p.nyusha_date
               and person_dir(out_dir, p).exists()]
    if fixable:
        names = [f"{p.name}({defaults.fmt_md(p.nyusha_date)}入社)" for p in fixable]
        idx = st.selectbox("対象者", range(len(fixable)), format_func=lambda i: names[i])
        if st.button("♻️ この方のPDFとメール下書きを作り直す"):
            p = fixable[idx]
            try:
                with st.spinner("変換中…"):
                    folder = refresh_person(p, out_dir, company, subject_tpl, body_tpl,
                                            shataku_text=shataku_text)
                st.success(f"✅ 作り直しました → `{folder}`")
            except Exception as e:
                st.error(f"❌ {e}")
    else:
        st.write("(まだ書類を作成した方がいません)")
