# -*- coding: utf-8 -*-
"""入社書類作成アプリ(Streamlit)。

スプレッドシート(新入社員一覧)を読み込み、一人ずつ
  入社のご案内 / 雇用契約書 / 入社辞令 (+ 転居者には社宅案内一式)
を Word/Excel と PDF で生成し、添付付きメール下書き(.eml)まで作成する。
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

import requests
import streamlit as st

from generator import defaults
from generator.data import load_people
from generator.pipeline import generate_person, person_dir, refresh_person

st.set_page_config(page_title="入社書類作成アプリ", page_icon="📄", layout="wide")

DEFAULT_OUTPUT = Path.home() / "Desktop" / "入社書類"


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
        raise ValueError(
            "ダウンロードできませんでした。共有設定が「リンクを知っている全員」に"
            "なっていない場合は、スプレッドシートを「ファイル→ダウンロード→"
            "Microsoft Excel(.xlsx)」で保存し、ファイルアップロードをご利用ください。")
    return r.content


def _set_sheet(data: bytes):
    """シートを読み込む。内容が変わったときは選択状態をリセットする。"""
    if st.session_state.get("sheet_bytes") != data:
        st.session_state["sheet_bytes"] = data
        st.session_state.pop("bulk_select", None)
        st.session_state.pop("people_editor", None)


# ---------------------------------------------------------------- 画面
st.title("📄 入社書類作成アプリ")
st.caption("スプレッドシートの新入社員一覧から、入社案内・雇用契約書・入社辞令(転居者には社宅案内も)と"
           "メール下書きを一括作成します。")

with st.sidebar:
    st.header("⚙️ 基本設定")
    out_dir_str = st.text_input("出力先フォルダ", value=str(DEFAULT_OUTPUT))
    out_dir = Path(out_dir_str).expanduser()
    saved = load_settings(out_dir)

    hakko_date = st.date_input("案内状の発行日(日付欄)", value=dt.date.today(),
                               format="YYYY/MM/DD")
    default_year = st.number_input(
        "入社日の年(シートに年の記載がない場合に使用)",
        value=dt.date.today().year, min_value=2020, max_value=2100, step=1)

    st.divider()
    st.subheader("会社情報")
    company = dict(defaults.COMPANY)
    company.update(saved.get("company", {}))
    company["担当者名"] = st.text_input("担当者名(案内状の差出人)", company["担当者名"])
    company["問い合わせ"] = st.text_input("問い合わせ先(メール署名)", company["問い合わせ"])
    company["社長名"] = st.text_input("代表者名(辞令)", company["社長名"])
    company["辞令配属先"] = st.text_input("辞令の配属先", company["辞令配属先"])

# --- 1. スプレッドシート読み込み -------------------------------------------
st.header("1️⃣ 新入社員一覧の読み込み")
tab_url, tab_file = st.tabs(["🔗 GoogleスプレッドシートURL", "📁 Excelファイル"])
with tab_url:
    sheet_url = st.text_input(
        "スプレッドシートのURL",
        value=saved.get("sheet_url", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...")
    if st.button("読み込む", type="primary", key="load_url") and sheet_url:
        try:
            _set_sheet(fetch_google_sheet(sheet_url))
            st.session_state["sheet_url"] = sheet_url
            st.success("読み込みました")
        except Exception as e:
            st.error(str(e))
with tab_file:
    up = st.file_uploader("スプレッドシートをxlsx形式でダウンロードしたファイル",
                          type=["xlsx"])
    if up is not None:
        _set_sheet(up.getvalue())

if "sheet_bytes" not in st.session_state:
    st.info("スプレッドシートを読み込むと、続きの手順が表示されます。")
    st.stop()

try:
    people = load_people(st.session_state["sheet_bytes"], int(default_year))
except Exception as e:
    st.error(f"一覧を読み取れませんでした: {e}")
    st.stop()
if not people:
    st.warning("新入社員の行が見つかりませんでした(2行目がヘッダー、3行目以降がデータの前提です)。")
    st.stop()

# --- 2. 対象者の選択 --------------------------------------------------------
st.header("2️⃣ 対象者の確認・選択")
st.caption("チェックが付いている方の分だけ作成します。**入社日が過ぎている方**と、"
           "シートで**グレーアウトされている方**(対応済み)は自動でチェックが外れます。"
           "「社宅案内」はシートの「転居」列から自動判定していますが、ここで変更できます。")

bc1, bc2, _ = st.columns([1, 1, 4])
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
        "役職": p.yakushoku,
        "等級": p.grade,
        "メールアドレス": p.email,
        "社宅案内": p.shataku,
        "状態": ("✅対応済み" if p.done else
                 "入社日経過" if p.nyusha_date and p.nyusha_date < today else ""),
        "注意": " / ".join(p.warnings),
    })
edited = st.data_editor(
    rows, hide_index=True, width="stretch", key="people_editor",
    disabled=["氏名", "入社日", "所属", "役職", "等級", "メールアドレス", "状態", "注意"],
    column_config={
        "作成": st.column_config.CheckboxColumn(help="書類とメールを作成する"),
        "社宅案内": st.column_config.CheckboxColumn(help="社宅案内とマニュアルを添付する"),
    })

selected = []
for p, row in zip(people, edited):
    p.shataku = bool(row["社宅案内"])
    if row["作成"]:
        if p.warnings:
            st.error(f"{p.name}: {' / '.join(p.warnings)} — 修正するまで作成できません")
        else:
            selected.append(p)

# --- 3. 入社回ごとの設定 ----------------------------------------------------
st.header("3️⃣ 入社式の情報(入社日ごと)")
st.caption("入社案内に差し込む情報です。日程はシートから自動計算した値が入っています。")

cohort_settings = saved.get("cohorts", {})
cohorts = {}
for d in sorted({p.nyusha_date for p in selected}):
    key = d.isoformat()
    base = defaults.cohort_defaults(d)
    ends = [p.tresen_end for p in selected if p.nyusha_date == d and p.tresen_end]
    if ends:
        base["チェックアウト日"] = defaults.fmt_md(max(ends))
    base.update(cohort_settings.get(key, {}))
    with st.expander(f"📅 {defaults.fmt_full_youbi(d)} 入社", expanded=False):
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

# --- 4. メール文面 ----------------------------------------------------------
st.header("4️⃣ メール文面")
with st.expander("✉️ 件名・本文のテンプレートを確認・編集する", expanded=False):
    st.caption("{氏名} {入社日} {入社日full} {添付一覧} {社宅段落} {会社名} {部署名} "
               "{担当者名} {問い合わせ} が差し込まれます。")
    subject_tpl = st.text_input(
        "件名", saved.get("mail_subject", defaults.MAIL_SUBJECT_DEFAULT))
    body_tpl = st.text_area(
        "本文", saved.get("mail_body", defaults.MAIL_BODY_DEFAULT), height=380)

# --- 5. 作成 ----------------------------------------------------------------
st.header("5️⃣ 作成")
col_a, col_b = st.columns([1, 2])
with col_a:
    run = st.button(f"🚀 {len(selected)}名分の書類とメール下書きを作成",
                    type="primary", disabled=not selected)
with col_b:
    st.caption(f"出力先: {out_dir}(入社日ごと・氏名ごとのフォルダに保存されます)")

if run:
    save_settings(out_dir, {
        "company": company,
        "cohorts": cohorts,
        "mail_subject": subject_tpl,
        "mail_body": body_tpl,
        "sheet_url": st.session_state.get("sheet_url", ""),
    })
    results, errors = [], []
    bar = st.progress(0.0)
    status = st.empty()
    for i, p in enumerate(selected):
        try:
            status.info(f"{p.name} さんの書類を作成中…")
            folder = generate_person(
                p, out_dir, cohorts[p.nyusha_date.isoformat()], company,
                subject_tpl, body_tpl, hakko_date=hakko_date)
            results.append((p, folder))
        except Exception as e:
            errors.append((p, str(e)))
        bar.progress((i + 1) / len(selected))
    status.empty()
    if results:
        st.success(f"✅ {len(results)}名分を作成しました")
        for p, folder in results:
            marker = "(社宅案内あり)" if p.shataku else ""
            st.write(f"- **{p.name}** {marker} → `{folder}`")
    for p, msg in errors:
        st.error(f"❌ {p.name}: {msg}")
    st.info("💡 メール送信: 各フォルダの「メール下書き.eml」をダブルクリックすると、"
            "宛先・件名・本文・添付が入った状態でメールソフトが開きます。"
            "Gmailをお使いの場合は「メール本文.txt」の内容を貼り付け、PDFを添付してください。")

# --- 6. 修正があったとき ----------------------------------------------------
st.header("6️⃣ 内容を修正したいとき")
st.caption("出力フォルダの中の Word / Excel を直接修正して保存 → 下のボタンを押すと、"
           "PDFへの変換とメール下書きの作り直しだけを行います(Word/Excelは上書きしません)。")

fixable = [p for p in selected if person_dir(out_dir, p).exists()]
if fixable:
    names = [f"{p.name}({defaults.fmt_md(p.nyusha_date)}入社)" for p in fixable]
    idx = st.selectbox("対象者", range(len(fixable)), format_func=lambda i: names[i])
    if st.button("♻️ この方のPDFとメール下書きを作り直す"):
        p = fixable[idx]
        try:
            with st.spinner("変換中…"):
                folder = refresh_person(p, out_dir, company, subject_tpl, body_tpl)
            st.success(f"✅ 作り直しました → `{folder}`")
        except Exception as e:
            st.error(f"❌ {e}")
else:
    st.caption("(まだ書類を作成した方がいません)")
