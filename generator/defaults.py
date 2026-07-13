# -*- coding: utf-8 -*-
"""日付の整形、入社回(コホート)ごとの既定値、メール文面の既定テンプレート。"""
from __future__ import annotations

import datetime as dt

WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]


def fmt_full(d: dt.date) -> str:
    """2026年8月3日"""
    return f"{d.year}年{d.month}月{d.day}日"


def fmt_full_youbi(d: dt.date) -> str:
    """2026年8月3日（月）"""
    return f"{fmt_full(d)}（{WEEKDAYS[d.weekday()]}）"


def fmt_md(d: dt.date) -> str:
    """8月3日"""
    return f"{d.month}月{d.day}日"


def fmt_range(start: dt.date, end: dt.date) -> str:
    """8月4日～7日 / 9月29日～10月3日"""
    if start == end:
        return fmt_md(start)
    if start.month == end.month:
        return f"{fmt_md(start)}～{end.day}日"
    return f"{fmt_md(start)}～{fmt_md(end)}"


def fmt_reiwa(d: dt.date) -> str:
    """令和8年7月6日"""
    year = d.year - 2018
    return f"令和{'元' if year == 1 else year}年{d.month}月{d.day}日"


def cohort_defaults(nyusha_date: dt.date) -> dict:
    """入社日ごとに変わる設定の既定値。画面上で編集できる。"""
    return {
        "集合時間": "AM 8:10頃～8:30",
        "入社式場所": "北九州本社(福岡県北九州市小倉南区上葛原2-18-50)",
        "トレセン住所": "北九州市小倉北区神岳2丁目5-29",
        "ホテル名": "サンスカイホテル小倉",
        "ホテル住所": "福岡県北九州市小倉北区幸町 2-1",
        "チェックイン日": fmt_md(nyusha_date - dt.timedelta(days=1)),
        # チェックアウトは通常トレセン研修最終日。人ごとの研修終了日が
        # 取れる場合はアプリ側でその日付を初期値にする。
        "チェックアウト日": "",
    }


COMPANY = {
    "会社名": "株式会社資さん",
    "部署名": "人事採用課",
    "担当者名": "古川　和幸",
    "社長名": "崎田 晴義",
    "問い合わせ": "人事採用課　大野　093-932-4757",
    "辞令配属先": "営業本部",
}

MAIL_SUBJECT_DEFAULT = "【{会社名}】入社のご案内({入社日}入社)"

MAIL_BODY_DEFAULT = """{氏名} 様

いつもお世話になっております。
株式会社資さん 人事採用課の古川です。

入社のご案内を送付いたします。
お手数ですが、内容をご確認いただけますでしょうか。
{社宅段落}
もしご不明な点やご質問がございましたら、添付資料に記載の連絡先までお問い合わせください。
ご確認のほど、よろしくお願いいたします。

株式会社　資さん
人事採用課　採用担当
古川　和幸  Kazuyuki Furukawa
email：kazuyuki.furukawa@sukesan.co.jp
mobile：070-1444-6910
"""

# 社宅案内を添付する人にだけ本文に入る段落({社宅段落}に差し込まれる)
MAIL_SHATAKU_PARAGRAPH_DEFAULT = """また、ご転居を伴うご入社となりますので、「社宅利用申込のご案内」と「社宅システム入力マニュアル」も添付しております。
社宅のご利用をご希望の場合は、ご案内に記載の手順に沿って、お早めにお申し込みをお願いいたします。
※ご入居希望日は、入社日の10日前以降でご指定いただけます。"""
