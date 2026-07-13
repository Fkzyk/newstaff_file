#!/bin/bash
# 入社書類作成アプリ 起動用 (Mac)
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt
python3 -m streamlit run app.py
