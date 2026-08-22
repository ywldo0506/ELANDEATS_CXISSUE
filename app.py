import os
import csv
import io
import time
import sqlite3
from datetime import datetime
from functools import wraps
from urllib.parse import quote

import requests
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_from_directory, flash, Response, jsonify
)
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 기본 설정
# ----------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Render에 배포할 때는 영구 디스크 경로(/data)를 쓰고,
# 로컬에서 테스트할 때는 프로젝트 폴더 안의 data 폴더를 씁니다.
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(BASE_DIR, "data"))
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
DB_PATH = os.path.join(DATA_DIR, "reports.db")

os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "heic"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key-please")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 사진 여러 장 업로드 대비 25MB

# 관리자 페이지 비밀번호 (Render 배포시 환경변수로 꼭 바꿔주세요!)
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "eland2026")

BRANDS = ["애슐리 퀸즈", "로운", "피자몰 뷔페", "피자몰 V2", "자연별곡", "리미니", "프랑제리", "반궁", "델리BY애슐리", "기타"]
INCIDENT_TYPES = ["이물질", "해충", "식중독·장염", "화상·낙상 등 안전사고", "시설파손", "기타"]
STATUS_LIST = ["접수", "처리중", "완료"]
TIME_SLOTS = ["평일 런치", "평일 디너", "주말 런치", "주말 디너"]

# ----------------------------------------------------------------------
# 매장코드-매장명 매칭용 구글시트 CSV 연동
# ----------------------------------------------------------------------
# Render 환경변수 STORE_SHEET_CSV_URL 에 "웹에 게시(Publish to web) CSV" 링크를 넣어주세요.
STORE_SHEET_CSV_URL = os.environ.get("STORE_SHEET_CSV_URL", "")

# 시트 헤더에 이런 이름들이 있으면 각각 브랜드/매장코드/매장명 컬럼으로 인식합니다.
# 시트의 실제 헤더가 다르면 이 리스트에 추가해주세요.
BRAND_COL_ALIASES = ["브랜드", "브랜드명"]
CODE_COL_ALIASES = ["매장코드", "점코드", "코드"]
NAME_COL_ALIASES = ["매장명", "지점명", "점명"]

_store_cache = {"data": [], "loaded_at": 0}
STORE_CACHE_TTL_SECONDS = 300  # 5분마다 시트 다시 읽기


def _find_col(fieldnames, aliases):
    if not fieldnames:
        return None
    normalized = {f.strip(): f for f in fieldnames if f}
    for alias in aliases:
        if alias in normalized:
            return normalized[alias]
    return None


def load_store_master(force=False):
    """구글시트 CSV에서 브랜드/매장코드/매장명 목록을 읽어옵니다.
    5분 이내 재요청이면 캐시된 값을 그대로 씁니다."""
    now = time.time()
    if not force and _store_cache["data"] and (now - _store_cache["loaded_at"] < STORE_CACHE_TTL_SECONDS):
        return _store_cache["data"], None

    if not STORE_SHEET_CSV_URL:
        return [], "STORE_SHEET_CSV_URL 환경변수가 설정되어 있지 않습니다."

    try:
        resp = requests.get(STORE_SHEET_CSV_URL, timeout=10)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        reader = csv.DictReader(io.StringIO(resp.text))

        brand_col = _find_col(reader.fieldnames, BRAND_COL_ALIASES)
        code_col = _find_col(reader.fieldnames, CODE_COL_ALIASES)
        name_col = _find_col(reader.fieldnames, NAME_COL_ALIASES)

        if not code_col or not name_col:
            return [], (
                f"시트 헤더에서 매장코드/매장명 컬럼을 찾지 못했습니다. "
                f"현재 헤더: {reader.fieldnames}"
            )

        rows = []
        for row in reader:
            code = (row.get(code_col) or "").strip()
            name = (row.get(name_col) or "").strip()
            brand = (row.get(brand_col) or "").strip() if brand_col else ""
            if code and name:
                rows.append({"brand": brand, "store_code": code, "store_name": name})

        _store_cache["data"] = rows
        _store_cache["loaded_at"] = now
