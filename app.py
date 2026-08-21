import os
import csv
import io
import time
import sqlite3
from datetime import datetime
from functools import wraps

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

BRANDS = ["애슐리퀸즈", "로운", "피자몰 뷔페", "피자몰V2", "자연별곡", "리미니", "프랑제리", "반궁", "기타"]
INCIDENT_TYPES = ["이물질", "식중독·장염", "화상·낙상 등 안전사고", "시설파손", "기타"]
STATUS_LIST = ["접수", "처리중", "완료"]

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
        return rows, None
    except Exception as e:
        # 시트 요청이 실패해도 이전 캐시가 있으면 그걸로 계속 서비스합니다.
        if _store_cache["data"]:
            return _store_cache["data"], f"최신 시트 로드 실패(이전 캐시 사용 중): {e}"
        return [], f"시트를 불러오지 못했습니다: {e}"


# ----------------------------------------------------------------------
# DB 준비
# ----------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            brand TEXT NOT NULL,
            store_code TEXT,
            store_name TEXT NOT NULL,
            incident_datetime TEXT NOT NULL,
            incident_type TEXT NOT NULL,
            description TEXT NOT NULL,
            customer_name TEXT,
            customer_phone TEXT,
            action_taken TEXT,
            reporter_name TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT '접수',
            photo_filenames TEXT,
            created_at TEXT NOT NULL
        )
    """)
    # 기존에 store_code 컬럼 없이 만들어진 DB가 있으면 추가해줍니다.
    try:
        conn.execute("ALTER TABLE reports ADD COLUMN store_code TEXT")
    except sqlite3.OperationalError:
        pass  # 이미 있으면 무시
    conn.commit()
    conn.close()


init_db()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ----------------------------------------------------------------------
# 관리자 로그인 보호
# ----------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login", next=request.path))
        return f(*args, **kwargs)
    return decorated


# ----------------------------------------------------------------------
# 매장코드 -> 매장명 자동완성 API (폼에서 JS로 호출)
# ----------------------------------------------------------------------
@app.route("/api/store-lookup")
def store_lookup():
    code = request.args.get("code", "").strip()
    stores, error = load_store_master()
    if not code:
        return jsonify({"found": False, "error": error})
    for s in stores:
        if s["store_code"] == code:
            return jsonify({"found": True, "store_name": s["store_name"], "brand": s["brand"]})
    return jsonify({"found": False, "error": error})


# ----------------------------------------------------------------------
# 매장용: 사고 제출 폼
# ----------------------------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def submit_report():
    if request.method == "POST":
        brand = request.form.get("brand", "").strip()
        store_code = request.form.get("store_code", "").strip()
        store_name = request.form.get("store_name", "").strip()
        incident_datetime = request.form.get("incident_datetime", "").strip()
        incident_type = request.form.get("incident_type", "").strip()
        description = request.form.get("description", "").strip()
        customer_name = request.form.get("customer_name", "").strip()
        customer_phone = request.form.get("customer_phone", "").strip()
        action_taken = request.form.get("action_taken", "").strip()
        reporter_name = request.form.get("reporter_name", "").strip()

        errors = []
        if not brand:
            errors.append("브랜드를 선택해주세요.")
        if not store_code:
            errors.append("매장코드를 입력해주세요.")
        if not store_name:
            errors.append("매장코드를 확인해서 매장명이 자동으로 표시되어야 합니다. 코드를 다시 확인해주세요.")
        if not incident_datetime:
            errors.append("사고 발생일시를 입력해주세요.")
        if not incident_type:
            errors.append("사고 유형을 선택해주세요.")
        if not description:
            errors.append("사고 내용을 입력해주세요.")
        if not reporter_name:
            errors.append("작성자명을 입력해주세요.")

        if errors:
            for e in errors:
                flash(e, "error")
            return render_template(
                "submit.html", brands=BRANDS, incident_types=INCIDENT_TYPES,
                form=request.form
            )

        # 사진 저장
        saved_filenames = []
        files = request.files.getlist("photos")
        for file in files:
            if file and file.filename and allowed_file(file.filename):
                ext = file.filename.rsplit(".", 1)[1].lower()
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
                safe_name = secure_filename(f"{timestamp}_{file.filename}")
                file.save(os.path.join(UPLOAD_DIR, safe_name))
                saved_filenames.append(safe_name)

        conn = get_db()
        conn.execute("""
            INSERT INTO reports (
                brand, store_code, store_name, incident_datetime, incident_type, description,
                customer_name, customer_phone, action_taken, reporter_name,
                status, photo_filenames, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            brand, store_code, store_name, incident_datetime, incident_type, description,
            customer_name, customer_phone, action_taken, reporter_name,
            "접수", ",".join(saved_filenames), datetime.now().isoformat()
        ))
        conn.commit()
        conn.close()

        return render_template("submit_success.html")

    return render_template("submit.html", brands=BRANDS, incident_types=INCIDENT_TYPES, form={})


# ----------------------------------------------------------------------
# 관리자: 로그인 / 로그아웃
# ----------------------------------------------------------------------
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == ADMIN_PASSWORD:
            session["is_admin"] = True
            next_url = request.args.get("next") or url_for("admin_list")
            return redirect(next_url)
        flash("비밀번호가 올바르지 않습니다.", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login"))


# ----------------------------------------------------------------------
# 관리자: 목록 + 필터
# ----------------------------------------------------------------------
@app.route("/admin")
@login_required
def admin_list():
    brand_filter = request.args.get("brand", "")
    type_filter = request.args.get("incident_type", "")
    status_filter = request.args.get("status", "")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    query = "SELECT * FROM reports WHERE 1=1"
    params = []
    if brand_filter:
        query += " AND brand = ?"
        params.append(brand_filter)
    if type_filter:
        query += " AND incident_type = ?"
        params.append(type_filter)
    if status_filter:
        query += " AND status = ?"
        params.append(status_filter)
    if date_from:
        query += " AND incident_datetime >= ?"
        params.append(date_from)
    if date_to:
        query += " AND incident_datetime <= ?"
        params.append(date_to + "T23:59")
    query += " ORDER BY incident_datetime DESC"

    conn = get_db()
    rows = conn.execute(query, params).fetchall()
    conn.close()

    return render_template(
        "admin_list.html", reports=rows, brands=BRANDS,
        incident_types=INCIDENT_TYPES, status_list=STATUS_LIST,
        filters={
            "brand": brand_filter, "incident_type": type_filter,
            "status": status_filter, "date_from": date_from, "date_to": date_to
        }
    )


# ----------------------------------------------------------------------
# 관리자: 상태 변경
# ----------------------------------------------------------------------
@app.route("/admin/status/<int:report_id>", methods=["POST"])
@login_required
def update_status(report_id):
    new_status = request.form.get("status")
    if new_status in STATUS_LIST:
        conn = get_db()
        conn.execute("UPDATE reports SET status = ? WHERE id = ?", (new_status, report_id))
        conn.commit()
        conn.close()
    return redirect(url_for("admin_list", **request.args))


# ----------------------------------------------------------------------
# 관리자: 업로드된 사진 보기
# ----------------------------------------------------------------------
@app.route("/admin/photo/<filename>")
@login_required
def view_photo(filename):
    return send_from_directory(UPLOAD_DIR, secure_filename(filename))


# ----------------------------------------------------------------------
# 관리자: 엑셀 다운로드 (현재 필터 조건 그대로 반영)
# ----------------------------------------------------------------------
@app.route("/admin/export")
@login_required
def export_excel():
    brand_filter = request.args.get("brand", "")
    type_filter = request.args.get("incident_type", "")
    status_filter = request.args.get("status", "")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    query = "SELECT * FROM reports WHERE 1=1"
    params = []
    if brand_filter:
        query += " AND brand = ?"
        params.append(brand_filter)
    if type_filter:
        query += " AND incident_type = ?"
        params.append(type_filter)
    if status_filter:
        query += " AND status = ?"
        params.append(status_filter)
    if date_from:
        query += " AND incident_datetime >= ?"
        params.append(date_from)
    if date_to:
        query += " AND incident_datetime <= ?"
        params.append(date_to + "T23:59")
    query += " ORDER BY incident_datetime DESC"

    conn = get_db()
    rows = conn.execute(query, params).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "사고보고"

    headers = [
        "번호", "브랜드", "매장코드", "매장명", "사고발생일시", "사고유형", "사고내용",
        "고객명", "고객연락처", "매장조치내용", "작성자", "처리상태",
        "사진개수", "접수시각"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1E2761", end_color="1E2761", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        photo_count = len([p for p in (row["photo_filenames"] or "").split(",") if p])
        ws.append([
            row["id"], row["brand"], row["store_code"], row["store_name"], row["incident_datetime"],
            row["incident_type"], row["description"], row["customer_name"],
            row["customer_phone"], row["action_taken"], row["reporter_name"],
            row["status"], photo_count, row["created_at"]
        ])

    # 열 너비 자동 조정 (대략)
    widths = [6, 12, 10, 14, 18, 18, 40, 12, 14, 30, 10, 10, 8, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"BK사고보고_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
