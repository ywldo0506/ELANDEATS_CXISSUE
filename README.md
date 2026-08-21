# 매장 BK(사고) 보고 TOOL

매장에서 사고(이물질/식중독/안전사고 등)를 웹 폼으로 입력하면,
CX팀은 관리자 페이지에서 목록으로 보고 엑셀로 다운로드할 수 있는 도구입니다.

## 화면 구성
- `/` : 매장에서 사고 보고를 입력하는 화면 (로그인 필요 없음, 매장 누구나 접속 가능)
- `/admin` : CX팀만 보는 관리자 화면 (비밀번호 필요, 목록 확인 + 필터 + 엑셀 다운로드)

---

## Render에 배포하는 방법 (기존 cx-portal과 완전히 똑같은 방식이에요)

1. 이 폴더를 GitHub 새 저장소로 올려주세요 (예: `bk-report-tool`)
2. Render 대시보드에서 **New + → Web Service** 선택
3. 방금 만든 GitHub 저장소 연결
4. 아래처럼 설정
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. **환경변수(Environment Variables)** 추가 (아주 중요해요!)
   - `ADMIN_PASSWORD` = CX팀만 아는 비밀번호로 바꿔주세요 (기본값 `eland2026`은 예시일 뿐이니 꼭 바꿔주세요)
   - `SECRET_KEY` = 아무 랜덤 문자열 (예: `eland-cx-2026-secret`)
   - `DATA_DIR` = `/data` (아래 6번의 디스크 경로와 맞춰야 해요)
   - `STORE_SHEET_CSV_URL` = 매장코드/매장명이 들어있는 구글시트의 "웹에 게시 CSV" 링크
     (시트 첫 줄 헤더가 `브랜드`, `매장코드`, `매장명` 이어야 자동으로 인식돼요.
     다른 이름을 쓰신다면 `app.py` 상단의 `BRAND_COL_ALIASES` / `CODE_COL_ALIASES` / `NAME_COL_ALIASES`
     목록에 그 이름을 추가해주세요.)
6. **Disks** 탭에서 영구 디스크(Persistent Disk) 추가
   - Mount Path: `/data`
   - 이렇게 해야 서버가 재시작돼도 사진이랑 데이터가 사라지지 않아요 (cx-portal에서 하셨던 것과 동일)
7. 배포 완료되면 매장에는 `/` 주소를, CX팀 관리자는 `/admin` 주소를 안내해주시면 됩니다.

## 나중에 CX portal에 합칠 때
- 이 앱의 `templates`, `static/style.css`, DB 구조를 그대로 cx-portal의 Flask 앱에
  블루프린트(blueprint)나 라우트로 옮겨 붙이면 됩니다. 구조가 동일한 Flask라서
  큰 변경 없이 합칠 수 있어요.

## 로컬에서 미리 확인해보고 싶을 때
```bash
pip install -r requirements.txt
python app.py
```
브라우저에서 `http://localhost:5000` 접속하면 폼 화면이 보입니다.
