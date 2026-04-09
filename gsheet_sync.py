"""
정산 관리 — 구글 스프레드시트 동기화 모듈
"""

import os

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLUMNS = ["ID", "회차", "공연일", "지급대상", "지급구분", "금액",
           "정산상태", "지급요청일", "지급완료일", "은행명", "예금주",
           "계좌번호", "증빙", "비고"]


class SettlementSheetSync:
    """구글 스프레드시트 읽기/쓰기"""

    def __init__(self, credentials_path=None, credentials_dict=None,
                 spreadsheet_id=None):
        if not HAS_GSPREAD:
            raise ImportError("gspread 패키지가 필요합니다.")

        self.spreadsheet_id = spreadsheet_id

        if credentials_dict:
            creds = Credentials.from_service_account_info(
                credentials_dict, scopes=SCOPES)
        elif credentials_path and os.path.exists(credentials_path):
            creds = Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
        else:
            raise FileNotFoundError("구글 서비스 계정 인증 정보가 없습니다.")

        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)

    def _get_or_create_sheet(self, title, rows=500, cols=20):
        try:
            return self.spreadsheet.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            return self.spreadsheet.add_worksheet(
                title=title, rows=rows, cols=cols)

    def _clear_and_write(self, ws, data):
        ws.clear()
        if data:
            ws.update(data, value_input_option="RAW")

    # ═══════════════════════════════════════════
    #  업로드
    # ═══════════════════════════════════════════

    def upload_all(self, dm):
        # 1) 정산내역
        ws = self._get_or_create_sheet("정산내역")
        data = [COLUMNS]
        for rec in dm.records:
            data.append([rec.get(c, "") or "" for c in COLUMNS])
        self._clear_and_write(ws, data)

        # 2) 예산기준
        ws2 = self._get_or_create_sheet("예산기준")
        self._clear_and_write(ws2, [
            ["편성목", "연간예산", "출연료배정", "행사진행인력배정"],
            ["행사실비보상금",
             dm.budget.get("연간예산", 0),
             dm.budget.get("출연료배정", 0),
             dm.budget.get("행사진행인력배정", 0)],
        ])

        # 3) 지급구분
        ws3 = self._get_or_create_sheet("지급구분")
        pt_data = [["지급구분"]]
        for pt in dm.pay_types:
            pt_data.append([pt])
        self._clear_and_write(ws3, pt_data)

        self._cleanup_default_sheets()

    # ═══════════════════════════════════════════
    #  다운로드
    # ═══════════════════════════════════════════

    def download_all(self, dm):
        # 1) 정산내역
        try:
            ws = self.spreadsheet.worksheet("정산내역")
            rows = ws.get_all_values()
            dm.records = []
            if len(rows) > 1:
                headers = rows[0]
                for row in rows[1:]:
                    if all(v == "" for v in row):
                        continue
                    rec = {}
                    for i, h in enumerate(headers):
                        rec[h] = row[i] if i < len(row) else ""
                    # 금액을 정수로
                    try:
                        rec["금액"] = int(float(str(rec.get("금액", 0)).replace(",", "") or 0))
                    except (ValueError, TypeError):
                        rec["금액"] = 0
                    # ID를 정수로
                    try:
                        rec["ID"] = int(rec.get("ID", 0))
                    except (ValueError, TypeError):
                        pass
                    dm.records.append(rec)
        except gspread.exceptions.WorksheetNotFound:
            pass

        # 2) 예산기준
        try:
            ws = self.spreadsheet.worksheet("예산기준")
            rows = ws.get_all_values()
            for row in rows[1:]:
                if row and row[0] == "행사실비보상금":
                    dm.budget["연간예산"] = int(float(str(row[1]).replace(",", "") or 0))
                    dm.budget["출연료배정"] = int(float(str(row[2]).replace(",", "") or 0))
                    dm.budget["행사진행인력배정"] = int(float(str(row[3]).replace(",", "") or 0))
        except (gspread.exceptions.WorksheetNotFound, Exception):
            pass

        # 3) 지급구분
        try:
            ws = self.spreadsheet.worksheet("지급구분")
            rows = ws.get_all_values()
            pts = [row[0] for row in rows[1:] if row and row[0]]
            if pts:
                dm.pay_types = pts
        except gspread.exceptions.WorksheetNotFound:
            pass

    def _cleanup_default_sheets(self):
        try:
            sheets = self.spreadsheet.worksheets()
            if len(sheets) > 1:
                for s in sheets:
                    if s.title in ("Sheet1", "시트1"):
                        self.spreadsheet.del_worksheet(s)
                        break
        except Exception:
            pass
