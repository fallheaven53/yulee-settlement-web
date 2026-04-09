"""
정산 관리 — 데이터 매니저 (구글 시트 기반)
"""

COLUMNS = ["ID", "회차", "공연일", "지급대상", "지급구분", "금액",
           "정산상태", "지급요청일", "지급완료일", "은행명", "예금주",
           "계좌번호", "증빙", "비고"]

PAY_TYPES_DEFAULT = ["출연료", "행사진행인력", "음향 오퍼", "사회자", "기타"]
STATUSES = ["미지급", "지급예정", "지급요청중", "지급완료"]
BANKS = ["국민은행", "신한은행", "우리은행", "하나은행", "농협",
         "기업은행", "SC제일은행", "부산은행", "대구은행", "경남은행",
         "광주은행", "전북은행", "제주은행", "카카오뱅크", "케이뱅크",
         "토스뱅크", "새마을금고", "신협", "우체국"]


def parse_int(v):
    try:
        return int(float(str(v).replace(",", "").strip() or 0))
    except (ValueError, TypeError):
        return 0


def fmt_won(v):
    return f"{parse_int(v):,}원"


class SettlementManager:
    """정산 데이터 관리 (구글 시트 연동)"""

    def __init__(self, gsheet_sync=None):
        self.gsheet = gsheet_sync
        self.records: list = []
        self.budget: dict = {"연간예산": 0, "출연료배정": 0, "행사진행인력배정": 0}
        self.pay_types: list = list(PAY_TYPES_DEFAULT)
        self.load()

    def load(self):
        if not self.gsheet:
            return
        try:
            self.gsheet.download_all(self)
        except Exception as e:
            print(f"[구글시트 로드 실패] {e}")

    def save(self):
        if not self.gsheet:
            return
        try:
            self.gsheet.upload_all(self)
        except Exception as e:
            print(f"[구글시트 저장 실패] {e}")

    # ── CRUD ──
    def next_id(self):
        if not self.records:
            return 1
        return max((parse_int(r.get("ID", 0)) for r in self.records), default=0) + 1

    def add(self, rec):
        rec["ID"] = self.next_id()
        self.records.append(rec)
        self.save()

    def update(self, rec_id, new_rec):
        for i, r in enumerate(self.records):
            if str(r.get("ID")) == str(rec_id):
                new_rec["ID"] = r["ID"]
                self.records[i] = new_rec
                break
        self.save()

    def delete(self, rec_id):
        self.records = [r for r in self.records if str(r.get("ID")) != str(rec_id)]
        self.save()

    # ── 집계 ──
    def calc_summary(self, recs=None):
        if recs is None:
            recs = self.records
        total_amt = sum(parse_int(r.get("금액", 0)) for r in recs)
        unpaid = [r for r in recs if r.get("정산상태") == "미지급"]
        scheduled = [r for r in recs if r.get("정산상태") == "지급예정"]
        request = [r for r in recs if r.get("정산상태") == "지급요청중"]
        done = [r for r in recs if r.get("정산상태") == "지급완료"]
        return {
            "total_cnt": len(recs), "total_amt": total_amt,
            "unpaid_cnt": len(unpaid),
            "unpaid_amt": sum(parse_int(r.get("금액", 0)) for r in unpaid),
            "sched_cnt": len(scheduled),
            "sched_amt": sum(parse_int(r.get("금액", 0)) for r in scheduled),
            "req_cnt": len(request),
            "req_amt": sum(parse_int(r.get("금액", 0)) for r in request),
            "done_cnt": len(done),
            "done_amt": sum(parse_int(r.get("금액", 0)) for r in done),
        }

    def calc_by_paytype(self):
        result = {}
        for r in self.records:
            pt = r.get("지급구분", "") or ""
            amt = parse_int(r.get("금액", 0))
            if pt not in result:
                result[pt] = {"total": 0, "done": 0}
            result[pt]["total"] += amt
            if r.get("정산상태") == "지급완료":
                result[pt]["done"] += amt
        return result

    def calc_by_target(self):
        result = {}
        for r in self.records:
            t = r.get("지급대상", "") or ""
            if t not in result:
                result[t] = {"total": 0, "done": 0, "cnt": 0, "done_cnt": 0}
            amt = parse_int(r.get("금액", 0))
            result[t]["total"] += amt
            result[t]["cnt"] += 1
            if r.get("정산상태") == "지급완료":
                result[t]["done"] += amt
                result[t]["done_cnt"] += 1
        return result

    def calc_monthly(self):
        monthly = {i: 0 for i in range(1, 13)}
        for r in self.records:
            if r.get("정산상태") == "지급완료":
                d = str(r.get("지급완료일", ""))
                if d and len(d) >= 7:
                    try:
                        m = int(d[5:7])
                        if 1 <= m <= 12:
                            monthly[m] += parse_int(r.get("금액", 0))
                    except (ValueError, IndexError):
                        pass
        return monthly
