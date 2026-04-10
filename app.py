"""
율이공방 — 토요상설공연 정산 관리 (Streamlit 웹앱)
구글 스프레드시트 연동 · 비밀번호 보호
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
from io import BytesIO

from data_manager import (
    SettlementManager, COLUMNS, PAY_TYPES_DEFAULT,
    STATUSES, BANKS, parse_int, fmt_won,
)

# ── 페이지 설정 ──
st.set_page_config(
    page_title="토요상설공연 정산 관리",
    page_icon="💰",
    layout="wide",
)

st.markdown("""
<style>
    .main .block-container { max-width: 1200px; padding-top: 1rem; }
    .metric-card {
        background: #313244; border-radius: 8px; padding: 14px;
        text-align: center; margin: 4px;
    }
    .metric-value { font-size: 1.6rem; font-weight: bold; color: #cdd6f4; }
    .metric-label { font-size: 0.8rem; color: #a6adc8; }
    .status-done { color: #4CAF50; }
    .status-req { color: #E65100; }
    .status-unpaid { color: #C62828; }
    .status-sched { color: #1565C0; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  비밀번호 보호
# ══════════════════════════════════════════════════════════════

def check_password():
    if "app_password" not in st.secrets:
        return True
    if st.session_state.get("authenticated"):
        return True
    pwd = st.text_input("비밀번호를 입력하세요", type="password", key="pwd_input")
    if st.button("로그인", key="btn_login"):
        if pwd == st.secrets["app_password"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False


# ══════════════════════════════════════════════════════════════
#  데이터 연결
# ══════════════════════════════════════════════════════════════

@st.cache_resource
def get_dm():
    gsheet = None
    try:
        from gsheet_sync import SettlementSheetSync
        if "gcp_service_account" in st.secrets:
            gsheet = SettlementSheetSync(
                credentials_dict=dict(st.secrets["gcp_service_account"]),
                spreadsheet_id=st.secrets["spreadsheet_id"],
            )
    except Exception as e:
        st.sidebar.warning(f"구글 시트 연결 실패: {e}")
    return SettlementManager(gsheet_sync=gsheet)


def reload_dm():
    get_dm.clear()
    st.rerun()


def load_target_dates():
    """출연단체 DB에서 당해 연도 단체명 → [(회차, 공연일)] 매핑"""
    if "target_dates" in st.session_state:
        return st.session_state["target_dates"]
    result = {}
    target_list = []
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if "gcp_service_account" not in st.secrets:
            return result
        if "performer_spreadsheet_id" not in st.secrets or not st.secrets["performer_spreadsheet_id"]:
            return result
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets"])
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["performer_spreadsheet_id"])

        cur_year = str(datetime.now().year)
        # 단체ID → 단체명
        ws1 = sh.worksheet("단체정보")
        rows = ws1.get_all_values()
        id_to_name = {}
        for row in rows[1:]:
            if row[0] and row[1]:
                id_to_name[row[0].strip()] = row[1].strip()

        # 출연이력
        ws2 = sh.worksheet("출연이력")
        rows2 = ws2.get_all_values()
        for row in rows2[1:]:
            if len(row) > 4 and row[2].strip() == cur_year:
                tid = row[1].strip()
                name = id_to_name.get(tid, "")
                rnd = row[3].strip()
                date = row[4].strip()
                if name and date:
                    result.setdefault(name, []).append((rnd, date))
                    if name not in target_list:
                        target_list.append(name)
    except Exception:
        pass
    st.session_state["target_dates"] = result
    st.session_state["target_list_db"] = sorted(target_list)
    return result


# ══════════════════════════════════════════════════════════════
#  탭 1: 정산 등록·관리
# ══════════════════════════════════════════════════════════════

def render_tab_records():
    dm = get_dm()
    target_dates = load_target_dates()
    db_targets = st.session_state.get("target_list_db", [])
    existing_targets = sorted({r.get("지급대상", "") for r in dm.records if r.get("지급대상", "")})
    all_targets = sorted(set(db_targets + existing_targets))

    # ── 입력 폼 ──
    if st.session_state.get("edit_mode"):
        _render_edit_form(dm, all_targets, target_dates)
    else:
        _render_add_form(dm, all_targets, target_dates)

    # ── 필터 바 ──
    st.markdown("---")
    fc1, fc2, fc3, fc4, fc5 = st.columns([1, 1, 1, 1, 1])
    with fc1:
        flt_status = st.selectbox("정산상태", ["전체"] + STATUSES, key="flt_status")
    with fc2:
        flt_pt = st.selectbox("지급구분", ["전체"] + dm.pay_types, key="flt_pt")
    with fc3:
        rnds = sorted({str(r.get("회차", "")) for r in dm.records if r.get("회차", "")})
        flt_rnd = st.selectbox("회차", ["전체"] + rnds, key="flt_rnd")
    with fc4:
        flt_month = st.selectbox("월", ["전체"] + [f"{m}월" for m in range(1, 13)], key="flt_month")
    with fc5:
        flt_unpaid = st.checkbox("미지급 건만", key="flt_unpaid")

    # 필터 적용
    recs = list(dm.records)
    if flt_status != "전체":
        recs = [r for r in recs if r.get("정산상태") == flt_status]
    if flt_pt != "전체":
        recs = [r for r in recs if r.get("지급구분") == flt_pt]
    if flt_rnd != "전체":
        recs = [r for r in recs if str(r.get("회차", "")) == flt_rnd]
    if flt_month != "전체":
        m = flt_month.replace("월", "")
        recs = [r for r in recs if str(r.get("공연일", ""))[5:7].lstrip("0") == m]
    if flt_unpaid:
        recs = [r for r in recs if r.get("정산상태") in ("미지급", "지급예정")]

    # ── 정산 목록 테이블 ──
    if recs:
        table_data = []
        for r in recs:
            table_data.append({
                "ID": r.get("ID", ""),
                "회차": r.get("회차", ""),
                "공연일": r.get("공연일", ""),
                "지급대상": r.get("지급대상", ""),
                "지급구분": r.get("지급구분", ""),
                "금액": f"{parse_int(r.get('금액', 0)):,}",
                "정산상태": r.get("정산상태", ""),
                "지급완료일": r.get("지급완료일", ""),
                "비고": r.get("비고", ""),
            })
        df = pd.DataFrame(table_data)

        # 행 선택
        sel = st.dataframe(
            df, use_container_width=True, hide_index=True,
            on_select="rerun", selection_mode="single-row",
            key="rec_table"
        )

        # 선택된 행 → 수정/삭제 버튼
        selected_rows = sel.get("selection", {}).get("rows", [])
        if selected_rows:
            sel_idx = selected_rows[0]
            if sel_idx >= len(recs):
                st.rerun()
                return
            sel_rec = recs[sel_idx]
            bc1, bc2 = st.columns(2)
            with bc1:
                if st.button("수정", key="btn_edit", use_container_width=True):
                    st.session_state["edit_mode"] = True
                    st.session_state["edit_id"] = sel_rec.get("ID")
                    st.rerun()
            with bc2:
                if st.button("삭제", key="btn_del", use_container_width=True, type="primary"):
                    dm.delete(sel_rec.get("ID"))
                    st.success("삭제 완료!")
                    st.rerun()
    else:
        st.info("해당 조건의 정산 내역이 없습니다.")

    # ── 상태바 ──
    s = dm.calc_summary(recs)
    st.markdown(
        f"**표시: {s['total_cnt']}건 · {s['total_amt']:,}원** | "
        f"<span class='status-unpaid'>미지급 {s['unpaid_cnt']}건 {s['unpaid_amt']:,}원</span> | "
        f"<span class='status-sched'>지급예정 {s['sched_cnt']}건 {s['sched_amt']:,}원</span> | "
        f"<span class='status-req'>지급요청중 {s['req_cnt']}건 {s['req_amt']:,}원</span> | "
        f"<span class='status-done'>지급완료 {s['done_cnt']}건 {s['done_amt']:,}원</span>",
        unsafe_allow_html=True
    )

    # ── 내보내기 ──
    with st.expander("엑셀 내보내기"):
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            if st.button("전체 내보내기", key="exp_full", use_container_width=True):
                _export_excel(dm.records, "정산_전체", include_account=True)
        with ec2:
            if st.button("계좌 제외 내보내기", key="exp_noacct", use_container_width=True):
                _export_excel(dm.records, "정산_계좌제외", include_account=False)
        with ec3:
            unpaid_recs = [r for r in dm.records
                           if r.get("정산상태") in ("미지급", "지급예정", "지급요청중")]
            if st.button("미지급 내보내기", key="exp_unpaid", use_container_width=True):
                _export_excel(unpaid_recs, "정산_미지급", include_account=True)


def _render_add_form(dm, all_targets, target_dates):
    with st.expander("정산 등록", expanded=True):
        with st.form("add_form"):
            r1c1, r1c2, r1c3 = st.columns([1, 1, 2])
            with r1c1:
                rnd = st.selectbox("회차", [str(i) for i in range(1, 25)], key="add_rnd")
            with r1c2:
                perf_date = st.text_input("공연일", key="add_date", placeholder="YYYY-MM-DD")
            with r1c3:
                target = st.selectbox("지급대상", [""] + all_targets, key="add_target")

            r2c1, r2c2, r2c3 = st.columns(3)
            with r2c1:
                paytype = st.selectbox("지급구분", dm.pay_types, key="add_pt")
            with r2c2:
                amt = st.text_input("금액", key="add_amt", placeholder="숫자")
            with r2c3:
                status = st.selectbox("정산상태", STATUSES, key="add_status")

            r3c1, r3c2, r3c3 = st.columns(3)
            with r3c1:
                req_date = st.text_input("지급요청일", key="add_reqdate", placeholder="YYYY-MM-DD")
            with r3c2:
                done_date = st.text_input("지급완료일", key="add_donedate", placeholder="YYYY-MM-DD")
            with r3c3:
                bank = st.selectbox("은행명", [""] + BANKS, key="add_bank")

            r4c1, r4c2, r4c3 = st.columns(3)
            with r4c1:
                holder = st.text_input("예금주", key="add_holder")
            with r4c2:
                acct = st.text_input("계좌번호", key="add_acct")
            with r4c3:
                evid = st.text_input("증빙", key="add_evid")

            note = st.text_input("비고", key="add_note")

            if st.form_submit_button("등록", use_container_width=True, type="primary"):
                if not target:
                    st.warning("지급대상을 선택하세요.")
                elif not amt:
                    st.warning("금액을 입력하세요.")
                else:
                    # 공연일 자동 매칭
                    final_date = perf_date
                    if not final_date and target:
                        entries = target_dates.get(target, [])
                        for r, d in entries:
                            if r == rnd:
                                final_date = d
                                break
                        if not final_date and entries:
                            final_date = entries[0][1]

                    rec = {
                        "회차": rnd, "공연일": final_date,
                        "지급대상": target, "지급구분": paytype,
                        "금액": parse_int(amt), "정산상태": status,
                        "지급요청일": req_date, "지급완료일": done_date,
                        "은행명": bank, "예금주": holder,
                        "계좌번호": acct, "증빙": evid, "비고": note,
                    }
                    dm.add(rec)
                    st.success("등록 완료!")
                    st.rerun()


def _render_edit_form(dm, all_targets, target_dates):
    target_id = st.session_state.get("edit_id")
    rec = None
    for r in dm.records:
        if str(r.get("ID")) == str(target_id):
            rec = r
            break
    if rec is None:
        st.session_state["edit_mode"] = False
        st.rerun()
        return

    st.markdown(f"##### 정산 수정 — ID {target_id}")
    with st.form(f"edit_form_{target_id}"):
        r1c1, r1c2, r1c3 = st.columns([1, 1, 2])
        with r1c1:
            rnds = [str(i) for i in range(1, 25)]
            rnd_val = str(rec.get("회차", "1"))
            rnd = st.selectbox("회차", rnds,
                                index=rnds.index(rnd_val) if rnd_val in rnds else 0,
                                key=f"ed_rnd_{target_id}")
        with r1c2:
            perf_date = st.text_input("공연일", value=str(rec.get("공연일", "")),
                                       key=f"ed_date_{target_id}")
        with r1c3:
            tgt_val = str(rec.get("지급대상", ""))
            opts = [""] + all_targets
            if tgt_val and tgt_val not in opts:
                opts.append(tgt_val)
                opts.sort()
            target = st.selectbox("지급대상", opts,
                                   index=opts.index(tgt_val) if tgt_val in opts else 0,
                                   key=f"ed_target_{target_id}")

        r2c1, r2c2, r2c3 = st.columns(3)
        with r2c1:
            pt_val = str(rec.get("지급구분", ""))
            paytype = st.selectbox("지급구분", dm.pay_types,
                                    index=dm.pay_types.index(pt_val) if pt_val in dm.pay_types else 0,
                                    key=f"ed_pt_{target_id}")
        with r2c2:
            amt = st.text_input("금액", value=str(rec.get("금액", "")),
                                 key=f"ed_amt_{target_id}")
        with r2c3:
            st_val = str(rec.get("정산상태", "미지급"))
            status = st.selectbox("정산상태", STATUSES,
                                   index=STATUSES.index(st_val) if st_val in STATUSES else 0,
                                   key=f"ed_status_{target_id}")

        r3c1, r3c2, r3c3 = st.columns(3)
        with r3c1:
            req_date = st.text_input("지급요청일", value=str(rec.get("지급요청일", "")),
                                      key=f"ed_reqdate_{target_id}")
        with r3c2:
            done_date = st.text_input("지급완료일", value=str(rec.get("지급완료일", "")),
                                       key=f"ed_donedate_{target_id}")
        with r3c3:
            bank_val = str(rec.get("은행명", ""))
            bank_opts = [""] + BANKS
            bank = st.selectbox("은행명", bank_opts,
                                 index=bank_opts.index(bank_val) if bank_val in bank_opts else 0,
                                 key=f"ed_bank_{target_id}")

        r4c1, r4c2, r4c3 = st.columns(3)
        with r4c1:
            holder = st.text_input("예금주", value=str(rec.get("예금주", "")),
                                    key=f"ed_holder_{target_id}")
        with r4c2:
            acct = st.text_input("계좌번호", value=str(rec.get("계좌번호", "")),
                                  key=f"ed_acct_{target_id}")
        with r4c3:
            evid = st.text_input("증빙", value=str(rec.get("증빙", "")),
                                  key=f"ed_evid_{target_id}")

        note = st.text_input("비고", value=str(rec.get("비고", "")),
                              key=f"ed_note_{target_id}")

        ebc1, ebc2 = st.columns(2)
        with ebc1:
            if st.form_submit_button("수정 저장", use_container_width=True, type="primary"):
                new_rec = {
                    "회차": rnd, "공연일": perf_date,
                    "지급대상": target, "지급구분": paytype,
                    "금액": parse_int(amt), "정산상태": status,
                    "지급요청일": req_date, "지급완료일": done_date,
                    "은행명": bank, "예금주": holder,
                    "계좌번호": acct, "증빙": evid, "비고": note,
                }
                dm.update(target_id, new_rec)
                st.session_state["edit_mode"] = False
                st.success("수정 완료!")
                st.rerun()
        with ebc2:
            if st.form_submit_button("취소", use_container_width=True):
                st.session_state["edit_mode"] = False
                st.rerun()


def _export_excel(records, filename, include_account=True):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "정산내역"
    cols = list(COLUMNS)
    if not include_account:
        cols = [c for c in cols if c not in ("은행명", "예금주", "계좌번호")]
    ws.append(cols)
    for rec in records:
        ws.append([rec.get(c, "") for c in cols])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button(
        f"{filename}.xlsx 다운로드", buf,
        file_name=f"{filename}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ══════════════════════════════════════════════════════════════
#  탭 2: 예산 크로스체크
# ══════════════════════════════════════════════════════════════

def render_tab_crosscheck():
    dm = get_dm()

    st.markdown("#### 예산 기준")
    with st.form("budget_form"):
        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            annual = st.text_input("행사실비보상금(301-10) 연간 예산",
                                    value=str(dm.budget.get("연간예산", 0) or ""),
                                    key="bud_annual")
        with bc2:
            perf_bud = st.text_input("출연료 배정",
                                      value=str(dm.budget.get("출연료배정", 0) or ""),
                                      key="bud_perf")
        with bc3:
            staff_bud = st.text_input("행사진행인력 배정",
                                       value=str(dm.budget.get("행사진행인력배정", 0) or ""),
                                       key="bud_staff")
        if st.form_submit_button("예산 저장", use_container_width=True):
            dm.budget["연간예산"] = parse_int(annual)
            dm.budget["출연료배정"] = parse_int(perf_bud)
            dm.budget["행사진행인력배정"] = parse_int(staff_bud)
            dm.save()
            st.success("예산 기준 저장 완료!")
            st.rerun()

    # ── 크로스체크 테이블 ──
    st.markdown("#### 예산 크로스체크")
    by_pt = dm.calc_by_paytype()
    budget = dm.budget

    def _row(label, bud, pt_key):
        total = by_pt.get(pt_key, {}).get("total", 0)
        done = by_pt.get(pt_key, {}).get("done", 0)
        pending = total - done
        remain = bud - total if bud else 0
        rate = round(total / bud * 100, 1) if bud else 0
        return {
            "구분": label,
            "예산(원)": f"{bud:,}" if bud else "-",
            "정산등록합계": f"{total:,}",
            "지급완료합계": f"{done:,}",
            "미지급+요청중": f"{pending:,}",
            "잔액": f"{remain:,}",
            "집행률(%)": f"{rate}%",
        }

    rows = [
        _row("출연료", budget.get("출연료배정", 0), "출연료"),
        _row("행사진행인력", budget.get("행사진행인력배정", 0), "행사진행인력"),
    ]

    # 합계
    total_bud = budget.get("출연료배정", 0) + budget.get("행사진행인력배정", 0)
    total_reg = sum(by_pt.get(k, {}).get("total", 0) for k in ("출연료", "행사진행인력"))
    total_done = sum(by_pt.get(k, {}).get("done", 0) for k in ("출연료", "행사진행인력"))
    total_pending = total_reg - total_done
    total_remain = total_bud - total_reg
    total_rate = round(total_reg / total_bud * 100, 1) if total_bud else 0
    rows.append({
        "구분": "합계",
        "예산(원)": f"{total_bud:,}",
        "정산등록합계": f"{total_reg:,}",
        "지급완료합계": f"{total_done:,}",
        "미지급+요청중": f"{total_pending:,}",
        "잔액": f"{total_remain:,}",
        "집행률(%)": f"{total_rate}%",
    })

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── 월별 지급 현황 차트 ──
    st.markdown("#### 월별 지급 현황")
    monthly = dm.calc_monthly()
    months = [f"{m}월" for m in range(1, 13)]
    amts = [monthly[m] for m in range(1, 13)]

    # 누적
    cumulative = []
    cum = 0
    for a in amts:
        cum += a
        cumulative.append(cum)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=months, y=amts, name="월별 지급완료",
        marker_color="#2B3A67", opacity=0.85
    ))
    fig.add_trace(go.Scatter(
        x=months, y=cumulative, name="누적 지급",
        mode="lines+markers", yaxis="y2",
        line=dict(color="#E65100", width=2),
        marker=dict(size=7)
    ))
    fig.update_layout(
        yaxis=dict(title="월별 지급(원)"),
        yaxis2=dict(title="누적 지급(원)", overlaying="y", side="right"),
        template="plotly_dark",
        height=400,
        legend=dict(orientation="h", y=1.12),
    )
    st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════
#  탭 3: 단체별 정산 현황
# ══════════════════════════════════════════════════════════════

def render_tab_by_target():
    dm = get_dm()
    by_target = dm.calc_by_target()

    if not by_target:
        st.info("정산 데이터가 없습니다.")
        return

    st.markdown("#### 단체별 정산 요약")

    # 단체 필터
    targets = sorted(by_target.keys())
    flt_target = st.selectbox("단체 선택", ["전체"] + targets, key="flt_target")

    rows = []
    for name in targets:
        if flt_target != "전체" and name != flt_target:
            continue
        info = by_target[name]
        rate = round(info["done"] / info["total"] * 100, 1) if info["total"] else 0
        rows.append({
            "지급대상": name,
            "총출연횟수": info["cnt"],
            "총출연료": f"{info['total']:,}",
            "지급완료": f"{info['done']:,}",
            "미지급": f"{info['total'] - info['done']:,}",
            "정산완료율": f"{rate}%",
        })

    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # 선택된 단체의 상세
    if flt_target != "전체":
        st.markdown(f"#### {flt_target} — 상세 내역")
        detail_recs = [r for r in dm.records if r.get("지급대상") == flt_target]
        if detail_recs:
            detail_data = []
            for r in detail_recs:
                detail_data.append({
                    "ID": r.get("ID", ""),
                    "회차": r.get("회차", ""),
                    "공연일": r.get("공연일", ""),
                    "지급구분": r.get("지급구분", ""),
                    "금액": f"{parse_int(r.get('금액', 0)):,}",
                    "정산상태": r.get("정산상태", ""),
                    "지급완료일": r.get("지급완료일", ""),
                    "비고": r.get("비고", ""),
                })
            st.dataframe(pd.DataFrame(detail_data), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════
#  메인
# ══════════════════════════════════════════════════════════════

def main():
    if not check_password():
        return

    st.markdown("## 💰 토요상설공연 정산 관리")

    # 사이드바
    st.sidebar.markdown("### 데이터 관리")
    if st.sidebar.button("구글 시트 새로고침", use_container_width=True, key="btn_reload"):
        reload_dm()

    st.sidebar.markdown("---")
    with st.sidebar.expander("지급구분 관리"):
        dm = get_dm()
        st.write("현재: " + ", ".join(dm.pay_types))
        new_pt = st.text_input("추가할 지급구분", key="new_pt")
        if st.button("추가", key="btn_add_pt"):
            if new_pt.strip() and new_pt.strip() not in dm.pay_types:
                dm.pay_types.append(new_pt.strip())
                dm.save()
                st.success(f"'{new_pt.strip()}' 추가!")
                st.rerun()
        del_pt = st.selectbox("삭제할 지급구분",
                               [p for p in dm.pay_types if p not in PAY_TYPES_DEFAULT],
                               key="del_pt")
        if st.button("삭제", key="btn_del_pt"):
            if del_pt and del_pt in dm.pay_types:
                dm.pay_types.remove(del_pt)
                dm.save()
                st.rerun()

    tab1, tab2, tab3 = st.tabs([
        "📋 정산 등록·관리",
        "📊 예산 크로스체크",
        "👥 단체별 정산 현황",
    ])

    with tab1:
        render_tab_records()
    with tab2:
        render_tab_crosscheck()
    with tab3:
        render_tab_by_target()


if __name__ == "__main__":
    main()
