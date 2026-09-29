import traceback
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from collector import DartError, collect, default_week
from excel_writer import write_excel

st.set_page_config(page_title="주식연계채권 주간 집계", layout="wide")
st.title("주식연계채권 주간 발행현황 (CB · EB · BW)")
st.caption("DART 주요사항보고서 원본 공시만 수집 · [기재정정]/[발행조건확정]/[첨부정정] 등 정정류 제외 · 유가증권/코스닥")

try:
    key = st.secrets.get("DART_API_KEY", "")
except Exception:  # secrets 파일이 없을 때
    key = ""
if not key:
    key = st.text_input("DART API 키", type="password", help="Streamlit secrets에 DART_API_KEY로 저장해두면 입력하지 않아도 됩니다.")

s0, e0 = default_week()
c1, c2 = st.columns(2)
start = c1.date_input("시작일(목)", s0)
end = c2.date_input("종료일(수)", e0)
if start.weekday() != 3 or end.weekday() != 2:
    st.info("목요일~수요일이 아닌 기간입니다. 그대로 수집은 가능합니다.")

if st.button("수집 시작", type="primary", disabled=not key):
    log_box = st.empty()
    bar = st.progress(0.0)
    logs = []

    def log(m):
        logs.append(m)
        log_box.caption(" · ".join(logs[-3:]))

    try:
        recs, excluded, errors = collect(key, start, end, log=log, progress=bar.progress)
    except DartError as e:
        st.error(str(e))
        st.stop()
    except Exception:
        st.error("예상치 못한 오류")
        st.code(traceback.format_exc())
        st.stop()

    st.session_state["result"] = (recs, excluded, errors, start, end)

if "result" in st.session_state:
    recs, excluded, errors, s, e = st.session_state["result"]
    n_flag = sum(1 for r in recs if r.flags)
    st.success(f"{len(recs)}건 수집 · 확인필요 {n_flag}건 · 제외 {len(excluded)}건")

    st.download_button(
        "엑셀 다운로드",
        write_excel(recs, s, e),
        file_name=f"주식연계채권 발행현황_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    df = pd.DataFrame([{
        "공시일": r.공시일, "발행사": r.발행사, "종류": r.종류, "총액(백만)": r.총액,
        "Coupon": r.coupon, "YTM": r.ytm, "YTP": r.ytp, "Put": f"{r.put_start}/{r.put_cycle}",
        "Call": f"{r.call_start}~{r.call_end}/{r.call_cycle} {r.call_pct}", "투자자": r.투자자,
        "확인필요": ", ".join(r.flags), "원문": r.link,
    } for r in recs])
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True,
                     column_config={"원문": st.column_config.LinkColumn("원문", display_text="열기")})

    if errors:
        with st.expander(f"수집 실패 {len(errors)}건", expanded=True):
            for x in errors:
                st.write(x)
    if excluded:
        with st.expander(f"제외된 공시 {len(excluded)}건"):
            st.dataframe(pd.DataFrame(excluded)[["rcept_dt", "corp_name", "report_nm", "사유"]],
                         use_container_width=True, hide_index=True)
