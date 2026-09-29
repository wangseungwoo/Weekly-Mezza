"""
주식연계채권(CB/EB/BW) 주간 발행 공시 수집기 — DART OpenAPI 기반

흐름
 1) list.json (공시검색) : 기간 내 주요사항보고서(B) 전체 페이지 수집
 2) 보고서명으로 CB/EB/BW 발행결정만 선별, 대괄호 접두어 붙은 정정류([기재정정]/[발행조건확정]/[첨부정정] 등) 제외
 3) 유가증권(Y)/코스닥(K)만 남김
 4) 상세 API(cvbdIsDecsn/exbdIsDecsn/bdwtIsDecsn)에서 rcept_no로 정확히 매칭해 구조화 필드 수집
 5) 공시 원문(document.xml)에서 투자자, Put/Call, 리픽싱, 할증률 등 파싱
 6) 파싱이 애매한 항목은 flags 에 기록 → 엑셀에서 노란색 표시

CLI 테스트:  python collector.py --key <API키> --start 2026-09-17 --end 2026-09-23
"""
from __future__ import annotations

import io
import re
import time
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Any, Callable, Optional

import requests
import warnings

from bs4 import BeautifulSoup, NavigableString, XMLParsedAsHTMLWarning

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

BASE = "https://opendart.fss.or.kr/api"

KIND_RULES = [  # (보고서명 키워드, 종류코드, 상세 API)
    ("신주인수권부사채권발행결정", "BW", "bdwtIsDecsn"),
    ("교환사채권발행결정", "EB", "exbdIsDecsn"),
    ("전환사채권발행결정", "CB", "cvbdIsDecsn"),
]
MARKET = {"Y": "유가증권", "K": "코스닥"}

STATUS_MSG = {
    "010": "등록되지 않은 API 키입니다.",
    "011": "사용할 수 없는 API 키입니다.",
    "012": "접근할 수 없는 IP입니다.",
    "020": "API 요청 한도를 초과했습니다(일 20,000건).",
    "021": "조회 가능한 회사 수 초과",
    "100": "요청 파라미터가 부적절합니다.",
    "101": "부적절한 접근입니다.",
    "800": "DART 시스템 점검 중입니다.",
    "900": "DART 쪽 정의되지 않은 오류입니다.",
    "901": "API 키 사용기간이 만료되었습니다.",
}


class DartError(Exception):
    pass


# ───────────────────────────── 공통 유틸 ─────────────────────────────
def to_num(s: Any) -> Optional[float]:
    if s is None:
        return None
    s = str(s).strip().replace(",", "").replace("%", "")
    if s in ("", "-", "―", "N/A"):
        return None
    try:
        return float(s)
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None


_DATE_PATS = [
    r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일",
    r"(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})",
    r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)",
]


def to_date(s: Any) -> Optional[date]:
    if s is None:
        return None
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    s = str(s)
    for p in _DATE_PATS:
        m = re.search(p, s)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                continue
    return None


def all_dates(s: str) -> list[date]:
    out = []
    for m in re.finditer(r"(\d{4})\s*(?:년|[.\-/])\s*(\d{1,2})\s*(?:월|[.\-/])\s*(\d{1,2})", s):
        try:
            out.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass
    return out


def months_between(d1: date, d2: date) -> int:
    return round((d2 - d1).days / 30.4375)


def pct(s: Any) -> Optional[float]:
    """'3.0' → 0.03"""
    v = to_num(s)
    return None if v is None else round(v / 100, 6)


def first(d: dict, *keys: str) -> Any:
    for k in keys:
        v = d.get(k)
        if v not in (None, "", "-"):
            return v
    return None


# ───────────────────────────── API 호출 ─────────────────────────────
class Dart:
    def __init__(self, key: str, sleep: float = 0.12, log: Callable[[str], None] = print):
        self.key = key.strip()
        self.s = requests.Session()
        self.sleep = sleep
        self.log = log
        self.calls = 0

    def _get(self, path: str, params: dict, raw: bool = False):
        params = {"crtfc_key": self.key, **params}
        for attempt in range(4):
            try:
                r = self.s.get(f"{BASE}/{path}", params=params, timeout=30)
                self.calls += 1
                time.sleep(self.sleep)
            except requests.RequestException as e:
                if attempt == 3:
                    raise DartError(f"DART 접속 실패: {e}")
                time.sleep(2 * (attempt + 1))
                continue
            if raw:
                # document.xml 은 정상일 때 zip, 오류일 때 xml/json 메시지
                if r.content[:2] == b"PK":
                    return r.content
                m = re.search(rb"<status>(\d+)</status>", r.content) or re.search(rb'"status"\s*:\s*"(\d+)"', r.content)
                code = m.group(1).decode() if m else "?"
                if code == "020" and attempt < 3:
                    time.sleep(5)
                    continue
                raise DartError(f"원문 다운로드 실패({code}) {STATUS_MSG.get(code, '')}")
            j = r.json()
            st = j.get("status")
            if st == "000":
                return j
            if st == "013":  # 조회 데이터 없음
                return {"status": "013", "list": [], "total_page": 0}
            if st == "020" and attempt < 3:
                time.sleep(5)
                continue
            raise DartError(f"[{st}] {STATUS_MSG.get(st, j.get('message', ''))}")
        raise DartError("DART 응답 없음")

    def list_filings(self, start: date, end: date) -> list[dict]:
        out, page = [], 1
        while True:
            j = self._get("list.json", {
                "bgn_de": start.strftime("%Y%m%d"),
                "end_de": end.strftime("%Y%m%d"),
                "pblntf_ty": "B",          # 주요사항보고서
                "last_reprt_at": "N",      # Y로 하면 정정 후 원본이 사라짐 → 반드시 N
                "page_no": page,
                "page_count": 100,         # 기본값 10 → 페이지 누락의 흔한 원인
                "sort": "date", "sort_mth": "asc",
            })
            out += j.get("list", [])
            total = int(j.get("total_page") or 0)
            self.log(f"공시목록 {page}/{max(total, 1)} 페이지")
            if page >= total:
                break
            page += 1
        return out

    def detail(self, api: str, corp_code: str, rcept_dt: str) -> list[dict]:
        j = self._get(f"{api}.json", {"corp_code": corp_code, "bgn_de": rcept_dt, "end_de": rcept_dt})
        return j.get("list", [])

    def document_text(self, rcept_no: str) -> str:
        z = zipfile.ZipFile(io.BytesIO(self._get("document.xml", {"rcept_no": rcept_no}, raw=True)))
        names = sorted(z.namelist(), key=lambda n: (not n.startswith(rcept_no), n))
        return "\n".join(xml_to_text(z.read(n)) for n in names)


# ───────────────────────────── 원문 → 텍스트 ─────────────────────────────
def xml_to_text(b: bytes) -> str:
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            s = b.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        s = b.decode("utf-8", "ignore")
    soup = BeautifulSoup(s, "lxml")
    for tr in soup.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th", "te", "tu"])]
        tr.replace_with(NavigableString("\n" + " | ".join(cells) + "\n"))
    txt = soup.get_text("\n")
    txt = re.sub(r"[ \t 　]+", " ", txt)
    txt = re.sub(r"\n\s*\n+", "\n", txt)
    return txt


def row_value(text: str, label: str) -> Optional[str]:
    """표 형태 줄('라벨 | ... | 값')에서 마지막 칸 값을 돌려줌"""
    for line in text.split("\n"):
        if "|" in line and re.search(label, line):
            cells = [c.strip() for c in line.split("|") if c.strip()]
            if len(cells) >= 2:
                return cells[-1]
    return None


# ───────────────────────────── 섹션 추출 ─────────────────────────────
PUT_KW = r"조기\s*상환\s*청구권|Put\s*Option|풋\s*옵션"
CALL_KW = r"매도\s*청구권|Call\s*Option|콜\s*옵션"
REFIX_KW = r"시가\s*하락에\s*따른|리픽싱|Refixing|(?:전환|교환|행사)\s*가액의\s*조정"
STOP_KW = r"합병\s*관련|기타\s*투자\s*판단|자금조달의\s*목적|특정인에\s*대한"


def section(text: str, kw: str, stop_kws: list[str], maxlen: int = 5000) -> str:
    """키워드가 여러 번 나오면(표의 '참조' 문구 등) 내용이 가장 풍부한 구간을 선택"""
    best, best_score = "", -1
    for m in re.finditer(kw, text, re.I):
        s = m.start()
        w = text[s:s + maxlen]
        cut = len(w)
        for sk in stop_kws:
            m2 = re.search(sk, w[40:], re.I)
            if m2:
                cut = min(cut, m2.start() + 40)
        w = w[:cut]
        score = len(re.findall(r"개월", w)) * 2 + len(re.findall(r"%", w)) + len(all_dates(w))
        if score > best_score:
            best, best_score = w, score
    return best


def is_none_section(w: str) -> bool:
    head = w[:250]
    return (not w) or bool(re.search(r"해당\s*(사항)?\s*없음|부여하지\s*않|없습니다", head))


def _num_after(pat: str, w: str) -> Optional[float]:
    m = re.search(pat, w, re.I | re.S)
    return float(m.group(1)) if m else None


def parse_rate(w: str, prefer: str) -> Optional[float]:
    """수익률(%) → 소수. prefer: 선호 라벨 정규식"""
    for p in (
        prefer + r"[^%\d]{0,40}?(\d{1,2}(?:\.\d+)?)\s*%",
        r"연\s*(?:복리\s*|단리\s*)?(\d{1,2}(?:\.\d+)?)\s*%",
        r"(\d{1,2}(?:\.\d+)?)\s*%\s*\(?\s*(?:연|3개월|분기|월)?\s*(?:단위)?\s*복리",
    ):
        v = _num_after(p, w)
        if v is not None and 0 <= v <= 30:
            return round(v / 100, 6)
    return None


def start_month(w: str) -> Optional[int]:
    v = _num_after(r"발행\s*일?\s*(?:로부터|부터|이후|후)\s*(\d{1,3})\s*개월", w)
    return int(v) if v is not None else None


def cycle_month(w: str) -> Optional[int]:
    if re.search(r"매\s*월(?!\s*\d)|매\s*1\s*개월", w):
        return 1
    v = _num_after(r"매\s*(\d{1,2})\s*개월", w)
    return int(v) if v is not None else None


def schedule_months(w: str, issue: Optional[date]) -> list[int]:
    """옵션 일정표(청구기간 시작|종료|상환일|율)의 각 행 마지막 날짜 → 발행일 대비 개월수"""
    if not issue:
        return []
    ms = []
    for line in w.split("\n"):
        if "|" not in line:
            continue
        ds = all_dates(line)
        if ds:
            m = months_between(issue, ds[-1])
            if 0 < m < 240 and (not ms or m > ms[-1]):
                ms.append(m)
    return ms


def call_ratio(w: str) -> Optional[float]:
    for m in re.finditer(r"(\d{1,3}(?:\.\d+)?)\s*%", w):
        v = float(m.group(1))
        pre, post = w[max(0, m.start() - 25):m.start()], w[m.end():m.end() + 25]
        if v <= 0 or v >= 100:
            continue
        if re.search(r"연\s*$|연\s*복리\s*$|이율|수익률|이자", pre) or re.search(r"^\s*\)?\s*(?:\(?\s*(?:연|3개월|분기))?\s*[^%]{0,6}복리", post):
            continue
        if re.search(r"초과|해당|한도|범위|이내|이하|까지", pre + post):
            return round(v / 100, 6)
    return None


# ───────────────────────────── 투자자 ─────────────────────────────
def parse_investors(text: str) -> list[str]:
    lines = text.split("\n")
    names: list[str] = []
    for i, line in enumerate(lines):
        if "|" in line and re.search(r"대상자", line) and re.search(r"관계|총액|금액|권면", line):
            for l2 in lines[i + 1:i + 80]:
                if "|" not in l2:
                    if names:
                        break
                    continue
                cells = [c.strip() for c in l2.split("|")]
                nm = cells[0] if cells else ""
                if not nm or re.search(r"합\s*계|^계$|^-$|대상자", nm):
                    if re.search(r"합\s*계|^계$", nm):
                        break
                    continue
                names.append(nm)
            if names:
                break
    return names


def clean_investor(n: str) -> str:
    m = re.search(r"([가-힣A-Za-z0-9&]+?자산운용)", n.replace(" ", ""))
    if m:
        return m.group(1)
    n = re.sub(r"\(주\)|㈜|주식회사|유한회사|유한책임회사", "", n)
    n = re.sub(r"\s+", " ", n).strip(" ,")
    return n


def format_investors(raw: list[str]) -> str:
    seen, cleaned = set(), []
    for r in raw:
        c = clean_investor(r)
        if c and c not in seen:
            seen.add(c)
            cleaned.append(c)
    amcs = [c for c in cleaned if c.endswith("자산운용")]
    if len(amcs) <= 1:
        return ", ".join(cleaned)
    # 자산운용사 여러 곳이면 '수성, 라이노스, 이지스, 신한자산운용' 식으로 묶음
    grouped = [a[:-4] for a in amcs[:-1]] + [amcs[-1]]
    out, placed = [], False
    for c in cleaned:
        if c in amcs:
            if not placed:
                out += grouped
                placed = True
        else:
            out.append(c)
    return ", ".join(out)


# ───────────────────────────── 레코드 ─────────────────────────────
@dataclass
class Rec:
    rcept_no: str
    report_nm: str
    kind: str
    공시일: Optional[date] = None
    납입일: Optional[date] = None
    발행사: str = ""
    종류: str = ""
    총액: Optional[float] = None
    투자자: str = ""
    법인유형: str = ""
    코드: str = ""
    coupon: Any = None
    ytm: Any = None
    ytp: Any = "-"
    put_start: Any = "-"
    put_cycle: Any = "-"
    ytc: Any = "-"
    call_start: Any = "-"
    call_end: Any = "-"
    call_cycle: Any = "-"
    call_pct: Any = "-"
    만기일: Optional[date] = None
    만기: str = ""
    전환개시: Optional[date] = None
    할증율: Any = None
    refix_down: Any = "-"
    refix_down_par: Any = "-"
    refix_up: Any = "-"
    refix_cycle: Any = "-"
    신고서: str = ""
    회차: str = ""
    주관사_API: str = ""
    flags: list[str] = field(default_factory=list)
    snippets: dict = field(default_factory=dict)

    @property
    def link(self) -> str:
        return f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={self.rcept_no}"


def classify(report_nm: str):
    if report_nm.strip().startswith("["):  # [기재정정] [발행조건확정] [첨부정정] 등 모두 제외
        return None
    nm = report_nm.replace(" ", "")
    for kw, kind, api in KIND_RULES:
        if kw in nm:
            return kind, api
    return None


def build_record(dart: Dart, item: dict, kind: str, api: str) -> Rec:
    r = Rec(rcept_no=item["rcept_no"], report_nm=item["report_nm"], kind=kind)
    r.공시일 = to_date(item["rcept_dt"])
    r.발행사 = item["corp_name"]
    r.법인유형 = MARKET.get(item.get("corp_cls", ""), item.get("corp_cls", ""))
    r.코드 = "A" + (item.get("stock_code") or "").strip()

    # 1) 상세 API — rcept_no 로 정확히 매칭 (같은 날 여러 회차 대비)
    d: dict = {}
    try:
        rows = dart.detail(api, item["corp_code"], item["rcept_dt"])
        d = next((x for x in rows if x.get("rcept_no") == r.rcept_no), {})
        if not d:
            r.flags.append("상세API 미반영(원문으로 대체)")
    except DartError as e:
        r.flags.append(f"상세API 오류: {e}")

    # 2) 원문
    text = ""
    try:
        text = dart.document_text(r.rcept_no)
    except (DartError, zipfile.BadZipFile) as e:
        r.flags.append(f"원문 오류: {e}")

    fill_record(r, d, text)
    return r


def fill_record(r: Rec, d: dict, text: str) -> None:
    kind = r.kind
    price_lbl = {"CB": r"전환\s*가액", "EB": r"교환\s*가액", "BW": r"행사\s*가액"}[kind]

    m = re.search(r"회\s*차\s*\|\s*(\d+)", text)
    r.회차 = str(first(d, "bd_tm") or (m.group(1) if m else ""))
    r.주관사_API = str(first(d, "rpmcmp") or "")

    # 발행방법
    mth = str(first(d, "bdis_mthn") or row_value(text, r"사채\s*발행\s*방법") or "")
    if "공모" in mth:
        r.종류, r.신고서 = f"공모{kind}", "제출"
    elif "사모" in mth:
        r.종류, r.신고서 = f"사모{kind}", "면제"
    else:
        r.종류, r.신고서 = kind, ""
        r.flags.append("공모/사모 구분 확인")
    if re.search(r"해외", r.report_nm + mth) or first(d, "ovis_fta"):
        r.flags.append("해외발행 여부 확인")

    # 금액
    fta = to_num(first(d, "bd_fta") or row_value(text, r"권면\s*(?:\(전자등록\))?\s*총액"))
    r.총액 = round(fta / 1_000_000) if fta else None
    if r.총액 is None:
        r.flags.append("전자등록총액")

    # 이율
    r.coupon = pct(first(d, "bd_intr_ex") or row_value(text, r"표면\s*이자율"))
    r.ytm = pct(first(d, "bd_intr_sf") or row_value(text, r"만기\s*이자율"))
    if r.coupon is None:
        r.flags.append("Coupon")
    if r.ytm is None:
        r.flags.append("YTM")

    # 일자
    r.납입일 = to_date(first(d, "pymd") or row_value(text, r"납\s*입\s*일"))
    r.만기일 = to_date(first(d, "bd_mtd") or row_value(text, r"사채\s*만기\s*일"))
    r.전환개시 = to_date(first(d, "cvrqpd_bgd", "exrqpd_bgd", "expd_bgd")
                     or row_value(text, r"(?:청구|행사)\s*기간.*시작\s*일|^\s*시작\s*일"))
    for k, v in (("납입일", r.납입일), ("만기일", r.만기일), ("전환개시", r.전환개시)):
        if v is None:
            r.flags.append(k)
    if r.납입일 and r.만기일:
        m = months_between(r.납입일, r.만기일)
        r.만기 = f"{m // 12}년" if m % 12 == 0 else f"{m}개월"

    # 투자자
    inv = parse_investors(text)
    r.투자자 = format_investors(inv)
    if not r.투자자:
        r.flags.append("투자자")

    issue = r.납입일

    # Put
    w = section(text, PUT_KW, [CALL_KW, REFIX_KW, STOP_KW])
    r.snippets["put"] = w[:1500]
    if is_none_section(w):
        pass  # '-' 유지
    else:
        sm = start_month(w)
        cy = cycle_month(w)
        sched = schedule_months(w, issue)
        if sm is None and sched:
            sm = sched[0]
        if cy is None and len(sched) >= 2:
            cy = sched[1] - sched[0]
        r.put_start = sm if sm is not None else None
        r.put_cycle = cy if cy is not None else None
        r.ytp = parse_rate(w, r"(?:조기\s*상환|보장)\s*수익률")
        if r.ytp is None and r.ytm is not None:
            r.ytp = r.ytm
            r.flags.append("YTP(YTM으로 대체)")
        for k, v in (("Put 시작", r.put_start), ("Put 주기", r.put_cycle), ("YTP", r.ytp)):
            if v is None:
                r.flags.append(k)

    # Call
    w = section(text, CALL_KW, [PUT_KW, REFIX_KW, STOP_KW])
    r.snippets["call"] = w[:1500]
    if not is_none_section(w):
        sm = start_month(w)
        em = _num_after(r"개월\s*(?:이\s*)?(?:되는|경과하는|경과한)?\s*(?:날|일)?\s*(?:부터|로부터)[^.。]{0,30}?(\d{1,3})\s*개월", w)
        cy = cycle_month(w)
        sched = schedule_months(w, issue)
        if sm is None and sched:
            sm = sched[0]
        if em is None and sched:
            em = sched[-1]
        if cy is None and len(sched) >= 2:
            cy = sched[1] - sched[0]
        r.call_start = sm
        r.call_end = int(em) if em is not None else None
        r.call_cycle = cy
        r.call_pct = call_ratio(w)
        r.ytc = parse_rate(w, r"(?:매도\s*청구|매매\s*가격|조기\s*상환)?\s*수익률|이율")
        for k, v in (("Call 시작", r.call_start), ("Call 종료", r.call_end), ("Call 주기", r.call_cycle),
                     ("Call %", r.call_pct), ("YTC", r.ytc)):
            if v is None:
                r.flags.append(k)

    # 할증율
    m = re.search(r"할증\s*(?:률|율)\s*[:：]?\s*(-?\d{1,3}(?:\.\d+)?)\s*%", text)
    if m:
        r.할증율 = round(float(m.group(1)) / 100, 6)
    else:
        m = re.search(r"기준\s*(?:주가|가격|가액)의\s*(\d{2,3}(?:\.\d+)?)\s*%", text)
        if m:
            r.할증율 = round(float(m.group(1)) / 100 - 1, 6)
        else:
            r.flags.append("할증율")

    # Refixing
    cvp = to_num(first(d, "cv_prc", "ex_prc") or row_value(text, price_lbl + r"\s*\(원|^\s*\S*\s*" + price_lbl))
    low = to_num(first(d, "act_mktprcfl_cvprc_lwtrsprc") or row_value(text, r"최저\s*조정\s*가액"))
    w = section(text, REFIX_KW, [PUT_KW, CALL_KW, STOP_KW], maxlen=3000)
    r.snippets["refix"] = w[:1500]
    if low and cvp:
        r.refix_down = round(low / cvp, 2)
        cy = cycle_month(w)
        r.refix_cycle = f"매{cy}개월" if cy else None
        if cy is None:
            r.flags.append("Refixing 주기")
        if re.search(r"상향\s*(?:하여|조정)", w) and not re.search(r"상향\s*조정\s*(?:하지|은\s*하지|되지)", w):
            r.refix_up = 1
            r.flags.append("상향Refixing(자동감지)")
    elif cvp is None:
        r.flags.append("전환가액")
    # 70% 미만 리픽싱(주총 특별결의 등)은 형태가 다양해 표시만
    if re.search(r"70\s*%\s*미만", w) and re.search(r"특별\s*결의|액면", w):
        r.flags.append("하향Refixing(액면가o) 확인")


# ───────────────────────────── 메인 ─────────────────────────────
def collect(key: str, start: date, end: date, log: Callable[[str], None] = print,
            progress: Callable[[float], None] = lambda p: None):
    dart = Dart(key, log=log)
    filings = dart.list_filings(start, end)

    targets, excluded = [], []
    for it in filings:
        c = classify(it["report_nm"])
        nm = it["report_nm"].replace(" ", "")
        is_elb = any(kw in nm for kw, _, _ in KIND_RULES)
        if not is_elb:
            continue
        if c is None:
            excluded.append({**it, "사유": "정정/확정류"})
            continue
        if it.get("corp_cls") not in MARKET:
            excluded.append({**it, "사유": f"시장구분 {it.get('corp_cls')}"})
            continue
        targets.append((it, *c))
    log(f"대상 {len(targets)}건 / 제외 {len(excluded)}건")

    recs, errors = [], []
    for i, (it, kind, api) in enumerate(targets):
        try:
            recs.append(build_record(dart, it, kind, api))
        except Exception as e:  # 한 건 실패로 전체가 멈추지 않게
            errors.append(f"{it['corp_name']} {it['rcept_no']}: {e}")
        progress((i + 1) / max(len(targets), 1))
    recs.sort(key=lambda r: (r.공시일 or date.min, r.rcept_no))
    log(f"API 호출 {dart.calls}회")
    return recs, excluded, errors


def default_week(today: Optional[date] = None) -> tuple[date, date]:
    """직전(또는 오늘) 수요일로 끝나는 목~수 주간"""
    today = today or date.today()
    end = today - timedelta(days=(today.weekday() - 2) % 7)
    return end - timedelta(days=6), end


if __name__ == "__main__":
    import argparse
    from excel_writer import write_excel

    ap = argparse.ArgumentParser()
    ap.add_argument("--key", required=True)
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--out", default="주식연계채권_발행현황.xlsx")
    a = ap.parse_args()
    s, e = default_week()
    s = to_date(a.start) or s
    e = to_date(a.end) or e
    recs, exc, errs = collect(a.key, s, e)
    with open(a.out, "wb") as f:
        f.write(write_excel(recs, s, e))
    print(f"{len(recs)}건 저장 → {a.out}")
    for x in errs:
        print("오류:", x)
