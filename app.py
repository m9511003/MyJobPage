import html
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
CSV_CANDIDATES = [
    BASE_DIR / "data" / "SB_projects_demo.csv",
    BASE_DIR / "data" / "SB_project_demo.csv",  # 실제 폴더에 있는 파일명
]
RESOURCE_CANDIDATES = [
    BASE_DIR / "data" / "SB_resources_demo.csv",
    BASE_DIR / "data" / "SB_resource_demo.csv",  # 실제 폴더에 있는 파일명
]
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
REQUIRED_COLS = [
    "의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격",
    "담당엔지니어", "접수일", "완료예정일", "진행상태", "시료상태", "메모",
]
FILTER_COLS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
INITIAL_CATEGORIES = ["EMC", "무선(RF)", "인증서류", "SAR"]
DEFAULT_BASE_DATE = date(2026, 10, 23)

S_DONE = "완료"
S_LATE = "지연 후보(일정 확인 필요)"
S_TODAY = "오늘 마감"
S_PLANNED = "진행 예정"
S_CHECK = "일정 확인 필요"

EDU_NOTICE = "교육용 가상 데이터 기준 — 실제 업무 시스템 아님"
XLSX_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
SESSION_NOTICE = (
    "카테고리·메모 등 변경 내용은 이 브라우저 세션(session_state)에만 보관되며, "
    "세션이 종료되면 사라질 수 있습니다. 원본 CSV는 수정되지 않습니다."
)
RESOURCE_NOTICE = "현황 조회용이며 예약·시험 재개를 승인하는 기능이 아님"

NAVY = "#0B2A5B"
BLUE = "#1F6FEB"
STATUS_COLORS = ["#9FB8E0", "#6C9BE0", BLUE, "#1B4DB1", "#D9534F", NAVY]
RESOURCE_COLORS = {"사용 가능": "#1E8E5A", "사용 중": BLUE, "점검 중": "#D9534F"}


# ───────────────────────── 데이터 로딩 ─────────────────────────

def load_csv(source) -> pd.DataFrame:
    """CSV를 UTF-8(BOM 포함) 우선, 실패하면 CP949로 읽는다. 경로/업로드 객체 공용."""
    for enc in ("utf-8-sig", "cp949"):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            return pd.read_csv(source, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 인코딩을 읽을 수 없습니다.")


def first_existing(paths):
    """후보 경로 중 실제로 존재하는 첫 파일을 반환한다(없으면 None)."""
    return next((p for p in paths if p.exists()), None)


def init_state(df: pd.DataFrame, source_name: str):
    """데이터 소스가 바뀌었을 때만 session_state를 초기화한다(카테고리·원본 데이터 보관)."""
    if st.session_state.get("source_name") == source_name:
        return
    for k in list(st.session_state.keys()):
        if k.startswith(("f_", "memo_input_")):
            del st.session_state[k]
    st.session_state["source_name"] = source_name
    st.session_state["df"] = df
    extra = [v for v in df["시험분야"].dropna().unique() if v not in INITIAL_CATEGORIES]
    st.session_state["categories"] = INITIAL_CATEGORIES + extra


# ───────────────────────── 일정 상태 계산 ─────────────────────────

def schedule_status(progress, remaining):
    """진행상태와 잔여일수(NaN 가능)로 일정 상태를 판단한다."""
    if progress == "완료":
        return S_DONE
    if pd.isna(remaining):
        return S_CHECK
    if remaining < 0:
        return S_LATE
    if remaining == 0:
        return S_TODAY
    return S_PLANNED


def add_schedule_columns(df: pd.DataFrame, base_ts: pd.Timestamp) -> pd.DataFrame:
    """날짜 컬럼을 파싱하고 잔여일수·일정 상태 열을 추가한다(빈칸/형식오류는 NaT)."""
    df = df.copy()
    df["접수일"] = pd.to_datetime(df["접수일"], errors="coerce")
    df["완료예정일"] = pd.to_datetime(df["완료예정일"], errors="coerce")
    df["잔여일수"] = (df["완료예정일"] - base_ts).dt.days.astype("Int64")
    df["일정 상태"] = [
        schedule_status(p, r) for p, r in zip(df["진행상태"], df["잔여일수"])
    ]
    return df


def to_display(df: pd.DataFrame) -> pd.DataFrame:
    """표/엑셀 출력용: 날짜를 문자열로 바꾸고 컬럼 순서를 정리한다."""
    view = df.copy()
    view["접수일"] = view["접수일"].dt.strftime("%Y-%m-%d")
    view["완료예정일"] = view["완료예정일"].dt.strftime("%Y-%m-%d")
    return view[REQUIRED_COLS[:9] + ["잔여일수", "일정 상태"] + REQUIRED_COLS[9:]]


def highlight_late(row):
    """지연 후보 행을 빨간색 계열로 강조하는 스타일."""
    style = "background-color: #FDE2E2; color: #9B1C1C;" if row["일정 상태"] == S_LATE else ""
    return [style] * len(row)


# ───────────────────────── 카테고리 관리 (콜백) ─────────────────────────

def _set_msg(kind: str, text: str):
    st.session_state["cat_msg"] = (kind, text)


def _reset_cat_widgets():
    """카테고리 변경 후 선택지가 바뀌므로 관련 위젯 상태를 비운다."""
    for k in ("ren_target", "del_target", "del_reassign"):
        st.session_state.pop(k, None)


def cb_add_category():
    """시험분야를 새로 등록한다(빈 값·중복 거부). 필터에는 선택된 상태로 추가."""
    name = st.session_state.get("new_cat", "").strip()
    cats = st.session_state["categories"]
    if not name:
        return _set_msg("error", "등록할 시험분야 이름을 입력하세요.")
    if name in cats:
        return _set_msg("error", f"'{name}'은(는) 이미 있는 시험분야입니다.")
    cats.append(name)
    sel = st.session_state.get("f_시험분야")
    if sel is not None:
        sel.append(name)
    st.session_state["new_cat"] = ""
    _reset_cat_widgets()
    _set_msg("success", f"'{name}' 등록 완료")


def cb_rename_category():
    """시험분야 이름을 바꾸고 의뢰 데이터·필터 선택값도 함께 갱신한다."""
    old = st.session_state.get("ren_target")
    new = st.session_state.get("ren_new", "").strip()
    cats = st.session_state["categories"]
    if old is None:
        return _set_msg("error", "이름을 바꿀 시험분야를 선택하세요.")
    if not new:
        return _set_msg("error", "새 이름을 입력하세요.")
    if new == old:
        return _set_msg("error", "현재 이름과 같습니다.")
    if new in cats:
        return _set_msg("error", f"'{new}'은(는) 이미 있는 시험분야입니다.")
    cats[cats.index(old)] = new
    d = st.session_state["df"]
    n = int((d["시험분야"] == old).sum())
    d.loc[d["시험분야"] == old, "시험분야"] = new
    sel = st.session_state.get("f_시험분야")
    if sel is not None:
        st.session_state["f_시험분야"] = [new if v == old else v for v in sel]
    st.session_state["ren_new"] = ""
    _reset_cat_widgets()
    _set_msg("success", f"'{old}' → '{new}' 변경 완료 (의뢰 {n}건 함께 갱신)")


def cb_delete_category():
    """사용 중이지 않은 시험분야만 삭제한다."""
    name = st.session_state.get("del_target")
    d = st.session_state["df"]
    if name is None:
        return _set_msg("error", "삭제할 시험분야를 선택하세요.")
    if (d["시험분야"] == name).any():
        return _set_msg("error", f"'{name}'은(는) 사용 중이라 바로 삭제할 수 없습니다.")
    st.session_state["categories"].remove(name)
    _reset_cat_widgets()
    _set_msg("success", f"'{name}' 삭제 완료")


def cb_reclassify_and_delete():
    """사용 중인 시험분야의 의뢰를 다른 분야로 재분류한 뒤 삭제한다."""
    name = st.session_state.get("del_target")
    target = st.session_state.get("del_reassign")
    if name is None or target is None or target == name:
        return _set_msg("error", "재분류할 다른 시험분야를 선택하세요.")
    d = st.session_state["df"]
    n = int((d["시험분야"] == name).sum())
    d.loc[d["시험분야"] == name, "시험분야"] = target
    st.session_state["categories"].remove(name)
    _reset_cat_widgets()
    _set_msg("success", f"의뢰 {n}건을 '{target}'(으)로 재분류하고 '{name}'을(를) 삭제했습니다.")


def render_category_manager():
    """사이드바 '카테고리 관리' 영역: 등록 / 이름 변경 / 삭제(재분류 포함)."""
    d = st.session_state["df"]
    cats = st.session_state["categories"]
    st.sidebar.header("카테고리 관리")
    st.sidebar.caption("시험분야 등록 · 이름 변경 · 삭제")
    st.sidebar.warning(SESSION_NOTICE)

    msg = st.session_state.pop("cat_msg", None)
    if msg:
        getattr(st.sidebar, msg[0])(msg[1])

    with st.sidebar.expander("등록", expanded=False):
        st.text_input("새 시험분야", key="new_cat")
        st.button("등록", on_click=cb_add_category, key="btn_add")

    with st.sidebar.expander("이름 변경", expanded=False):
        st.selectbox("대상 시험분야", cats, key="ren_target")
        st.text_input("새 이름", key="ren_new")
        st.button("이름 변경", on_click=cb_rename_category, key="btn_ren")
        st.caption("관련 의뢰 데이터의 시험분야도 함께 바뀝니다.")

    with st.sidebar.expander("삭제", expanded=False):
        target = st.selectbox("삭제할 시험분야", cats, key="del_target")
        used = int((d["시험분야"] == target).sum()) if target else 0
        if target and used > 0:
            st.error(
                f"'{target}'은(는) 의뢰 {used}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                "아래에서 다른 분야로 재분류한 뒤 삭제하세요."
            )
            others = [c for c in cats if c != target]
            st.selectbox("재분류할 시험분야", others, key="del_reassign")
            st.button(
                "재분류 후 삭제", on_click=cb_reclassify_and_delete,
                key="btn_reclass", disabled=not others,
            )
        else:
            st.button("삭제", on_click=cb_delete_category, key="btn_del", disabled=not target)


# ───────────────────────── 필터 ─────────────────────────

def render_filters(df: pd.DataFrame):
    """사이드바 '필터' 영역을 그리고 (필터 결과 DataFrame, 선택 내역 dict)를 반환한다."""
    st.sidebar.header("필터")
    mask = pd.Series(True, index=df.index)
    chosen = {}
    for col in FILTER_COLS:
        if col == "시험분야":
            options = list(st.session_state["categories"])
        else:
            options = list(df[col].dropna().unique())
        if col == "진행상태":
            options = [s for s in STATUSES if s in options] + [o for o in options if o not in STATUSES]
        key = f"f_{col}"
        if key not in st.session_state:
            st.session_state[key] = list(options)
        st.session_state[key] = [v for v in st.session_state[key] if v in options]
        picked = st.sidebar.multiselect(col, options, key=key)
        mask &= df[col].isin(picked)
        chosen[col] = "전체" if set(picked) == set(options) else (", ".join(picked) or "(선택 없음)")
    return df[mask], chosen


# ───────────────────────── 화면 조각 ─────────────────────────

def card_html(label: str, value, extra_class: str = "") -> str:
    """요약 카드 1개의 HTML을 만든다."""
    return (
        f'<div class="card {extra_class}"><div class="label">{html.escape(str(label))}</div>'
        f'<div class="value">{value}</div></div>'
    )


def render_summary(fdf: pd.DataFrame, base_date: date):
    """전체 건수 + 진행상태별 건수 카드와 일정 상태 요약 문구를 그린다."""
    cols = st.columns(len(STATUSES) + 1)
    cols[0].markdown(card_html("전체 의뢰", len(fdf), "total"), unsafe_allow_html=True)
    for col, status in zip(cols[1:], STATUSES):
        col.markdown(card_html(status, int((fdf["진행상태"] == status).sum())), unsafe_allow_html=True)
    st.caption(
        f"기준일 {base_date:%Y-%m-%d} · 지연 후보 {int((fdf['일정 상태'] == S_LATE).sum())}건 · "
        f"오늘 마감 {int((fdf['일정 상태'] == S_TODAY).sum())}건 · "
        f"일정 확인 필요 {int((fdf['일정 상태'] == S_CHECK).sum())}건"
    )


def render_charts(fdf: pd.DataFrame):
    """시험분야별 건수 막대그래프와 진행상태 비율 도넛 차트."""
    c1, c2 = st.columns(2)
    with c1:
        by_field = fdf["시험분야"].value_counts().rename_axis("시험분야").reset_index(name="건수")
        bar = (
            alt.Chart(by_field).mark_bar(color=BLUE)
            .encode(
                x=alt.X("시험분야:N", sort="-y", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("건수:Q", axis=alt.Axis(tickMinStep=1)),
                tooltip=["시험분야", "건수"],
            )
            .properties(title="시험분야별 건수", height=300)
        )
        st.altair_chart(bar, width="stretch")
    with c2:
        by_status = fdf["진행상태"].value_counts().rename_axis("진행상태").reset_index(name="건수")
        donut = (
            alt.Chart(by_status).mark_arc(innerRadius=60)
            .encode(
                theta="건수:Q",
                color=alt.Color("진행상태:N", scale=alt.Scale(domain=STATUSES, range=STATUS_COLORS)),
                tooltip=["진행상태", "건수"],
            )
            .properties(title="진행상태 비율", height=300)
        )
        st.altair_chart(donut, width="stretch")


def render_resources():
    """시험 자원 현황 카드. CSV 값을 그대로 보여 주며 상태를 코드가 바꾸지 않는다."""
    st.subheader("시험 자원 현황")
    path = first_existing(RESOURCE_CANDIDATES)
    if path is None:
        st.info("data 폴더에서 SB_resources_demo.csv 파일을 찾지 못해 자원 현황을 표시하지 않습니다.")
        return
    res = load_csv(path)
    if not {"자원", "상태"} <= set(res.columns):
        st.info("자원 파일에 '자원', '상태' 컬럼이 없어 표시하지 않습니다.")
        return
    cols = st.columns(max(len(res), 1))
    for col, (_, r) in zip(cols, res.iterrows()):
        status = str(r["상태"])
        color = RESOURCE_COLORS.get(status, "#50607A")
        note = "" if pd.isna(r.get("비고")) else html.escape(str(r["비고"]))
        col.markdown(
            f'<div class="card" style="border-top-color:{color}">'
            f'<div class="label">{html.escape(str(r["자원"]))}</div>'
            f'<div class="value" style="color:{color}">{html.escape(status)}</div>'
            f'<div class="label">{note}</div>'
            f'<div class="label" style="font-size:0.72rem;margin-top:6px">{RESOURCE_NOTICE}</div></div>',
            unsafe_allow_html=True,
        )
    st.caption(f"{RESOURCE_NOTICE}. 데이터 파일: {path.name}")


def cb_save_memo(req_no: str):
    """입력창의 메모를 session_state의 의뢰 데이터에 저장한다."""
    d = st.session_state["df"]
    d.loc[d["의뢰번호"] == req_no, "메모"] = st.session_state.get(f"memo_input_{req_no}", "")
    st.toast(f"{req_no} 메모를 저장했습니다 (세션 내 보관)")


def render_detail(fdf: pd.DataFrame):
    """의뢰번호 선택 → 상세 정보 표시, 메모는 입력창+저장 버튼으로 수정."""
    st.subheader("의뢰 상세 조회")
    req_no = st.selectbox("의뢰번호", fdf["의뢰번호"].tolist(), key="detail_req")
    row = fdf[fdf["의뢰번호"] == req_no].iloc[0]
    due = "-" if pd.isna(row["완료예정일"]) else row["완료예정일"].strftime("%Y-%m-%d")
    a, b, c = st.columns(3)
    a.markdown(f"**고객사**  \n{row['고객사']}")
    b.markdown(f"**제품 · 모델**  \n{row['제품명']} · {row['모델명']}")
    c.markdown(f"**시험분야**  \n{row['시험분야']}")
    a.markdown(f"**적용규격**  \n{row['적용규격']}")
    b.markdown(f"**담당엔지니어**  \n{row['담당엔지니어']}")
    c.markdown(f"**시료상태**  \n{row['시료상태']}")
    a.markdown(f"**완료예정일**  \n{due}")
    b.markdown(f"**진행상태 / 일정 상태**  \n{row['진행상태']} / {row['일정 상태']}")
    st.text_area("담당자 메모", value=str(row["메모"]), key=f"memo_input_{req_no}", height=100)
    st.button("메모 저장", on_click=cb_save_memo, args=(req_no,), key="btn_memo")


# ───────────────────────── Excel 다운로드 ─────────────────────────

def _safe_cell(v):
    """엑셀 수식으로 해석될 수 있는 문자열(= + - @ 시작)은 앞에 '를 붙여 막는다."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def _write_sheet(writer, name: str, frame: pd.DataFrame):
    """시트 1개를 쓴다: 1행 교육용 안내문구, 3행부터 표, 헤더 스타일·열 너비 적용."""
    from openpyxl.styles import Alignment, Font, PatternFill

    frame = frame.map(_safe_cell) if not frame.empty else frame
    frame.to_excel(writer, sheet_name=name, index=False, startrow=2)
    ws = writer.sheets[name]
    ws["A1"] = XLSX_NOTICE
    ws["A1"].font = Font(bold=True, color="9B1C1C", size=12)
    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
        cell.alignment = Alignment(horizontal="center")
    for i, col in enumerate(frame.columns, start=1):
        width = max([len(str(col))] + [len(str(v)) for v in frame[col].tolist()] + [8])
        ws.column_dimensions[ws.cell(row=3, column=i).column_letter].width = min(width * 1.6 + 2, 60)


def build_excel(data: pd.DataFrame, base_date: date, scope: str, filters: dict) -> bytes:
    """프로젝트목록/상태별집계/담당자별집계/분석조건 4개 시트의 xlsx 바이트를 만든다."""
    listing = to_display(data)

    counts = data["진행상태"].value_counts()
    order = STATUSES + [s for s in counts.index if s not in STATUSES]
    by_status = pd.DataFrame({"진행상태": order, "건수": [int(counts.get(s, 0)) for s in order]})
    by_status["비율(%)"] = (by_status["건수"] / max(len(data), 1) * 100).round(1)

    pivot = pd.crosstab(data["담당엔지니어"], data["진행상태"]).reindex(columns=order, fill_value=0)
    pivot["합계"] = pivot.sum(axis=1)
    pivot["지연 후보"] = data[data["일정 상태"] == S_LATE].groupby("담당엔지니어").size().reindex(pivot.index, fill_value=0)
    by_owner = pivot.reset_index()

    cond_rows = [
        ("기준일", base_date.strftime("%Y-%m-%d")),
        ("다운로드 일시", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("대상", scope),
        ("대상 건수", len(data)),
    ] + [(f"필터: {k}", v) for k, v in filters.items()]
    cond_rows.append(("일정 상태 기준", "완료=완료 / 잔여일수<0=지연 후보 / 0=오늘 마감 / 그 외=진행 예정 / 날짜 이상=일정 확인 필요"))
    conditions = pd.DataFrame(cond_rows, columns=["항목", "값"])

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _write_sheet(writer, "프로젝트목록", listing)
        _write_sheet(writer, "상태별집계", by_status)
        _write_sheet(writer, "담당자별집계", by_owner)
        _write_sheet(writer, "분석조건", conditions)
    return buf.getvalue()


def render_downloads(full: pd.DataFrame, fdf: pd.DataFrame, base_date: date, filters: dict):
    """전체 목록 / 현재 필터 결과 xlsx 다운로드 버튼(파일명에 오늘 날짜 포함)."""
    today = datetime.now().strftime("%Y%m%d")
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    no_filter = {k: "전체" for k in filters}
    c1, c2, _ = st.columns([1, 1, 2])
    c1.download_button(
        "전체 목록 .xlsx 다운로드",
        data=build_excel(full, base_date, "전체 목록", no_filter),
        file_name=f"standardbank_project_status_{today}.xlsx", mime=mime, key="dl_full",
    )
    c2.download_button(
        "현재 필터 결과 .xlsx 다운로드",
        data=build_excel(fdf, base_date, "현재 필터 결과", filters) if not fdf.empty else b"",
        file_name=f"standardbank_project_status_filtered_{today}.xlsx", mime=mime,
        key="dl_filtered", disabled=fdf.empty,
    )


# ───────────────────────── 메인 ─────────────────────────

CSS = f"""
<style>
.stApp {{ background: #FFFFFF; color: #1B2A41; }}
[data-testid="stHeader"] {{ background: #FFFFFF; }}
h1, h2, h3 {{ color: {NAVY}; }}
.notice {{
    background: #EAF1FF; color: {NAVY}; border-left: 5px solid {BLUE};
    padding: 10px 16px; border-radius: 6px; font-weight: 600; margin-bottom: 12px;
}}
.card {{
    background: #FFFFFF; border: 1px solid #D5E0F2; border-top: 4px solid {BLUE};
    border-radius: 8px; padding: 12px 14px; text-align: center;
}}
.card.total {{ border-top-color: {NAVY}; background: {NAVY}; }}
.card.total .label, .card.total .value {{ color: #FFFFFF; }}
.card .label {{ font-size: 0.85rem; color: #50607A; }}
.card .value {{ font-size: 1.9rem; font-weight: 700; color: {NAVY}; line-height: 1.3; }}
</style>
"""


def main():
    st.set_page_config(page_title="인증 프로젝트 대시보드", page_icon="📋", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(f'<div class="notice">{EDU_NOTICE}</div>', unsafe_allow_html=True)
    st.title("인증 프로젝트 접수 · 진행현황 대시보드")
    st.caption(SESSION_NOTICE)

    raw, source_name = None, None
    csv_path = first_existing(CSV_CANDIDATES)
    if csv_path is not None:
        raw, source_name = load_csv(csv_path), str(csv_path)
        st.caption(f"데이터 파일: {csv_path.relative_to(BASE_DIR)}")
    else:
        st.warning(
            "data 폴더에서 SB_projects_demo.csv 파일을 찾지 못했습니다. "
            "아래에서 CSV 파일을 직접 업로드해 주세요."
        )
        uploaded = st.file_uploader("의뢰 목록 CSV 업로드", type="csv")
        if uploaded is not None:
            try:
                raw, source_name = load_csv(uploaded), f"upload:{uploaded.name}:{uploaded.size}"
            except Exception as e:
                st.error(f"파일을 읽는 중 문제가 발생했습니다: {e}")
    if raw is None:
        st.stop()

    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        st.error(f"필수 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()
    raw = raw[REQUIRED_COLS].copy()
    raw["메모"] = raw["메모"].fillna("").astype(str)
    init_state(raw, source_name)

    base_date = st.date_input("기준일", value=DEFAULT_BASE_DATE)
    df = add_schedule_columns(st.session_state["df"], pd.Timestamp(base_date))

    # 사이드바: 필터 영역 → 구분선 → 카테고리 관리 영역
    fdf, chosen = render_filters(df)
    st.sidebar.divider()
    render_category_manager()

    render_resources()

    st.subheader("진행 요약")
    if fdf.empty:
        st.info("조건에 맞는 인증 프로젝트가 없습니다")
        render_downloads(df, fdf, base_date, chosen)
        st.stop()
    render_summary(fdf, base_date)

    st.subheader("현황 차트")
    render_charts(fdf)

    st.subheader("전체 의뢰 목록")
    st.dataframe(to_display(fdf).style.apply(highlight_late, axis=1), width="stretch", hide_index=True)
    render_downloads(df, fdf, base_date, chosen)

    render_detail(fdf)


if __name__ == "__main__":
    main()
