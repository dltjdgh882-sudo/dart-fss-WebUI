"""
=============================================================================
DART 기업 재무제표 일괄 검색 / 다운로드 Web UI  (Streamlit)
=============================================================================
search_dart_corp.py 를 import 하여 GUI 환경에서 다음 기능을 제공합니다.

  1. 기업 검색 & 선택  – 키워드 검색 → 결과에서 체크 → 선택 목록에 추가
  2. 공시정보 조회      – 선택 기업들의 재무보고서 목록 일괄 조회
  3. 일괄 다운로드      – 선택 기업들의 재무제표를 순차 추출 → 기업별 엑셀 저장

실행 방법:
    streamlit run web_ui.py
=============================================================================
"""

import importlib
import os
import subprocess
import sys
import traceback
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Streamlit 자동 설치
# ---------------------------------------------------------------------------
try:
    import streamlit as st
except Exception:
    print("streamlit 패키지를 설치합니다…", file=sys.stderr)
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "streamlit"])
        st = importlib.import_module("streamlit")
        print("streamlit 설치 완료.", file=sys.stderr)
    except Exception as _err:
        print(f"streamlit 자동 설치 실패: {_err}", file=sys.stderr)
        print("  pip install streamlit  명령으로 직접 설치하세요.", file=sys.stderr)
        raise

# ---------------------------------------------------------------------------
# search_dart_corp.py 의 함수 import
# ---------------------------------------------------------------------------
from search_dart_corp import (
    DEFAULT_API_KEY,
    download_corp_list,
    extract_financial_statement,
    get_available_financial_reports,
    search_corp_keyword,
    set_api_key,
    download_original_document,
)

# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------
FS_TYPE_LABELS = {
    "bs": "재무상태표 (BS)",
    "is": "손익계산서 (IS)",
    "cis": "포괄손익계산서 (CIS)",
    "cf": "현금흐름표 (CF)",
}
REPORT_TYPE_LABELS = {
    "annual": "연간 (Annual)",
    "half": "반기 (Half)",
    "quarter": "분기 (Quarter)",
}
MARKET_LABELS = {
    "KSP": "코스피 (KOSPI)",
    "KDQ": "코스닥 (KOSDAQ)",
    "KNX": "코넥스 (KONEX)",
    "ETC": "기타 (ETC)",
}
# UI 키 → DART API corp_cls 코드 매핑
MARKET_TO_DART_CODE = {
    "KSP": "Y",
    "KDQ": "K",
    "KNX": "N",
    "ETC": "E",
}
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SAVE_DIR = "fsdata"

# ---------------------------------------------------------------------------
# Session State 초기화
# ---------------------------------------------------------------------------

def _init_session():
    """Streamlit session_state 기본값 설정."""
    defaults = {
        "api_key": os.environ.get("DART_API_KEY") or DEFAULT_API_KEY or "",
        "search_keyword": "",
        "search_results": [],          # List[Dict]
        "selected_corps": [],           # List[Dict]  — 선택된 기업 목록
        "batch_logs": [],               # List[str]   — 일괄 다운로드 로그
        "batch_running": False,
        "batch_files": [],              # List[Tuple[str, str]]  — (표시명, 절대경로)
        "reports_cache": {},            # corp_code -> List[Dict]
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ---------------------------------------------------------------------------
# 유틸 함수
# ---------------------------------------------------------------------------

def _corp_label(corp: Dict[str, Any]) -> str:
    """기업 표시 문자열: 기업명 (기업코드) [종목코드]"""
    stock = corp.get("stock_code") or ""
    stock_suffix = f"  [{stock}]" if stock else ""
    return f"{corp['corp_name']} ({corp['corp_code']}){stock_suffix}"


def _safe_filename(name: str) -> str:
    """파일명으로 사용 가능하도록 특수문자 제거."""
    import re
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def _add_corps(codes: List[str]):
    """검색 결과에서 선택된 기업 코드를 selected_corps 에 추가."""
    existing = {c["corp_code"] for c in st.session_state.selected_corps}
    for corp in st.session_state.search_results:
        if corp["corp_code"] in codes and corp["corp_code"] not in existing:
            st.session_state.selected_corps.append(corp)
            existing.add(corp["corp_code"])


def _remove_corp(code: str):
    """선택 목록에서 기업 제거."""
    st.session_state.selected_corps = [
        c for c in st.session_state.selected_corps if c["corp_code"] != code
    ]
    st.session_state.reports_cache.pop(code, None)


# ---------------------------------------------------------------------------
# 사이드바 — API 키 / 기업 검색 / 선택
# ---------------------------------------------------------------------------

def _render_sidebar():
    with st.sidebar:
        st.markdown("## ⚙️ 설정")

        # --- API Key ---
        st.text_input(
            "DART API 키",
            key="api_key",
            type="password",
            help="미입력 시 기본 API 키가 사용됩니다.",
        )

        st.divider()

        # --- 기업 검색 ---
        st.markdown("## 🔍 기업 검색")
        st.text_input("검색 키워드", key="search_keyword", placeholder="예: 삼성")

        # 시장 필터 — 고정 라벨 + 체크박스 (기본값: 모두 선택)
        st.markdown("**시장 필터**")
        market_cols = st.columns(len(MARKET_LABELS))
        active_markets = []
        for idx, (code, label) in enumerate(MARKET_LABELS.items()):
            with market_cols[idx]:
                if st.checkbox(code, value=True, key=f"mkt_{code}", help=label):
                    active_markets.append(code)

        if st.button("🔎 검색", use_container_width=True):
            _do_search(active_markets)

        # --- 검색 결과 표시 & 선택 (스크롤 리스트 + 체크박스) ---
        if st.session_state.search_results:
            st.markdown(f"**{len(st.session_state.search_results)}건 검색됨**")

            # 체크박스 상태를 session_state 에서 관리
            if "search_checks" not in st.session_state:
                st.session_state.search_checks = {}

            # 검색 결과가 바뀌면 체크 상태 초기화
            result_codes = {c["corp_code"] for c in st.session_state.search_results}
            old_codes = set(st.session_state.search_checks.keys())
            if result_codes != old_codes:
                st.session_state.search_checks = {c: False for c in result_codes}

            # 스크롤 가능한 컨테이너
            with st.container(height=320):
                # 헤더
                h1, h2, h3 = st.columns([2, 3, 1])
                h1.markdown("**기업코드**")
                h2.markdown("**기업명** *(클릭→추가)*")
                h3.markdown("**선택**")

                for corp in st.session_state.search_results:
                    c1, c2, c3 = st.columns([2, 3, 1])
                    c1.code(corp["corp_code"], language=None)
                    # 기업명 클릭 시 즉시 추가
                    if c2.button(
                        corp["corp_name"],
                        key=f"add1_{corp['corp_code']}",
                        use_container_width=True,
                    ):
                        _add_corps([corp["corp_code"]])
                        st.toast(f"✅ {corp['corp_name']} 추가됨")
                        st.rerun()
                    # 일괄 추가용 체크박스
                    st.session_state.search_checks[corp["corp_code"]] = c3.checkbox(
                        "추가",
                        value=st.session_state.search_checks.get(corp["corp_code"], False),
                        key=f"chk_{corp['corp_code']}",
                        label_visibility="collapsed",
                    )

            # 선택된 코드 수집
            checked_codes = [
                code for code, checked in st.session_state.search_checks.items() if checked
            ]
            st.caption(f"{len(checked_codes)}개 선택됨")

            if st.button("➕ 선택 기업 일괄 추가", use_container_width=True):
                if not checked_codes:
                    st.warning("추가할 기업을 체크하세요.")
                else:
                    _add_corps(checked_codes)
                    # 체크 초기화
                    st.session_state.search_checks = {c: False for c in result_codes}
                    st.success(f"{len(checked_codes)}개 기업 추가 완료")
                    st.rerun()

        st.divider()
        st.caption(f"선택된 기업 수: **{len(st.session_state.selected_corps)}**")


def _do_search(active_markets: List[str] = None):
    """사이드바 검색 버튼 콜백."""
    kw = st.session_state.search_keyword.strip()
    if not kw:
        st.sidebar.error("검색 키워드를 입력하세요.")
        return
    if active_markets is None:
        active_markets = list(MARKET_LABELS.keys())
    # UI 키를 DART API corp_cls 코드로 변환
    dart_codes = [MARKET_TO_DART_CODE.get(m, m) for m in active_markets]
    try:
        set_api_key(st.session_state.api_key or None)
        market = "".join(sorted(set(dart_codes)))
        results = search_corp_keyword(
            kw,
            api_key=st.session_state.api_key or None,
            market=market,
            max_results=100,
            force_update=False,
        )
        st.session_state.search_results = results or []
        if not results:
            st.sidebar.warning(f'"{kw}" 에 해당하는 기업이 없습니다.')
    except Exception as err:
        st.sidebar.error(f"검색 실패: {err}")


# ---------------------------------------------------------------------------
# 탭 1 — 선택 기업 관리
# ---------------------------------------------------------------------------

def _render_tab_selected():
    st.markdown("### 📋 선택된 기업 목록")

    corps = st.session_state.selected_corps
    if not corps:
        st.info("사이드바에서 기업을 검색한 뒤 추가하세요.")
        return

    # 전체 초기화 버튼
    col_top1, col_top2, _ = st.columns([1, 1, 4])
    with col_top1:
        if st.button("🗑️ 전체 초기화"):
            st.session_state.selected_corps = []
            st.session_state.reports_cache = {}
            st.rerun()
    with col_top2:
        st.markdown(f"**총 {len(corps)}개 기업**")

    # 기업 목록 테이블
    for idx, corp in enumerate(corps):
        col_name, col_code, col_stock, col_cls, col_del = st.columns(
            [3, 2, 2, 1, 1]
        )
        col_name.markdown(f"**{corp['corp_name']}**")
        col_code.code(corp["corp_code"])
        col_stock.write(corp.get("stock_code") or "—")
        col_cls.write(corp.get("corp_cls") or "—")
        if col_del.button("❌", key=f"del_{corp['corp_code']}"):
            _remove_corp(corp["corp_code"])
            st.rerun()



def _download_report_ui(corp_name: str, rcept_no: str, report_nm: str):
    """공시 보고서 원본을 다운로드하여 세션 상태에 저장하는 UI 헬퍼 함수."""
    try:
        set_api_key(st.session_state.api_key or None)
        save_dir_name = st.session_state.get("batch_save_dir", DEFAULT_SAVE_DIR)
        save_dir = os.path.join(SCRIPT_DIR, save_dir_name)
        os.makedirs(save_dir, exist_ok=True)
        
        safe_corp_name = _safe_filename(corp_name)
        safe_report_nm = _safe_filename(report_nm)
        save_path = os.path.join(save_dir, f"{safe_corp_name}_{safe_report_nm}_{rcept_no}.zip")
        
        with st.spinner("원본 문서 다운로드 중..."):
            saved_filepath = download_original_document(
                rcept_no=rcept_no,
                save_path=save_path,
                api_key=st.session_state.api_key or None
            )
        
        if "downloaded_docs" not in st.session_state:
            st.session_state.downloaded_docs = {}
        st.session_state.downloaded_docs[rcept_no] = saved_filepath
        st.toast(f"✅ 다운로드 완료: {os.path.basename(saved_filepath)}")
        st.rerun()
    except Exception as err:
        st.error(f"다운로드 실패: {err}")


# ---------------------------------------------------------------------------
# 탭 2 — 공시정보 조회
# ---------------------------------------------------------------------------

def _render_tab_reports():
    st.markdown("### 📊 공시정보 일괄 조회")

    corps = st.session_state.selected_corps
    if not corps:
        st.info("먼저 기업을 선택하세요.")
        return

    col1, col2 = st.columns(2)
    with col1:
        bgn = st.date_input(
            "조회 시작일",
            value=date(date.today().year - 1, 1, 1),
            key="report_bgn",
        )
    with col2:
        end = st.date_input(
            "조회 종료일",
            value=date.today(),
            key="report_end",
        )

    if st.button("📥 공시 보고서 조회", use_container_width=True):
        bgn_str = bgn.strftime("%Y%m%d")
        end_str = end.strftime("%Y%m%d")

        progress_bar = st.progress(0, text="조회 준비 중…")
        total = len(corps)

        for i, corp in enumerate(corps):
            code = corp["corp_code"]
            name = corp["corp_name"]
            progress_bar.progress(
                (i) / total,
                text=f"조회 중: {name} ({i+1}/{total})",
            )
            try:
                set_api_key(st.session_state.api_key or None)
                reports = get_available_financial_reports(
                    code, bgn_str, end_str,
                    api_key=st.session_state.api_key or None,
                )
                st.session_state.reports_cache[code] = reports or []
            except Exception as err:
                st.session_state.reports_cache[code] = []
                st.warning(f"{name}: 조회 실패 — {err}")

        progress_bar.progress(1.0, text="조회 완료!")

    # 결과 표시
    if st.session_state.reports_cache:
        for corp in corps:
            code = corp["corp_code"]
            name = corp["corp_name"]
            reports = st.session_state.reports_cache.get(code)
            if reports is None:
                continue
            with st.expander(f"📄 {name} ({code}) — {len(reports)}건", expanded=False):
                if not reports:
                    st.write("해당 기간 보고서 없음")
                else:
                    import pandas as pd
                    df = pd.DataFrame(reports)
                    st.dataframe(df, use_container_width=True)
                    
                    st.markdown("#### 📂 공시 원본보고서 개별 다운로드")
                    st.caption("아래 목록에서 원문 보고서(ZIP)를 다운로드해 로컬 서버에 저장하고, 브라우저 다운로드 버튼을 활성화합니다.")
                    
                    for idx, report in enumerate(reports):
                        rcept_no = report.get("rcept_no")
                        report_nm = report.get("report_nm")
                        bgn_de = report.get("bgn_de", "N/A")
                        
                        col_r1, col_r2, col_r3 = st.columns([5, 2, 2])
                        col_r1.markdown(f"**{idx+1}. {report_nm}**")
                        col_r2.caption(f"접수번호: `{rcept_no}`\n일자: {bgn_de}")
                        
                        if "downloaded_docs" not in st.session_state:
                            st.session_state.downloaded_docs = {}
                            
                        btn_key = f"btn_dl_{rcept_no}"
                        
                        if rcept_no in st.session_state.downloaded_docs:
                            saved_filepath = st.session_state.downloaded_docs[rcept_no]
                            if os.path.exists(saved_filepath):
                                with open(saved_filepath, "rb") as f:
                                    col_r3.download_button(
                                        label="⬇️ 브라우저 저장",
                                        data=f.read(),
                                        file_name=os.path.basename(saved_filepath),
                                        mime="application/zip",
                                        key=f"web_dl_{rcept_no}",
                                        use_container_width=True,
                                    )
                            else:
                                del st.session_state.downloaded_docs[rcept_no]
                                if col_r3.button("📥 원본 다운로드", key=btn_key, use_container_width=True):
                                    _download_report_ui(name, rcept_no, report_nm)
                        else:
                            if col_r3.button("📥 원본 다운로드", key=btn_key, use_container_width=True):
                                _download_report_ui(name, rcept_no, report_nm)


# ---------------------------------------------------------------------------
# 탭 3 — 일괄 다운로드 (핵심 기능)
# ---------------------------------------------------------------------------

def _render_tab_batch():
    st.markdown("### ⬇️ 재무제표 일괄 다운로드")

    corps = st.session_state.selected_corps
    if not corps:
        st.info("먼저 기업을 선택하세요.")
        return

    # --- 옵션 영역 ---
    st.markdown("#### 추출 옵션")

    opt_col1, opt_col2 = st.columns(2)

    with opt_col1:
        fs_types = st.multiselect(
            "재무제표 유형",
            options=list(FS_TYPE_LABELS.keys()),
            format_func=lambda k: FS_TYPE_LABELS[k],
            default=list(FS_TYPE_LABELS.keys()),
            key="batch_fs_types",
        )
        report_tp = st.selectbox(
            "보고서 유형",
            options=list(REPORT_TYPE_LABELS.keys()),
            format_func=lambda k: REPORT_TYPE_LABELS[k],
            key="batch_report_tp",
        )
        separate = st.toggle("개별재무제표 (OFF: 연결재무제표)", value=False, key="batch_separate")

    with opt_col2:
        bgn = st.date_input(
            "시작일",
            value=date(date.today().year - 1, 1, 1),
            key="batch_bgn",
        )
        end = st.date_input(
            "종료일",
            value=date.today(),
            key="batch_end",
        )
        lang = st.selectbox(
            "언어", ["ko", "en"],
            format_func=lambda v: "한국어" if v == "ko" else "English",
            key="batch_lang",
        )
        save_dir_input = st.text_input(
            "저장 폴더명",
            value=DEFAULT_SAVE_DIR,
            key="batch_save_dir",
            help="스크립트 위치 기준 하위 폴더명 (예: fsdata, output 등)",
        )
        # 스크립트 위치 기준 절대경로로 변환
        save_dir = os.path.join(SCRIPT_DIR, save_dir_input)
        st.caption(f"📁 저장 경로: `{save_dir}`")

    st.divider()

    # --- 대상 기업 미리보기 ---
    st.markdown("#### 대상 기업")
    corp_names = [f"**{c['corp_name']}** `{c['corp_code']}`" for c in corps]
    st.markdown(" · ".join(corp_names))
    st.caption(f"총 {len(corps)}개 기업의 재무제표를 추출합니다.")

    st.divider()

    # --- 저장 폴더 존재 확인 ---
    dir_exists = os.path.isdir(save_dir)
    if not dir_exists:
        st.warning(f"📂 저장 폴더가 존재하지 않습니다: `{save_dir}`")
        if st.button("📁 폴더 생성", key="create_dir_btn"):
            try:
                os.makedirs(save_dir, exist_ok=True)
                st.success(f"폴더 생성 완료: `{save_dir}`")
                st.rerun()
            except Exception as e:
                st.error(f"폴더 생성 실패: {e}")

    # --- 일괄 추출 실행 ---
    if st.button(
        "🚀 일괄 추출 시작",
        use_container_width=True,
        type="primary",
        disabled=not dir_exists,
    ):
        _run_batch_extract(
            corps=corps,
            fs_types=tuple(fs_types) if fs_types else ("bs", "is", "cis", "cf"),
            report_tp=report_tp,
            separate=separate,
            bgn_de=bgn.strftime("%Y%m%d"),
            end_de=end.strftime("%Y%m%d"),
            lang=lang,
            save_dir=save_dir,
        )

    # --- 이전 로그 표시 ---
    if st.session_state.batch_logs:
        st.markdown("#### 실행 로그")
        for log in st.session_state.batch_logs:
            st.markdown(log)

    # --- 다운로드 버튼 ---
    if st.session_state.batch_files:
        st.divider()
        st.markdown("#### 📥 파일 다운로드")
        st.caption("추출된 재무제표 파일을 브라우저로 다운로드합니다.")
        for display_name, fpath in st.session_state.batch_files:
            if os.path.isfile(fpath):
                with open(fpath, "rb") as fp:
                    st.download_button(
                        label=f"⬇️ {display_name}",
                        data=fp.read(),
                        file_name=os.path.basename(fpath),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{os.path.basename(fpath)}",
                    )


def _run_batch_extract(
    corps: List[Dict],
    fs_types: Tuple[str, ...],
    report_tp: str,
    separate: bool,
    bgn_de: str,
    end_de: str,
    lang: str,
    save_dir: str,
):
    """선택된 모든 기업에 대해 재무제표를 순차 추출 → 개별 엑셀 저장."""

    os.makedirs(save_dir, exist_ok=True)
    saved_files: List[Tuple[str, str]] = []  # (표시명, 절대경로)

    total = len(corps)
    logs: List[str] = []
    success_count = 0
    fail_count = 0

    progress_bar = st.progress(0, text="준비 중…")
    status_area = st.empty()
    log_area = st.container()

    set_api_key(st.session_state.api_key or None)

    for i, corp in enumerate(corps):
        code = corp["corp_code"]
        name = corp["corp_name"]

        progress_bar.progress(
            i / total,
            text=f"추출 중: {name} ({i+1}/{total})",
        )
        status_area.info(f"⏳ **{name}** ({code}) 재무제표 추출 중…")

        try:
            fs = extract_financial_statement(
                corp_code=code,
                bgn_de=bgn_de,
                end_de=end_de if end_de else None,
                fs_tp=fs_types,
                separate=separate,
                report_tp=report_tp,
                lang=lang,
                api_key=st.session_state.api_key or None,
            )

            # 저장
            safe_name = _safe_filename(name)
            filename = f"{safe_name}_{code}.xlsx"
            filepath = os.path.join(save_dir, filename)

            if fs is not None and hasattr(fs, "save"):
                fs.save(filepath)
                saved_files.append((f"{name} ({code})", filepath))
                msg = f"✅ **{name}** → `{filepath}`"
                success_count += 1
            elif fs is not None:
                # save() 가 없는 경우 — DataFrame 류라면 to_excel 시도
                try:
                    import pandas as pd
                    if isinstance(fs, pd.DataFrame):
                        fs.to_excel(filepath, index=False)
                        saved_files.append((f"{name} ({code})", filepath))
                        msg = f"✅ **{name}** → `{filepath}` (DataFrame)"
                        success_count += 1
                    else:
                        msg = f"⚠️ **{name}** — 추출 완료, 저장 불가 (save/to_excel 미지원)"
                        fail_count += 1
                except Exception:
                    msg = f"⚠️ **{name}** — 추출 완료, 저장 불가"
                    fail_count += 1
            else:
                msg = f"⚠️ **{name}** — 추출 결과가 비어있습니다."
                fail_count += 1

        except Exception as err:
            msg = f"❌ **{name}** — 오류: {err}"
            fail_count += 1

        logs.append(msg)
        with log_area:
            st.markdown(msg)

    # 완료
    progress_bar.progress(1.0, text="완료!")

    # 요약
    summary = f"### 📊 일괄 추출 완료\n\n"
    summary += f"- **성공**: {success_count}건\n"
    summary += f"- **실패/건너뜀**: {fail_count}건\n"
    summary += f"- **저장 폴더**: `{os.path.abspath(save_dir)}`\n"

    status_area.success(summary)
    st.session_state.batch_logs = logs
    st.session_state.batch_files = saved_files


# ---------------------------------------------------------------------------
# 메인 레이아웃
# ---------------------------------------------------------------------------

def main():
    _init_session()

    st.set_page_config(
        page_title="DART 재무제표 일괄 다운로드",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # --- 페이지 헤더 ---
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.5rem; }
        [data-testid="stSidebar"] { min-width: 410px; width: 410px; }
        /* 사이드바 접기(collapse) 버튼 숨김 — 항상 표시 */
        button[kind="headerNoPadding"],
        [data-testid="stSidebar"] button[data-testid="stBaseButton-headerNoPadding"],
        [data-testid="collapsedControl"] {
            display: none !important;
            visibility: hidden !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("📊 DART 재무제표 일괄 다운로드")
    st.caption("search_dart_corp.py 기반 · 여러 기업의 재무제표를 한 번에 검색·추출·저장합니다.")

    # --- 사이드바 ---
    _render_sidebar()

    # --- 메인 탭 ---
    tab1, tab2, tab3 = st.tabs([
        "📋 선택 기업 관리",
        "📊 공시정보 조회",
        "⬇️ 일괄 다운로드",
    ])

    with tab1:
        _render_tab_selected()

    with tab2:
        _render_tab_reports()

    with tab3:
        _render_tab_batch()

    # --- 푸터 ---
    st.divider()
    st.caption("DART OpenAPI · dart-fss · Streamlit")


if __name__ == "__main__":
    main()
