"""
=============================================================================
DART 기업 공시정보 일괄 검색 / 다운로드 Web UI  (Streamlit)
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
# DART API 공시유형 코드 매핑
PBLNTF_TY_LABELS = {
    "전체 (All)": None,
    "정기공시 (A)": "A",
    "주요사항보고 (B)": "B",
    "발행공시 (C)": "C",
    "지분공시 (D)": "D",
    "기타공시 (E)": "E",
    "외부감사관련 (F)": "F",
    "펀드공시 (G)": "G",
    "자산유동화 (H)": "H",
    "거래소공시 (I)": "I",
    "공정위공시 (J)": "J",
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
        "selected_corps": [],           # List[Dict]  — 검색대상 기업 선택
        "batch_logs": [],               # List[str]   — 일괄 다운로드 로그
        "batch_running": False,
        "batch_files": [],              # List[Tuple[str, str]]  — (표시명, 절대경로)
        "reports_cache": {},            # corp_code -> List[Dict]
        "selected_reports_by_corp": {}, # corp_code -> pd.DataFrame
        "xml_batch_zip_path": None,
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


def _download_report_ui(corp_name: str, corp_code: str, rcept_no: str, report_nm: str):
    """공시 보고서 원본을 다운로드하여 세션 상태에 저장하는 UI 헬퍼 함수."""
    try:
        set_api_key(st.session_state.api_key or None)
        save_dir = os.path.join(SCRIPT_DIR, DEFAULT_SAVE_DIR)
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
        st.session_state.active_expander = corp_code
        st.session_state[f"expander_{corp_code}"] = True
        st.toast(f"✅ 다운로드 완료: {os.path.basename(saved_filepath)}")
        st.rerun()
    except Exception as err:
        st.error(f"다운로드 실패: {err}")


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
    st.markdown("### 📋 검색대상 기업 선택")

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

    # 기업 목록 테이블 헤더 추가
    st.divider()
    h_name, h_code, h_stock, h_cls, h_del = st.columns([3, 2, 2, 1, 1])
    h_name.markdown("**🏢 기업명**")
    h_code.markdown("**🔑 DART 코드**")
    h_stock.markdown("**📈 종목코드**")
    h_cls.markdown("**🏛️ 시장구분**")
    h_del.markdown("**❌ 삭제**")
    st.divider()

    # 기업 목록 테이블 데이터 행
    for idx, corp in enumerate(corps):
        col_name, col_code, col_stock, col_cls, col_del = st.columns(
            [3, 2, 2, 1, 1]
        )
        col_name.markdown(f"**{corp['corp_name']}**")
        col_code.code(corp["corp_code"])
        col_stock.write(corp.get("stock_code") or "—")
        # 시장구분 코드 변환 (Y->KSP, K->KDQ, N->KNX, E->ETC)
        raw_cls = corp.get("corp_cls") or ""
        cls_mapping = {"Y": "KSP", "K": "KDQ", "N": "KNX", "E": "ETC"}
        display_cls = cls_mapping.get(raw_cls, raw_cls) or "—"
        col_cls.write(display_cls)
        if col_del.button("❌", key=f"del_{corp['corp_code']}", use_container_width=True):
            _remove_corp(corp["corp_code"])
            st.rerun()


# ---------------------------------------------------------------------------
# 탭 2 — 공시정보 조회
# ---------------------------------------------------------------------------

def _render_tab_reports():
    if "selected_reports_by_corp" not in st.session_state:
        st.session_state.selected_reports_by_corp = {}
    if "xml_batch_zip_path" not in st.session_state:
        st.session_state.xml_batch_zip_path = None

    st.markdown("### 📊 공시정보 원본 조회")

    corps = st.session_state.selected_corps
    if not corps:
        st.info("먼저 기업을 선택하세요.")
        return

    col1, col2, col3 = st.columns(3)
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
    with col3:
        pblntf_ty_sel = st.selectbox(
            "공시 유형",
            options=list(PBLNTF_TY_LABELS.keys()),
            key="report_pblntf_ty",
            help="검색할 공시 서류의 대분류 유형입니다.",
        )
        pblntf_ty_code = PBLNTF_TY_LABELS[pblntf_ty_sel]

    if st.button("📥 공시 보고서 조회(기업별 최대 100건/1회)", use_container_width=True):
        st.session_state.active_expander = None
        st.session_state.xml_batch_zip_path = None
        st.session_state.selected_reports_by_corp = {}
        for key in list(st.session_state.keys()):
            if key.startswith("df_editor_") or key.startswith("expander_"):
                del st.session_state[key]
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
                    pblntf_ty=pblntf_ty_code,
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
            
            # key 매개변수를 지정하여 expander의 열림/닫힘 상태를 st.session_state[f"expander_{code}"]와 자동으로 동기화합니다.
            # Rerun(체크박스 클릭 등) 시에도 사용자가 조작한 상태가 완전히 그대로 보존됩니다.
            expander_ctx = st.expander(
                f"📄 {name} ({code}) — {len(reports)}건", 
                key=f"expander_{code}"
            )
            with expander_ctx:
                if not reports:
                    st.write("해당 기간 보고서 없음")
                else:
                    import pandas as pd
                    # 복사본을 만들어 원본 데이터 훼손 방지 및 공시 뷰어 링크 추가
                    df = pd.DataFrame(reports).copy()
                    df["공시 뷰어"] = df["rcept_no"].apply(lambda rcp: f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp}")
                    
                    # 불필요한 컬럼 삭제 (보고서코드, 시작일, 종료일)
                    cols_to_drop = ["reprt_code", "bgn_de", "end_de"]
                    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors="ignore")
                    
                    # 1. 뷰어 바로가기 오른쪽에 체크박스 열(선택) 추가
                    df["선택"] = False
                    
                    st.markdown("#### 📂 공시 목록 및 원문 다운로드 선택")
                    
                    # 데이터프레임 컬럼 시각화 커스텀 설정 (표 내부에 클릭 뷰어 삽입 및 체크박스 열 추가)
                    column_config = {
                        "공시 뷰어": st.column_config.LinkColumn(
                            "🌐 뷰어 바로가기",
                            help="클릭하시면 금융감독원 DART 공시 뷰어 새 창이 열립니다.",
                            validate="^https://.*",
                            display_text="🌐 열기"
                        ),
                        "rcept_no": st.column_config.TextColumn("접수번호"),
                        "report_nm": st.column_config.TextColumn("보고서명"),
                        "corp_name": st.column_config.TextColumn("회사명"),
                        "선택": st.column_config.CheckboxColumn(
                            "선택",
                            help="체크 시 하단에서 일괄 다운로드 할 수 있습니다.",
                            default=False
                        )
                    }
                    
                    # st.data_editor를 사용하여 체크박스 선택 허용
                    edited_df = st.data_editor(
                        df,
                        column_config=column_config,
                        use_container_width=True,
                        disabled=["rcept_no", "report_nm", "corp_name", "공시 뷰어"],
                        key=f"df_editor_{code}"
                    )
                    
                    # 세션에 편집 정보 저장
                    st.session_state.selected_reports_by_corp[code] = edited_df

        # --- 공시정보 원문(XML) 일괄 다운로드 영역 ---
        st.divider()
        st.markdown("### 📥 공시정보 원문(XML) 일괄 다운로드")
        st.caption("위의 공시 목록 표에서 **'선택'** 체크박스를 체크한 보고서들의 원문(XML) ZIP 파일들을 한 번에 일괄 다운로드하여 하나의 ZIP 파일로 결합합니다.")
        
        # 선택된 보고서 목록 수집
        selected_items = []
        if "selected_reports_by_corp" in st.session_state:
            for c_code, ed_df in st.session_state.selected_reports_by_corp.items():
                if ed_df is not None and "선택" in ed_df.columns:
                    checked_rows = ed_df[ed_df["선택"] == True]
                    for _, row in checked_rows.iterrows():
                        selected_items.append({
                            "corp_code": c_code,
                            "corp_name": row.get("corp_name"),
                            "rcept_no": row.get("rcept_no"),
                            "report_nm": row.get("report_nm")
                        })
                        
        st.write(f"선택된 보고서 개수: **{len(selected_items)}** 개")
        
        # 다운로드 시작 버튼 및 진행률 표시를 위한 레이아웃
        col_btn, col_status = st.columns([1, 2])
        
        with col_btn:
            btn_disabled = len(selected_items) == 0
            start_download = st.button(
                "🚀 원문 일괄 다운로드 시작",
                use_container_width=True,
                disabled=btn_disabled,
                type="primary",
                key="xml_batch_download_btn"
            )
            
        status_placeholder = col_status.empty()
        
        if start_download:
            status_placeholder.info("⏳ 다운로드 준비 중...")
            try:
                # 1. 루트 폴더에 "temp" 폴더 생성
                TEMP_DIR = os.path.join(SCRIPT_DIR, "temp")
                import shutil
                if os.path.exists(TEMP_DIR):
                    try:
                        shutil.rmtree(TEMP_DIR)
                    except Exception:
                        for filename in os.listdir(TEMP_DIR):
                            file_path = os.path.join(TEMP_DIR, filename)
                            try:
                                if os.path.isfile(file_path) or os.path.islink(file_path):
                                    os.unlink(file_path)
                                elif os.path.isdir(file_path):
                                    shutil.rmtree(file_path)
                            except Exception:
                                pass
                os.makedirs(TEMP_DIR, exist_ok=True)
                
                # 2. 체크된 보고서 원본 다운로드 수행
                total_count = len(selected_items)
                for idx, item in enumerate(selected_items):
                    c_name = item["corp_name"]
                    r_name = item["report_nm"]
                    rcp_no = item["rcept_no"]
                    
                    status_placeholder.info(f"⏳ [{idx + 1}/{total_count}] {c_name} - {r_name} 다운로드 중...")
                    
                    safe_c_name = _safe_filename(c_name)
                    safe_r_name = _safe_filename(r_name)
                    filename = f"{safe_c_name}_{safe_r_name}_{rcp_no}.zip"
                    file_save_path = os.path.join(TEMP_DIR, filename)
                    
                    # search_dart_corp의 download_original_document 호출
                    download_original_document(
                        rcept_no=rcp_no,
                        save_path=file_save_path,
                        api_key=st.session_state.api_key or None
                    )
                
                # 3. zip 파일로 압축
                status_placeholder.info("📦 ZIP 파일 결합 압축 진행 중...")
                import zipfile
                final_zip_path = os.path.join(SCRIPT_DIR, "original_documents_batch.zip")
                if os.path.exists(final_zip_path):
                    try:
                        os.remove(final_zip_path)
                    except Exception:
                        pass
                        
                with zipfile.ZipFile(final_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root_dir, _, files in os.walk(TEMP_DIR):
                        for file in files:
                            file_path = os.path.join(root_dir, file)
                            zipf.write(file_path, os.path.relpath(file_path, TEMP_DIR))
                
                st.session_state.xml_batch_zip_path = final_zip_path
                status_placeholder.success("✅ 일괄 다운로드 및 압축 성공!")
                st.rerun()
            except Exception as err:
                status_placeholder.error(f"❌ 다운로드 오류 발생: {err}")
                
        # 최종 zip파일 웹 브라우저 다운로드 버튼 노출
        if st.session_state.get("xml_batch_zip_path") and os.path.isfile(st.session_state.xml_batch_zip_path):
            st.markdown("#### 🎁 다운로드 준비 완료")
            with open(st.session_state.xml_batch_zip_path, "rb") as f:
                st.download_button(
                    label="💾 압축파일 다운로드 (original_documents_batch.zip)",
                    data=f.read(),
                    file_name="original_documents_batch.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

# ---------------------------------------------------------------------------
# 탭 3 — 일괄 추출
# ---------------------------------------------------------------------------

def _render_tab_batch():
    st.markdown("### ⬇️ 재무제표 일괄 추출")

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
        # 기본 저장 경로 고정 (루트 폴더 내 fsdata)
        save_dir = os.path.join(SCRIPT_DIR, DEFAULT_SAVE_DIR)

    st.divider()

    # --- 대상 기업 미리보기 ---
    st.markdown("#### 대상 기업")
    corp_names = [f"**{c['corp_name']}** `{c['corp_code']}`" for c in corps]
    st.markdown(" · ".join(corp_names))
    st.caption(f"총 {len(corps)}개 기업의 재무제표를 추출합니다.")

    st.divider()

    # --- 일괄 추출 실행 ---
    if st.button(
        "🚀 일괄 추출 시작",
        use_container_width=True,
        type="primary",
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

    # 완료
    progress_bar.progress(1.0, text="완료!")

    # 요약
    summary = f"### 📊 일괄 추출 완료\n\n"
    summary += f"- **성공**: {success_count}건\n"
    summary += f"- **실패/건너뜀**: {fail_count}건\n"

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
    st.title("📊 DART 공시정보 검색 / 재무제표 추출")
    st.caption("Dart Open API를 활용 · 여러 기업의 공시자료 / 재무제표를 한 번에 검색·추출·저장합니다.")

    # --- 사이드바 ---
    _render_sidebar()

    # --- 메인 탭 ---
    tab1, tab2, tab3 = st.tabs([
        "📋 검색대상 기업 선택",
        "📊 공시정보 원본 조회",
        "⬇️ 재무제표 일괄 추출",
    ])

    with tab1:
        _render_tab_selected()

    with tab2:
        _render_tab_reports()

    with tab3:
        _render_tab_batch()

    # --- 푸터 ---
    st.divider()
    st.caption("thx to DART OpenAPI · dart-fss · Streamlit")


if __name__ == "__main__":
    main()
