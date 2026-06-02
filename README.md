# 📊 DART 재무제표 일괄 다운로드 서비스 (Streamlit)

이 프로젝트는 **DART OpenAPI**와 **dart-fss** 라이브러리를 활용하여, 여러 기업의 공시정보를 실시간으로 조회하고 지정한 유형의 재무제표(BS, IS, CIS, CF)를 일괄 추출하여 엑셀 파일로 저장 및 다운로드할 수 있는 웹 어플리케이션입니다.

스트림릿 링크 : https://dart-fss-webui.streamlit.app/

---

## 🚀 주요 기능
1. **기업 검색 & 선택**
   - 코스피(KSP), 코스닥(KDQ), 코넥스(KNX), 기타(ETC) 시장 필터링 지원.
   - 키워드 검색 후 기업명을 **클릭**하여 즉시 선택 목록에 추가하거나 체크박스를 통한 일괄 추가 가능.
2. **공시정보 조회**
   - 선택된 기업들의 특정 기간 내 공시 보고서 목록 일괄 조회
   - 뷰어미리보기, 원본(XML)다운로드.
3. **재무제표 일괄 다운로드**
   - 선택한 모든 기업의 재무제표 개별 추출 진행.
   - 추출 후 **브라우저를 통한 다이렉트 엑셀 파일 다운로드** 지원.

## 📂 파일 구조
* `web_ui.py`: Streamlit 기반 프론트엔드 및 사용자 인터랙션 로직
* `DART_OpenAPI.py`: DART OpenAPI 파싱 및 엑셀 다운로드 라이브러리 모듈
* `requirements.txt`: Streamlit Cloud 환경 종속성 리스트

* thx to https://github.com/josw123/dart-fss
