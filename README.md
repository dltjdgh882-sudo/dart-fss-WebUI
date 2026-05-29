# 📊 DART 재무제표 일괄 다운로드 서비스 (Streamlit)

이 프로젝트는 **DART OpenAPI**와 **dart-fss** 라이브러리를 활용하여, 여러 기업의 공시정보를 실시간으로 조회하고 지정한 유형의 재무제표(BS, IS, CIS, CF)를 일괄 추출하여 엑셀 파일로 저장 및 다운로드할 수 있는 웹 어플리케이션입니다.

---

## 🚀 주요 기능
1. **기업 검색 & 선택** (사이드바)
   - 코스피(KSP), 코스닥(KDQ), 코넥스(KNX), 기타(ETC) 시장 필터링 지원.
   - 키워드 검색 후 기업명을 **클릭**하여 즉시 선택 목록에 추가하거나 체크박스를 통한 일괄 추가 가능.
2. **공시정보 조회** (탭 2)
   - 선택된 기업들의 특정 기간 내 공시 보고서 목록 일괄 조회 및 DataFrame 미리보기.
3. **재무제표 일괄 다운로드** (탭 3)
   - 선택한 모든 기업의 재무제표 개별 추출 진행.
   - 추출 성공 시 서버 폴더 저장 외에도 **브라우저를 통한 다이렉트 엑셀 파일 다운로드** 지원 (클라우드 배포 호환성 확보).

---

## 🛠️ 로컬 설치 및 실행 방법

### 1. 가상환경 구성 및 패키지 설치
```bash
# 가상환경 생성 및 활성화 (Windows 기준)
python -m venv .venv
.venv\Scripts\activate

# 의존성 패키지 설치
pip install -r requirements.txt
```

### 2. 어플리케이션 실행
```bash
streamlit run web_ui.py
```
실행 후 브라우저에서 `http://localhost:8501`로 자동 연결됩니다.

---

## ☁️ Streamlit Community Cloud 배포 가이드

1. **GitHub 저장소 생성 및 푸시**:
   - `.gitignore` 설정에 따라 `.venv`, `__pycache__`, `corp_list_cache.json` 및 추출된 엑셀 파일 폴더(`fsdata/`, `samsung/`) 등은 업로드에서 제외됩니다.
   - 저장소에는 아래 **필수 파일**들만 업로드하면 됩니다.

2. **Streamlit Cloud 로그인**:
   - [Streamlit Community Cloud](https://share.streamlit.io/)에 접속하여 GitHub 계정으로 로그인합니다.

3. **New App 배포**:
   - `Repository`, `Branch`, `Main file path`를 입력합니다.
     - **Main file path**: `web_ui.py`
   - **Advanced settings** (중요):
     - `DART_API_KEY` 환경 변수를 사용하거나 Secrets 탭에 API 키를 사전 설정할 수 있습니다. 
     - 예시 (Secrets 탭 설정):
       ```toml
       DART_API_KEY = "내_DART_API_키_값"
       ```

4. **Deploy!** 버튼을 누르면 패키지(`requirements.txt`) 자동 설치 후 몇 분 내에 온라인에 배포가 완료됩니다.

---

## 📂 파일 구조
* `web_ui.py`: Streamlit 기반 프론트엔드 및 사용자 인터랙션 로직
* `search_dart_corp.py`: DART OpenAPI 파싱 및 엑셀 다운로드 핵심 라이브러리 모듈
* `requirements.txt`: Streamlit Cloud 환경에서 설치할 종속성 리스트
* `.gitignore`: 무겁거나 보안이 필요한 임시 파일 업로드 방지 파일
