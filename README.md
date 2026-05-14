# 🏢 아파트 실거래가 조회 앱

국토교통부 공공데이터포털 **RTMSDataSvcAptTrade** API를 사용해 서울·경기 시군구 단위 아파트 매매 실거래가를 조회하고 Excel·HTML로 다운로드하는 Streamlit 웹앱.

- 시도 → 시군구 드롭다운 (서울 25 + 경기 31 = 56개)
- 시작월 ~ 종료월 다중 월 자동 누적 호출
- 단지명·법정동 키워드 검색 + 평형·가격대 슬라이더 필터
- 결과: 단지·평형별 요약 (Sheet1) + 전체 거래내역 (Sheet2)
- 다운로드: Excel(.xlsx) · 단일 HTML 파일
- 캐시: 동일 시군구·월 24시간 (API 일일 한도 절약)

---

## 🌐 배포 (Streamlit Community Cloud, 무료)

### 1. GitHub 저장소 준비
1. [github.com](https://github.com) 가입·로그인
2. 우측 상단 **+ → New repository**
3. Repository name: 자유 (예: `apt-trade-app`)
4. Public 또는 Private 어느 쪽이든 OK
5. **README/.gitignore 추가 옵션은 모두 OFF** (이미 로컬에 있음)
6. Create

### 2. 코드 푸시
```bash
cd "C:\Users\ajfxo\Desktop\부동산 실거래가 앱"
git remote add origin https://github.com/<본인계정>/<repo이름>.git
git branch -M main
git push -u origin main
```

### 3. Streamlit Cloud 연결
1. [share.streamlit.io](https://share.streamlit.io) 가입 (GitHub 계정으로 로그인)
2. **New app** → 위에서 만든 repo·branch(`main`)·`app.py` 선택
3. **Advanced settings → Secrets** 에 아래 한 줄 붙여넣기:
   ```toml
   MOLIT_API_KEY = "여기에_국토부_일반인증키"
   ```
4. **Deploy** → 1~3분 후 `https://<앱이름>.streamlit.app` 발급

> ⚠️ **절대 GitHub에 API 키를 커밋하지 마세요.** 키는 오직 Streamlit Cloud Secrets 또는 로컬 `.env`에만 둡니다. `.gitignore`가 자동으로 차단합니다.

---

## 💻 로컬 실행

### 최초 1회
```bash
python -m pip install -r requirements.txt
```
또는 `setup_최초1회.bat` 더블클릭.

### 평소 사용
`부동산앱_시작.bat` 더블클릭 → 자동으로 `http://localhost:8501` 열림.

### 로컬 API 키
다음 중 하나:
- `C:\Users\ajfxo\.claude\.env` 파일에 `MOLIT_API_KEY=...` (지호님 PC 기본 설정)
- 또는 같은 폴더의 `.streamlit/secrets.toml.example` 을 `secrets.toml`로 복사 후 키 입력

---

## 📁 파일 구조

```
부동산 실거래가 앱/
├── app.py                      # 메인 Streamlit 앱
├── lawd_codes.py               # 시도·시군구 LAWD_CD 매핑
├── exporters.py                # Excel·HTML 변환
├── requirements.txt            # 의존성 (Streamlit Cloud 자동 인식)
├── runtime.txt                 # Python 버전 지정
├── .gitignore                  # 민감 파일 제외
├── .streamlit/
│   └── secrets.toml.example    # 로컬 secrets 예시
├── setup_최초1회.bat
├── 부동산앱_시작.bat
├── README.md                   # 이 파일
└── README.txt                  # 비개발자용 한글 사용설명서
```

---

## 🔑 국토부 API 키 발급

[data.go.kr](https://www.data.go.kr) → "아파트 매매 실거래가 자료" 검색 → **활용신청** → 승인 후 **일반 인증키 (Decoding)** 사용.

일일 한도: 기본 10,000건. 캐시로 동일 조건 재호출은 차감되지 않음.

---

## 📊 데이터 출처

- 국토교통부 실거래가 공개시스템 (동일 데이터)
- 계약일 기준 신고분. 신고 의무 30일이므로 최근 1~2개월은 누적 중일 수 있음
