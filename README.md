# 심의자료 자동화 웹 애플리케이션

삼성증권 광고 소재의 심의자료 PPT를 자동 생성하는 웹 애플리케이션입니다.

## 실행 방법

### 로컬 개발
```bash
# 가상환경 생성
python -m venv .venv
.venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt

# 환경변수 설정
copy .env.example .env
# .env 파일에 DROPBOX_REFRESH_TOKEN 설정

# 서버 실행
uvicorn app.main:app --reload --port 8000
```

### Docker
```bash
docker build -t autoexamine-web .
docker run -p 8080:8080 --env-file .env autoexamine-web
```

## 사용법
1. 브라우저에서 http://localhost:8000 접속
2. 키워드 입력 (예: usp-dm-1st)
3. [생성하기] 클릭
4. PPT 파일 다운로드
