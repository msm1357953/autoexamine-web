"""
환경 설정 모듈
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# Dropbox 설정
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN", "")
DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY", "0tyemk1osl7a0x1")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET", "0ltp3dkj081hoos")

# Dropbox 경로
DROPBOX_BASE_PATH = "/광고사업부/4. 광고주/삼성증권/2. 업무/03. 소재/03. 전체 Vari 심의자료/■ 파일 취합본 new/파이썬용"
DROPBOX_OUTPUT_PATH = "/광고사업부/4. 광고주/삼성증권/2. 업무/03. 소재/03. 전체 Vari 심의자료/준법,협회 심의자료 ppt/파이썬 결과"

# Google Sheets 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1lIn2bd7bwlWpKedwmswJieTBlfzTkMNxMNmNw5dX_Bo/edit"
OBJECT_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/13yRJbD6THRP5ZS-jTSPoKtdj_ULcPaFTkv9h8eJYR7E/edit"

# 프로젝트 경로
BASE_DIR = Path(__file__).parent.parent
CREDENTIALS_DIR = BASE_DIR / "credentials"
TEMPLATE_PATH = BASE_DIR / "template5.pptx"

# 이미지 확장자
IMAGE_EXTENSIONS = [".jpg", ".png"]
