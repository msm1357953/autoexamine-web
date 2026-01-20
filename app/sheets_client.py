"""
Google Sheets API 클라이언트 모듈
"""
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from typing import Optional, Dict, List, Any
from pathlib import Path

from . import config


class SheetsClient:
    """Google Sheets API 클라이언트"""
    
    def __init__(self):
        self.gc = self._authorize()
        self._df_cache: Dict[str, List[pd.DataFrame]] = {}
        self._df_obj_cache: Dict[str, List[pd.DataFrame]] = {}
    
    def _authorize(self) -> gspread.Client:
        """Google Sheets 인증"""
        scope = ['https://spreadsheets.google.com/feeds']
        
        # credentials 폴더에서 JSON 파일 찾기
        json_files = list(config.CREDENTIALS_DIR.glob("*.json"))
        if not json_files:
            raise FileNotFoundError(f"No JSON credentials found in {config.CREDENTIALS_DIR}")
        
        json_file = json_files[0]
        credentials = Credentials.from_service_account_file(str(json_file), scopes=scope)
        return gspread.authorize(credentials)
    
    def _load_spreadsheet_data(self, url: str, cache_key: str) -> List[pd.DataFrame]:
        """스프레드시트 데이터 로드 (캐싱)"""
        if cache_key in self._df_cache:
            return self._df_cache[cache_key]
        
        doc = self.gc.open_by_url(url)
        worksheet_list = doc.worksheets()
        
        dfs = []
        for ws in worksheet_list:
            data = ws.get_all_values()
            df = pd.DataFrame(data=data)
            dfs.append(df)
        
        self._df_cache[cache_key] = dfs
        return dfs
    
    def get_text_assets(self, material_name: str) -> Dict[str, Any]:
        """
        소재명으로 텍스트 에셋 추출
        Returns: {
            'google_range_list': [...],  # 구글 AC 텍스트 30개
            'meta_range_list': [...],    # META 텍스트
            'meta_caution': str          # META 유의문구
        }
        """
        dfs = self._load_spreadsheet_data(config.SPREADSHEET_URL, "main")
        
        result = {
            'google_range_list': [],
            'meta_range_list': [],
            'meta_caution': ''
        }
        
        # 소재명이 있는 시트 찾기
        worksheet_df = None
        cell_creative_row = None
        cell_creative_col = None
        
        for df in dfs:
            find_result = df.isin([material_name])
            series_result = find_result.any()
            if series_result.any():
                worksheet_df = df
                
                # 소재 위치 찾기
                alias_cols = list(series_result[series_result == True].index)
                for col in alias_cols:
                    rows = list(find_result[col][find_result[col] == True].index)
                    if rows:
                        cell_creative_row = rows[0]
                        cell_creative_col = col
                        break
                break
        
        if worksheet_df is None or cell_creative_row is None:
            return result
        
        # 구글 AC 텍스트 추출
        try:
            # "구글 AC" 위치 찾기
            find_google = worksheet_df.isin(["구글 AC"])
            series_google = find_google.any()
            google_cols = list(series_google[series_google == True].index)
            
            if google_cols:
                cell_google_col = google_cols[0]
                
                # "광고 제목1" 위치 찾기
                col_data = worksheet_df[cell_google_col][:cell_creative_row + 1]
                title_rows = col_data.index[col_data == "광고 제목1"]
                if len(title_rows) > 0:
                    text_row_address = title_rows[-1]
                    
                    # 30개 텍스트 추출
                    for i in range(30):
                        try:
                            text = worksheet_df.iloc[text_row_address + i, cell_google_col + 1]
                            result['google_range_list'].append(text)
                        except:
                            result['google_range_list'].append('')
        except Exception as e:
            print(f"Error extracting Google AC text: {e}")
        
        # META 텍스트 추출
        try:
            find_meta = worksheet_df.isin(["META"])
            series_meta = find_meta.any()
            meta_cols = list(series_meta[series_meta == True].index)
            
            if meta_cols:
                cell_meta_col = meta_cols[0]
                
                # "광고 제목" 위치 찾기
                col_data = worksheet_df[cell_meta_col][:cell_creative_row + 1]
                title_rows = col_data.index[col_data == "광고 제목"]
                if len(title_rows) > 0:
                    text_row_address = title_rows[-1]
                    
                    # 8개 텍스트 추출
                    for i in range(8):
                        try:
                            text = worksheet_df.iloc[text_row_address + i, cell_meta_col + 1]
                            result['meta_range_list'].append(text)
                        except:
                            result['meta_range_list'].append('')
                    
                    # 유의문구 추출
                    caution_col = worksheet_df[cell_meta_col][text_row_address:]
                    caution_rows = caution_col.index[caution_col == "유의 문구"]
                    if len(caution_rows) > 0:
                        result['meta_caution'] = worksheet_df.iloc[caution_rows[0], cell_meta_col + 1]
        except Exception as e:
            print(f"Error extracting META text: {e}")
        
        return result
    
    def get_object_assets(self, materials: List[str]) -> pd.DataFrame:
        """
        오브젝트형 텍스트 에셋 추출
        Returns: DataFrame with columns for each platform
        """
        dfs = self._load_spreadsheet_data(config.OBJECT_SPREADSHEET_URL, "object")
        
        columns = [
            "카카오_비즈보드_메인카피",
            "카카오_비즈보드_서브카피",
            "카카오_비즈보드(몰로코,애피어)_메인카피",
            "토스_혜택탭_메인문구",
            "토스_혜택탭_보조문구",
            "토스_모먼트탭_메인문구1",
            "토스_모먼트탭_메인문구2",
            "토스_모먼트탭_보조문구",
            "네이버GFA_네이티브_광고문구",
            "네이버GFA_네이티브_설명문구1",
            "네이버GFA_네이티브_설명문구2",
            "네이버GFA_네이티브_설명문구3",
            "네이버GFA_커뮤니케이션애드_광고문구1",
            "네이버GFA_커뮤니케이션애드_광고문구2",
            "당근_당근네이티브_광고 제목",
            "당근_당근네이티브_심의필 문구",
            "버즈빌_카카오금융_광고 제목",
        ]
        
        df_result = pd.DataFrame(index=materials, columns=columns)
        
        for material in materials:
            worksheet_df = None
            cell_creative_row = None
            
            # 소재명이 있는 시트 찾기
            for df in dfs:
                find_result = df.isin([material])
                series_result = find_result.any()
                if series_result.any():
                    worksheet_df = df
                    alias_cols = list(series_result[series_result == True].index)
                    for col in alias_cols:
                        rows = list(find_result[col][find_result[col] == True].index)
                        if rows:
                            cell_creative_row = rows[0]
                            break
                    break
            
            if worksheet_df is None or cell_creative_row is None:
                continue
            
            # 각 플랫폼별 텍스트 추출
            platform_keywords = {
                "비즈보드": ["카카오_비즈보드_메인카피", "카카오_비즈보드_서브카피"],
                "비즈보드(몰로코,애피어)": ["카카오_비즈보드(몰로코,애피어)_메인카피"],
                "혜택탭": ["토스_혜택탭_메인문구", "토스_혜택탭_보조문구"],
                "모먼트탭": ["토스_모먼트탭_메인문구1", "토스_모먼트탭_메인문구2", "토스_모먼트탭_보조문구"],
                "네이티브": ["네이버GFA_네이티브_광고문구", "네이버GFA_네이티브_설명문구1", 
                          "네이버GFA_네이티브_설명문구2", "네이버GFA_네이티브_설명문구3"],
                "커뮤니케이션애드 (=콘텍스트)": ["네이버GFA_커뮤니케이션애드_광고문구1", "네이버GFA_커뮤니케이션애드_광고문구2"],
                "당근 네이티브": ["당근_당근네이티브_광고 제목", None, "당근_당근네이티브_심의필 문구"],
                "버즈빌 카카오 금융": ["버즈빌_카카오금융_광고 제목"],
            }
            
            for keyword, col_names in platform_keywords.items():
                try:
                    find_kw = worksheet_df.isin([keyword])
                    series_kw = find_kw.any()
                    kw_cols = list(series_kw[series_kw == True].index)
                    
                    if kw_cols:
                        cell_kw_col = kw_cols[0]
                        
                        for i, col_name in enumerate(col_names):
                            if col_name is not None:
                                try:
                                    value = worksheet_df.iloc[cell_creative_row - 1, cell_kw_col + i]
                                    df_result.loc[material, col_name] = value
                                except:
                                    pass
                except Exception as e:
                    print(f"Error extracting {keyword}: {e}")
        
        df_result.fillna('', inplace=True)
        return df_result


# 싱글톤 인스턴스
_client: Optional[SheetsClient] = None

def get_sheets_client() -> SheetsClient:
    """Sheets 클라이언트 싱글톤"""
    global _client
    if _client is None:
        _client = SheetsClient()
    return _client
