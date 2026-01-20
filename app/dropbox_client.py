"""
Dropbox API 클라이언트 모듈 - 캐싱 및 병렬 다운로드 지원
"""
import dropbox
from dropbox.files import FileMetadata, FolderMetadata
from dropbox.common import PathRoot
from io import BytesIO
from typing import Optional, List, Dict, Callable
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

from . import config

# 팀 Dropbox root_namespace_id
ROOT_NAMESPACE_ID = "12114515089"


class DropboxClient:
    """Dropbox API 클라이언트 (캐싱 및 병렬 다운로드 지원)"""
    
    def __init__(self):
        self.access_token: Optional[str] = None
        self._dbx: Optional[dropbox.Dropbox] = None
        self._image_cache: Dict[str, BytesIO] = {}  # 이미지 캐시
        self._folder_cache: Dict[str, str] = {}     # 최신 폴더 경로 캐시
        self._refresh_access_token()
    
    def _refresh_access_token(self):
        """Refresh token을 사용해 access token 갱신"""
        response = requests.post(
            "https://api.dropboxapi.com/oauth2/token",
            data={
                "refresh_token": config.DROPBOX_REFRESH_TOKEN,
                "grant_type": "refresh_token",
                "client_id": config.DROPBOX_APP_KEY,
                "client_secret": config.DROPBOX_APP_SECRET,
            }
        )
        if response.status_code == 200:
            self.access_token = response.json().get("access_token")
        else:
            raise Exception(f"Failed to refresh Dropbox token: {response.text}")
    
    @property
    def dbx(self) -> dropbox.Dropbox:
        """Dropbox 클라이언트 인스턴스 (팀 Dropbox용 path_root 설정)"""
        if not self.access_token:
            self._refresh_access_token()
        if self._dbx is None:
            base_dbx = dropbox.Dropbox(self.access_token)
            self._dbx = base_dbx.with_path_root(PathRoot.root(ROOT_NAMESPACE_ID))
        return self._dbx
    
    def get_materials_list(self, keyword: Optional[str] = None) -> Dict[str, List[str]]:
        """키워드로 소재 폴더 목록 조회 (병렬 + 캐싱)"""
        # 캐시 키 생성
        cache_key = f"materials_{keyword or 'all'}"
        if hasattr(self, '_materials_cache') and cache_key in self._materials_cache:
            cached = self._materials_cache[cache_key]
            if keyword:
                return {k: v for k, v in cached.items() if keyword in k}
            return cached
        
        materials_sizes = {}
        
        try:
            result = self.dbx.files_list_folder(config.DROPBOX_BASE_PATH)
            
            # 폴더 목록 추출
            folders = [
                entry.name for entry in result.entries 
                if isinstance(entry, FolderMetadata)
            ]
            
            # 키워드 필터링
            if keyword:
                folders = [f for f in folders if keyword in f]
            
            # 병렬로 각 폴더 정보 조회
            def get_folder_info(folder_name):
                material_path = f"{config.DROPBOX_BASE_PATH}/{folder_name}"
                latest_folder = self._get_latest_date_folder(material_path)
                if not latest_folder:
                    return folder_name, []
                sizes = self._get_image_sizes(latest_folder)
                return folder_name, sizes
            
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {executor.submit(get_folder_info, f): f for f in folders}
                
                for future in as_completed(futures):
                    folder_name, sizes = future.result()
                    if sizes:
                        materials_sizes[folder_name] = sizes
            
            # 캐시 저장 (전체 목록일 때만)
            if not keyword:
                if not hasattr(self, '_materials_cache'):
                    self._materials_cache = {}
                self._materials_cache[cache_key] = materials_sizes
                        
        except Exception as e:
            print(f"Error listing materials: {e}")
        
        return materials_sizes
    
    def _get_latest_date_folder(self, material_path: str) -> Optional[str]:
        """최신 날짜 폴더 경로 반환 (캐싱)"""
        if material_path in self._folder_cache:
            return self._folder_cache[material_path]
        
        try:
            result = self.dbx.files_list_folder(material_path)
            date_folders = [
                entry.name for entry in result.entries 
                if isinstance(entry, FolderMetadata)
            ]
            
            if not date_folders:
                return None
            
            latest_date = max(date_folders)
            full_path = f"{material_path}/{latest_date}"
            self._folder_cache[material_path] = full_path
            return full_path
            
        except Exception as e:
            print(f"Error finding date folder: {e}")
            return None
    
    def _get_image_sizes(self, folder_path: str) -> List[str]:
        """폴더 내 이미지 사이즈 목록"""
        sizes = []
        
        try:
            result = self.dbx.files_list_folder(folder_path)
            
            for entry in result.entries:
                if isinstance(entry, FileMetadata):
                    for ext in config.IMAGE_EXTENSIONS:
                        if entry.name.lower().endswith(ext):
                            size = entry.name.rsplit(".", 1)[0]
                            sizes.append(size)
                            break
                            
        except Exception as e:
            print(f"Error listing images: {e}")
        
        return sizes
    
    def download_image(self, material: str, size: str) -> Optional[BytesIO]:
        """이미지 다운로드 (캐싱 지원)"""
        cache_key = f"{material}/{size}"
        
        # 캐시 확인
        if cache_key in self._image_cache:
            cached = self._image_cache[cache_key]
            cached.seek(0)
            return BytesIO(cached.getvalue())  # 복사본 반환
        
        material_path = f"{config.DROPBOX_BASE_PATH}/{material}"
        latest_folder = self._get_latest_date_folder(material_path)
        
        if not latest_folder:
            return None
        
        for ext in config.IMAGE_EXTENSIONS:
            file_path = f"{latest_folder}/{size}{ext}"
            try:
                _, response = self.dbx.files_download(file_path)
                img_bytes = BytesIO(response.content)
                self._image_cache[cache_key] = img_bytes
                return BytesIO(response.content)
            except dropbox.exceptions.ApiError:
                continue
        
        return None
    
    def preload_images(self, materials: List[str], sizes: List[str], 
                       progress_callback: Optional[Callable] = None) -> int:
        """
        이미지 병렬 프리로드 (캐싱)
        Args:
            materials: 소재 목록
            sizes: 다운로드할 사이즈 목록
            progress_callback: 진행 콜백 (current, total, message)
        Returns:
            다운로드된 이미지 수
        """
        download_tasks = []
        
        # 다운로드할 이미지 목록 생성
        for material in materials:
            for size in sizes:
                cache_key = f"{material}/{size}"
                if cache_key not in self._image_cache:
                    download_tasks.append((material, size))
        
        total = len(download_tasks)
        if total == 0:
            return 0
        
        downloaded = 0
        
        def download_one(task):
            material, size = task
            return material, size, self.download_image(material, size)
        
        # 병렬 다운로드 (최대 5개 동시)
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(download_one, task): task for task in download_tasks}
            
            for future in as_completed(futures):
                material, size, result = future.result()
                downloaded += 1
                
                if progress_callback:
                    progress_callback(downloaded, total, f"📥 {material} - {size}")
        
        return downloaded
    
    def clear_cache(self):
        """캐시 초기화"""
        self._image_cache.clear()
        self._folder_cache.clear()
    
    def upload_ppt(self, ppt_bytes: BytesIO, filename: str) -> Optional[str]:
        """PPT 파일을 Dropbox에 업로드"""
        upload_path = f"{config.DROPBOX_OUTPUT_PATH}/{filename}"
        
        try:
            ppt_bytes.seek(0)
            self.dbx.files_upload(
                ppt_bytes.read(),
                upload_path,
                mode=dropbox.files.WriteMode.overwrite
            )
            return upload_path
        except Exception as e:
            print(f"Error uploading PPT: {e}")
            return None


# 싱글톤 인스턴스
_client: Optional[DropboxClient] = None

def get_dropbox_client() -> DropboxClient:
    """Dropbox 클라이언트 싱글톤"""
    global _client
    if _client is None:
        _client = DropboxClient()
    return _client
