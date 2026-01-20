# Cloud Run 배포 가이드

## 1. 사전 준비

### GCP 프로젝트 설정
```bash
# GCP 프로젝트 ID (본인 프로젝트로 변경)
PROJECT_ID="your-gcp-project-id"
REGION="asia-northeast3"  # 서울
SERVICE_NAME="autoexamine-web"
```

### gcloud CLI 로그인
```bash
gcloud auth login
gcloud config set project $PROJECT_ID
```

---

## 2. Docker 이미지 빌드 및 푸시

```bash
# autoexamine-web 폴더로 이동
cd C:\Users\MADUP\Documents\seokmin\autoexamine\autoexamine-web

# Artifact Registry에 이미지 빌드 & 푸시
gcloud builds submit --tag gcr.io/$PROJECT_ID/$SERVICE_NAME
```

---

## 3. Cloud Run 배포

```bash
gcloud run deploy $SERVICE_NAME \
  --image gcr.io/$PROJECT_ID/$SERVICE_NAME \
  --platform managed \
  --region $REGION \
  --allow-unauthenticated \
  --memory 1Gi \
  --timeout 300 \
  --set-env-vars "DROPBOX_REFRESH_TOKEN=ZWDTQqtw9S0AAAAAAAAAATm5rUisU_CVqrhccLJqR6OyJ-ckvlffAVTSJB-C0dep,DROPBOX_APP_KEY=0tyemk1osl7a0x1,DROPBOX_APP_SECRET=0ltp3dkj081hoos"
```

### 주요 옵션 설명
- `--allow-unauthenticated`: 인증 없이 접속 허용
- `--memory 1Gi`: PPT 생성에 충분한 메모리
- `--timeout 300`: 5분 타임아웃 (PPT 생성 시간 고려)
- `--set-env-vars`: Dropbox 환경변수 설정

---

## 4. 배포 확인

배포 후 표시되는 URL로 접속:
```
Service URL: https://autoexamine-web-xxxxx-an.a.run.app
```

---

## 5. 로그 확인

```bash
gcloud run logs read $SERVICE_NAME --region $REGION
```

---

## 보안 참고사항

⚠️ 현재 Dropbox 토큰이 환경변수에 직접 노출됨.
프로덕션 환경에서는 **Secret Manager** 사용 권장:

```bash
# Secret 생성
echo -n "your-token" | gcloud secrets create dropbox-refresh-token --data-file=-

# Cloud Run에서 Secret 사용
gcloud run deploy $SERVICE_NAME \
  --set-secrets "DROPBOX_REFRESH_TOKEN=dropbox-refresh-token:latest"
```
