#!/bin/bash
# Google Cloud Function 배포 스크립트
# 사전 요구: gcloud CLI 설치 및 인증

PROJECT_ID="gen-lang-client-0367740438"
FUNCTION_NAME="update_final_sheets"
RUNTIME="python312"
TOPIC_NAME="high-school-tracking-update"
REGION="asia-northeast1"  # 서울

echo "=========================================================="
echo " Google Cloud Function 배포"
echo "=========================================================="

# Step 1: Pub/Sub 토픽 생성
echo ""
echo "[1/3] Pub/Sub 토픽 생성 중..."
gcloud pubsub topics create $TOPIC_NAME \
  --project=$PROJECT_ID 2>/dev/null || echo "  → 토픽이 이미 존재합니다."

# Step 2: Cloud Function 배포
echo ""
echo "[2/3] Cloud Function 배포 중..."
gcloud functions deploy $FUNCTION_NAME \
  --gen2 \
  --runtime=$RUNTIME \
  --region=$REGION \
  --source=. \
  --entry-point=update_final_sheets \
  --trigger-topic=$TOPIC_NAME \
  --service-account-email=school-bot@gen-lang-client-0367740438.iam.gserviceaccount.com \
  --set-env-vars="SPREADSHEET_ID=14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0" \
  --allow-unauthenticated \
  --project=$PROJECT_ID

# Step 3: 배포 확인
echo ""
echo "[3/3] 배포 상태 확인 중..."
gcloud functions describe $FUNCTION_NAME \
  --gen2 \
  --region=$REGION \
  --project=$PROJECT_ID

echo ""
echo "=========================================================="
echo " 배포 완료!"
echo ""
echo "테스트 방법:"
echo "  gcloud pubsub topics publish $TOPIC_NAME \\"
echo "    --message='test' --project=$PROJECT_ID"
echo ""
echo "로그 확인:"
echo "  gcloud functions logs read $FUNCTION_NAME --gen2 \\"
echo "    --region=$REGION --project=$PROJECT_ID"
echo "=========================================================="
