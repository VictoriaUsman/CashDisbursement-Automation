#!/bin/bash
set -e

PROJECT_ID="consolidation-495808"
IMAGE="gcr.io/$PROJECT_ID/cash-disbursement"
REGION="asia-southeast1"
SERVICE="cash-disbursement"

echo "==> Storing service account in Secret Manager..."
gcloud secrets create service-account-json \
  --data-file=service_account.json \
  --project=$PROJECT_ID 2>/dev/null || \
gcloud secrets versions add service-account-json \
  --data-file=service_account.json \
  --project=$PROJECT_ID

echo "==> Building and pushing image..."
gcloud builds submit --tag $IMAGE --project=$PROJECT_ID

echo "==> Deploying to Cloud Run..."
gcloud run deploy $SERVICE \
  --image $IMAGE \
  --platform managed \
  --region $REGION \
  --allow-unauthenticated \
  --set-secrets="SERVICE_ACCOUNT_JSON=service-account-json:latest" \
  --port 8080 \
  --project=$PROJECT_ID

echo "==> Done! Your app is live."
