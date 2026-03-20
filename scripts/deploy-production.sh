#!/usr/bin/env bash
set -euo pipefail

TARGET_BRANCH="main"
SERVER_USER="root"
SERVER_HOST="200.58.126.177"
SERVER_PORT="5935"
SERVER_APP_DIR="/var/www/plan-de-visitas"
SERVER_PM2_APP="plan-de-visitas"
PUBLIC_URL="https://visitasms.online"

CURRENT_BRANCH="$(git branch --show-current)"

if [[ "$CURRENT_BRANCH" != "$TARGET_BRANCH" ]]; then
  echo "Este deploy solo publica ${TARGET_BRANCH}. Rama actual: ${CURRENT_BRANCH}."
  echo "Cambiá a ${TARGET_BRANCH} antes de seguir."
  exit 1
fi

if [[ -n "$(git status --porcelain)" ]]; then
  if [[ $# -lt 1 ]]; then
    echo "Hay cambios sin commitear."
    echo "Uso:"
    echo "  ./scripts/deploy-production.sh \"Mensaje del commit\""
    exit 1
  fi

  COMMIT_MESSAGE="$1"
  echo "Guardando cambios locales..."
  git add .
  git commit -m "$COMMIT_MESSAGE"
fi

echo "Subiendo ${TARGET_BRANCH} a GitHub..."
git push origin "$TARGET_BRANCH"

echo "Actualizando servidor..."
ssh -p "$SERVER_PORT" "${SERVER_USER}@${SERVER_HOST}" "
  set -e
  cd '${SERVER_APP_DIR}'
  git fetch origin
  git reset --hard origin/${TARGET_BRANCH}
  pm2 restart '${SERVER_PM2_APP}'
  pm2 status
  curl -I '${PUBLIC_URL}'
"

echo "Deploy terminado."
