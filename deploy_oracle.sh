#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-$HOME/painel-emendas}"
REPO_URL="${REPO_URL:-https://github.com/victorramon88-dev/painel-emendas.git}"

sudo apt-get update
sudo apt-get install -y ca-certificates curl git

sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg

. /etc/os-release
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu ${VERSION_CODENAME} stable" \
  | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

sudo apt-get update
sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
sudo usermod -aG docker "$USER"

if [ ! -d "$APP_DIR/.git" ]; then
  git clone "$REPO_URL" "$APP_DIR"
else
  git -C "$APP_DIR" pull --ff-only
fi

cd "$APP_DIR"
sudo docker compose up -d --build

echo "Aplicacao iniciada."
echo "Acesse: http://IP_PUBLICO_DA_VM:8000"
echo "Se necessario, libere a porta 8000 na Oracle e no firewall da VM."
