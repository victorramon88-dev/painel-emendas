#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-$HOME/painel-emendas}"
REPO_URL="${REPO_URL:-https://github.com/victorramon88-dev/painel-emendas.git}"

. /etc/os-release

if command -v apt-get >/dev/null 2>&1; then
  sudo apt-get update
  sudo apt-get install -y ca-certificates curl git gnupg

  sudo install -m 0755 -d /etc/apt/keyrings
  if [ ! -f /etc/apt/keyrings/docker.gpg ]; then
    curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
  fi
  sudo chmod a+r /etc/apt/keyrings/docker.gpg

  echo \
    "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu ${VERSION_CODENAME} stable" \
    | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

  sudo apt-get update
  sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
elif command -v dnf >/dev/null 2>&1; then
  sudo dnf install -y dnf-utils ca-certificates curl git
  sudo dnf config-manager --add-repo https://download.docker.com/linux/centos/docker-ce.repo
  sudo dnf install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
else
  echo "Distribuicao Linux sem apt-get ou dnf. Instale Docker e Git manualmente."
  exit 1
fi

sudo systemctl enable --now docker
sudo usermod -aG docker "$USER" || true

if [ ! -d "$APP_DIR/.git" ]; then
  git clone "$REPO_URL" "$APP_DIR"
else
  git -C "$APP_DIR" pull --ff-only
fi

cd "$APP_DIR"
sudo docker compose up -d --build

if command -v firewall-cmd >/dev/null 2>&1; then
  sudo firewall-cmd --permanent --add-port=8000/tcp || true
  sudo firewall-cmd --reload || true
fi

echo "Aplicacao iniciada."
echo "Acesse: http://IP_PUBLICO_DA_VM:8000"
echo "Se necessario, libere a porta 8000 na Oracle e no firewall da VM."
