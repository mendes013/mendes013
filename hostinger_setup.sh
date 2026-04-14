#!/usr/bin/env bash
# hostinger_setup.sh
# Script de configuração para deploy da integração Claude na VPS Hostinger (Ubuntu 24.04)
# Uso: bash hostinger_setup.sh

set -euo pipefail

APP_DIR="/opt/claude-api"
SERVICE_NAME="claude-api"
PYTHON_BIN="python3"
VENV_DIR="$APP_DIR/.venv"

echo "=== Configurando integração Claude na Hostinger VPS ==="

# 1. Atualiza pacotes do sistema
echo "[1/7] Atualizando pacotes do sistema..."
apt-get update -qq
apt-get install -y python3 python3-pip python3-venv git curl

# 2. Cria diretório da aplicação
echo "[2/7] Criando diretório da aplicação em $APP_DIR..."
mkdir -p "$APP_DIR"

# 3. Copia os arquivos do projeto para /opt/claude-api
echo "[3/7] Copiando arquivos do projeto..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cp "$SCRIPT_DIR/claude_excel_integration.py" "$APP_DIR/"
cp "$SCRIPT_DIR/api_server.py"               "$APP_DIR/"
cp "$SCRIPT_DIR/requirements.txt"             "$APP_DIR/"

# 4. Cria ambiente virtual e instala dependências
echo "[4/7] Criando ambiente virtual Python e instalando dependências..."
$PYTHON_BIN -m venv "$VENV_DIR"
"$VENV_DIR/bin/pip" install --upgrade pip -q
"$VENV_DIR/bin/pip" install -r "$APP_DIR/requirements.txt" -q

# 5. Cria arquivo de variáveis de ambiente
echo "[5/7] Configurando variáveis de ambiente..."
ENV_FILE="/etc/claude-api.env"
if [ ! -f "$ENV_FILE" ]; then
    cat > "$ENV_FILE" <<EOF
# Variáveis de ambiente para a API Claude
# Edite este arquivo e insira sua chave da API Anthropic
ANTHROPIC_API_KEY=sk-ant-SUBSTITUA_PELA_SUA_CHAVE
API_HOST=0.0.0.0
API_PORT=8000
EOF
    chmod 600 "$ENV_FILE"
    echo "    ATENÇÃO: edite $ENV_FILE e insira sua ANTHROPIC_API_KEY antes de iniciar o serviço."
else
    echo "    Arquivo $ENV_FILE já existe, mantendo configuração existente."
fi

# 6. Instala o serviço systemd
echo "[6/7] Instalando serviço systemd ($SERVICE_NAME)..."
cp "$SCRIPT_DIR/claude-api.service" "/etc/systemd/system/$SERVICE_NAME.service"
systemctl daemon-reload
systemctl enable "$SERVICE_NAME"

# 7. Resumo
echo ""
echo "=== Instalação concluída! ==="
echo ""
echo "Próximos passos:"
echo "  1. Edite /etc/claude-api.env e defina ANTHROPIC_API_KEY"
echo "  2. Inicie o serviço:  systemctl start $SERVICE_NAME"
echo "  3. Verifique o status: systemctl status $SERVICE_NAME"
echo "  4. Acesse a API:       http://72.60.12.66:8000/docs"
echo ""
echo "Comandos úteis:"
echo "  journalctl -u $SERVICE_NAME -f   # logs em tempo real"
echo "  systemctl restart $SERVICE_NAME   # reiniciar serviço"
echo "  systemctl stop $SERVICE_NAME      # parar serviço"
