# Claude Excel Integration

Integração entre **Claude AI** e planilhas **Excel** usando a [API da Anthropic](https://docs.anthropic.com/).

## Funcionalidades

- **Analisar planilhas** — envia dados de uma aba Excel ao Claude com um prompt personalizado
- **Extração estruturada** — obtém respostas em JSON conforme um esquema definido
- **Processar linhas** — envia cada linha individualmente ao Claude e salva as respostas na planilha
- **Gerar relatórios** — pede ao Claude que crie tabelas completas e as salva em Excel com formatação

## Instalação

```bash
pip install -r requirements.txt
```

## Configuração

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

## Uso rápido

```python
from claude_excel_integration import ClaudeExcelIntegration

integration = ClaudeExcelIntegration()

# Analisar uma planilha existente
resposta = integration.analyze_sheet(
    filepath="vendas.xlsx",
    prompt="Quais são os 3 produtos com maior faturamento?",
    system="Você é um analista de dados especialista em varejo.",
)
print(resposta)

# Processar cada linha e salvar resultado
resultados = integration.process_rows(
    filepath="clientes.xlsx",
    row_prompt_template="Classifique o cliente {Nome} com score {Score} como VIP, Regular ou Inativo.",
    output_column="classificacao",
)
integration.save_results(resultados, "clientes_classificados.xlsx")

# Gerar relatório Excel com Claude
integration.generate_report(
    prompt="Crie um relatório de resumo mensal de vendas com totais e médias.",
    output_path="relatorio.xlsx",
    context_file="vendas.xlsx",
)
```

## Exemplos completos

Execute `example.py` para ver os quatro demos em ação:

```bash
python example.py
```

## API

### `ClaudeExcelIntegration(api_key, model)`

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `api_key` | `ANTHROPIC_API_KEY` | Chave da API Anthropic |
| `model` | `claude-opus-4-6` | Modelo Claude a usar |

### Métodos principais

| Método | Descrição |
|--------|-----------|
| `read_sheet(filepath, sheet_name)` | Lê aba Excel → lista de dicts |
| `sheet_to_markdown(filepath, sheet_name)` | Converte aba em tabela Markdown |
| `analyze_sheet(filepath, prompt, ...)` | Analisa planilha com Claude |
| `analyze_sheet_structured(filepath, prompt, json_schema, ...)` | Retorna JSON estruturado |
| `process_rows(filepath, row_prompt_template, ...)` | Processa cada linha com Claude |
| `save_results(data, output_path, ...)` | Salva lista de dicts em Excel |
| `append_results_to_sheet(source, results, ...)` | Adiciona coluna de respostas no Excel original |
| `generate_report(prompt, output_path, ...)` | Gera relatório Excel via Claude |

## Deploy na Hostinger VPS

A integração pode ser executada como uma **API REST** na VPS Hostinger (Ubuntu 24.04).

### Pré-requisitos

- VPS Hostinger com Ubuntu 24.04 (ex: `ssh root@72.60.12.66`)
- Chave da API Anthropic (`sk-ant-...`)

### Instalação

```bash
# 1. Clone o repositório na VPS
ssh root@72.60.12.66
git clone https://github.com/mendes013/mendes013.git /tmp/claude-setup
cd /tmp/claude-setup

# 2. Execute o script de instalação (como root)
bash hostinger_setup.sh
```

O script realiza automaticamente:
- Instala Python 3 e dependências do sistema
- Cria `/opt/claude-api/` com ambiente virtual
- Instala as dependências Python
- Registra o serviço `claude-api` no systemd

### Configuração

```bash
# Edite o arquivo de variáveis de ambiente e insira sua chave
nano /etc/claude-api.env
```

```env
ANTHROPIC_API_KEY=sk-ant-SUA_CHAVE_AQUI
API_HOST=0.0.0.0
API_PORT=8000
```

### Iniciar o serviço

```bash
systemctl start claude-api
systemctl status claude-api
```

### Endpoints da API

| Método | Rota | Descrição |
|--------|------|-----------|
| `GET` | `/health` | Verifica se o serviço está rodando |
| `POST` | `/analyze` | Analisa uma planilha Excel com um prompt |
| `POST` | `/analyze-structured` | Análise com resposta JSON estruturada |
| `POST` | `/generate-report` | Gera um relatório Excel com o Claude |

Documentação interativa disponível em: `http://72.60.12.66:8000/docs`

### Exemplo de uso via curl

```bash
# Verificar saúde
curl http://72.60.12.66:8000/health

# Analisar planilha
curl -X POST http://72.60.12.66:8000/analyze \
  -F "file=@vendas.xlsx" \
  -F "prompt=Quais são os 3 produtos com maior faturamento?"
```

### Logs e manutenção

```bash
journalctl -u claude-api -f      # logs em tempo real
systemctl restart claude-api      # reiniciar serviço
systemctl stop claude-api         # parar serviço
```

---

## Dependências

- [anthropic](https://pypi.org/project/anthropic/) >= 0.49.0
- [openpyxl](https://pypi.org/project/openpyxl/) >= 3.1.2
- [fastapi](https://pypi.org/project/fastapi/) >= 0.115.0
- [uvicorn](https://pypi.org/project/uvicorn/) >= 0.30.0
- [python-multipart](https://pypi.org/project/python-multipart/) >= 0.0.9
