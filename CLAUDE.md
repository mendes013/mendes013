# Claude Code – Configuração de Colaboração (Cowork)

Este arquivo configura o ambiente de colaboração para sessões Claude Code neste repositório.

## Acesso Cowork

Para colaborar com outros usuários usando Claude Code:

1. **Na Web (claude.ai/code):** Acesse claude.ai/code com sua conta Pro e abra este repositório.
2. **No Aplicativo:** Use o aplicativo Claude no celular e conecte-se ao repositório via GitHub.

## Permissões de Colaboração

```bash
# Ferramentas permitidas para colaboradores
allow: Read, Edit, Write, Bash, Glob, Grep
```

## Comandos disponíveis

- Análise de planilhas Excel com Claude AI
- Integração com a API Anthropic
- Processamento de dados em batch

## Configuração do ambiente

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
```

## Modelo padrão

- **Modelo:** `claude-opus-4-6`
- **Conta necessária:** Claude Pro ou superior
