"""
Exemplos de uso da integração Claude + Excel.

Execute com:
    ANTHROPIC_API_KEY=sk-ant-... python example.py
"""

import os
from pathlib import Path
from claude_excel_integration import ClaudeExcelIntegration


def criar_planilha_exemplo(caminho: str) -> None:
    """Cria uma planilha de vendas de exemplo para os demos."""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendas"

    ws.append(["Produto", "Categoria", "Vendas_Jan", "Vendas_Fev", "Vendas_Mar", "Preco_Unitario"])
    ws.append(["Notebook Pro", "Eletrônicos", 120, 95, 110, 3500.00])
    ws.append(["Mouse Wireless", "Acessórios", 340, 280, 310, 89.90])
    ws.append(["Teclado Mecânico", "Acessórios", 210, 195, 230, 299.00])
    ws.append(["Monitor 27\"", "Eletrônicos", 85, 70, 90, 1299.00])
    ws.append(["Webcam HD", "Acessórios", 450, 510, 480, 189.90])
    ws.append(["SSD 1TB", "Armazenamento", 320, 290, 340, 399.00])
    ws.append(["Headset Gamer", "Acessórios", 175, 160, 200, 249.90])

    wb.save(caminho)
    print(f"Planilha de exemplo criada: {caminho}")


def demo_analise_simples(integration: ClaudeExcelIntegration, planilha: str) -> None:
    """Demo 1: análise livre da planilha."""
    print("\n=== Demo 1: Análise Simples ===")
    resposta = integration.analyze_sheet(
        filepath=planilha,
        prompt=(
            "Analise os dados de vendas abaixo e forneça:\n"
            "1. Os 3 produtos com maior faturamento total (jan+fev+mar × preço).\n"
            "2. A tendência geral de vendas (crescimento, queda ou estável).\n"
            "3. Uma recomendação de negócio baseada nos dados."
        ),
        system="Você é um analista de dados especialista em varejo. Responda em português.",
    )
    print(resposta)


def demo_analise_estruturada(integration: ClaudeExcelIntegration, planilha: str) -> None:
    """Demo 2: extração de insights estruturados em JSON."""
    print("\n=== Demo 2: Análise Estruturada (JSON) ===")
    schema = {
        "type": "object",
        "properties": {
            "melhor_produto": {"type": "string"},
            "pior_produto": {"type": "string"},
            "faturamento_total": {"type": "number"},
            "insights": {
                "type": "array",
                "items": {"type": "string"},
            },
        },
    }
    resultado = integration.analyze_sheet_structured(
        filepath=planilha,
        prompt="Analise as vendas e preencha o JSON com os campos solicitados. Calcule o faturamento total = soma(vendas × preço) de todos os produtos nos 3 meses.",
        json_schema=schema,
    )
    import json
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


def demo_processar_linhas(integration: ClaudeExcelIntegration, planilha: str) -> None:
    """Demo 3: processa cada linha individualmente e salva resultado."""
    print("\n=== Demo 3: Processar Linhas Individualmente ===")
    template = (
        "Produto: {Produto} | Categoria: {Categoria} | "
        "Vendas: jan={Vendas_Jan}, fev={Vendas_Fev}, mar={Vendas_Mar} | Preço: R${Preco_Unitario}\n"
        "Em uma frase curta, classifique o desempenho deste produto como Excelente, Bom, Regular ou Fraco "
        "e justifique brevemente."
    )

    resultados = integration.process_rows(
        filepath=planilha,
        row_prompt_template=template,
        output_column="avaliacao_claude",
        system="Você é um analista de vendas. Responda em português, de forma concisa.",
    )

    saida = "resultados_por_produto.xlsx"
    integration.save_results(resultados, saida)
    print(f"Resultados salvos em: {saida}")
    for r in resultados:
        print(f"  {r['Produto']}: {r['avaliacao_claude'][:80]}...")


def demo_gerar_relatorio(integration: ClaudeExcelIntegration, planilha: str) -> None:
    """Demo 4: Claude gera um relatório Excel do zero."""
    print("\n=== Demo 4: Gerar Relatório Excel com Claude ===")
    saida = "relatorio_gerado.xlsx"
    integration.generate_report(
        prompt=(
            "Com base nos dados de vendas fornecidos, crie uma tabela de ranking de produtos "
            "com as colunas: Posição, Produto, Faturamento_Total (R$), Tendência, Recomendação. "
            "Calcule o faturamento total como soma dos 3 meses × preço unitário. "
            "Ordene por faturamento decrescente."
        ),
        output_path=saida,
        system="Você é um analista de dados. Responda em português.",
        context_file=planilha,
    )
    print(f"Relatório gerado: {saida}")


def main() -> None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("Erro: defina a variável de ambiente ANTHROPIC_API_KEY")
        return

    integration = ClaudeExcelIntegration(api_key=api_key)
    planilha = "vendas_exemplo.xlsx"

    if not Path(planilha).exists():
        criar_planilha_exemplo(planilha)

    demo_analise_simples(integration, planilha)
    demo_analise_estruturada(integration, planilha)
    demo_processar_linhas(integration, planilha)
    demo_gerar_relatorio(integration, planilha)

    print("\n=== Todos os demos concluídos! ===")


if __name__ == "__main__":
    main()
