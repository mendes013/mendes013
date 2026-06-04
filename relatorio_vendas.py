"""Gera um relatório simples de vendas do dia em arquivo de texto."""

vendas = [
    {"produto": "Carne", "quantidade": 3, "valor_unitario": 25.00},
    {"produto": "Frango", "quantidade": 5, "valor_unitario": 12.50},
]

total = sum(item["quantidade"] * item["valor_unitario"] for item in vendas)

with open("relatorio_vendas.txt", "w", encoding="utf-8") as arquivo:
    arquivo.write(f"Total do dia: R$ {total:.2f}\n")

print(f"Relatório gerado! Total: R$ {total:.2f}")
