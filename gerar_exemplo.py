"""Generate a sample workbook with fake tickets."""

from __future__ import annotations

from pathlib import Path

from excel_handler import ExcelHandler


CHAMADOS_EXEMPLO = [
    {
        "id": "CH-2025-0001",
        "data_hora": "29/04/2025 08:15",
        "solicitante": "Carlos Mendes",
        "setor": "Financeiro",
        "problema": "Computador nao liga apos atualizacao do Windows.",
        "prioridade": "Alta",
        "status": "Aberto",
    },
    {
        "id": "CH-2025-0002",
        "data_hora": "29/04/2025 09:02",
        "solicitante": "Ana Lima",
        "setor": "RH",
        "problema": "Impressora travando ao imprimir documentos PDF.",
        "prioridade": "Média",
        "status": "Em andamento",
    },
    {
        "id": "CH-2025-0003",
        "data_hora": "29/04/2025 09:45",
        "solicitante": "Roberto Silva",
        "setor": "Vendas",
        "problema": "Sistema de CRM nao abre e mostra erro 502.",
        "prioridade": "Alta",
        "status": "Aberto",
    },
    {
        "id": "CH-2025-0004",
        "data_hora": "29/04/2025 10:30",
        "solicitante": "Juliana Costa",
        "setor": "Administrativo",
        "problema": "Criacao de novo usuario para funcionario contratado.",
        "prioridade": "Baixa",
        "status": "Resolvido",
    },
    {
        "id": "CH-2025-0005",
        "data_hora": "29/04/2025 11:10",
        "solicitante": "Marcos Souza",
        "setor": "TI",
        "problema": "Servidor de arquivos muito lento para toda a equipe.",
        "prioridade": "Alta",
        "status": "Em andamento",
    },
]


def main() -> None:
    arquivo_saida = Path("chamados_exemplo.xlsx")
    if arquivo_saida.exists():
        arquivo_saida.unlink()

    excel = ExcelHandler(str(arquivo_saida))
    for chamado in CHAMADOS_EXEMPLO:
        excel.salvar_chamado(chamado)
    excel.fechar()

    print(f"Planilha de exemplo gerada: {arquivo_saida}")


if __name__ == "__main__":
    main()
