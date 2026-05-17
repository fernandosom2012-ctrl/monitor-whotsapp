"""Project configuration."""

CONFIG = {
    # WhatsApp
    "nome_grupo": "🟣 CHAMADOS T.I HRJL",
    "trigger_chamado": "#SOL. DE MANUTENÇÃO ",

    # Excel
    "arquivo_excel": "chamados.xlsx",

    # Ticket id
    "prefixo_id": "CH",

    # Browser and monitoring
    "intervalo_segundos": 10,
    "timeout_login": 120,
    "timeout_busca_grupo": 20,

    # Optional local ChromeDriver path
    "chromedriver_path": "",
}


COLUNAS = [
    {"campo": "id", "titulo": "ID Chamado", "largura": 18},
    {"campo": "data_hora", "titulo": "Data/Hora", "largura": 18},
    {"campo": "solicitante", "titulo": "Solicitante", "largura": 25},
    {"campo": "setor", "titulo": "Setor", "largura": 20},
    {"campo": "problema", "titulo": "Descrição", "largura": 50},
    {"campo": "prioridade", "titulo": "Prioridade", "largura": 12},
    {"campo": "status", "titulo": "Status", "largura": 14},
]
