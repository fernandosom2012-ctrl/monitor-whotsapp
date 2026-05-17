# WhatsApp Monitor - Chamados para Excel

Automacao em Python que monitora um grupo do WhatsApp Web e registra chamados em uma planilha Excel.

## Requisitos

- Python 3.10+
- Google Chrome instalado
- Dependencias em `requirements.txt`

## Instalacao

```bash
pip install -r requirements.txt
```

## Configuracao

Edite `config.py` antes de executar:

```python
CONFIG = {
    "nome_grupo": "🟣 CHAMADOS T.I HRJL",
    "trigger_chamado": "#SOL. DE MANUTENÇÃO ",
    "arquivo_excel": "chamados.xlsx",
    "prefixo_id": "CH",
    "intervalo_segundos": 10,
}
```

Observacoes:

- O nome do grupo precisa ser exatamente igual ao do WhatsApp.
- O trigger eh comparado de forma tolerante: com ou sem `#`, com ou sem acentos e ignorando espacos extras.
- Se quiser informar um driver manualmente, voce pode adicionar `chromedriver_path` no `CONFIG`.

## Formato das mensagens

Exemplo aceito:

```text
SOL. DE MANUTENÇÃO
Solicitante: James Sousa
Setor: Posto SPA
Prioridade: Alta
Descrição do problema:
Impressora sem sair impressões
```

Campos reconhecidos:

- Solicitante: `Nome`, `Solicitante`
- Setor: `Setor`, `Departamento`, `Área`, `Area`
- Problema: `Problema`, `Descrição`, `Descricao`, `Assunto`
- Prioridade: `Prioridade`, `Urgência`, `Urgencia`

Se a prioridade nao vier preenchida, o sistema tenta detectar por palavras-chave.

## Como executar

```bash
python monitor.py
```

Fluxo esperado:

1. O Chrome abre no WhatsApp Web.
2. Se necessario, voce escaneia o QR Code.
3. O script abre o grupo configurado.
4. As mensagens ja visiveis sao sincronizadas para evitar duplicidade ao iniciar.
5. O monitor passa a capturar apenas novas mensagens.

Para encerrar, use `Ctrl+C`.

## Planilha gerada

O arquivo `chamados.xlsx` possui:

- Aba `Chamados` com os registros
- Aba `Resumo` com contadores por status e prioridade

## Arquivos principais

- `monitor.py`: automacao do WhatsApp Web
- `excel_handler.py`: escrita e formatacao do Excel
- `config.py`: configuracoes
- `gerar_exemplo.py`: gera uma planilha de exemplo

## Observacoes importantes

- A sessao do WhatsApp fica salva em `whatsapp_profile/`.
- O projeto tenta usar um `chromedriver` local ou em cache antes de baixar um novo.
- Mudancas no HTML do WhatsApp Web podem exigir ajuste de seletores no futuro.
