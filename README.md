# Book de fotos em PowerPoint

Automação do relatório semanal: tira as fotos das apresentações de campo da semana e preenche os slides de fotos do relatório, sem mexer no resto do arquivo.

## Scripts

| Arquivo | O que faz |
|---|---|
| `scripts/extrair_fotos.py` | Lê os decks de cada tipo de atividade e extrai as fotos com loja, data e colaborador para um índice |
| `scripts/contato.py` | Gera folhas de contato com miniaturas e identificadores para escolher as fotos |
| `scripts/inserir_fotos.py` | Insere as fotos escolhidas nos slides certos, corrigindo a rotação pelo EXIF, e atualiza as legendas |
| `scripts/corrigir_rotulos.py`, `scripts/grafico_racks_q4.py` | Ajustes de legenda e um gráfico do trimestre |
| `scripts/exportar.ps1` | Exporta o resultado pelo PowerPoint |

## Stack

Python, python-pptx, Pillow, pandas, PowerShell.

---
Projeto feito no trabalho, publicado sem fotos nem apresentações do cliente.
