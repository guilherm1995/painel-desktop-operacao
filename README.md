# Painel Desktop — consulta operacional consolidada

> Aplicação desktop que junta planilha, Google Sheets e consulta de status de conexão numa tela só, para quem atende no balcão.

![status](https://img.shields.io/badge/status-portfolio-blue)
![licença](https://img.shields.io/badge/licen%C3%A7a-MIT-green)

## O que faz

Antes, responder "esse contrato está em garantia?" exigia abrir três coisas: a
planilha de chamados, o formulário de conveniência e o sistema de status de
conexão. Este painel junta as três numa tela, com cache, e responde em segundos.

## Como está organizado

Uma aplicação CustomTkinter de arquivo único (`painel_operacional.py`), dividida
internamente em:

- `DataCache` — cache com TTL e detecção de ociosidade
- carregadores (`carregar_excel`, `carregar_google_sheet`, `carregar_csv_robusto`)
- normalizadores (data, contrato, cidade, nome de coluna)
- `consultar_autenticador_status` — status de conexão em lote
- `PainelApp` — a interface

## Decisões que valem comentário

**Cache com TTL e ociosidade.** As fontes são lentas e mudam pouco. O cache
guarda por 5 minutos e para de renovar quando ninguém está usando, para não
martelar a planilha durante a noite.

**Normalização defensiva.** A função `encontrar_coluna` aceita uma lista de nomes
possíveis para a mesma coluna. Planilha mantida por várias pessoas muda de
cabeçalho sem avisar; quebrar porque "Nº Contrato" virou "Número do Contrato"
não é aceitável numa ferramenta de balcão.

**Consulta em lote.** O status de conexão é consultado para a lista inteira de
contratos de uma vez, não um a um. A diferença é de minutos para segundos.

**Log de exceção não tratada.** A função `handle_exception` captura o que escapou
e grava com stack trace. Aplicação desktop que fecha sozinha sem deixar rastro é
impossível de diagnosticar remotamente.

## Rodando

```bash
pip install -r requirements.txt
python painel_operacional.py
```

Precisa de credenciais de conta de serviço do Google com acesso à planilha.

## Aviso

Este repositório é uma versão de portfólio, extraída de um sistema que rodou em
produção. Foi anonimizado antes da publicação: nomes de empresas, domínios
internos, credenciais, sessões de mensageria e dados de clientes foram
substituídos por valores de exemplo. Os arquivos de configuração são gabaritos,
não os valores reais de operação.

O código está aqui como referência técnica. Para rodar de verdade, é preciso
apontar as variáveis de ambiente e os configs para um ambiente próprio.

## Licença

MIT — veja [LICENSE](LICENSE). Copyright (c) 2026 Guilherme da Silva dos Santos.
