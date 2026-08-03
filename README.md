# Mesa da Reunião — Conselhos Metropolitanos

Extensão (Google Apps Script) para o Google Sheets que transforma uma planilha em um painel operacional completo para conduzir reuniões de instalação e apuração dos Conselhos Metropolitanos (Palmas, Araguaína e Gurupi). Cobre desde o controle de presença até a votação, apuração de quórum, geração de ata e um dashboard visual — com suporte a presença ao vivo via Zoom.

## O que o projeto faz

**Controle de presença**
- Presença manual, lançada diretamente nas abas de cada conselho (`Inst_*`, `Vice_*`, `Reg_*`).
- Presença ao vivo via Zoom: um Cloudflare Worker recebe os webhooks do Zoom (entrada/saída de participante, início/fim de reunião) e repassa para este Apps Script, publicado como Web App (`doGet`/`doPost`).
- As duas fontes convivem na mesma sessão: quem entra pelo Zoom é somado a quem foi lançado manualmente como plano de contingência ("Plano B"), sem que uma fonte apague a outra.
- Nomes captados no Zoom são casados com os nomes oficiais da planilha por uma tabela de apelidos (`ALIASES_ZOOM`), com correspondência aproximada por ente/titular/suplente.

**Votação**
- Abertura de votação por pauta (ex.: eleição da Vice-Presidência, Regimento Interno), com tira-foto (snapshot) de quem está apto a votar no momento.
- Lançamento em massa ("Todos presentes favoráveis") ou exceções pontuais por linha.
- Cálculo de quórum e resultado considerando o peso de voto de cada ente (município/estado).

**Dashboard e ata**
- Botão que gera/atualiza uma aba `Dashboard_Reuniao` com cartões de resumo, gráfico de pizza da distribuição de votos, gráfico de presença por segmento, mapa de presença por município e o log dos últimos votos.
- Geração automática do texto da ata a partir dos dados apurados.

**Painel lateral (Sidebar.html)**
- Interface HTML aberta como barra lateral ou modal na planilha, usada pela mesa da reunião para: selecionar conselho/pauta, abrir votação, colar lista de presença do Zoom como plano B, acompanhar status da presença ao vivo, buscar/registrar exceções de voto e consultar o resumo em tempo real.

**Depuração e auditoria**
- Abas ocultas de log (`_LOG_ZOOM`, `_LOG_ZOOM_AO_VIVO`, `_LOG_VOTOS`, `_DEBUG_PIPEDREAM`, `_DEBUG_MATCH_ZOOM`, `_ERROS_ZOOM`, `_PENDENCIAS_ZOOM`) para rastrear eventos recebidos, tentativas de casamento de nome e pendências não identificadas.

## Estrutura

- `Code.gs` — toda a lógica de backend: menu do Sheets, votação, presença, integração com Zoom/Cloudflare, dashboard e geração de ata.
- `Sidebar.html` — interface do painel lateral usado durante a reunião.
- `appsscript.json` — manifesto do projeto (runtime V8 e configuração de Web App). **Confira as configurações reais de implantação no seu projeto Apps Script antes de reimplantar**, pois esse arquivo foi recriado manualmente e pode não refletir 100% o que está publicado hoje.

## Segredos e configuração

O código não tem nenhum token, senha ou chave gravada diretamente — apenas nomes de propriedades (`PropertiesService`) que são preenchidas em tempo de execução pelos menus "Configurar Zoom ao vivo" etc. Antes de reaproveitar este projeto:

1. Abra a planilha → Extensões → Apps Script.
2. Rode `Configurar Zoom ao vivo` para gravar o Secret Token do Zoom e gerar o token público do webhook.
3. Em **Implantar → Gerenciar implantações**, publique o Web App e aponte a URL `/exec` no Cloudflare Worker (variável `APPS_SCRIPT_URL`).

## Instalação

1. Crie um projeto Apps Script vinculado a uma cópia da planilha (Extensões → Apps Script).
2. Cole o conteúdo de `Code.gs` no arquivo `Code.gs` do projeto.
3. Crie um arquivo HTML chamado `Sidebar` e cole o conteúdo de `Sidebar.html`.
4. Salve, recarregue a planilha e use o menu **REUNIÃO** que aparece na barra de menus.
