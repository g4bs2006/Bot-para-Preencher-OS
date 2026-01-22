# Automação de Ordens de Serviço - Neovero (Orbis)

Status: Em Produção (v1.0)
Stack: Python 3.10+ | Playwright | Polars | Pydantic

Este projeto consiste em uma solução de Engenharia de Automação para abertura e gestão em massa de Ordens de Serviço (O.S.) no sistema Neovero. A arquitetura foi desenhada para superar limitações comuns em sistemas legados, como instabilidade de sessão, carregamento via iframes e performance na leitura de dados.

## Diferenciais Técnicos

1. Alta Performance de Dados: Utiliza a biblioteca Polars (baseada em Rust) para leitura e validação instantânea de planilhas Excel, garantindo velocidade superior ao Pandas.
2. Page Object Model (POM): O código segue estritamente o padrão POM. Cada tela do sistema (Login, Menu, Equipamento, OS) possui sua própria classe, facilitando a manutenção e isolando a lógica de interação.
3. Algoritmo "Caçador de Iframes": Implementação de busca recursiva que varre o frame principal e todos os iframes filhos (cross-origin ou não) para localizar elementos dinâmicos ou ocultos.
4. Estratégia "Terra Arrasada" (Stability): Para evitar vazamento de contexto e erros de sessão entre execuções, o sistema realiza um reload forçado da página (F5) e limpeza de memória após cada transação, garantindo um ambiente limpo para a próxima O.S.
5. Tipagem e Validação: Uso do Pydantic para validação rigorosa dos dados de entrada antes da execução do navegador.


1. Pré-requisitos
Certifique-se de ter o Python 3.10+ instalado.

2. Instalação das Dependências
Execute no terminal:
pip install -r requirements.txt
playwright install chromium

3. Configuração de Ambiente
Crie um arquivo chamado '.env' na raiz do projeto e insira suas credenciais:
NEOVERO_URL="https://orbis.neovero.com/login"
NEOVERO_USER="seu_usuario"
NEOVERO_PASS="sua_senha"

## Execução

1. Prepare os dados:
Coloque sua planilha Excel devidamente formatada na pasta `data/input/` com o nome `dados.xlsx`.

2. Inicie a automação:
python src/main.py

O sistema iniciará o processo de login, varredura de equipamentos e preenchimento das ordens. O progresso pode ser acompanhado via terminal, com logs detalhados de sucesso, avisos (skip) e falhas.

## Tratamento de Erros e Logs

O projeto utiliza a biblioteca Loguru para registro de atividades.
- Logs de Execução: Exibidos no terminal em tempo real.
- Screenshots de Erro: Em caso de falha crítica (ex: elemento não encontrado), um print da tela é salvo automaticamente em `data/logs/` para facilitar o debug.
- Logs de Arquivo: Um histórico completo é salvo em `data/logs/execution.log`.

## Testes

Para validar as regras de negócio e a leitura de dados sem abrir o navegador:
pytest tests/
