
# Controle de Estoque e Vendas - Monster Alpha Suplements

## Descrição

Este programa é uma aplicação de controle de estoque e vendas, desenvolvida para gerenciar a entrada e saída de produtos, registrar informações detalhadas sobre os produtos, gerar relatórios financeiros e manter um histórico de todas as transações. A interface gráfica é construída usando Tkinter, e os dados são armazenados em arquivos Excel, com a opção de gerar relatórios em PDF com marca d'água personalizada.

## Funcionalidades

- Registrar novos produtos e suas entradas no estoque.
- Registrar saídas de produtos e calcular o saldo das vendas.
- Pesquisar produtos pelo nome e gerenciar o estoque com sugestão automática de nomes e sabores.
- Gerar relatórios mensais de gastos e lucro/prejuízo em formato PDF, com marca d'água personalizada.
- Excluir registros de produtos quando necessário.

## Pré-requisitos

- Python 3.x instalado.
- Bibliotecas necessárias listadas no `requirements.txt`.

## Instruções para Instalação e Execução

1. **Instale as Dependências:**
   - Navegue até o diretório do projeto no terminal.
   - Execute o comando para instalar as dependências:
     ```bash
     pip install -r requirements.txt
     ```

2. **Crie o Executável:**
   - No terminal, execute o comando:
     ```bash
     pyinstaller --onefile --windowed --add-data "wsalpha.png;." --add-data "wsalpha.ico;." --icon=lobo.ico monster.py
     ```

3. **Executando o Programa:**
   - Após a criação do executável, navegue até o diretório `dist` onde o executável foi gerado.
   - Execute o programa clicando duas vezes no arquivo `monster.exe`.

## Notas

- Certifique-se de que os arquivos `wsalpha.png` e `wsalpha.ico` estão no mesmo diretório que o executável.
- O arquivo `relatorio_mensal.pdf` será gerado automaticamente na área de trabalho no início de cada mês ou quando solicitado manualmente.

## Problemas Comuns

- **Erro de Importação:** Verifique se todas as bibliotecas necessárias estão instaladas.
- **Erro de Caminho:** Certifique-se de que todos os arquivos (imagens) estão no local correto e que os caminhos estão configurados corretamente no código.
- **Ícone do Executável:** Se o ícone exibido não for o esperado, verifique o cache de ícones do Windows e recrie o atalho na área de trabalho.
