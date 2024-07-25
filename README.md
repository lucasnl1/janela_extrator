
# Projeto de Consulta e Relatório de Banco de Dados

## Descrição

Este projeto é uma aplicação Python que permite a consulta de dados em um banco de dados SQL Server, a exibição dos resultados e a exportação desses dados para um arquivo Excel. Além disso, a aplicação possibilita o envio do arquivo Excel por e-mail diretamente pela interface gráfica.

## Funcionalidades

1. **Consulta ao Banco de Dados:**
   - Conexão segura com um banco de dados SQL Server usando `pyodbc`.
   - Consulta personalizada com base em parâmetros fornecidos pelo usuário, como fornecedor, natureza, localização e período.

2. **Exportação para Excel:**
   - Geração de relatórios em formato Excel (.xlsx) com formatação de cabeçalhos.
   - Salvamento automático dos arquivos em um diretório especificado.

3. **Envio de E-mail:**
   - Envio do arquivo Excel gerado para um e-mail especificado pelo usuário.
   - Suporte a servidores SMTP para envio seguro.

## Tecnologias Utilizadas

- Python
- `pyodbc` para conexão ao banco de dados SQL Server
- `pandas` para manipulação de dados
- `openpyxl` para criação e formatação de arquivos Excel
- `tkinter` para criação da interface gráfica
- `tkcalendar` para seleção de datas

## Requisitos

- Python 3.x
- Bibliotecas Python necessárias (podem ser instaladas via `pip`):
  - `pyodbc`
  - `pandas`
  - `openpyxl`
  - `tkcalendar`

## Instalação

1. Clone o repositório:
   ```
   git clone <URL_DO_REPOSITORIO>
   cd <NOME_DO_DIRETORIO>
   ```

2. Instale as dependências:
   ```
   pip install pyodbc pandas openpyxl tkcalendar
   ```

3. Configure os detalhes do banco de dados e e-mail:
   - Edite as informações de conexão ao banco de dados e credenciais de e-mail no código, caso necessário.

## Uso

1. Execute o script principal para iniciar a interface gráfica:
   ```python
   python <nome_do_script>.py
   ```

2. Preencha os campos necessários na interface gráfica:
   - **Fornecedor, Natureza, Localização:** Selecione os valores desejados.
   - **Data Inicial e Data Final:** Selecione o intervalo de datas.
   - **Nome do Arquivo:** Insira o nome para o arquivo Excel.
   - **E-mail do Destinatário:** (Opcional) Insira o e-mail para enviar o arquivo.

3. Clique em "Consultar" para realizar a consulta no banco de dados.

4. Clique em "Salvar" para salvar os resultados em um arquivo Excel.

5. Clique em "Enviar" para enviar o arquivo por e-mail.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests para melhorar o projeto.
