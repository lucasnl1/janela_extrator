import pyodbc
import pandas as pd

import os

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from tkinter import *
from tkinter import messagebox
from tkcalendar import DateEntry
import tkinter as tk
from tkinter import ttk


#cria a conecxão ao banco
def get_db_connection():
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=srv-bd;'
            'DATABASE=Gestor;'
            'UID=ultrasyst;'
            'PWD=masterkey;'
            'Trusted_Connection=no;'
        )
        return conn
    #checa a conexão e retorna em caso de erro
    except pyodbc.Error as e:
        print("Erro na conexão:", e)
        return None
    
#conecta ao banco
def consultar_bd(fornecedor, natureza, localizacao, periodo_In, periodo_fin):
    conn = get_db_connection()
    #checa e retorna caso a conexão falhe
    if conn is None:
        print("Falha na conexão a banco de dados.")
        return []
    #executa o SQL
    try:
        cur = conn.cursor()
        query = '''
        SELECT 
        LOCALIZACAO, 
        DATA,
        CLI_NOME,
        CLI_CPF,
        UF,
        CIDADE,
        NATUREZA,
        VENDEDOR,
        PRODUTO,
        QUANTIDADE, 
        VALOR_TOTAL
    FROM VW_FATURAMENTO_DETALHADO
    JOIN CLIENTE C ON CLI_CODI = COD_CLIENTE_GERAL
    WHERE FOR_CODI = ?
    AND NATUREZA IN ({})
    AND LOCALIZACAO = ?
    AND CONVERT(DATETIME, DATA, 103) BETWEEN ? AND ?
    ORDER BY DATA DESC
        '''
         # Formatando corretamente as listas de parâmetros
        query = query.format(','.join(['?']*len(natureza)))
        parametros = [fornecedor] + natureza + [localizacao] + [periodo_In, periodo_fin]
        cur.execute(query, parametros)

        #retorna os dados recolhidos no select
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows
    #retorna em caso de erro no recolhimento dos dados
    except pyodbc.Error as e:
        print("Erro ao juntar os dados data:", e)
        return []
    
#botão consultar recolhendo os dados da caixa de entrada   
def on_consulta():
    fornecedor = forn_var.get()
    natureza = nat_var.get().split(',')  # Separando a string em lista
    localizacao = loc_var.get()
    periodo_In = periodo_ini_var.get()
    periodo_fin = periodo_fin_var.get()
    resultados = consultar_bd(fornecedor, natureza, localizacao, periodo_In, periodo_fin)
    exibir_resultados(resultados)  # Passa os resultados para exibir na função exibir_resultados
      
#botão salvar em excel recolhendo os dados da caixa de entrada   
def on_salva():
    fornecedor = forn_var.get()
    natureza = nat_var.get().split(',')  # Separando a string em lista
    localizacao = loc_var.get()
    periodo_In = periodo_ini_var.get()
    periodo_fin = periodo_fin_var.get()
    resultados = consultar_bd(fornecedor, natureza, localizacao, periodo_In, periodo_fin)
    filename = filename_var.get()
    if not filename:
        messagebox.showwarning("Atenção", "Por favor, insira o nome para o arquivo")
        return
    filepath = f"extraido/{filename}.xlsx"
    save_to_excel(resultados, filepath) # Passando resultados para serem salvos na função save_to_excel

#exibe os 5 primeiros resultados para validação prévia
def exibir_resultados(resultados):
    janela_resultados = Toplevel(root)
    janela_resultados.title("Resultados da Consulta")
     
    for i in range(min(5, len(resultados))):
        row = resultados[i]
        ttk.Label(janela_resultados, text=f"Resultado {i+1}: {row}").grid(column=0, row=i, padx=10, pady=5)
    
    global resultados_df
    colunas = [
        'DISTRIBUIDOR', 'DATA FATURAMENTO', 'RAZÃO SOCIAL', 'CNPJ', 'ESTADO',
        'MUNICÍPIO', 'NATUREZA', 'VENDEDOR', 'DESCR ITEM', 'QTDE', 'VALOR TOTAL'
    ]
    resultados_df = pd.DataFrame(resultados, columns=colunas)

# Cria a pasta caso não exista e salva o arquivo em xls para excell
def save_to_excel(rows, filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    # Converte as linhas em tuplas
    data = [tuple(row) for row in rows]

    df = pd.DataFrame(data, columns=[
        'DISTRIBUIDOR', 'DATA FATURAMENTO', 'RAZÃO SOCIAL', 'CNPJ', 'ESTADO', 'MUNICÍPIO', 'NATUREZA', 'VENDEDOR', 'DESCR ITEM', 'QTDE', 'VALOR TOTAL BRUTO'])
    df.to_excel(filename, index=False)
    workbook = load_workbook(filename)
    sheet = workbook.active
    # Seta a formatação do texto e cores da planilha
    red_fill = PatternFill(start_color="FFFF5050", end_color="FFFF5050", fill_type="solid")
    white_font = Font(color="FFFFFF", name="Aptos Narrow", bold=True)
    for cell in sheet[1]:
        cell.fill = red_fill
        cell.font = white_font
    
    workbook.save(filename)
    print(f"Aqruivo salvo como: {filename}")
    messagebox.showinfo("Successo", f"Aqruivo salvo como: {filename}")
 
# Botão enviar email
def on_enviar():
    smtp_server = 'smtps.uhserver.com'
    smtp_port = 465
    smtp_user = 'ti@veteagro.com.br'
    smtp_password = 'Veteagro@16'
    from_addr = 'ti@veteagro.com.br'
    to_addr = email_var.get()  # Obtém o e-mail digitado na caixa de texto

    #retorna caso a caixa de email esteja vazia
    if not to_addr:
        messagebox.showwarning("Atenção", "Por favor, insira o e-mail do destinatário.")
        return

    subject = f'RELATORIO SELL OUT {filename_var.get()}'
    body = f'Segue em anexo relatório {filename_var.get()} para análise e redirecionamento'
    file_path = f'extraido/{filename_var.get()}.xlsx'
    
    send_email_with_attachment(smtp_server, smtp_port, smtp_user, smtp_password, from_addr, to_addr, subject, body, file_path)
# Envia o email
def send_email_with_attachment(smtp_server, smtp_port, smtp_user, smtp_password, from_addr, to_addr, subject, body, file_path):
    try:
        # Cria a mensagem
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = subject

        # Adiciona o corpo do email
        msg.attach(MIMEText(body, 'plain'))

        # Anexa o arquivo
        with open(file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
            msg.attach(part)

        # Configura o servidor SMTP e envia o email
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(from_addr, to_addr, msg.as_string())

        print("Email enviado com sucesso!")
        messagebox.showinfo("Sucesso", "Email enviado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar o email: {e}")

# Inicia a interface gráfica
root = tk.Tk()
root.title("Consulta ao Banco de Dados")

# Listas predefinidas
fornecedor = ['00155', '00116', '00290', '01076', '00993', '00790', '01069', '00207', '00224', '00262', '00940']
natureza = ['VEN', 'BOV', 'DVE']
localizacao = ['000', '001', '002', '004']

# Variáveis para armazenar as seleções
forn_var = tk.StringVar()
nat_var = tk.StringVar()
loc_var = tk.StringVar()
periodo_ini_var = tk.StringVar()
periodo_fin_var = tk.StringVar()
filename_var = StringVar()
email_var = StringVar()

# Caixas de listas
ttk.Label(root, text="Fornecedor:").grid(column=0, row=0, padx=10, pady=5)
fornecedor_combobox = ttk.Combobox(root, textvariable=forn_var, values=[f"{codigo} - {nome}" for codigo, nome in fornecedor])
fornecedor_combobox.grid(column=1, row=0, padx=10, pady=5)

ttk.Label(root, text="Natureza:").grid(column=0, row=1, padx=10, pady=5)
ttk.Combobox(root, textvariable=nat_var, values=natureza).grid(column=1, row=1, padx=10, pady=5)

ttk.Label(root, text="Localização:").grid(column=0, row=2, padx=10, pady=5)
ttk.Combobox(root, textvariable=loc_var, values=localizacao).grid(column=1, row=2, padx=10, pady=5)

ttk.Label(root, text="Data Inicial:").grid(column=0, row=4, padx=10, pady=5)
data_ini_entry = DateEntry(root, textvariable=periodo_ini_var, date_pattern='yyyy-mm-dd')
data_ini_entry.grid(column=1, row=4, padx=10, pady=5)

ttk.Label(root, text="Data Final:").grid(column=0, row=5, padx=10, pady=5)
data_fin_entry = DateEntry(root, textvariable=periodo_fin_var, date_pattern='yyyy-mm-dd')
data_fin_entry.grid(column=1, row=5, padx=10, pady=5)

# Botão para realizar a consulta
ttk.Button(root, text="Consultar", command=on_consulta).grid(column=1, row=6, padx=10, pady=5)


# Label e Entry para o nome do arquivo
Label(root, text="Nome do arquivo:").grid(column=0, row=7, padx=10, pady=5)
Entry(root, textvariable=filename_var).grid(column=1, row=7, padx=10, pady=5)

# Botão para salvar
ttk.Button(root, text="Salvar", command=on_salva).grid(column=1, row=8, padx=10, pady=5)
# Esta caixa só aparece após clicar em salvar
# Entry para o email a ser enviado com o arquivo
ttk.Label(root, text="deseja enviar o arquivo por email?").grid(column=0, row=9, padx=10, pady=5)
ttk.Entry(root, textvariable=email_var).grid(column=1, row=9, padx=10, pady=5)
# Botão para enviar
ttk.Button(root, text="enviar", command=on_enviar).grid(column=0, row=10, padx=10, pady=5)
ttk.Button(root, text="cancelar", command=root.destroy).grid(column=1, row=10, padx=10, pady=5)

# Executa a interface gráfica
root.mainloop()