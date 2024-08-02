#pip install PyPDF2 pandas selenium chromedriver_autoinstaller tabula-py openpyxl jpype1


import time
import ctypes
import PyPDF2
import os
import pandas as pd
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import simpledialog
import tabula
import shutil

# Instalar o ChromeDriver e obter o caminho
chromedriver_autoinstaller.install()

# Ler dados da planilha
planilha = pd.read_excel(r"C:\vagas\planilhaCIE.xlsx", sheet_name='p1')
codigocie = planilha['codigo'].tolist()

# Configurar o WebDriver
pasta_download = r"C:\vagas\baixados"
chrome_options = webdriver.ChromeOptions()
prefs = {"download.default_directory": pasta_download}
chrome_options.add_experimental_option("prefs", prefs)

# Configurar o serviço do WebDriver
servico = Service()

# Exibir mensagem de início
# mensagem = "Realizando os Downloads, esse processo pode demorar"
# titulo = "INICIO"
# ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 1)

# Digitar dados de login
root = tk.Tk()
root.withdraw()
nomelogin = simpledialog.askstring("Insira seu login:", "Login na SED")
senhalogin = simpledialog.askstring("Insira sua Senha:", "Senha na SED")

# Abrir o navegador
navegador = webdriver.Chrome(service=servico, options=chrome_options)

# Site
navegador.get("https://sed.educacao.sp.gov.br/")
time.sleep(5)

# Logar no site
navegador.find_element('xpath', '//*[@id="name"]').send_keys(nomelogin)
navegador.find_element('xpath', '//*[@id="senha"]').send_keys(senhalogin)
time.sleep(0.5)
navegador.find_element('xpath', '//*[@id="botaoEntrar"]').click()

# Selecionar o perfil na SED
time.sleep(2)
navegador.find_element('xpath', '//*[@id="sedUiModalWrapper_1body"]/ul/li[2]/a').click()
time.sleep(2)

# Fechar um aviso
# navegador.find_element('xpath', '//*[@id="sedUiModalWrapper_1close"]').click()
# time.sleep(3)

# Digitar matricular aluno
navegador.find_element('xpath', '//*[@id="decorMenuFilterTxt"]').send_keys("Matricular Aluno")
time.sleep(3)
navegador.find_element('xpath', '/html/body/div[3]/div/div/aside/div/ul/li/a').click()
time.sleep(3)

# Digitar o código CIE, pesquisar e fazer o download
for cie in codigocie:
    # Usar o filtro do CIE
    time.sleep(2)
    navegador.find_element('xpath', '/html/body/div[3]/div/div/main/div[1]/form/div/div/div/fieldset/div[1]/div/div/button/span').click()
    time.sleep(3)
    navegador.find_element('xpath', '/html/body/div[3]/div/div/main/div[1]/form/div/div/div/fieldset/div[1]/div/div/div/div[2]/ul/li[3]/a/span').click()
    time.sleep(3)
    navegador.find_element('xpath', '//*[@id="codigoEscolaCIE"]').send_keys(str(cie))
    time.sleep(4)
    navegador.find_element('xpath', '//*[@id="btnPesquisar"]').click()
    time.sleep(4)
    navegador.find_element('xpath', '//*[@id="btnPesquisar"]').click()
    time.sleep(4)
    # Geral pdf
    navegador.find_element('xpath', '/html/body/div[3]/div/div/main/div[2]/div/div/div[1]/button[5]').click()
    time.sleep(4)
    navegador.find_element('xpath', '/html/body/div[5]/div/div/div[3]/button[1]').click()
    time.sleep(10)
    # Limpar
    navegador.find_element('xpath', '/html/body/div[3]/div/div/main/div[1]/form/div/div/div/fieldset/div[6]/button[2]').click()
    time.sleep(4)

# Fechar navegador
navegador.quit()

# Código para juntar os PDFs
juntar = PyPDF2.PdfMerger()
pasta_juntar = r"C:\vagas\baixados"
saida_pdf = r"C:\vagas\juntoMatricula.pdf"

if not os.path.exists(pasta_juntar):
    print(f"A pasta '{pasta_juntar}' não existe.")
else:
    lista_arquivos = [arquivo for arquivo in os.listdir(pasta_juntar) if arquivo.lower().endswith(".pdf")]

    if not lista_arquivos:
        print(f"A pasta '{pasta_juntar}' não contém arquivos PDF.")
    else:
        for arquivo in lista_arquivos:
            caminho_arquivo = os.path.join(pasta_juntar, arquivo)
            juntar.append(caminho_arquivo)

        juntar.write(saida_pdf)
time.sleep(2)

# Exibir mensagem de início
mensagem = "Arquivo em PDF pronto. Agora Vamos trabalhar de PDF para Excel"
titulo = "PDF Para Excel"
ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 1)

############### para excel########
##############

# Função para corrigir texto com parágrafos sem espaços
def fix_paragraphs(text):
    # Verificar se o valor é uma string
    if isinstance(text, str):
        # Adicionar espaços entre parágrafos (separados por '\n')
        return ' '.join(text.split('\n'))
    else:
        # Se não for uma string, retornar o valor original
        return text

# Leitura do PDF
lista_tabelas = tabula.read_pdf(r"C:\vagas\juntoMatricula.pdf", pages="all", lattice=True)

# Criar um DataFrame a partir de cada tabela extraída
dataframes = []
for tabela in lista_tabelas:
    df = pd.DataFrame(tabela)
    dataframes.append(df)

# Salvar os DataFrames em um arquivo Excel
with pd.ExcelWriter(r"C:\vagas\separados.xlsx", engine='openpyxl') as writer:
    for i, df in enumerate(dataframes, 1):
        df.to_excel(writer, sheet_name=f'Tabela_{i}', index=False)

# Leitura do arquivo Excel com todas as planilhas
excel_file = pd.ExcelFile(r"C:\vagas\separados.xlsx")

# Inicializar um DataFrame vazio
combined_df = pd.DataFrame()

# Iterar sobre todas as planilhas no arquivo Excel e concatená-las
for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(excel_file, sheet_name)
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# Lista de colunas com parágrafos sem espaços
columns_to_fix = ['Escola', 'Tipo de Ensino', 'Tipo de Classe', 'Turma']  # Substitua pelos nomes reais das colunas afetadas

# Corrigir texto com parágrafos sem espaços para cada coluna
for column in columns_to_fix:
    combined_df[column] = combined_df[column].apply(fix_paragraphs)

# Salvar o DataFrame combinado em um novo arquivo Excel
combined_df.to_excel(r"C:\vagas\compilado.xlsx", index=False)

# Reordenar as colunas
combined_df = combined_df[['Série', 'Turma'] + [col for col in combined_df.columns if col not in ['Série', 'Turma']]]

# Salvar o DataFrame combinado em um novo arquivo Excel
combined_df.to_excel(r"C:\vagas\compilado.xlsx", index=False)

# Alterar nomes da planilha
# Carregar o arquivo compilado.xlsx
planilhaNova = pd.read_excel(r'C:\vagas\compilado.xlsx')

# Função para atualizar os valores da coluna 'Série' com base na coluna 'Turma'
def atualizar_serie(turma):
    if '2° ANO' in turma:
        return '2i'
    elif '3° ANO' in turma:
        return '3i'
    elif '4° ANO' in turma:
        return '4i'
    elif '5° ANO' in turma:
        return '5i'
    else:
        return turma  # Se a turma não contiver nenhuma das palavras, manter o valor original

# Filtrar as linhas que atendem às condições
condicoes = planilhaNova['Turma'].str.contains('2° ANO|3° ANO|4° ANO|5° ANO', na=False)

# Aplicar a função apenas às linhas filtradas e atribuir o resultado à coluna 'Série'
planilhaNova.loc[condicoes, 'Série'] = planilhaNova.loc[condicoes, 'Turma'].apply(atualizar_serie)

# Salvar as alterações de volta ao arquivo compilado.xlsx
planilhaNova.to_excel(r"C:\vagas\final.xlsx", index=False)

# Exibir mensagem de conclusão
mensagem = "Trabalho Concluído"
titulo = "Final"
ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 1)




# Definir caminhos dos arquivos a serem deletados
arquivos_a_deletar = [r"C:\vagas\compilado.xlsx",r"C:\vagas\separados.xlsx",r"C:\vagas\juntoMatricula.pdf"]

# Deletar os arquivos especificados
for arquivo in arquivos_a_deletar:
    if os.path.exists(arquivo):
        os.remove(arquivo)

# Deletar a pasta 'baixados' e seu conteúdo
pasta_baixados = r"C:\vagas\baixados"
if os.path.exists(pasta_baixados):
    shutil.rmtree(pasta_baixados)

# Exibir mensagem de conclusão
mensagem = "Arquivos temporários deletados. Trabalho Concluído"
titulo = "Final"
ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 1)
