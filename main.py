### importação das bibliotecas que serão utilizadas para o projeto ###

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from winotify import Notification, audio
import time
import pandas as pd
import warnings
import os
import openpyxl
import smtplib

warnings.filterwarnings("ignore")

#--------------

# lendo o arquivo "Produto" para descobrir o produto a ser procurado, preço mínimo e preço máximo

with open("./Produto.txt", "r", encoding="utf-8") as arquivo:
    arquivo_lido = arquivo.read()
    arquivo_lido_tratado = arquivo_lido.split(",")
    produto = arquivo_lido_tratado[0].strip()
    preco_min = float(arquivo_lido_tratado[1].strip())
    preco_max = float(arquivo_lido_tratado[2].strip())
    arquivo.close()

#--------------

# baixando a versão mais recente do ChromeDriver e instanciando o navegador

service = Service(ChromeDriverManager().install())

browser = webdriver.Chrome(service=service)

url = "https://www.zoom.com.br/"

# abrindo o navegador e pesquisando pelo produto
 
browser.get(url)

browser.maximize_window()

barra_pesquisa = browser.find_element(By.XPATH, "//input[@placeholder='Digite sua busca...']")

barra_pesquisa.send_keys(produto)

btn_pesquisa = browser.find_element(By.XPATH, "//button[@type='button']")

btn_pesquisa.click()

time.sleep(5)

#--------------

# criando os arquivos em .xlsx para armazenar todos os produtos encontrados e àqueles que estejam dentro da faixa de preço desejada

arq_todos_produtos = openpyxl.Workbook()

planilha_todos_produtos = arq_todos_produtos.active

arq_produtos_selecionados = openpyxl.Workbook()

planilha_produtos_selecionados = arq_produtos_selecionados.active

#--------------

# descobrindo a quantidade de páginas que será necessário percorrer e salvando o link na lista lst_paginas

lst_paginas = list()

total_paginas = browser.find_elements(By.CLASS_NAME, "Paginator_pageLink__GsWrn")

for i, pagina in enumerate(total_paginas):

    if i > 1 and i < len(total_paginas) - 1:
        lst_paginas.append(pagina.get_attribute("href"))

#--------------

# criação da função buscar_elementos para realizar a raspagem dos dados

def buscar_elementos():

    global produtos
    global sites
    global precos
    global parcelamentos
    global link_anuncios

    produtos = browser.find_elements(By.TAG_NAME, "h2")

    sites = browser.find_elements(By.TAG_NAME, "h3")

    precos = browser.find_elements(By.XPATH, "//p[@data-testid='product-card::price']")

    parcelamentos = browser.find_elements(By.XPATH, "//p[@data-testid='product-card::installment']")

    link_anuncios = browser.find_elements(By.CLASS_NAME, "SearchCard_ProductCard_Inner__7JhKb")

#--------------

buscar_elementos()

x = 1

dict_produtos = dict()

# percorre a primeira página do site

for i in range(len(produtos)):

    preco = precos[i].text.replace("R$", "").strip()
    posicao_casa_decimal = preco.find(",")
    preco_tratado = int(preco[:posicao_casa_decimal].replace(".", ""))

    if ("via" in sites[i].text):

        if preco_min <= preco_tratado <= preco_max:
            produtos_selecionados = [
                "Site: " +  sites[i].text[16:],
                "Preco: " + precos[i].text,
                "Parcelamento: " + parcelamentos[i].text,
                "Link do anúncio: " + link_anuncios[i].get_attribute("href")
            ]
            dict_produtos[produtos[i].text] = produtos_selecionados

        preencher_cabecalho = [["Produto", "Site", "Preco", "Parcelamento", "Link do anúncio"]]

        gravar_dados = [[
            produtos[i].text, 
            sites[i].text[16:], 
            precos[i].text, 
            parcelamentos[i].text, 
            link_anuncios[i].get_attribute("href")
        ]]

        if x == 1:
            for linha in preencher_cabecalho:
                planilha_todos_produtos.append(linha)
                planilha_produtos_selecionados.append(linha)
            for linha in gravar_dados:
                planilha_todos_produtos.append(linha)
                if preco_min <= preco_tratado <= preco_max:
                    planilha_produtos_selecionados.append(linha)
        else:
            for linha in gravar_dados:
                planilha_todos_produtos.append(linha)
                if preco_min <= preco_tratado <= preco_max:
                    planilha_produtos_selecionados.append(linha)
    x += 1

# verifica se há mais de uma página a ser percorrida

if len(lst_paginas) > 0:

    for pagina in lst_paginas:

        browser.execute_script(f"window.open('{pagina}', '_blank')")

        browser.switch_to.window(browser.window_handles[1])

        buscar_elementos()

        for i in range(len(produtos)):

            preco = precos[i].text.replace("R$", "").strip()
            posicao_casa_decimal = preco.find(",")
            preco_tratado = int(preco[:posicao_casa_decimal].replace(".", ""))

            if ("via" in sites[i].text):

                if preco_min <= preco_tratado <= preco_max:
                    produtos_selecionados = [
                        "Site: " +  sites[i].text[16:],
                        "Preco: " + precos[i].text,
                        "Parcelamento: " + parcelamentos[i].text,
                        "Link do anúncio: " + link_anuncios[i].get_attribute("href")
                    ]
                    dict_produtos[produtos[i].text] = produtos_selecionados

                gravar_dados = [[
                    produtos[i].text, 
                    sites[i].text[16:], 
                    precos[i].text, 
                    parcelamentos[i].text, 
                    link_anuncios[i].get_attribute("href")
                ]]

                for linha in gravar_dados:
                    planilha_todos_produtos.append(linha)
                    if preco_min <= preco_tratado <= preco_max:
                        planilha_produtos_selecionados.append(linha)
                

        browser.close()

        browser.switch_to.window(browser.window_handles[0])

browser.quit()

# salva os arquivos

arq_todos_produtos.save(r"./todos_produtos.xlsx")

arq_produtos_selecionados.save(r"./produtos_selecionados.xlsx")

#--------------

df = pd.read_excel(r"./produtos_selecionados.xlsx")

# verifica se o arquivo produtos_selecionados contém algum dado para disparar o e-mail

if df.shape[0] >= 1:

#--------------

    # declaração das variáveis

    smtp_servidor = "insira o smtp de e-mail aqui"

    smtp_porta = 587

    login = "insira seu e-mail aqui"

    senha = "insira sua senha aqui"

    # cria e starta a conexão com o servidor

    servidor = smtplib.SMTP(smtp_servidor, smtp_porta)

    servidor.starttls()

    servidor.login(login, senha)

    # cria o e-mail

    email = MIMEMultipart()
    email["From"] = "insira seu e-mail aqui"
    email["To"] = "insira o e-mail do destinatário aqui"
    email["Subject"] = "RPA Python"

    caminho_anexo = r"./produtos_selecionados.xlsx"
    anexo = open(caminho_anexo, "rb")

    obj_anexo = MIMEBase('application', 'octet-stream')
    obj_anexo.set_payload(anexo.read())
    encoders.encode_base64(obj_anexo)
    obj_anexo.add_header('Content-Disposition', f'{anexo}', filename=os.path.basename(caminho_anexo))
    anexo.close()

    # faz com que o Windows notifique caso seja encontrado algum produto dentro da faixa de preço informada
  
    win_notificacao = Notification(
                            app_id="RPA Python",
                            title=f"{produto} encontrado (a)",
                            msg="Verifique seu e-mail para mais informações",
                            duration="long",
                            icon=fr"{os.getcwd()}\RPA.jpg"
                            )
    win_notificacao.set_audio(sound=audio.Mail, loop=False)
    win_notificacao.show()

    # se houver apenas um produto encontrado informa os detalhes do produto no corpo do e-mail, caso contrário anexa o arquivo em excel com todos os produtos encontrados

    if len(dict_produtos) == 1:
        mensagem = f"""
        <p> "insira um nome aqui", o bot RPA Python encontrou o produto que você está procurando dentro da faixa de preço informada. </p>
        <p> Segue abaixo as informações do produto localizado: </p>
        <p> <b> Produto: {list(dict_produtos.keys())[0]} </b> </p>
        <p> <b> {dict_produtos[list(dict_produtos.keys())[0]][0]} </b> </p>
        <p> <b> {dict_produtos[list(dict_produtos.keys())[0]][1]} </b> </p>
        <p> <b> {dict_produtos[list(dict_produtos.keys())[0]][2]} </b> </p>
        <p> <b> {dict_produtos[list(dict_produtos.keys())[0]][3]} </b> </p>
        """
    else:
        mensagem = f"""
        <p> "insira um nome aqui", o bot RPA Python encontrou alguns produtos que você está procurando dentro da faixa de preço informada. </p> 
        <p> Segue em anexo as informações dos produtos localizados. </p>
        """
        email.attach(obj_anexo)

    email.attach(MIMEText(mensagem, "html"))

    # dispara o e-mail e fecha a conexão com o servidor

    servidor.sendmail(email["From"], email["To"], email.as_string().encode("utf-8"))
    servidor.quit()