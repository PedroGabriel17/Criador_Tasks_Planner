# ================================================================== importações do app
from asyncio.windows_events import NULL
from cgitb import text
from distutils import command
from msilib.schema import ComboBox
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from turtle import color
from datetime import date
import requests
import json

# =====================================================================================

# =================================================================== importações token
import selenium
import time
from asyncio import sleep
from xml.dom.minidom import TypeInfo
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# =====================================================================================



# ===================================== FUNÇÃO PEGA TOKEN =============================
def catch_token():
    # =====================================================================================
    # ===================================== CODIGO PEGA TOKEN =============================
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"

    # configs para rodar sem abrir a janela
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument(f'user-agent={user_agent}')
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument("--disable-extensions")
    options.add_argument("--proxy-server='direct://'")
    options.add_argument("--proxy-bypass-list=*")
    options.add_argument("--start-maximized")
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--no-sandbox')
    navegador = webdriver.Chrome(options=options)
    # PARA RODAR COM A JANELA ABERTA, REMOVER O options=options

    # entra no Graph Explorer
    navegador.get('https://developer.microsoft.com/en-us/graph/graph-explorer')

    # captura botão de login
    button_login = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[2]/div[1]/div[2]/div[1]/div/button')))

    # click botão login com CTRL apertado pra abrir em outra aba
    ActionChains(navegador) \
    .key_down(Keys.CONTROL) \
    .click(button_login) \
    .key_up(Keys.CONTROL) \
    .perform()

    time.sleep(3)

    # recebe aba1 e aba2
    window1 = navegador.window_handles[0]
    window2 = navegador.window_handles[1]

    time.sleep(1)

    # muda pra aba2
    navegador.switch_to.window(window2)

    # print("====================================================================================================================================================================")
    # print(navegador.current_url)
    # print("====================================================================================================================================================================")

    # preenche email e clica em next
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]'))).send_keys("---------SEU EMAIL CORPORATIVO---------")
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'idSIButton9'))).click()

    time.sleep(1)

    # preenche senha e clica em next
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))).send_keys("---------SUA SENHA---------")
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idSIButton9"]'))).click()

    # clica no botão (Não)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idBtn_Back"]'))).click()

    # muda pra primeira aba
    navegador.switch_to.window(window1)

    time.sleep(2)

    # clica no botão de Token
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Pivot27-Tab3"]'))).click()

    # Clica no botão de copiar
    token = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div/div/div/div/label')))
    token = token.text
    print(token)
    return token

token = catch_token()

#=====================================================================================================
#================================= função que pega pessoa responsável ================================
#=====================================================================================================
def catch_people():
    if people_response.get() == "":
        return ""
        pass
    else:
        url_people = "https://graph.microsoft.com/v1.0/users/"+people_response.get()
        people = requests.get(url_people, headers={"Authorization": token})
        if people.status_code == 200:
            people = json.loads(people.content)
            return people["id"]
        else:
            return people_response.get()
#=====================================================================================================






#=====================================================================================================
#===================================== função que pega ultima task ===================================
#=====================================================================================================
def last_task_num():
    get_numbers = requests.get("https://graph.microsoft.com/v1.0/planner/plans/---------ID DO PLANO---------/tasks", headers={"Authorization": token})
    value_processed = json.loads(get_numbers.content)
    numero_projeto = value_processed["@odata.count"]
    return numero_projeto
#=====================================================================================================






#=====================================================================================================
#===================================== função que pega o ano atual ===================================
#=====================================================================================================
def take_year():
    counter = 0
    year = ""
    current_year = date.today()

    for x in str(current_year):
        if counter < 4:
            year = year + x
        else:
            break
        counter += 1
    return year
#======================================================================================================


#==========================================================================================================
#===================================== função que posta Task no Planner ===================================
#==========================================================================================================
def post_planner():
    tipo_projeto = type_task.get()
    nome_projeto = name_task.get()
    palavra_chave = key_word.get()
    category = category_input.get()
    numero_projeto = last_task_num()
    response_people = catch_people()
    start_date = start_date_input.get()
    duetime = duetime_input.get()
    year = take_year()
    counter = 1

    if nome_projeto == "" or palavra_chave == "" or tipo_projeto == "":
        tkinter.messagebox.showinfo( "ALERTA", "A Task deve possuir Tipo, Nome e Produto/Sistema/Projeto!")
    else:

        list_category = [
            "Comunicação",
            "Suporte",
            "Desenvolvimento",
            "Gestão",
            "Equipamento",
            "Especial"
        ]

        for x in list_category:
            if category == x:
                category = "category"+ str(counter)
            counter += 1

        # coloca o ano após o "planejamento anual"
        if tipo_projeto == "Planejamento Anual":
            tipo_projeto = tipo_projeto
        elif tipo_projeto == "Demanda corrente":
            tipo_projeto = tipo_projeto+" "+year


        # formata data de início
        if start_date == "":
            start_date = None
        else:
            start_date = start_date[len(start_date)-4]+start_date[len(start_date)-3]+start_date[len(start_date)-2]+start_date[len(start_date)-1]+"-"+start_date[len(start_date)-7]+start_date[len(start_date)-6]+"-"+start_date[len(start_date)-10]+start_date[len(start_date)-9]+"T00:00:00.0000-03:00"

        # formata data de prazo
        if duetime == "":
            duetime = None
        else:
            duetime = duetime[len(duetime)-4]+duetime[len(duetime)-3]+duetime[len(duetime)-2]+duetime[len(duetime)-1]+"-"+duetime[len(duetime)-7]+duetime[len(duetime)-6]+"-"+duetime[len(duetime)-10]+duetime[len(duetime)-9]+"T00:00:00.0000-03:00"

        # tratamento do nome com o padrão estabelecido
        nome_total = str(tipo_projeto)+" | "+str(nome_projeto)+" | "+palavra_chave+" | "+str(numero_projeto)+"/2022"


        assignments = {}
        appliedCategories = {}

        if response_people != "":
            assignments = {
                response_people: {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !"
                }
            }
        if category != "":
            appliedCategories = {category: True}
            

        json_schema = {
                "planId": "---------ID DO PLANO---------",
                "title": nome_total,
                "assignments": assignments,
                "dueDateTime": duetime,
                "startDateTime": start_date,
                "appliedCategories": appliedCategories
        }


        response = requests.post("https://graph.microsoft.com/v1.0/planner/tasks", json=json_schema, headers={"Authorization": token})
        
        if response.status_code == 201:
            tkinter.messagebox.showinfo( "Resultado", "Task criada com sucesso!")
        else:
            tkinter.messagebox.showinfo( "Resultado", "Não Foi possível criar a Task, verifique as informações digitadas ou reinicie o app")



# ======================== INÍCIO DA TELA ========================

screen = Tk()

screen.title("Criador de Tasks")
screen.geometry('800x500')
screen.configure(background='#58ade0')
screen.minsize(width='700', height='400')
screen.maxsize(width='900', height='600')


frame = Frame(screen)
frame.place(relx=0, rely=0, relwidth=1, relheight=1)
frame.configure(background="#58ade0")

# =============================== INPUTS E LABELS ===============================

# SELECT TIPO DA TASK
type_task = ttk.Combobox(screen, values=["Planejamento Anual", "Demanda corrente"])
type_task.current(0)
type_task.place(relx=0.02, rely=0.2, relwidth=0.29, relheight=0.06)
# LABEL
label_type_task = Label(frame, text="Tipo do Projeto", background="#58ade0", fg='white')
label_type_task.config(font=('Arial', 13))
label_type_task.place(relx=0.02, rely=0.13)

# INPUT NOME DA TASK
name_task = Entry(frame, borderwidth=0)
name_task.pack()
name_task.place(relx=0.355, rely=0.2, relwidth=0.29, relheight=0.06)
# LABEL
label_name_task = Label(frame, text="Nome da Tarefa", background="#58ade0", fg='white')
label_name_task.config(font=('Arial', 13))
label_name_task.place(relx=0.355, rely=0.13)


# INPUT PALAVRA-CHAVE
key_word = Entry(frame, borderwidth=0)
key_word.pack()
key_word.place(relx=0.69, rely=0.2, relwidth=0.29, relheight=0.06)
# LABEL
key_word_label = Label(frame, text="Palavra-Chave", background="#58ade0", fg='white')
key_word_label.config(font=('Arial', 13))
key_word_label.place(relx=0.69, rely=0.13)



# INPUT CATEGORIA
category_input = ttk.Combobox(
    frame, values=[
            "Comunicação",
            "Suporte",
            "Desenvolvimento",
            "Gestão",
            "Equipamento",
            "Especial"
        ]
)
category_input.current(0)
category_input.place(relx=0.02, rely=0.4, relwidth=0.45, relheight=0.06)
# LABEL
category_input_label = Label(frame, text="Categoria", background="#58ade0", fg='white')
category_input_label.config(font=('Arial', 13))
category_input_label.place(relx=0.02, rely=0.33)

# INPUT PESSOA RESPONSÁVEL
people_response = Entry(frame, borderwidth=0)
people_response.pack()
people_response.place(relx=0.53, rely=0.4, relwidth=0.45, relheight=0.06)
# LABEL
label_people_response = Label(frame, text="Email Pessoa Responsável", background="#58ade0", fg='white')
label_people_response.config(font=('Arial', 13))
label_people_response.place(relx=0.53, rely=0.33)

# INPUT DATA DE INÍCIO
start_date_input = Entry(frame, borderwidth=0)
start_date_input.pack()
start_date_input.place(relx=0.02, rely=0.60, relwidth=0.45, relheight=0.06)
# LABEL
label_start_date = Label(frame, text="Data de início", background="#58ade0", fg='white')
label_start_date.config(font=('Arial', 13))
label_start_date.place(relx=0.02, rely=0.53)

# INPUT PRAZO
duetime_input = Entry(frame, borderwidth=0)
duetime_input.pack()
duetime_input.place(relx=0.53, rely=0.60, relwidth=0.45, relheight=0.06)
# LABEL
label_duetime = Label(frame, text="Prazo", background="#58ade0", fg='white')
label_duetime.config(font=('Arial', 13))
label_duetime.place(relx=0.53, rely=0.53)

# BOTÃO CRIAR TASK
create_task = Button(frame, command=post_planner, text='Criar Task', font=('Arial', 13), background="white", borderwidth=1, relief='raised', cursor='hand2')
create_task.pack()
create_task.place(relx=0.4, rely=0.8, relwidth=0.2, relheight=0.08)

screen.mainloop()