from trello import TrelloClient
import openpyxl
from datetime import datetime, timedelta
import string
import xlrd
from openpyxl.formula.translate import Translator
from termcolor import colored

primeiro_dia_sprint = int(input("Primeiro dia da sprint"))
mes_comeco_sprint = int(input("Mês começo sprint"))
data_inicial_sprint = datetime(year=2020, month=mes_comeco_sprint, day=primeiro_dia_sprint, hour=0, minute=0, second=0)
n_sprint = int(input("Numero da sprint"))
print(colored("Buscando dados do servidor...", 'cyan'))
xfile = None
client = TrelloClient(
    api_key='9c22087ebfb0dc02c4856c4018431db0',
    api_secret='61f5031756b1df74e5f1329173bfeae8057ea8345694d121e54ed3d7da4d8eb6',
    token='169cb6d5602b7539816db5087d44b2fc97794577ccf19ded656333cf2aed975f',
    token_secret='169cb6d5602b7539816db5087d44b2fc97794577ccf19ded656333cf2aed975f'
)

all_boards = client.list_boards()
antena = all_boards[2]


do = antena.list_lists()[0].id
doing = antena.list_lists()[1].id
done = antena.list_lists()[2].id

lista_do = antena.get_list(do)
lista_doing = antena.get_list(doing)
lista_done = antena.get_list(done)
print(colored("Separando lista de tarefas...", 'blue'))
total_tarefas = 0

def nomeMembro(id):
    global client
    return client.get_member(id).full_name

def getLabel(idLabel):
    global client
    global antena
    return client.get_label(idLabel, antena.id).name

def toDate(data):
    lista = data.split('-')
    datetime_object = datetime(year=int(lista[0]), month=int(lista[1]), day=int(lista[2][:3].replace('T', '')), hour=0, minute=0, second=0)
    return datetime_object

print(colored("Ordenando lista de 'DO'...", 'green'))
for card in lista_do.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1

print(colored("Ordenando lista de 'DOING'... ", 'cyan'))
for card in lista_doing.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1

print(colored("Ordenando lista de 'DONE'...", 'blue'))
for card in lista_done.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1

loc = 'C:\\Users\\Aleixo\\Documents\\fatec\\lpbd\\antena.xlsx'

class Task:
    def __init__(self, data, linha):
        self.data = data
        self.linha = linha

def preencherDadosExcel():
    global loc
    global xfile
    global data_inicial_sprint
    global lista_do
    global lista_doing
    global lista_done
    global total_tarefas
    print(colored("Abrindo arquivo xlsx...", 'green'))
    if xfile == None:
        xfile = openpyxl.load_workbook(loc)
    sheet = xfile['Plan1']
    sheet['A1'] = "Data de início:"
    sheet['B1'] = fmtData(data_inicial_sprint)

    sheet['D1'] = "Data atual:"
    sheet['E1'] = fmtData(datetime.now())

    sheet['A2'] = "Duração (d):"
    sheet['B2'] = abs((datetime.now() - data_inicial_sprint).days)

    sheet['B4'] = "Total de horas"
    sheet['A4'] = "Atividades"
    print(colored("Escrevendo labels...", 'cyan'))
    letras = list(string.ascii_lowercase)
    letra = 2
    total_horas = 2 * total_tarefas
    remover_por_task = 2
    data_inicial_sprint_c = data_inicial_sprint
    for x in range(0, total_tarefas):
        sheet[letras[letra] + '4'] = fmtData(data_inicial_sprint_c)
        letra = letra + 1
        data_inicial_sprint_c = data_inicial_sprint_c + timedelta(days=1)
    print(colored("Escrevendo datas...", 'blue'))
    letra = 2
    linha = 5
    data_cache = datetime.now()
    listaSetarHora = []
    print(colored("Escrevendo tarefas 1/3...", 'green'))
    for card in lista_done.list_cards():
        sheet['A' + str(linha)] = card.name
        if card.is_due_complete:
            dataLinha = fmtData(toDate(card.due + ""))
            listaSetarHora.append(Task(dataLinha, linha))
        linha = linha + 1
    print(colored("Escrevendo tarefas 2/3...", 'cyan'))
    for card in lista_doing.list_cards():
        sheet['A' + str(linha)] = card.name
        linha = linha + 1
    print(colored("Escrevendo tarefas 3/3...", 'blue'))
    for card in lista_do.list_cards():
        sheet['A' + str(linha)] = card.name
        linha = linha + 1

    print(colored("Escrevendo formulas...", 'green'))
    sheet['A' + str(linha)] = "Restante"
    linha = linha + 1
    sheet['A' + str(linha)] = "Estimado"
    sheet['B' + str(linha)] = total_horas
    sheet['B' + str(linha - 1)] = "=SUM(" + "B5:" + "B" + str(linha - 2) + ")"
    c = 2
    anterior = total_horas
    print(colored("Escrevendo e aplicando valores & funções...", 'cyan'))
    for x in range(2, total_tarefas + 1):
        total_horas = total_horas - remover_por_task
        sheet[letras[c] + str(linha)] = total_horas
        soma = 0
        for y in range(5, linha - 1):
            v = sheet[letras[c] + str(y)].value
            soma = sumData(soma, v)        
        anterior = anterior - soma
        if soma > 0 or x == 2:
            sheet[letras[c] + str(linha - 1)] = anterior
        c = c + 1
        
    print(colored("Escrevendo horas...", 'blue'))
    for x in range(5, linha - 1):
        sheet['B' + str(x)] = remover_por_task

    xfile.save(loc)
    if xfile == None:
        xfile = openpyxl.load_workbook(loc)
    sheet = xfile['Plan1']

    wb = xlrd.open_workbook(loc) 
    folha = wb.sheet_by_index(0) 
    c = 1
    while c < total_tarefas:
        for d in listaSetarHora:
            if d.data == folha.cell_value(3, c):
                sheet[letras[c] + str(d.linha)] = remover_por_task
        c = c + 1
    print(colored("Finalizado", 'magenta'))
    xfile.save(loc)

def sumData(soma, valor):
    if valor != None and valor != "" and valor != 'Cell' and valor != 'None':
        return soma + int(valor)
    return soma

def fmtData(data):
    return str(data.day) + "/" + str(data.month) + "/" + str(data.year)

preencherDadosExcel()