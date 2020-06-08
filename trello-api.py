from trello import TrelloClient
import json
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt

primeiro_dia_sprint = int(input("Primeiro dia da sprint"))
mes_comeco_sprint = int(input("Mês começo sprint"))
data_inicial_sprint = datetime(year=2020, month=mes_comeco_sprint, day=primeiro_dia_sprint, hour=0, minute=0, second=0)
n_sprint = int(input("Numero da sprint"))

client = TrelloClient(
    api_key='9c22087ebfb0dc02c4856c4018431db0',
    api_secret='61f5031756b1df74e5f1329173bfeae8057ea8345694d121e54ed3d7da4d8eb6',
    token='169cb6d5602b7539816db5087d44b2fc97794577ccf19ded656333cf2aed975f',
    token_secret='169cb6d5602b7539816db5087d44b2fc97794577ccf19ded656333cf2aed975f'
)

all_boards = client.list_boards()
last_board = all_boards[2]


do = last_board.list_lists()[0].id
doing = last_board.list_lists()[1].id
done = last_board.list_lists()[2].id

lista_do = last_board.get_list(do)
lista_doing = last_board.get_list(doing)
lista_done = last_board.get_list(done)

def nomeMembro(id):
    global client
    return client.get_member(id).full_name


def getLabel(idLabel):
    global client
    global last_board
    return client.get_label(idLabel, last_board.id).name

def toDate(data):
    lista = data.split('-')
    datetime_object = datetime(year=int(lista[0]), month=int(lista[1]), day=int(lista[2][:3].replace('T', '')), hour=0, minute=0, second=0)
    return datetime_object





total_tarefas = 0

for card in lista_do.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1

for card in lista_doing.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1

for card in lista_done.list_cards():
    sprint = getLabel(card.idLabels[0])
    if sprint == ('Sprint ' + str(n_sprint)):
        total_tarefas = total_tarefas + 1


lista_dias_sprint = []
for single_date in (data_inicial_sprint + timedelta(n) for n in range(total_tarefas)):
    lista_dias_sprint.append(str(single_date.day)+'-' +str(single_date.month))

class TaskCompleta:
    def __init__(self, nome, qtd):
        self.nome = nome
        self.qtd = qtd

lista_completos = []


def addTaskCompleta(nome):
    global lista_completos
    i = 0
    existe = False
    if len(lista_completos) == 0:
            lista_completos.append(TaskCompleta(nome, 1))
    else:
        for t in lista_completos:
            if t.nome == nome:
                existe = True
                lista_completos[i].qtd = lista_completos[i].qtd + 1

        if existe == False:
            lista_completos.append(TaskCompleta(nome, 1))
        i = i + 1

for card in lista_done.list_cards():
    sprint = getLabel(card.idLabels[0])
    for data in lista_dias_sprint:
        if sprint == ('Sprint ' + str(n_sprint)):
            d = toDate(card.due)
            d2 = str(d.day) + '-' + str(d.month)
            if (d2) == data:
                addTaskCompleta(d2)
            


lista_conc = []
ultimo = total_tarefas
lista_restantes = []

for x in range(0, total_tarefas):
    lista_restantes.append(total_tarefas - x)

print(str(len(lista_restantes)))
print(str(len(lista_conc)))
print(str(len(lista_dias_sprint)))


for data in lista_dias_sprint:
    existe = False
    for x in lista_completos:
        if x.nome == data:
            ultimo = ultimo - x.qtd
            lista_conc.append(ultimo)
            existe = True
            
    if existe == False:
        lista_conc.append(ultimo)
            
print(lista_conc)

df = pd.DataFrame({
   'Restantes': lista_restantes,
   'Finalizadas': lista_conc
   }, index=lista_dias_sprint)


lines = df.plot.line()
plt.show()