import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

contatos = openpyxl.load_workbook('contatos.xlsx')
pagina_contatos = contatos['Contatos']

for linha in pagina_contatos.iter_rows(min_row=2):
    telefone_completo = str(linha[2].value)
    
    # Verifica se o número começa com '5531' (DDD 31)
    if telefone_completo.startswith('5531'):
        print(f'Ignorando número com DDD 31: {telefone_completo}')
        continue
    
    mensagem = '*🌲Caro cliente, entre os dias 21/12/2023 a 01/01/2024 estaremos de recesso. Por gentileza entrar em contato até o dia 20/12/2023 para que possamos atender as demandas necessárias de fechamento de ponto, crachás e relógios. Pedimos que antecipem suas dúvidas e fechamento evitando que fiquem sem o atendimento necessário.*🌲'
    
    try:
        link_mensagem = f'https://web.whatsapp.com/send?phone={telefone_completo}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem)
        sleep(10)    
        pyautogui.click(2763, 682)
        sleep(8)
        pyautogui.hotkey('ctrl', 'w')
        sleep(8)
    except:
        print(f'Não foi possível enviar mensagem para {telefone_completo}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{telefone_completo}\n')
