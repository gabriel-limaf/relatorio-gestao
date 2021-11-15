import pyautogui
import pyperclip
from time import sleep
import webbrowser
import PySimpleGUI as sg  # package PySimpleGui
import pandas as pd  # package pandas
from openpyxl import load_workbook
df1 = pd.DataFrame()


def menu():  # Janela 1
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Bem-vindo(a) ao RPA Bitrix e Conferência de planilha !!\n'
                       'O que deseja fazer?:\n')],
              [sg.Button('RPA'), sg.Button('Conferir planilha'), sg.Button('Cancelar')],
              [sg.Text('\nIndicium Tech - 2021', size=[75, 5], justification='center')]]
    return sg.Window('Menu', layout=layout, finalize=True, size=(500, 150))


def erro():  # Janela 2
    sg.theme('DarkRed')
    layout = [[sg.Text('Favor verificar:\n'
                       '\n1) Planilha deve estar fechada\n'
                       '2) Extensão do arquivo de saída estar em formato .xlsx\n'
                       '3) Dados na planilha fora do padrão\n'
                       '4) Não há dados a ser processado\n')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('ERRO', layout=layout, size=(500, 200), finalize=True)


def sucesso():  # Janela 3
    sg.theme('DarkGreen')
    layout = [[sg.Text('Processo realizado com sucesso !')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('SUCESSO', layout=layout, size=(300, 100), finalize=True)


def escolher_arquivo():  # Janela 4
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Caminho do Arquivo extraído do Bitrix')],
              [sg.Input(), sg.FileBrowse(key='-ENTRADA1-', file_types=(('Text Files', '*.xls'),
                                                                       ('Text Files', '*.xlsx')))],
              [sg.Text('Caminho do Arquivo do Google Sheets')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('Conferir planilha'), sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('Conferencia de backlog google sheets x Bitrix', layout=layout, finalize=True,)


def rpa():  # Janela 5
    sg.theme('Dark Blue 3')
    layout = [
              [sg.Text('Caminho do Arquivo do Google Sheets')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('OK'), sg.Button('Voltar'), sg.Button('Cancelar')],
             [sg.Text('\nIndicium Tech - 2021', size=[75, 5], justification='center')]]
    return sg.Window('RPA - Lançamento no Bitrix', layout=layout, finalize=True, size=(450, 150))


janela1, janela2, janela3, janela4, janela5 = menu(), None, None, None, None

while True:
    window, event, values = sg.read_all_windows()
    # Operações no MENU
    if window == janela1 and event == sg.WINDOW_CLOSED:
        break
    if window == janela1 and event == 'Cancelar':
        break
    if window == janela1 and event == 'RPA':
        janela1.close()
        janela5 = rpa()
    if window == janela1 and event == 'Conferir planilha':
        janela1.close()
        janela4 = escolher_arquivo()
    if window == janela2 and event == 'Voltar':
        janela2.close()
        janela1 = menu()
    if window == janela2 and event == 'Cancelar':
        break
    if window == janela2 and event == sg.WINDOW_CLOSED:
        break
    if window == janela3 and event == 'Voltar':
        janela3.close()
        janela1 = menu()
    if window == janela3 and event == 'Cancelar':
        break
    if window == janela3 and event == sg.WINDOW_CLOSED:
        break
    if window == janela4 and event == sg.WINDOW_CLOSED:
        break
    if window == janela4 and event == 'Cancelar':
        break
    if window == janela4 and event == 'Voltar':
        janela4.close()
        janela1 = menu()
    if window == janela4 and event == 'Conferir planilha' \
            and (values['-ENTRADA1-'] == '' or values['-SAIDA-'] == ''
                 or values['-ENTRADA1-'] == values['-SAIDA-']):
        janela4.close()
        janela2 = erro()
    if window == janela4 and event == 'Conferir planilha' \
            and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''
                 and values['-SAIDA-'] != values['-ENTRADA1-']):
        path_saida = values['-SAIDA-']
        lista = values['-ENTRADA1-'].split(';')
        if path_saida not in lista:
            for i in lista:
                path = i
                table = pd.read_html(path)[0]
                data = pd.DataFrame(table)
                df1 = df1.append(data)
# Inicio da comparação
            df1 = pd.DataFrame(df1, columns=['ID', 'Tarefa', 'Active', 'Deadline', 'Created by', 'Responsible person',
                                             'Status', 'Project', 'Created on', 'Closed on', 'Planned duration',
                                             'Time spent', 'Tags', 'Frente de Trabalho', 'Etapa', 'Natureza',
                                             'Sprint', 'Task', 'Description'])
            df2 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Tarefas'))
# Quebrar nome da task e Sprint
            for i, row in df1.iterrows():
                a = row['Tarefa'].find(']')
                b = row['Tarefa'][(a + 2):300]
                df1.loc[i, 'Task'] = b
                c = row['Tarefa'].find('Sprint')
                if c == 1:
                    d = row['Tarefa'][(c + 7):10].replace(' ', '')
                    df1.loc[i, 'Sprint'] = d
                else:
                    d = row['Tarefa'][(c + 7):c + 9].replace(']', '').replace(' ', '')
                    df1.loc[i, 'Sprint'] = d
# Comparar nome da task
            for i, row in df1.iterrows():
                task_sheets = df2.loc[df2['ID'] == row['ID'], 'Task']
                task_sheets = ''.join(task_sheets.values)
                task_bitrix = row['Task']
                if task_sheets == task_bitrix:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Task'] = 'Conforme Bitrix'
                else:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Task'] = 'Diferente Bitrix'
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, "Tarefas", header=True, index=False)
            writer.save()
# Comparar Frente de Trabalho
            for i, row in df1.iterrows():
                frente_sheets = df2.loc[df2['ID'] == row['ID'], 'Frente de Trabalho']
                frente_sheets = ''.join(frente_sheets)
                frente_bitrix = row['Frente de Trabalho']
                # print('Sheets: ' + str(frente_sheets) + ' Bitrix: ' + str(frente_bitrix))
                if frente_sheets == frente_bitrix:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Frente de Trabalho'] = 'Conforme Bitrix'
                else:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Frente de Trabalho'] = 'Diferente Bitrix'
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, "Tarefas", header=True, index=False)
            writer.save()
# Comparar Etapa
            for i, row in df1.iterrows():
                etapa_sheets = df2.loc[df2['ID'] == row['ID'], 'Etapa']
                etapa_sheets = ''.join(etapa_sheets)
                etapa_bitrix = row['Etapa']
                if etapa_sheets == etapa_bitrix:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Etapa'] = 'Conforme Bitrix'
                else:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Etapa'] = 'Diferente Bitrix'
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, "Tarefas", header=True, index=False)
            writer.save()
# Comparar Natureza
            for i, row in df1.iterrows():
                natureza_sheets = df2.loc[df2['ID'] == row['ID'], 'Natureza']
                natureza_sheets = ''.join(natureza_sheets)
                natureza_bitrix = row['Natureza']
                if natureza_sheets == natureza_bitrix:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Natureza'] = 'Conforme Bitrix'
                else:
                    df2.loc[df2['ID'] == row['ID'], 'Teste Natureza'] = 'Diferente Bitrix'
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, "Tarefas", header=True, index=False)
            writer.save()
            janela1.close()
            janela3 = sucesso()
        else:
            janela1.close()
            janela2 = erro()
    if window == janela5 and event == 'OK' and values['-SAIDA-'] != '':
        try:
            path_saida = values['-SAIDA-']
            df1 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Tarefas'))
            webbrowser.open('https://google.com.br')
            for i, row in df1.iterrows():
                id = (str(row['ID']))
                frente = (str(row['Frente de Trabalho']))
                etapa = (str(row['Etapa']))
                natureza = (str(row['Natureza']))
                rpa = (str(row['RPA']))
                if id != 'nan' and frente != 'nan' and etapa != 'nan' and natureza != 'nan' and rpa != 'Processado':
                    webbrowser.open('https://indicium.bitrix24.com/workgroups/group/' + str(row['Cod projeto']) +
                                    '/tasks/task/edit/' + str(row['ID']) + '/')
                    sleep(15)
                    pyautogui.click(700, 300)
                    pyautogui.hotkey('end'), sleep(1)
                    try:
                        pyautogui.click(265, 546), sleep(2)
                        pyautogui.hotkey('end'), sleep(1)
                        pyautogui.hotkey('ctrl', 'f'), sleep(1)
                        pyperclip.copy('Frente de Trabalho')
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.hotkey('tab'), sleep(1)
                        pyperclip.copy(str(row['Frente de Trabalho']))
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.hotkey('ctrl', 'f'), sleep(1)
                        pyperclip.copy('Etapa')
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.hotkey('esc'), sleep(1)
                        pyautogui.hotkey('tab'), sleep(1)
                        pyperclip.copy(str(row['Etapa']))
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.hotkey('ctrl', 'f'), sleep(1)
                        pyperclip.copy('Natureza')
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.hotkey('esc'), sleep(1)
                        pyautogui.hotkey('tab'), sleep(1)
                        pyperclip.copy(str(row['Natureza']))
                        pyautogui.hotkey('ctrl', 'v'), sleep(1)
                        pyautogui.click(300, 604), sleep(1)
                        sleep(15)
                        pyautogui.hotkey('ctrl', 'w')
                        df1.loc[i, 'RPA'] = 'Processado'
                        book = load_workbook(path_saida)
                        writer = pd.ExcelWriter(path_saida, engine='openpyxl')
                        writer.book = book
                        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                        df1.to_excel(writer, "Tarefas", header=True, index=False)
                        writer.save()
                    except:
                        sleep(10)
                        pyautogui.hotkey('ctrl', 'w')
                        df1.loc[i, 'RPA'] = 'ERRO'
                        book = load_workbook(path_saida)
                        writer = pd.ExcelWriter(path_saida, engine='openpyxl')
                        writer.book = book
                        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                        df1.to_excel(writer, "Tarefas", header=True, index=False)
                        writer.save()
                else:
                    sleep(10)
                    pyautogui.hotkey('ctrl', 'w')
                    df1.loc[i, 'RPA'] = 'Faltam dados de Frente/Etapa/Natureza'
                    book = load_workbook(path_saida)
                    writer = pd.ExcelWriter(path_saida, engine='openpyxl')
                    writer.book = book
                    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                    df1.to_excel(writer, "Tarefas", header=True, index=False)
                    writer.save()
            janela1.close()
            janela3 = sucesso()
        except:
            janela1.close()
            janela2 = erro()
    if window == janela5 and event == 'OK' and values['-SAIDA-'] == '':
        janela1.close()
        janela2 = erro()
    if window == janela5 and event == 'Voltar':
        janela5.close()
        janela1 = menu()
    if window == janela5 and event == 'Cancelar':
        break
    if window == janela5 and event == sg.WINDOW_CLOSED:
        break
