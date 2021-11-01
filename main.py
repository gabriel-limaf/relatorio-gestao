import PySimpleGUI as sg  # package PySimpleGui
import pandas as pd  # package pandas
from openpyxl import load_workbook  # package openpyxl, lxml, html5lib
df1 = pd.DataFrame()


def escolher_arquivo():  # Janela 1
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Arquivos de entrada')],
              [sg.Input(), sg.FilesBrowse(key='-ENTRADA1-', file_types=(('Text Files', '*.xls'),
                                                                        ('Text Files', '*.xlsx')))],
              [sg.Text('Caminho do arquivo de Gestão do backlog')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('Atualizar planilha de gestão'), sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('Atualização de relatório', layout=layout, finalize=True,)


def erro():  # Janela 2
    sg.theme('DarkRed')
    layout = [[sg.Text('Favor verificar:\n'
                       '\n1) Extensão do arquivo de entrada está conforme coletado no Bitrix24\n'
                       '2) Extensão do arquivo de saída está em formato .xlsx\n'
                       '3) Os arquivos selecionados são diferentes\n'
                       '4) Selecionou pelo menos um arquivo de entrada e o arquivo de saida\n')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('ERRO', layout=layout, size=(500, 200), finalize=True)


def sucesso():  # Janela 3
    sg.theme('DarkGreen')
    layout = [[sg.Text('Importação realizada com sucesso !')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('SUCESSO', layout=layout, size=(300, 100), finalize=True)


def convert_csv():  # Janela 4
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Caminho da pasta onde quer salvar')],
              [sg.Input(), sg.FolderBrowse(key='-ENTRADA1-')],
              [sg.Text('Caminho do arquivo de Gestão do backlog')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('OK'), sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('Conversor de aba para csv', layout=layout, finalize=True)


def match_taskid():  # Janela 5
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Caminho do arquivo de entrada')],
              [sg.Input(), sg.FileBrowse(key='-ENTRADA1-', file_types=(('Text Files', '*.xls'),
                                                                       ('Text Files', '*.xlsx')))],
              [sg.Text('Caminho do arquivo de Gestão do backlog')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('OK'), sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('Match para encontrar a taskid', layout=layout, finalize=True)


def menu():  # Janela 6
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Bem-vindo(a) ao aplicativo de relatório de gestão de projetos !!\n'
                       'O que deseja fazer?:\n')],
              [sg.Button('Match Task ID'), sg.Button('Atualizar backlog'),
               sg.Button('Gerar arquivo de importação'), sg.Button('Cancelar')],
              [sg.Text('\nIndicium Tech - 2021', size=[75, 5], justification='center')]]
    return sg.Window('Menu', layout=layout, finalize=True, size=(500, 150))


janela1, janela2, janela3, janela4, janela5, janela6 = None, None, None, None, None, menu()
while True:
    window, event, values = sg.read_all_windows()
    # Operações no MENU
    if window == janela6 and event == sg.WINDOW_CLOSED:
        break
    if window == janela6 and event == 'Cancelar':
        break
    if window == janela6 and event == 'Match Task ID':
        janela6.close()
        janela5 = match_taskid()
    if window == janela6 and event == 'Atualizar backlog':
        janela6.close()
        janela1 = escolher_arquivo()
    if window == janela6 and event == 'Gerar arquivo de importação':
        janela6.close()
        janela4 = convert_csv()
    # Operações Janela 1
    if window == janela1 and event == sg.WINDOW_CLOSED:
        break
    if window == janela1 and event == 'Cancelar':
        break
    if window == janela1 and event == 'Voltar':
        janela1.close()
        janela6 = menu()
    if window == janela1 and event == 'Atualizar planilha de gestão' \
            and (values['-ENTRADA1-'] == '' or values['-SAIDA-'] == ''
                 or values['-ENTRADA1-'] == values['-SAIDA-']):
        janela1.close()
        janela2 = erro()
    if window == janela1 and event == 'Atualizar planilha de gestão' \
            and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''
                 and values['-SAIDA-'] != values['-ENTRADA1-']):
        path_saida = values['-SAIDA-']
        lista = values['-ENTRADA1-'].split(';')
        if path_saida not in lista:
            try:
                for i in lista:
                    path = i
                    table = pd.read_html(path)[0]
                    data = pd.DataFrame(table)
                    df1 = df1.append(data)
                # Inicio da comparação
                df1 = pd.DataFrame(df1,
                                   columns=['ID', 'Tarefa', 'Active', 'Deadline', 'Created by', 'Responsible person',
                                            'Status', 'Project', 'Created on', 'Closed on', 'Planned duration',
                                            'Time spent', 'Tags', 'Etapa', 'Natureza', 'Frente de Trabalho'])
                df2 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Datas'))
                # Atualizar dataframe da planilha de gestão
                for i, row in df1.iterrows():
                    df2.loc[df2['ID'] == row['ID'], 'Tarefa'] = row['Tarefa']
                    df2.loc[df2['ID'] == row['ID'], 'Active'] = row['Active']
                    df2.loc[df2['ID'] == row['ID'], 'Deadline'] = row['Deadline']
                    df2.loc[df2['ID'] == row['ID'], 'Created by'] = row['Created by']
                    df2.loc[df2['ID'] == row['ID'], 'Responsible person'] = row['Responsible person']
                    df2.loc[df2['ID'] == row['ID'], 'Status'] = row['Status']
                    df2.loc[df2['ID'] == row['ID'], 'Project'] = row['Project']
                    df2.loc[df2['ID'] == row['ID'], 'Created on'] = row['Created on']
                    df2.loc[df2['ID'] == row['ID'], 'Closed on'] = row['Closed on']
                    df2.loc[df2['ID'] == row['ID'], 'Planned duration'] = row['Planned duration']
                    df2.loc[df2['ID'] == row['ID'], 'Time spent'] = row['Time spent']
                    df2.loc[df2['ID'] == row['ID'], 'Tags'] = row['Tags']
                    df2.loc[df2['ID'] == row['ID'], 'Etapa'] = row['Etapa']
                    df2.loc[df2['ID'] == row['ID'], 'Natureza'] = row['Natureza']
                    df2.loc[df2['ID'] == row['ID'], 'Frente de Trabalho'] = row['Frente de Trabalho']
                # Criar linha com novos dados
                for i, row in df1.iterrows():
                    if row['ID'] not in list(df2['ID']):
                        df2.loc[len(df2)] = list(row)
                # Fim da atualização
                book = load_workbook(path_saida)
                writer = pd.ExcelWriter(path_saida, engine='openpyxl')
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                df2.to_excel(writer, "Datas", header=True, index=False)
                writer.save()
                janela1.close()
                janela3 = sucesso()
            except:
                janela1.close()
                janela2 = erro()
        else:
            janela1.close()
            janela2 = erro()
    # Operações Janela 2
    if window == janela2 and event == 'Voltar':
        janela2.close()
        janela6 = menu()
    if window == janela2 and event == 'Cancelar':
        break
    if window == janela2 and event == sg.WINDOW_CLOSED:
        break
    # Operações Janela 3
    if window == janela3 and event == 'Voltar':
        janela3.close()
        janela6 = menu()
    if window == janela3 and event == 'Cancelar':
        break
    if window == janela3 and event == sg.WINDOW_CLOSED:
        break
    # Operações Janela 4
    if window == janela4 and event == 'OK' and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''):
        try:
            path_saida = values['-SAIDA-']
            path_salvar = values['-ENTRADA1-']
            df2 = pd.read_excel(path_saida, sheet_name='Backlog - importar')
            df2.to_csv(path_salvar + '\Tasks.csv', columns=['Name', 'Description', 'Important task',
                                                            'Responsible person',
                                                            'Created by', 'Participants',
                                                            'Observers', 'Deadline', 'Start task on',
                                                            'Complete task by',
                                                            'Responsible person can change deadline',
                                                            'Skip weekends and holidays', 'Approve task when completed',
                                                            'Derive task dates from subtask dates',
                                                            'Auto complete task when subtasks have been completed',
                                                            'Project', 'Task has time constraints',
                                                            'Task completion time, seconds', 'Checklist',
                                                            'Tags'], index=None, header=True, encoding='utf-8')
            janela4.close()
            janela3 = sucesso()
        except:
            janela4.close()
            janela2 = erro()
    if window == janela4 and event == sg.WINDOW_CLOSED:
        break
    if window == janela4 and event == 'Cancelar':
        break
    if window == janela4 and event == 'Voltar':
        janela4.close()
        janela6 = menu()
    if window == janela4 and event == 'OK' and (values['-ENTRADA1-'] == '' or values['-SAIDA-'] == ''):
        janela4.close()
        janela3 = erro()
    # Operações Janela 5
    if window == janela5 and event == sg.WINDOW_CLOSED:
        break
    if window == janela5 and event == 'Cancelar':
        break
    if window == janela5 and event == 'Voltar':
        janela5.close()
        janela6 = menu()
    if window == janela5 and event == 'OK' and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''
                                                and values['-SAIDA-'] != values['-ENTRADA1-']):
        path_saida = values['-SAIDA-']
        lista = values['-ENTRADA1-'].split(';')
        if path_saida not in lista:
            for i in lista:
                path = i
                table = pd.read_html(path)[0]
                data = pd.DataFrame(table)
                df1 = df1.append(data)
            df2 = pd.DataFrame((pd.read_excel(path_saida, sheet_name='Backlog - importar')))  #  Gestão
            for i, row in df1.iterrows():
                a = str(row['Tarefa']) + str(row['Responsible person']) + str(row['Created by']) + \
                    str(row['Deadline']) + str(row['Project']) + str(row['Planned duration'])
                df1.loc[i, 'Teste'] = a
            for i, row in df2.iterrows():
                b = str(row['Name']) + str(row['Responsible person']) + str(row['Created by']) + \
                    str(row['Deadline']) + str(row['Project']) + str(row['Task completion time, seconds'])
                df2.loc[i, 'Teste'] = b
            for i, row in df1.iterrows():
                df2.loc[df2['Teste'] == row['Teste'], 'ID'] = row['ID']
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, sheet_name="Backlog - importar", header=True, index=False)
            writer.save()
            janela5.close()
            janela3 = sucesso()
    if window == janela5 and event == 'OK' and (values['-ENTRADA1-'] == '' or values['-SAIDA-'] == ''
                                                or values['-ENTRADA1-'] == values['-SAIDA-']):
        janela5.close()
        janela2 = erro()
