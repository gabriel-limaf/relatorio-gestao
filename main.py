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
                       '\n1) Extensão do arquivo de entrada está conforme coletado no Bitrix24;\n'
                       '2) Extensão do arquivo de saída está em formato .xlsx;\n'
                       '3) Os arquivos selecionados são diferentes;\n'
                       '4) Selecionou pelo menos um arquivo de entrada e o arquivo de saida;\n'
                       '5) Planilhas utilizadas estão fechadas.')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('ERRO', layout=layout, size=(500, 200), finalize=True)


def sucesso():  # Janela 3
    sg.theme('DarkGreen')
    layout = [[sg.Text('Importação realizada com sucesso !')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('SUCESSO', layout=layout, size=(300, 100), finalize=True)


def menu():  # Janela 6
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Bem-vindo(a) ao aplicativo de relatório de gestão de projetos !!\n'
                       'O que deseja fazer?:\n')],
              [sg.Button('Atualizar backlog'), sg.Button('Cancelar')],
              [sg.Text('\nIndicium Tech - 2021', size=[75, 5], justification='center')]]
    return sg.Window('Menu', layout=layout, finalize=True, size=(500, 150))


janela1, janela2, janela3, janela6 = None, None, None, menu()
while True:
    window, event, values = sg.read_all_windows()
# Operações no MENU
    if window == janela6 and event == sg.WINDOW_CLOSED:
        break
    if window == janela6 and event == 'Cancelar':
        break
    if window == janela6 and event == 'Atualizar backlog':
        janela6.close()
        janela1 = escolher_arquivo()
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
#  Processo que atualiza a aba DATAS conforme dados extraidos dos projetos no Bitrix
    if window == janela1 and event == 'Atualizar planilha de gestão' \
            and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''
                 and values['-SAIDA-'] != values['-ENTRADA1-']):
        path_saida = values['-SAIDA-']
        lista = values['-ENTRADA1-'].split(';')
        if path_saida not in lista:
            for i in lista:
                path = i
                table = pd.read_html(path)[0]
                data = pd.DataFrame(table)
                # df1 = df1.append(data)
                df1 = pd.concat([df1, data])
# Inicio da comparação
            df1 = pd.DataFrame(df1,
                               columns=['ID', 'Tarefa', 'Active', 'Deadline', 'Created by', 'Responsible person',
                                        'Status', 'Project', 'Created on', 'Closed on', 'Planned duration',
                                        'Time spent', 'Tags', 'Campo CTI', 'Entrega', 'Produto',
                                        'Sprint', 'Task', 'Description'])
            df2 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Datas'))
# Atualizar dataframe da planilha de gestão (aba Datas)
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
                df2.loc[df2['ID'] == row['ID'], 'Campo CTI'] = row['Campo CTI']
                df2.loc[df2['ID'] == row['ID'], 'Entrega'] = row['Entrega']
                df2.loc[df2['ID'] == row['ID'], 'Produto'] = row['Produto']
# Criar linha com novos dados na aba Datas
            for i, row in df1.iterrows():
                if row['ID'] not in list(df2['ID']):
                    df2.loc[len(df2)] = list(row)
# Fim da atualização
            for i, row in df2.iterrows():
                a = row['Tarefa'].find(']')
                b = row['Tarefa'][(a+2):300]
                df2.loc[i, 'Task'] = b
                c = row['Tarefa'].find('Sprint')
                if c == 1:
                    d = row['Tarefa'][(c+7):10].replace(' ', '')
                    df2.loc[i, 'Sprint'] = str(d).zfill(2)  # teste
                else:
                    d = row['Tarefa'][(c+7):c+9].replace(']', '').replace(' ', '')
                    df2.loc[i, 'Sprint'] = str(d).zfill(2)  # teste
# Atualizar aba Sheets
            df3 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Sheets'))
            linhas = 0
            for i, row in df3.iterrows():
                linhas = linhas + 1
            for linhas, row in df3.iterrows():
                df3 = df3.drop(linhas)
            for i, row in df2.iterrows():
                df3['Sprint'] = df2['Sprint']
                df3['ID'] = df2['ID']
                df3['Campo CTI'] = df2['Campo CTI']
                df3['Entrega'] = df2['Entrega']
                df3['Produto'] = df2['Produto']
                df3['Task'] = df2['Task']
                df3['Responsavel'] = df2['Responsible person']
                df3['Status'] = df2['Status']
                df3['Inicio'] = df2['Created on']
                df3['Conclusao'] = df2['Closed on']
                df3['Horas Estimadas'] = df2['Planned duration']
                df3['Horas Executadas'] = df2['Time spent']
                df3['Projeto'] = df2['Project']
            for i, row in df3.iterrows():
                hrs_exec = str(row['Horas Executadas'])
                hrs_est = str(row['Horas Estimadas'])
                if hrs_exec != 'nan':
                    hh, mm, ss = hrs_exec.split(':')
                    dec_exec = (int(hh) * 3600 + int(mm) * 60 + int(ss)) / 3600
                    df3.loc[i, 'Tempo Efetivo'] = dec_exec
                if hrs_exec == 'nan':
                    df3.loc[i, 'Tempo Efetivo'] = ''
                if hrs_est != 'nan':
                    hh, mm, ss = hrs_est.split(':')
                    dec_est = (int(hh) * 3600 + int(mm) * 60 + int(ss)) / 3600
                    df3.loc[i, 'Tempo Estimado'] = dec_est
                if hrs_est == 'nan':
                    df3.loc[i, 'Tempo Estimado'] = ''
            book = load_workbook(path_saida)
            writer = pd.ExcelWriter(path_saida, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2.to_excel(writer, "Datas", header=True, index=False)
            df3.to_excel(writer, "Sheets", header=True, index=False)
            writer.save()
            janela1.close()
            janela3 = sucesso()
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
