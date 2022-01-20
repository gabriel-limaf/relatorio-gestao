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


def convert_csv():  # Janela 4
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Caminho da pasta onde quer salvar')],
              [sg.Input(), sg.FolderBrowse(key='-ENTRADA1-')],
              [sg.Text('Caminho do arquivo de Gestão do backlog')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Text('Em qual idioma esta sua planilha? Digite: PT ou EN')],
              [sg.InputText(key='-IDIOMA-')],
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

#  Processo que atualiza a aba DATAS conforme dados extraidos dos projetos no Bitrix
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
                                            'Time spent', 'Tags', 'Frente de Trabalho', 'Etapa', 'Natureza',
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
                    df2.loc[df2['ID'] == row['ID'], 'Frente de Trabalho'] = row['Frente de Trabalho']
                    df2.loc[df2['ID'] == row['ID'], 'Etapa'] = row['Etapa']
                    df2.loc[df2['ID'] == row['ID'], 'Natureza'] = row['Natureza']
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
                        #df2.loc[i, 'Sprint'] = d
                        df2.loc[i, 'Sprint'] = str(d).zfill(2)  # teste
                    else:
                        d = row['Tarefa'][(c+7):c+9].replace(']', '').replace(' ', '')
                        #df2.loc[i, 'Sprint'] = d
                        df2.loc[i, 'Sprint'] = str(d).zfill(2)  # teste
                    # print(c)

# Atualizar aba Sheets
                df3 = pd.DataFrame(pd.read_excel(path_saida, sheet_name='Sheets'))
               #print(df3)
                linhas = 0
                for i, row in df3.iterrows():
                    linhas = linhas + 1
               #print(linhas)
                for linhas, row in df3.iterrows():
                    df3 = df3.drop(linhas)
                   #print(df3)

                for i, row in df2.iterrows():
                    df3['Sprint'] = df2['Sprint']
                    df3['ID'] = df2['ID']
                    df3['Frente de Trabalho'] = df2['Frente de Trabalho']
                    df3['Etapa'] = df2['Etapa']
                    df3['Natureza'] = df2['Natureza']
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

#  Processo que gera o arquivo .CSV a partir da aba Backlog - importar para importação no Bitrix

    if window == janela4 and event == 'OK' and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != '') and \
            values['-IDIOMA-'] != '':
        try:
            path_saida = values['-SAIDA-']
            path_salvar = values['-ENTRADA1-']
            idioma = str(values['-IDIOMA-'])
            df2 = pd.read_excel(path_saida, sheet_name='Backlog - importar')
            if idioma == 'EN':
                df2.to_csv(path_salvar + '\Tasks.csv', columns=['Name', 'Description', 'Important task',
                                                                'Responsible person',
                                                                'Created by', 'Participants',
                                                                'Observers', 'Deadline',
                                                                'Responsible person can change deadline',
                                                                'Skip weekends and holidays', 'Approve task when completed',
                                                                'Derive task dates from subtask dates',
                                                                'Auto complete task when subtasks have been completed',
                                                                'Project', 'Task has time constraints',
                                                                'Task completion time, seconds', 'Checklist',
                                                                'Tags'], index=None, header=True, encoding='utf-8-sig',
                                                                 date_format='%d/%m/%Y %H:%M:%S')
            if idioma == 'PT':
                df2.to_csv(path_salvar + '\Tasks.csv', columns=['Nome', 'Descrição', 'Tarefa importante',
                                                                'Pessoa responsável',
                                                                'Criada por', 'Participantes',
                                                                'Observadores', 'Prazo final',
                                                                'A pessoa responsável pode alterar o prazo',
                                                                'Pular fins de semana e feriados',
                                                                'Verificar a tarefa após a conclusão',
                                                                'Derivar datas da tarefa das datas da subtarefa',
                                                                'Tarefa concluída automaticamente quando as subtarefas foram concluídas',
                                                                'Projeto', 'Tarefa tem restrições de tempo',
                                                                'Tempo de conclusão da tarefa, segundos', 'Lista de verificação',
                                                                'Marcadores'], index=None, header=True, encoding='utf-8-sig',
                                                                 date_format='%d/%m/%Y %H:%M:%S')
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

#  Processo que faz o match para encontrar a task id

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
