import PySimpleGUI as sg  # package PySimpleGui
import pandas as pd  # package pandas
from openpyxl import load_workbook  # package openpyxl, lxml, html5lib
x = pd.DataFrame()
each_rows = 1


def escolher_arquivo():
    sg.theme('Dark Blue 3')  # please make your creations colorful
    layout = [[sg.Text('Arquivos de entrada')],
              [sg.Input(), sg.FilesBrowse(key='-ENTRADA1-', file_types=(('Text Files', '*.xls'),
                                                                        ('Text Files', '*.xlsx')))],
              [sg.Text('Caminho do arquivo de Gestão do backlog')],
              [sg.Input(), sg.FileBrowse(key='-SAIDA-', file_types=(('Text Files', '*.xls'),
                                                                    ('Text Files', '*.xlsx')))],
              [sg.Button('OK')],
              [sg.Button('Cancelar')]
              ]
    return sg.Window('Atualização de relatório', layout=layout, finalize=True)


def erro():
    sg.theme('Dark Blue 3')  # please make your creations colorful
    layout = [[sg.Text('Favor verificar:\n'
                       '\n1) Extensão do arquivo de entrada está conforme coletado no Bitrix24\n'
                       '2) Extensão do arquivo de saída está em formato .xlsx\n'
                       '3) Os arquivos selecionados são diferentes\n'
                       '4) Selecionou pelo menos um arquivo de entrada e o arquivo de saida\n')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('ATENÇÃO', layout=layout, size=(500, 200), finalize=True)


def sucesso():
    sg.theme('Dark Blue 3')  # please make your creations colorful
    layout = [[sg.Text('Importação realizada com sucesso !')],
              [sg.Button('Voltar'), sg.Button('Cancelar')]]
    return sg.Window('ATENÇÃO', layout=layout, size=(300, 100), finalize=True)


janela1, janela2, janela3 = escolher_arquivo(), None, None
while True:
    window, event, values = sg.read_all_windows()
    if window == janela1 and event == sg.WINDOW_CLOSED:
        break
    if window == janela1 and event == 'Cancelar':
        break
    if window == janela1 and event == 'OK' and (values['-ENTRADA1-'] == '' or values['-SAIDA-'] == ''
                                                or values['-ENTRADA1-'] == values['-SAIDA-']):
        janela1.close()
        janela2 = erro()
    if window == janela2 and event == 'Voltar':
        janela2.close()
        janela1 = escolher_arquivo()
    if window == janela2 and event == 'Cancelar':
        break
    if window == janela3 and event == 'Voltar':
        janela3.close()
        janela1.close()
        janela1 = escolher_arquivo()
    if window == janela3 and event == 'Cancelar':
        break
    if window == janela1 and event == 'OK' and (values['-ENTRADA1-'] != '' and values['-SAIDA-'] != ''
                                                and values['-SAIDA-'] != values['-ENTRADA1-']):
        path_saida = values['-SAIDA-']
        lista = values['-ENTRADA1-'].split(';')
        if path_saida not in lista:
            try:
                for i in lista:
                    path = i
                    table = pd.read_html(path)[0]
                    data = pd.DataFrame(table)
                    x = x.append(data)
                plan = pd.read_excel(path_saida, sheet_name='Datas')
                rows = plan['ID'].count()
                rows = int(rows)
                wb = load_workbook(path_saida)
                ws = wb['Datas']
                print('antes do while\n')
                while each_rows <= rows:
                    ws.delete_rows(2)
                    each_rows = each_rows + 1
                wb.save(path_saida)
                print('antes do for\n')
                df_2 = pd.DataFrame(pd.read_excel(path_saida))
                book = load_workbook(path_saida)
                writer = pd.ExcelWriter(path_saida, engine='openpyxl')
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                x.to_excel(writer, "Datas", header=True, index=False)
                writer.save()
                janela1.close()
                janela3 = sucesso()
            except:
                janela1.close()
                janela2 = erro()
        else:
            janela1.close()
            janela2 = erro()
