import PySimpleGUI as sg
import sys
from xlsFlowX import dealwith_excel

if __name__ == '__main__':
    if len(sys.argv) == 1:
        event, values = sg.Window('CIP Excel to DV',
                                  [[sg.Text('请选择模块excel文件.')],
                                   [sg.In(), sg.FileBrowse(
                                       file_types=(("excel files", "*.xlsx"),))],
                                      [sg.Open(), sg.Cancel()]]).read(close=True)
        fname = values[0]
    else:
        fname = sys.argv[1]

if not fname:
    sg.popup("Cancel", "No filename supplied")
    raise SystemExit("Cancelling: no filename supplied")

else:
    # sg.popup('The filename you chose was', fname)
    if fname.endswith('.xlsx'):
        dealwith_excel(fname)
