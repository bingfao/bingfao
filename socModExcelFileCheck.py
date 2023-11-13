from xlsFlowX import checkModuleSheetVale
from openpyxl import load_workbook
import os

if __name__ == '__main__':
    moduleFileList=[]
    #遍历当前文件夹下的xlsx文件，然后处理
    with os.scandir('.') as it:
        for entry in it:
            entry_name=entry.name.lower()
            if not entry_name.startswith('.') and entry.is_file() and entry_name.endswith('.xlsx'):
                # print(entry.name)
                moduleFileList.append(entry_name)

    for mod_file in moduleFileList:
        wb = load_workbook(mod_file)
        ws = wb.active
        st_module_list, bCheckPass = checkModuleSheetVale(ws)
        if bCheckPass:
            print(f'{mod_file} checked Pass.')
        else:
            print(f'{mod_file} checked Failed.')