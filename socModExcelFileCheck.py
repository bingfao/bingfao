from xlsFlowX import checkModuleSheetVale
from openpyxl import load_workbook
import os

if __name__ == '__main__':
    moduleFileList = []
    # 遍历当前文件夹下的xlsx文件，然后处理
    with os.scandir('.') as it:
        for entry in it:
            entry_name = entry.name.lower()
            if not entry_name.startswith('.') and entry.is_file() and entry_name.endswith('.xlsx'):
                # print(entry.name)\
                ni = len(entry_name)-5
                ri = entry_name.rfind('_')
                if ri != -1:
                    # print(mod_name)
                    if ri < ni:
                        mod_name = entry_name[0:ri]
                        date_val = entry_name[ri+1:ni]
                        if date_val.isnumeric():
                            print('name: {0} , date: {1}'.format(
                                mod_name, date_val))
                    moduleFileList.append(entry_name)

    for mod_file in moduleFileList:
        wb = load_workbook(mod_file)
        ws = wb.active
        st_module_list, bCheckPass = checkModuleSheetVale(ws)
        if bCheckPass:
            print(f'{mod_file} checked Pass.')
        else:
            print(f'{mod_file} checked Failed.')
