from xlsFlowX import checkModuleSheetVale
from openpyxl import load_workbook

if __name__ == '__main__':
    moduleFileList=[]

    for mod_file in moduleFileList:
        wb = load_workbook(mod_file)
        ws = wb.active
        st_module_list, bCheckPass = checkModuleSheetVale(ws)
        if bCheckPass:
            print(f'{mod_file} checked Pass.')
        else:
            print(f'{mod_file} checked Failed.')