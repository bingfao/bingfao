from xlsFlowX import checkModuleSheetVale
from openpyxl import load_workbook
import os

if __name__ == '__main__':
    moduleFileList = []
    # 遍历当前文件夹下的xlsx文件，然后处理
    with os.scandir('.') as it:
        mod_dict = {}
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
                            if mod_name in mod_dict:
                                if mod_dict[mod_name] < date_val:
                                    mod_dict[mod_name] = date_val
                            else:
                                mod_dict[mod_name] = date_val

                    # moduleFileList.append(entry_name)
        for mod_name in mod_dict:
            mod_file = mod_name+'_'+mod_dict[mod_name]+'.xlsx'
            moduleFileList.append(mod_file)

    for mod_file in moduleFileList:
        wb = load_workbook(mod_file)
        ws = wb.active
        st_module_list, bCheckPass = checkModuleSheetVale(ws)
        if bCheckPass:
            print(f'{mod_file} checked Pass.')
        else:
            print(f'{mod_file} Check Failed. Please review the excel file and fix it.')
            filename = os.path.basename(mod_file)
            out_mark_xlsx_file = filename.replace('.xlsx', '_errMk.xlsx')
            # print(out_mark_xlsx_file)
            print("You can review the error mark file {0}.".format(
                out_mark_xlsx_file))
            wb.save(out_mark_xlsx_file)
