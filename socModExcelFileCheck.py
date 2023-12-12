from xlsFlowX import checkModuleSheetVale,output_C_moduleFile, output_SequenceSv_moduleFile,output_ralf_moduleFile,output_SV_moduleFile,outModuleFieldDefaultValueCheckCSrc
from openpyxl import load_workbook
import os


def HexVal(var):
    hxstr=hex(var)
    return hxstr[2:].upper()

if __name__ == '__main__':
    # print(HexVal(245))
    
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
                            # print('name: {0} , date: {1}'.format(mod_name, date_val))
                            if mod_name in mod_dict:
                                if mod_dict[mod_name] < date_val:
                                    mod_dict[mod_name] = date_val
                            else:
                                mod_dict[mod_name] = date_val

                    # moduleFileList.append(entry_name)
        for mod_name in mod_dict:
            mod_file = mod_name+'_'+mod_dict[mod_name]+'.xlsx'
            moduleFileList.append(mod_file)

    bAllChecPass=True
    soc_module_dict={}
    for mod_file in moduleFileList:
        print(mod_file)
        wb = load_workbook(mod_file)
        ws = wb.active
        modName, st_module_list, bCheckPass = checkModuleSheetVale(ws)
        if bCheckPass:
            print(f'{mod_file} checked Pass.')        
            soc_module_dict[modName]=st_module_list
        else:
            print(f'{mod_file} Check Failed. Please review the excel file and fix it.')
            bAllChecPass=False
            # filename = os.path.basename(mod_file)
            # out_mark_xlsx_file = filename.replace('.xlsx', '_errMk.xlsx')
            # # print(out_mark_xlsx_file)
            # print("You can review the error mark file {0}.".format(
            #     out_mark_xlsx_file))
            # wb.save(out_mark_xlsx_file)

    # if all check pass 生成相应的文件
    if bAllChecPass:
        soc_ralf_body_str=''
        soc_ralf_AHB_str='\n\tdomain AHB {\n\t\tbytes 4;\n'
        soc_ralf_AXI_str='\n\tdomain AXI {\n\t\tbytes 4;\n'
        for modName in soc_module_dict:
            print("module: "+modName)
            st_module_list=soc_module_dict[modName]
            mod_inst_count = len(st_module_list)
            if mod_inst_count:
                module_inst = st_module_list[0]
                # print('module name: {0}.'.format(modName))

                out_file_list = []
                out_file_name = output_C_moduleFile(
                    st_module_list, module_inst, modName)
                if out_file_name:
                    out_file_list.append(out_file_name)

                out_file_name = output_SV_moduleFile(module_inst, modName)
                if out_file_name:
                    out_file_list.append(out_file_name)

                ahb_pos = len(st_module_list)
                for index in range(mod_inst_count):
                    if st_module_list[index].bus_type:
                        ahb_pos = index
                        break
                    
                out_file_name = output_SequenceSv_moduleFile(st_module_list[0:ahb_pos], modName)
                if (out_file_name):
                    out_file_list.append(out_file_name)

                axi_len=mod_inst_count-ahb_pos
                modName_U=modName.upper()
                if ahb_pos>0:
                    for index in range(ahb_pos):
                        mod_inst=st_module_list[index]
                        hal_path=mod_inst.hdl_path
                        baseAddr = mod_inst.bus_baseAddr
                        if isinstance(baseAddr,int):
                            baseAddr = baseAddr & 0x1FFFFFFF
                        if hal_path and hal_path != 'NULL':
                            soc_ralf_AHB_str+=f'\t\tblock {modName} = {modName_U}{index} ({mod_inst.hdl_path}) @\'h{HexVal(baseAddr)} ;\n'
                        else:
                            soc_ralf_AHB_str+=f'\t\tblock {modName} = {modName_U}{index} @\'h{HexVal(baseAddr)} ;\n'
                    # if ahb_pos>1:
                    #     ahb_baseAddr_lst=[]
                    #     for index in range(ahb_pos):
                    #         ahb_baseAddr_lst.append(st_module_list[index].bus_baseAddr)
                    #     ahb_baseAddr_lst.sort()
                    #     bEqualDist=False
                    #     ahb_inst_len=len(ahb_baseAddr_lst)
                    #     if ahb_inst_len>2:
                    #         bEqualDist= (ahb_baseAddr_lst[-1]-ahb_baseAddr_lst[-2] == ahb_baseAddr_lst[1]-ahb_baseAddr_lst[0]) 
                    #     elif ahb_inst_len ==2:
                    #         bEqualDist=True
                    #     if bEqualDist:
                    #         nDist=ahb_baseAddr_lst[1]-ahb_baseAddr_lst[0]
                    #         soc_ralf_AHB_str+=f'\t\tblock {modName}[{ahb_inst_len}] @\'h{HexVal(ahb_baseAddr_lst[0])}+\'h{HexVal(nDist)} ;\n'
                    #     else:
                    #         for index in range(ahb_pos):
                    #             soc_ralf_AHB_str+=f'\t\tblock {modName}{index} @\'h{HexVal(st_module_list[index].bus_baseAddr)} ;\n'
                    # elif ahb_pos==1:
                    #     soc_ralf_AHB_str+=f'\t\tblock {modName} @\'h{HexVal(module_inst.bus_baseAddr)} ;\n'
                
                if axi_len>0:
                    for index in range(ahb_pos,mod_inst_count):
                        mod_inst=st_module_list[index]
                        hal_path=mod_inst.hdl_path
                        baseAddr = mod_inst.bus_baseAddr
                        if isinstance(baseAddr,int):
                            baseAddr = baseAddr & 0x1FFFFFFF
                        if hal_path and hal_path != 'NULL':
                            soc_ralf_AXI_str+=f'\t\tblock {modName} = {modName_U}{index} ({mod_inst.hdl_path}) @\'h{HexVal(baseAddr)} ;\n'
                        else:
                            soc_ralf_AXI_str+=f'\t\tblock {modName} = {modName_U}{index} @\'h{HexVal(baseAddr)} ;\n'
                    # module_inst_axi = st_module_list[-1]
                    # if axi_len ==1:
                    #     soc_ralf_AXI_str+=f'\t\tblock {modName} @\'h{HexVal(module_inst_axi.bus_baseAddr)} ;\n'
                    # elif axi_len>1:
                    #     axi_baseAddr_lst=[]
                    #     for index in range(ahb_pos,mod_inst_count):
                    #         axi_baseAddr_lst.append(st_module_list[index].bus_baseAddr)
                    #     axi_baseAddr_lst.sort()
                    #     bEqualDist=False
                    #     axi_inst_len=len(axi_baseAddr_lst)
                    #     if axi_inst_len>2:
                    #         bEqualDist= (axi_baseAddr_lst[-1]-axi_baseAddr_lst[-2] == axi_baseAddr_lst[1]-axi_baseAddr_lst[0]) 
                    #     elif axi_inst_len==2:
                    #         bEqualDist=True
                    #     if bEqualDist:
                    #         nDist=axi_baseAddr_lst[1]-axi_baseAddr_lst[0]
                    #         soc_ralf_AXI_str+=f'\t\tblock {modName}[{axi_inst_len}] @\'h{HexVal(axi_baseAddr_lst[0])}+\'h{HexVal(nDist)} ;\n'
                    #     else:
                    #         for index in range(ahb_pos,mod_inst_count):
                    #             soc_ralf_AXI_str+=f'\t\tblock {modName}{index} @\'h{HexVal(st_module_list[index].bus_baseAddr)} ;\n'
                    

                out_file_name = outModuleFieldDefaultValueCheckCSrc(
                    st_module_list[0:ahb_pos], modName)
                if out_file_name:
                    out_file_list.append(out_file_name)

                out_file_name = output_ralf_moduleFile(module_inst, modName)
                if out_file_name:
                    out_file_list.append(out_file_name)

                soc_ralf_body_str+=f'source {out_file_name}\n'


                # outModuleFieldDefaultValueCheckCSrc(st_module_list[0:1], modName)

                for out_file in out_file_list:
                    print('generate: '+out_file)
                
                print("module: "+modName+" Pass.")
                # 实例化各个module
        soc_ralf_AHB_str+='\t}\n'
        soc_ralf_AXI_str+='\t}\n'
        soc_ralf_body_str+='system soc {\n'
        out_soc_file_Name = './soc.ralf'
        with open(out_soc_file_Name, 'w+') as out_file:
            out_file.write(soc_ralf_body_str)
            out_file.write(soc_ralf_AHB_str)
            out_file.write(soc_ralf_AXI_str)

            out_file.write('}\n')
            out_file.close()
            print('generate: '+out_soc_file_Name)
        

        
        

