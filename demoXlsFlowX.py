# from openpyxl import load_workbook

# wb= load_workbook('./ahb_cfg_20230925.xlsx')
# ws = wb.active

# print(ws['A1'].value)
# print(ws['B1'].value)
# print(ws['C1'].value +':'+ ws['D1'].value)
# print(ws['A2'].value +':'+ ws['B2'].value)
# print(ws['C2'].value +':'+ ws['D2'].value)

import xlrd


char_arr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
busTypestr_arr = ("AHB", "AXI")
bitWidMask_arr = ('0x01', '0x03', '0x07', '0x0F', '0x1F', '0x3F', '0x7F', '0xFF', '0x01FF', '0x03FF', '0x07FF', '0x0FFF', '0x1FFF', '0x3FFF', '0x7FFF', '0xFFFF',
                  '0x01FFFF', '0x03FFFF', '0x07FFFF', '0x0FFFFF', '0x1FFFFF', '0x3FFFFF', '0x7FFFFF', '0xFFFFFF', '0x01FFFFFF', '0x03FFFFFF', '0x07FFFFFF', '0x0FFFFFFF', '0x1FFFFFFF', '0x3FFFFFFF', '0x7FFFFFFF', '0xFFFFFFFF')




class St_Filed_info:
    def __init__(self, name, attr):
        self.end_bit = 31
        self.start_bit = 0
        self.attribute = attr
        self.defaultValue = 0
        self.field_name = name
        self.field_comments = ''
        self.bModuleInit = False

    def field_info_str(self):
        out_str = f'fieldname: {self.field_name}, end_bit: {self.end_bit}, start_bit: {self.start_bit}, attribute: {self.attribute} \n , defaultValue: {hex(self.defaultValue)}, comments: {self.field_comments}'
        return out_str


class St_Reg_info:
    def __init__(self, name):
        self.field_list = []
        self.reg_name = name
        self.offset = 0
        self.bVirtual = False
        self.dim = 1   # 多组
        self.dimRange = 1  # 组间隔

    def field_count(self):
        return len(self.field_list)

    def is_fieldInReg(self, fd):
        bOut = fd in self.field_list
        return bOut

    def add_field(self, fd):
        self.field_list.append(fd)

    def getCHeaderString(self):
        return ""

    def reg_info_str(self):
        out_str = f'regname: {self.reg_name}, offset_addr: {hex(self.offset)}'
        for f in self.field_list:
            out_str += "\n"+f.field_info_str()
        return out_str


class St_Module_info:
    def __init__(self, name):
        self.module_name = name
        self.bus_baseAddr = 0
        self.addr_width = 32
        self.data_width = 32
        self.bus_type = 0
        self.reg_list = []

    def reg_count(self):
        return len(self.reg_list)

    def getAllFieldCount(self):
        nCount = 0
        for r in self.reg_list:
            nCount += r.field_count()
        return nCount

    def appendRegInfo(self, reginfo):
        self.reg_list.append(reginfo)

    def getCHeaderString(self):
        return ""

    def getCSourceString(self):
        return ""

    def module_info_str(self):
        out_str = f'moduleName: {self.module_name}, bus_type: {self.bus_type}, bus_addr: {hex(self.bus_baseAddr)}'
        for r in self.reg_list:
            out_str += "\n"+r.reg_info_str()
        return out_str


def checkModuleSheetVale(ws):
    print("Start Check Sheet Values.")
    modName = ws.cell(0, 1)
    baseAddr0 = ws.cell(0, 3)
    baseAddr1 = ws.cell(1, 5)
    data_width = ws.cell(1, 1)
    addr_with = ws.cell(1, 3)
    bCheckPass = True
    st_module_list = []
    if modName.ctype == xlrd.XL_CELL_EMPTY:
        print("ModuleName must be filled.")
        bCheckPass = False
    if baseAddr0.ctype == xlrd.XL_CELL_EMPTY and baseAddr1.ctype == xlrd.XL_CELL_EMPTY:
        print("baseAddr must be filled.")
        bCheckPass = False
    else:
        ahb_addr = baseAddr0.value
        ahb_addr_lst = ahb_addr.splitlines()
        for ahb in ahb_addr_lst:
            ahb_module = St_Module_info(modName.value)
            ahb_module.bus_baseAddr = int(ahb, 16)
            st_module_list.append(ahb_module)
        axi_addr = baseAddr1.value
        axi_addr_lst = axi_addr.splitlines()
        for axi in axi_addr_lst:
            axi_module = St_Module_info(modName.value)
            axi_module.bus_baseAddr = int(axi, 16)
            axi_module.bus_type = 1
            st_module_list.append(axi_module)

    if data_width.ctype == xlrd.XL_CELL_EMPTY:
        print("daa_width must be filled.")
        bCheckPass = False
    if addr_with.ctype == xlrd.XL_CELL_EMPTY:
        print("addr_witdh must be filled.")
        bCheckPass = False
    nRows = ws.nrows
    # nCols = ws.ncols
    regNameList = []
    i = 5
    laststartBit = 0
    while i < nRows:
        regNameCell = ws.cell(i, 0)
        regOffsetCell = ws.cell(i, 5)

        bNewRegName = False
        if regNameCell.ctype == xlrd.XL_CELL_TEXT:
            regName = regNameCell.value
            regName.strip()
            if len(regName) == 0:
                print(
                    "Cell[A"+str(i+1)+"] regName is empty string, Not allowed.")
                bCheckPass = False
            else:
                if not (regName in regNameList):
                    regNameList.append(regName)
                    bNewRegName = True
                    if regOffsetCell.ctype == xlrd.XL_CELL_EMPTY:
                        print(
                            "Cell[F"+str(i+1)+"] offset Addr must be filled.")
                        bCheckPass = False
                    elif regOffsetCell.ctype == xlrd.XL_CELL_TEXT:
                        reg_info = St_Reg_info(regName)
                        regOffset = regOffsetCell.value
                        # 待增加offset 值越来越大的规则判断

                        # print(regOffset)
                        if regOffset.find('0x') != 0:
                            print(
                                "Cell[F"+str(i+1)+"] offset Addr must be 0xFFFFFFF like hex string.")
                            bCheckPass = False
                        else:
                            reg_info.offset = int(regOffsetCell.value, 16)
                            if len(st_module_list):
                                module = st_module_list[-1]
                                if module.reg_count():
                                    lastOffset = module.reg_list[-1].offset
                                    if lastOffset >= reg_info.offset:
                                        print(
                                            "Cell[F"+str(i+1)+"] offset Addr must > last reg offset.")
                                        bCheckPass = False
                            for module in st_module_list:
                                module.appendRegInfo(reg_info)
                else:
                    print("Cell[A"+str(i+1)+"] regName repeated, Not allowed.")
                    bCheckPass = False

        # 处理每一行的 Field Info  这里需要重新考虑下
        bFiled_info_Pass = True
        col = 7
        while col < 11:
            ce = ws.cell(i, col)
            if ce.ctype == xlrd.XL_CELL_EMPTY:
                print(
                    "field_info cell[ " + char_arr[col] + str(i+1) + " ] must be filled.")
                bFiled_info_Pass = False
                bCheckPass = False
            elif ce.ctype == xlrd.XL_CELL_TEXT:
                str_ce = ce.value
                str_ce.strip()
                if len(str_ce) == 0:
                    print("field_info cell[ " + char_arr[col] +
                          str(i+1) + " ] is empty string, Not Allowed.")
                    bFiled_info_Pass = False
                    bCheckPass = False
            col += 1

        endBit_cell = ws.cell(i, 8)
        startBit_cell = ws.cell(i, 9)
        field_name = ws.cell(i, 7).value
        reg_info = st_module_list[-1].reg_list[-1]
        if field_name != 'reserved' and field_name in reg_info.field_list:
            print("Field Name not allow repeat at Row "+str(i+1))
            bFiled_info_Pass = False
            bCheckPass = False

        if (endBit_cell != xlrd.XL_CELL_EMPTY) and (startBit_cell.ctype != xlrd.XL_CELL_EMPTY):
            # 这里需要判断是否都是数字
            endBit = ws.cell(i, 8).value
            startBit = ws.cell(i, 9).value
            bEndbit_ok = False
            bStartbit_ok = False
            if endBit_cell.ctype == xlrd.XL_CELL_TEXT and endBit.isdecimal():
                endBit = int(endBit)
                bEndbit_ok = True
            if startBit_cell.ctype == xlrd.XL_CELL_TEXT and startBit.isdecimal():
                startBit = int(startBit)
                bStartbit_ok = True

            if endBit_cell.ctype == xlrd.XL_CELL_NUMBER:
                endBit = int(endBit)
                bEndbit_ok = True
            if startBit_cell.ctype == xlrd.XL_CELL_NUMBER:
                startBit = int(startBit)
                bStartbit_ok = True
                # print("EndBit:{0}, StartBit:{1}, LastStartBit: {2}".format(
                #     endBit, startBit, laststartBit))
            if bStartbit_ok and bEndbit_ok:
                # print("row: {0}, endPos: {1},startPos:{2},lastStartPos: {3}".format(str(i+1), endBit,startBit,laststartBit))
                if bNewRegName:
                    if endBit < startBit:
                        print("Field End Pos must >= Start Pos at Row "+str(i+1))
                        bCheckPass = False
                        bFiled_info_Pass = False
                    laststartBit = 31
                else:
                    if endBit >= laststartBit:
                        print(
                            "Field End Pos must < last row Start Pos at Row "+str(i+1))
                        bCheckPass = False
                        bFiled_info_Pass = False
                laststartBit = startBit
            else:
                print("Field End Pos and Start Pos at Row " +
                      str(i+1) + " must be Dec number string")
                bCheckPass = False
                bFiled_info_Pass = False

        if bFiled_info_Pass:
            # field_name = ws.cell(i, 7).value
            # if field_name != 'reserved' and field_name != 'unused':
            field_attr = ws.cell(i, 10).value
            field_inst = St_Filed_info(field_name, field_attr)
            field_inst.end_bit = int(ws.cell(i, 8).value)
            field_inst.start_bit = int(ws.cell(i, 9).value)
            if ws.cell(i, 11).ctype == xlrd.XL_CELL_TEXT:
                # print( ws.cell(i,11))
                field_inst.defaultValue = int(ws.cell(i, 11).value, 16)
            comments = ws.cell(i, 16)
            if comments.ctype == xlrd.XL_CELL_TEXT:   #
                field_inst.field_comments = comments.value
            moduleInit = ws.cell(i, 15)
            if moduleInit.ctype == xlrd.XL_CELL_NUMBER:   #
                field_inst.bModuleInit = (moduleInit.value == 1)
            reg_info.add_field(field_inst)

        i += 1
    if bCheckPass:
        print("Check Sheet:  result Pass")
    return st_module_list, bCheckPass



def dealwith_excel(xls_file):
#"UART_final_202301010.xls"
    with xlrd.open_workbook(xls_file) as book:
        # print("The number of worksheets is {0}".format(book.nsheets))
        # print("Worksheet name(s): {0}".format(book.sheet_names()))
        sh = book.sheet_by_index(0)
        # print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        # # print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
        # for rx in range(sh.nrows):
        #     print(sh.row(rx))
        # modName = sh.cell(0, 1).value
        # baseAddr = sh.cell(0, 3).value
        # data_width = sh.cell(1, 1).value
        # addr_with = sh.cell(1, 3).value

        st_module_list, bCheckPass = checkModuleSheetVale(sh)
        if bCheckPass:
            if len(st_module_list):
                module_inst = st_module_list[0]
                modName = module_inst.module_name
                print('module name: {0}.'.format(modName))
                out_C_file_Name = modName+'_reg'
                with open('./'+out_C_file_Name+'.h', 'w+') as out_file:
                    fileHeader = """// Autor: Auto generate by python From module excel\n
// Version: 0.0.2 X
// Description : struct define for module \n

// Waring: Do NOT Modify it !

#pragma once
"""
                    fileHeader += f'#ifndef  _CIP_MODULE_{modName}_DEFINE_\n'
                    fileHeader += f'#define  _CIP_MODULE_{modName}_DEFINE_\n\n'
                    fileHeader += """

#include <stdint.h>   //for use uint32_t type

#ifdef __cplusplus
extern "C" {
#endif 

"""

                    file_body_str = """#pragma pack(4)
typedef struct {
"""
    #     st_filed_info* pfield;
    # } st_module_{modName};"""
                    # 定义module的结构体
                    last_offset = 0
                    nRegReservedIndex = 0
                    field_define_str = """#ifdef _CIP_REG_POS_OPRATION_
////////////////////////////////////////////////////////////////////////
////////////////////define for field opration///////////////////////////
"""
                    for reg in module_inst.reg_list:
                        reg_offset = reg.offset
                        if reg_offset != last_offset:
                            # 增加占位
                            nRerived = (reg_offset-last_offset) / 4
                            n = 0
                            while n < nRerived:
                                file_body_str += f'\tvolatile uint32_t u_reg_reserved{nRegReservedIndex};\n'
                                nRegReservedIndex += 1
                                n += 1
                        last_offset = reg_offset + 4
                        field_count = reg.field_count()
                        if field_count:
                            file_body_str += "\tvolatile struct  {\n"
                            nFieldReservedIndex = 0
                            bReserved = False
                            field_index = field_count-1
                            field_bitPos = 0
                            while field_index != -1:
                                fd = reg.field_list[field_index]
                                if fd.start_bit != field_bitPos:
                                    # 需要补齐field
                                    file_body_str += f'\t\tunsigned fd_reserved{nFieldReservedIndex} : {fd.start_bit-field_bitPos} ;\n'
                                    nFieldReservedIndex += 1

                                field_bitPos = fd.end_bit+1
                                fd.field_comments = fd.field_comments.replace(
                                    '\n', ' ').replace('\r', ' ')
                                nBitWid = field_bitPos-fd.start_bit
                                if fd.field_name == 'reserved':
                                    fd.field_name = f'reserved{nFieldReservedIndex}'
                                    nFieldReservedIndex += 1
                                    bReserved = True
                                file_body_str += f'\t\tunsigned fd_{fd.field_name} : {nBitWid} ; /*{fd.field_comments} */\n'
                                field_index -= 1
                                if not bReserved:
                                    field_str_ = f'{reg.reg_name.upper()}_{fd.field_name.upper()}'
                                    field_define_str += f'//define for {field_str_}\n'
                                    field_define_str += f'#define \t {field_str_}_POS \t      {fd.start_bit}U\n'
                                    strfdMask = f'{bitWidMask_arr[nBitWid-1]}'
                                    field_define_str += f'#define \t {field_str_}_MSK \t      ((uint32_t){strfdMask} << {field_str_}_POS)\n'
                                    if fd.attribute.find('W') != -1:
                                        field_define_str += f'#define \t {field_str_}_SET(val) \t  ((uint32_t)((val) & {strfdMask}) << {field_str_}_POS)\n'

                                    field_define_str += f'#define \t {field_str_}_GET(val) \t  ((uint32_t)((val) & {field_str_}_MSK) >> {field_str_}_POS)\n'
                                    field_define_str += '\n\n'
                                # define QSPI_FCMDCR_NMDMYC_POS          7U
    # define QSPI_FCMDCR_NMDMYC_MSK          ((uint32_t)0x1F << QSPI_FCMDCR_NMDMYC_POS)
    # define QSPI_FCMDCR_NMDMYC              QSPI_FCMDCR_NMDMYC_MSK
    # define QSPI_FCMDCR_NMDMYC_SET(val)     ((uint32_t)((val) & 0x1F) << QSPI_FCMDCR_NMDMYC_POS)
    # define QSPI_FCMDCR_NMDMYC_GET(val)     ((uint32_t)((val) & QSPI_FCMDCR_NMDMYC_MSK) >> QSPI_FCMDCR_NMDMYC_POS)

                            file_body_str += "\t}\t"+f'st_reg_{reg.reg_name};\n'
                        else:
                            file_body_str += f'\tvolatile uint32_t u_reg_{reg.regname};\n'

                        # if reg.field_count() > 1:
                        #     # 定义为more
                        #     file_body_str += f'\tst_reg_info_field_More {reg.reg_name};\n'
                        #     reg_str += f'#define {modName}_reg_{reg.reg_name}  {reg_index}\n'
                        # else:
                        #     # 定义为one
                        #     file_body_str += f'\tst_reg_info_field_One {reg.reg_name};\n'
                        #     reg_str += f'#define {modName}_reg_{reg.reg_name}  {reg_index}\n'
                        # reg_index += 1
                        # field_index = 0
                        # for fd in reg.field_list:
                        #     field_str += f'#define {modName}_reg_{reg.reg_name}_{fd.field_name}  {field_index}\n'
                        #     field_index += 1
                        # field_str += f'//end of define of {reg.reg_name}\n\n'

                    # reg_str += "//endof define index for every reg_name\n\n"

                    file_body_str += "}"
                    file_body_str += f'st_module_info_{modName};\n'
                    file_body_str += "#pragma pack()\n\n"

                    # file_body_str += reg_str
                    # file_body_str += field_str

                    field_define_str += """
////////////////////define for field opration///////////////////////////

#endif // _CIP_REG_POS_OPRATION_


"""
                    file_body_str += '\n\n'+field_define_str

                    file_body_str += f'\n\n\n#define \t GET_{modName.upper()}_HANDLE   ( (st_module_info_{modName} *) base_addr)\n\n'

                    inst_str = """
////////////////////define for module instance///////////////////////////
"""
                    nbusAddrindex = 0
                    for mo in st_module_list:
                        inst_ = f'{modName}_{busTypestr_arr[mo.bus_type]}_baseAddr{nbusAddrindex}'
                        file_body_str += f'#define \t {inst_}  \t{hex(mo.bus_baseAddr)}\n'
                        inst_str += f'#define {modName.upper()}_{nbusAddrindex}  ( (st_module_info_{modName} *) {inst_})\n'
                        nbusAddrindex += 1

                    inst_str += """
////////////////////end of define for module instance///////////////////////////

"""
                    file_body_str += inst_str
                    file_body_str += f'\n#endif //endof  _CIP_MODULE_{modName}_DEFINE_\n'

                    file_body_str += """
#ifdef __cplusplus
}  //endof extern "C"
#endif

"""

                    out_file.write(fileHeader)
                    out_file.write(file_body_str)
                    out_file.close()

                out_svh_module_Name = modName.lower()+'_dut_cfg'
                with open('./'+out_svh_module_Name+'.svh', 'w+') as sv_file:
                    heder_str = f'_{modName.upper()}_DUT_CFG_SVH_'
                    file_str = F'`ifndef {heder_str}\n`define {heder_str}\n\n'

                    uvm_field_str = f'\n\t`uvm_object_utils_begin({out_svh_module_Name})\n'
                    uvm_fd_val_def_str = ""
                    val_def_strarr = ["// Autor: Auto generate by sv", "// Version: 0.0.2 X",
                                    "// Description : set module reg field random value", "// Waring: Do NOT Modify it !", "#pragma once"]
                    for str in val_def_strarr:
                        uvm_fd_val_def_str += f'\t\t$fdisplay(fd, "{str}" );\n'

                    uvm_fd_val_def_str += '\t\t$fdisplay(fd, "   " );\n\n'
                    file_str += f'class {out_svh_module_Name} extends uvm_object;\n\n'
                    for reg in module_inst.reg_list:
                        for fd in reg.field_list:
                            if fd.bModuleInit and fd.attribute.find('W') != -1:
                                reg_fd_name = f'{reg.reg_name}___{fd.field_name}'
                                uvm_field_str += f'\t\t`uvm_field_int({reg_fd_name}, UVM_ALL_ON)\n'
                                if fd.end_bit > fd.start_bit:
                                    file_str += f'\trand bit [{fd.end_bit-fd.start_bit}:0]  {reg_fd_name};\n'
                                else:
                                    file_str += f'\trand bit {reg_fd_name};\n'
                                fd_name_VAL = f'{reg_fd_name.upper()}_VALUE_'
                                fd_name_VAL = fd_name_VAL.ljust(48)
                                uvm_fd_val_def_str += f'\t\t$fdisplay(fd, "#define \t {fd_name_VAL}   0x%X",  {reg_fd_name});\n'
                    uvm_field_str += f'\t`uvm_object_utils_end\n'

                    uvm_field_str += f'\n\tfunction new(string name = "{out_svh_module_Name}");\n'
                    uvm_field_str += """\t\tsuper.new(name);
    endfunction:new

    virtual function void print_cfg_to_file();
        int fd;
"""
                    uvm_field_str += f'\t\tfd = $fopen("{modName}_dut_cfg.h");\n'

                    uvm_field_str += uvm_fd_val_def_str
                    uvm_field_str += """
        $fclose(fd);
    endfunction:print_cfg_to_file

"""
                    file_str += uvm_field_str
                    file_str += """endclass
`endif

"""
                    sv_file.write(file_str)
                    sv_file.close()

                # for module in st_module_list:
                #     print(module.module_info_str())

                    # 实例化各个module

        else:
            print("Check Failed. Please review the excel file and fix it.")


import PySimpleGUI as sg      
import sys

if len(sys.argv) == 1:
    event, values = sg.Window('CIP Excel to DV',
                    [[sg.Text('请选择模块excel文件.')],
                    [sg.In(), sg.FileBrowse()],
                    [sg.Open(), sg.Cancel()]]).read(close=True) # type: ignore
    fname = values[0]
else:
    fname = sys.argv[1]

if not fname:
    sg.popup("Cancel", "No filename supplied")
    raise SystemExit("Cancelling: no filename supplied")
else:
    # sg.popup('The filename you chose was', fname)
    if fname.endswith('.xls'):
        dealwith_excel(fname)