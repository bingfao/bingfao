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


class St_Filed_info:
    def __init__(self, name, attr):
        self.end_bit = 31
        self.start_bit = 0
        self.attribute = attr
        self.defaultValue = 0
        self.field_name = name
        self.field_comments = ''

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
    print("Check Sheet Values")
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
                        regOffset=regOffsetCell.value
                        # print(regOffset)
                        if regOffset.find('0x') != 0:
                            print("Cell[F"+str(i+1)+"] offset Addr must be 0xFFFFFFF like hex string.")
                            bCheckPass = False
                        else:
                            reg_info.offset = int(regOffsetCell.value, 16)
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
                if bNewRegName:
                    if endBit < startBit:
                        print("Field End Pos must >= Start Pos at Row "+str(i+1))
                        bCheckPass = False
                        bFiled_info_Pass = False
                    laststartBit = 31
                else:
                    if endBit >= laststartBit:
                        print("Field End Pos must < last row Start Pos at Row "+str(i+1))
                        bCheckPass = False
                        bFiled_info_Pass = False
                    laststartBit = startBit
            else:
                print("Field End Pos and Start Pos at Row " +
                      str(i+1) + " must be Dec number string")
                bCheckPass = False
                bFiled_info_Pass = False

        if bFiled_info_Pass:
            field_name = ws.cell(i, 7).value
            if field_name != 'reserved' and field_name != 'unused':
                field_attr = ws.cell(i, 10).value
                field_inst = St_Filed_info(field_name, field_attr)
                field_inst.end_bit = int(ws.cell(i, 8).value)
                field_inst.start_bit = int(ws.cell(i, 9).value)
                if ws.cell(i, 11).ctype == xlrd.XL_CELL_TEXT:
                    # print( ws.cell(i,11))
                    field_inst.defaultValue = int(ws.cell(i, 11).value, 16)
                if ws.cell(i, 15).ctype == xlrd.XL_CELL_TEXT:
                    field_inst.field_comments = ws.cell(i, 15).value
                reg_info = st_module_list[-1].reg_list[-1]
                reg_info.add_field(field_inst)

        i += 1
    print("Endof Check Sheet Values")
    return st_module_list, bCheckPass


with xlrd.open_workbook("UART_final_202301010.xls") as book:
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

        with open('./cip_module_common.h', 'w+') as common_file:
            str_file_common = """// Autor: Auto generate by python From module excel\n
// Version: 0.0.1
// Description : Common struct define for field_info and reg_info\n
// When You have more than one module excel, you will get the same file of this. 
// Only need one copy for all modules. !!!
// Waring: Do NOT Modify it !

#ifndef  _CIP_MODULE_REG_FILED_STRUCT_COMMON_DEFINE_
#define  _CIP_MODULE_REG_FILED_STRUCT_COMMON_DEFINE_

#ifndef  _REG_FIELD_ATTRIBUTE_ENUM_DEFINE_
#define  _REG_FIELD_ATTRIBUTE_ENUM_DEFINE_
typedef enum  _REG_FIELD_ATTR_ENUM_  {
    ro, rw, rc, rs, wrc, wrs, wc, ws, wsrc,
    wcrs, w1c, w1s, w1t, w0c, w0s, w0t, w1src,
    w1crs, w0src, w0crs, wo, woc, wos, w1, wo1
} _REG_FIELD_ATTR_ENUM ;
#endif // end  _REG_FIELD_ATTRIBUTE_ENUM_DEFINE_


#pragma pack(1)
typedef struct {
    unsigned char end_bit;
    unsigned char start_bit;
    _REG_FIELD_ATTR_ENUM attr;
    unsigned int default_value;
} st_filed_info;

typedef struct {
    unsigned int reg_value; 
    unsigned int offset; 
    st_filed_info filed_0;
} st_reg_info_field_One;

typedef struct {
    unsigned int reg_value; 
    unsigned int offset; 
    unsigned char field_count;
    st_filed_info* pfield;
} st_reg_info_field_More;

#pragma pack()

#endif _CIP_MODULE_REG_FILED_STRUCT_COMMON_DEFINE_
"""
            common_file.write(str_file_common)
            common_file.close()

        if len(st_module_list):
            module_inst = st_module_list[0]
            modName = module_inst.module_name
            print(modName)
            out_file_Name = modName+'_reg'

            with open('./'+out_file_Name+'.h', 'w+') as out_file:
                fileHeader = """// Autor: Auto generate by python From module excel\n
// Version: 0.0.1
// Description : struct define for module \n

// Waring: Do NOT Modify it !
"""
                fileHeader += f'#ifndef  _CIP_MODULE_{modName}_DEFINE_\n'
                fileHeader += f'#define  _CIP_MODULE_{modName}_DEFINE_\n\n'
                fileHeader += """
#include "./cip_module_common.h"

"""

                file_body_str = """#pragma pack(1)
typedef struct {
    unsigned char bus_type ; // 0 for AHB, 1 for AXI
    unsigned int base_addr ;
    unsigned int data_width;
    unsigned int addr_width;
"""
#     st_filed_info* pfield;
# } st_module_{modName};"""
                # 定义module的结构体
                reg_str = "//define index for every reg_name\n"
                reg_index = 0
                field_str = "//define all index for every reg.field\n"
                for reg in module_inst.reg_list:
                    field_str += f'//define index for {reg.reg_name}.\n'
                    if reg.field_count() > 1:
                        # 定义为more
                        file_body_str += f'\tst_reg_info_field_More {reg.reg_name};\n'
                        reg_str += f'#define {modName}_reg_{reg.reg_name}  {reg_index}\n'
                    else:
                        # 定义为one
                        file_body_str += f'\tst_reg_info_field_One {reg.reg_name};\n'
                        reg_str += f'#define {modName}_reg_{reg.reg_name}  {reg_index}\n'
                    reg_index += 1
                    field_index = 0
                    for fd in reg.field_list:
                        field_str += f'#define {modName}_reg_{reg.reg_name}_{fd.field_name}  {field_index}\n'
                        field_index += 1
                    field_str += f'//end of define of {reg.reg_name}\n\n'

                reg_str += "//endof define index for every reg_name\n\n"

                file_body_str += "}"
                file_body_str += f'st_module_info_{modName};\n'
                file_body_str += "#pragma pack()\n\n"

                file_body_str += reg_str
                file_body_str += field_str

                file_body_str += f'unsigned char get{modName}InstCount(unsigned char bustype);  //0 for AHB, 1 for AXI\n'
                file_body_str += f'st_module_info_{modName} * get{modName}Instance(unsigned char bustype,unsigned char index);  //return NULL if not existed.\n'
                file_body_str += f'void resetModule_{modName} (st_module_info_{modName} *);  //reset the module to the default value\n'
                file_body_str += f'void reset{modName}Reg_Byindex(st_module_info_{modName} * pModule,unsigned char reg_index);  //reset the reg to default value\n'
                file_body_str += f'void set{modName}Reg_Byindex(st_module_info_{modName} * pModule,unsigned char reg_index,unsigned int reg_val);    //set the reg value to the reg_val\n'
                file_body_str += f'void set{modName}RegField_Byindex(st_module_info_{modName} * pModule,unsigned char reg_index,unsigned char field_index, unsigned int field_val);  //set the reg field value to the field_val\n'

                out_file.write(fileHeader)
                out_file.write(file_body_str)
                out_file.close()

            # for module in st_module_list:
            #     print(module.module_info_str())

                # 实例化各个module

        # with open('./'+out_file_Name+'.h', 'w+') as out_file:
        #     fileHeader = "// Autor: Auto generate by python From module excel\n" + \
        #         "//Date : \n//Description : Struct C file for block " + \
        #         modName + " , registers, fileds\n\n"
        #     out_file.write(fileHeader)
        #     struct_name = "__STRUCT_"+modName+"_DEFINE__H__FILE__"
        #     str_Mode_struct = "#ifndef "+struct_name+"\n#define "+struct_name+"\n\n"
        #     out_file.write(str_Mode_struct)
        #     out_file.write("#endif  //"+struct_name)
        #     out_file.close()
    else:
        print("Check Failed. Please review the excel file and fix it.")
