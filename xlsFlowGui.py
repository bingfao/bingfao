import PySimpleGUI as sg
import sys
import os
from xlsFlowX import dealwith_excel

import paramiko


def sftp_upload_file(host, port, user, password, out_file_list, timeout=10):
    """
    上传文件，注意：不支持文件夹
    :param host: 主机名
    :param user: 用户名
    :param password: 密码
    :param server_path: 远程路径，比如：/home/sdn/tmp.txt
    :param local_path: 本地路径, 比如: D:/text.txt
    :param timeout: 超时时间(默认), 必须是int类型
    :return: bool
    """
    try:
        t = paramiko.Transport((host, int(port)))
        t.banner_timeout = timeout
        t.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        print("当前工作目录 : %s" % os.getcwd())
        # svr_path = './'
        for out_f in out_file_list:
            out_file_name = os.path.abspath(out_f)
            svr_file_path = os.path.normpath(out_f)
            print(svr_file_path)
            svr_file_path = svr_file_path.replace('\\', '__')
            print(svr_file_path)
            # print(out_file_name)
            sftp.put(out_file_name, svr_file_path)
        t.close()
        return True
    except Exception as e:
        print(e)
        return False


if __name__ == '__main__':
    layout = [[sg.Text('请选择模块excel文件.')],
              [sg.In(key='-fileName-'), sg.FileBrowse(
                  file_types=(("excel files", "*.xlsx"),))], [sg.Button('DoXlsXFlow')],  [sg.Text('SvrIp:', s=8), sg.In(key='-svrIp-', s=15, default_text='127.0.0.1'), sg.Text('SvrPort:', s=8), sg.In(key='-svrPort-', s=15, default_text='22')], [sg.Text('username:', s=8), sg.In(key='-username-', s=15), sg.Text('password:', s=8), sg.In(key='-usrpsw-', s=15, password_char='*')], [sg.Button('Upload')], [sg.CloseButton('Exit')]]
    window = sg.Window('CIP Excel to DV', layout)
    out_file_list = []
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Upload':
            svrIp = values['-svrIp-']
            svrPort = values['-svrPort-']
            usrName = values['-username-']
            usrPsw = values['-usrpsw-']
            if out_file_list and svrIp and usrName:
                sftp_upload_file(svrIp, svrPort, usrName,
                                 usrPsw, out_file_list)
        elif event == 'DoXlsXFlow':
            fname = values['-fileName-']
            if (fname):
                out_file_list = dealwith_excel(fname)

    window.close()


# if not fname:
#     sg.popup("Cancel", "No filename supplied")
#     raise SystemExit("Cancelling: no filename supplied")

# else:
#     # sg.popup('The filename you chose was', fname)
#     if fname.endswith('.xlsx'):
#         dealwith_excel(fname)
