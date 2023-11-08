import PySimpleGUI as sg
import sys
from xlsFlowX import dealwith_excel

# import paramiko


# def sftp_upload_file(host, user, password, server_path, local_path, timeout=10):
#     """
#     上传文件，注意：不支持文件夹
#     :param host: 主机名
#     :param user: 用户名
#     :param password: 密码
#     :param server_path: 远程路径，比如：/home/sdn/tmp.txt
#     :param local_path: 本地路径，比如：D:/text.txt
#     :param timeout: 超时时间(默认)，必须是int类型
#     :return: bool
#     """
#     try:
#         t = paramiko.Transport((host, 22))
#         t.banner_timeout = timeout
#         t.connect(username=user, password=password)
#         sftp = paramiko.SFTPClient.from_transport(t)
#         sftp.put(local_path, server_path)
#         t.close()
#         return True
#     except Exception as e:
#         print(e)
#         return False

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
