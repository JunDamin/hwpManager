import PySimpleGUI as sg
from Model import getHwpFileAddress, MergeHwp, MergeSection


sg.change_look_and_feel("TanBlue")


layout = [
    [sg.InputText(), sg.FolderBrowse(button_text="병합폴더선택")],
    [
        sg.Checkbox("하위폴더 포함", size=(10, 1), default=True),
        sg.Checkbox("구역병합", size=(10, 1), default=True),
    ],
    [sg.Output(size=(100, 30))],
    [
        sg.Button(button_text="병합하기", key="OK"),
        sg.Button(button_text="종료", key="Cancel"),
    ],
]

window = sg.Window("HWP Manager v.0.1", icon="icon\\email.ico", layout=layout)


while True:
    event, values = window.read()
    if event in (None, "Cancel"):
        break

    if event == "OK":
        if values[1]:
            fileList = getHwpFileAddress(values[0])
        else:
            fileList = os.listdir(values[0)

        hwpapi = MergeHwp(fileList, values[2])

window.close()
