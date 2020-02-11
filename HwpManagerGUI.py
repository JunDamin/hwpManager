import PySimpleGUI as sg
from Model import getHwpFileAddress, MergeHwp, MergeSection


sg.change_look_and_feel("TanBlue")


layout = [
    [sg.InputText(), sg.FolderBrowse(button_text="병합폴더선택")],
    [
        sg.Checkbox("하위폴더 포함", size=(10, 1), default=True),
        sg.Checkbox("구역병합", size=(10, 1), default=False),
    ],
    [sg.Output(size=(100, 30))],
    [
        sg.Button(button_text="병합하기", key="OK"),
        sg.Button(button_text="종료", key="Cancel"),
    ],
]

window = sg.Window("HWP Manager v.0.1", icon="icon\\email.ico", layout=layout)


openText = """
======================================================================
폴더를 고르고 병합하기를 클릭하면 한글이 실행되고 외부명령 실행여부를 묻습니다.
"모두허용"을 선택하시면 한글파일이 병합이 됩니다.

환경
1. 아래아한글이 설치되어 있어야 합니다.

주요
1. 하위폴더까지 포함할지 선택할 수 있습니다.
2. 구역을 병합할 지 선택할 수 있습니다.(단 한 파일이 하나의 구역을 가지고 있어야 합니다.)
======================================================================
"""

endText = """
======================================================================
{}개 생성 작업이 완료 되었습니다.
======================================================================
"""


window.read(timeout=10)
print(openText)

while True:
    event, values = window.read()
    if event in (None, "Cancel"):
        break

    if event == "OK":
        if values[1]:
            fileList = getHwpFileAddress(values[0])
        else:
            fileList = os.listdir(values[0])

        hwpapi = MergeHwp(fileList, values[2])

window.close()
