import PySimpleGUI as sg
from Model import getHwpFileAddress, MergeHwp, MergeSection
from manager import process


sg.change_look_and_feel("TanBlue")

merge_layout = [
    [sg.InputText(size=(70, 1)), sg.FolderBrowse(button_text="병합폴더선택", size=(20, 1))],
    [
        sg.Checkbox("하위폴더 포함", size=(10, 1), default=True),
        sg.Checkbox("구역병합", size=(10, 1), default=False),
    ],
    [sg.Button(button_text="병합하기", key="Merge"),],
]

split_layout = [
    [sg.InputText(size=(70, 1)), sg.FileBrowse(button_text="분리파일선택", size=(20, 1))],
    [sg.Button(button_text="분리하기", key="Split"),],
]


layout = [
    [sg.Output(size=(100, 30))],
    [
        sg.TabGroup(
            [
                [
                    sg.Tab("Merge HWPs", merge_layout, tooltip="여러 한글파일을 하나로 합칩니다."),
                    sg.Tab("Split a HWP", split_layout, tooltip="한 한글파일을 구역별로 쪼갭니다."),
                ]
            ],
            tooltip="한글파일을 합치거나 쪼갤 수 있습니다.",
        ),
    ],
    [sg.Button(button_text="종료", key="Cancel")],
]


window = sg.Window("HWP Manager v.0.2", icon="icon\\email.ico", layout=layout)


openText = """
======================================================================
"Merge HWPs"탭에서 폴더를 고르고 병합하기를 클릭하면 한글이 실행되고 외부명령 실행여부를 묻습니다.
"모두허용"을 선택하시면 한글파일이 병합이 됩니다.

"Split a HWP"탭에서 HWP파일을 선택하고 분리하기를 클릭하면 한글파일이 분리되서 하위 폴더에 생성됩니다.

환경
1. 아래아한글이 설치되어 있어야 합니다.

주요
1. 하위폴더까지 포함할지 선택할 수 있습니다.
2. 구역을 병합할 지 선택할 수 있습니다.(단 한 파일이 하나의 구역을 가지고 있어야 합니다.)
======================================================================
"""


horizontal_bar = """
======================================================================
"""
window.read(timeout=10)
print(openText)

while True:
    event, values = window.read()
    if event in (None, "Cancel"):
        break

    if event == "Merge":
        fileList = getHwpFileAddress(path=values["병합폴더선택"], sub_folder=values[1])
        hwp_count = MergeHwp(fileList, values[2])
        endText = f"""
======================================================================
{hwp_count}개 생성 작업이 완료 되었습니다.
======================================================================
"""
        print(endText)

    if event == "Split":
        print(horizontal_bar)
        print("분리를 시작합니다.")
        window.refresh()
        process(values["분리파일선택"], window)
        print("분리작업이 종료 되었습니다. ", horizontal_bar)
window.close()
