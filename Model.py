import os
from hwp_api_class import HwpApi


def getHwpFileAddress(path, neglect_keys=["_", "."], sub_folder=False):

    """ """

    hwpAddressList = []

    for root, dirs, files in os.walk(path):
        if sub_folder == False:
            dirs[:] = []
        dirs[:] = [d for d in dirs if not d[0] in neglect_keys]
        hwpList = [os.path.join(root, file) for file in files if file[-3:] == "hwp"]
        hwpAddressList.extend(hwpList)

    return hwpAddressList


def MergeHwp(file_list, mergeSection):
    hwpapi = HwpApi()
    # 파일 목록 만들기(역순으로 만들어야 끼어 넣기가 순서대로)
    FileList = file_list
    print("\n".join(FileList), "\n")
    FileList.reverse()
    HwpFileList = [file for file in FileList if file[-3:] == "hwp"]

    if len(HwpFileList) == 0:
        hwpapi.hwpQuit()
    print('"모두 허용"을 선택하시면 병합이 시작됩니다.')

    # 병합하기
    for i in HwpFileList:
        ongoing = hwpapi.HwpInsertFileFuction(i)
        print(ongoing)
    hwpapi.HWPAUTO.Run("Delete")  # 첫페이지(비어 있음) 삭제
    HwpPageNumber = hwpapi.HWPAUTO.CreateAction("PageNumPos")  # 페이지 번호 넣는 액션 생성
    HwpPageNumberPara = HwpPageNumber.CreateSet()  # 파라미터 세트 생성
    HwpPageNumberPara.SetItem("DrawPos", "BottomCenter")  # 파라미터 지정
    HwpPageNumber.GetDefault(HwpPageNumberPara)  # 파라미터 연결
    HwpPageNumber.Execute(HwpPageNumberPara)  # 실행
    print("작업이 완료되었습니다.")

    if mergeSection:
        MergeSection(hwpapi, len(HwpFileList) - 1)

    return hwpapi


def MergeSection(hwpapi, n):
    for i in range(n):
        hwpapi.Goto(2)
        hwpapi.HWPAUTO.Run("DeleteBack")
