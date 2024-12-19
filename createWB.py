import config
import utils
import os
import sys
import openpyxl as op

def get_current_directory():
    # 실행 파일이 위치한 폴더를 반환
    if getattr(sys, 'frozen', False):  # 실행 파일인 경우
        return os.path.dirname(sys.executable)
    else:  # 스크립트 실행인 경우
        return os.path.dirname(os.path.abspath(__file__))
    
def getFileName():
    utils.clear_console()
    print("파일을 확인합니다.")
    
    
    # 실행 파일의 실제 경로를 기준으로 설정
    current_directory = get_current_directory()
    print(f"현재 경로: {current_directory}")
    
    # 현재 폴더의 파일 및 디렉토리 목록 가져오기
    files_and_dirs = os.listdir(current_directory)

    # 파일만 가져오기
    files = [f for f in files_and_dirs if os.path.isfile(os.path.join(current_directory, f))]

    for file in files:
        if file == "soruce.xlsx":
            config.soruceFile = "soruce.xlsx"
            config.existFile = True
        elif file == "priceList.xlsx":
            config.priceListFile = "priceList.xlsx"
            config.existFile = True

    # 파일 상태 확인
    if not config.soruceFile:
        print("원본 파일이 없습니다.")
        config.existFile = False
    if not config.priceListFile:
        print("단가표 파일이 없습니다.")
        config.existFile = False



def setWorkBook():
    if config.existFile:
        print("워크북 로드를 시작합니다.")
        try:
            # 워크북 객체 생성
            print("워크북 로드 중...")
            config.srcWB = op.load_workbook(get_current_directory()+"/"+config.soruceFile)
            config.priceWB = op.load_workbook(get_current_directory()+"/"+config.priceListFile)
            config.successLoad = True
            print("워크북들이 성공적으로 로드되었습니다.")
        except FileNotFoundError as e:
            print(f"파일 로드 중 오류가 발생했습니다: {e}")
            config.successLoad = False
    else:
        print("파일이 없어서 워크북을 로드할 수 없습니다.")
        config.successLoad = False
        
