import createWB
import shipping
import utils
import config
import openpyxl as op


def main():
    createWB.getFileName()
    createWB.setWorkBook()
    while True:
        if not config.successLoad:
            print("정상적으로 워크북을 로드하지 못했습니다. 작업을 중단합니다.")
            break  # 함수 종료
        try:
            # 사용자 입력
            print("")
            print("")
            print("")
            user_input = int(input("0 : 종료\n1 : 정산 하기\n입력 : "))

            if user_input == 0:
                print("")
                print("")
                print("")
                print("프로그램을 종료합니다.")
                break  # 반복문 종료
            elif user_input == 1:
                if config.companyCount == 0:
                    shipping.getcustomerList()
                    selectNum()
                else:
                    utils.clear_console()
                    print("고객사 리스트:")
                    for idx, value in enumerate(config.unique_values, start=1):  # enumerate로 번호 추가, start=1로 시작 값 설정
                        print(f"{idx}. {value}")
                    print(f"업체 수 : {config.companyCount}")
                    selectNum()
            else:
                print("잘못된 입력입니다.")
        except ValueError:
            print("숫자를 입력하세요.")  # 숫자가 아닌 입력 처리

def selectNum():
    try:
        # 사용자 입력
        print("0 입력시 모든 업체 정산을 시작 합니다.")
        user_input = int(input("업체명 번호를 입력하세요 : "))
        shipping.selectCustomer(user_input)
    except ValueError:
        print("숫자를 입력하세요.")  # 숫자가 아닌 입력 처리

# 메인에서만 실행        
if __name__ == "__main__":
    main()

