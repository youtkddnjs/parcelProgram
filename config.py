


#원본
soruceFile = ""
#단가표
priceListFile = ""

#파일 존재 유무
existFile = False

#WorkBook 로드 유무
successLoad = True

# 워크북 전역 변수 선언
srcWB = None
srcWS = None
srcWSreturn = None
priceWB = None
settleMentWB = None

new_wb = None
new_ws = None
new_return_ws = None
new_settle_ws = None

companyName = ""
companyCount = 0
unique_values = None
target_value = ""

priceColumn = 0
priceRow = 0
price = 0
priceComplete = False

#반품 여부
existReturn = False