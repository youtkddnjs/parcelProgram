


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

returnCount = 0 

# 극소 상자
box_tiny = 0  # 가장 작은 크기의 상자

# 소형 상자
box_small = 0  # 소형 크기의 상자

# 중형 상자
box_medium = 0  # 중간 크기의 상자

# 대형1 상자
box_large1 = 0  # 대형 크기(1단계)

# 대형2 상자
box_large2 = 0  # 대형 크기(2단계)

# 이형 상자
box_irregular = 0  # 비표준 모양 또는 특수 크기의 상자