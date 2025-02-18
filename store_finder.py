import pandas as pd
from difflib import SequenceMatcher

# 데이터셋 불러오기
data = pd.read_csv("companies.csv")

# 유사도 계산 함수
def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# 거래처 검색 함수
def find_business_number(name, region):
    # 검색 결과를 저장할 변수
    result = None

    # 데이터셋 탐색
    for _, row in data.iterrows():
        dataset_name = row["이름"]
        dataset_region = row["지역"]

        # 이름과 지역이 정확히 일치하는지 확인
        if dataset_name == name and dataset_region == region:
            result = row["사업자등록번호"]
            break

        # 유사도를 활용하여 근사값 확인 (옵션)
        elif similarity(dataset_name, name) > 0.8 and similarity(dataset_region, region) > 0.8:
            result = row["사업자등록번호"]
            break

    # 결과 반환
    return result

# 테스트 실행
'''
name_input = "삼성전자"
region_input = "서울"
business_number = find_business_number(name_input, region_input)

if business_number:
    print(f"사업자등록번호: {business_number}")
else:
    print("검색 결과를 찾을 수 없습니다.")
'''
x = 'ㅅㅗㅁㅏㄴㅡㄹ'
y = 'ㅅㅗㅁㅏㄴㄴㅡ'
score = similarity(x, y)

print(score)
