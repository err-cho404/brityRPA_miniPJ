import OpenDartReader
my_api = input('DART OPEN API키를 입력해주십시오: ')
dart = OpenDartReader(my_api)
import FinanceDataReader as fdr # FinanceDataReader 모듈 불러오기
stock_list = fdr.StockListing('KRX') # 코스피, 코스닥, 코넥스 종목 모두 불러오기
stock_list = stock_list.loc[stock_list['Market'].isin(["KOSPI", "KOSDAQ"])] # Market 기준 필터링
stock_name_list = stock_list['Name'].tolist()# 선물과 같이 KOSPI에는 있으나 배당이 없을수밖에 없는 종목은 제외
import numpy as np
import pandas as pd
def find_financial_indicators(stock_name, year, indicators):
    report = dart.finstate(stock_name, year) # 데이터 가져오기
    if report is None: # 리포트가 없다면 (참고: 리포트가 없으면 None을 반환함)
        # 리포트가 없으면 당기, 전기, 전전기 값 모두 제거
        data = [[stock_name, year] + [np.nan] * len(indicators)]
        data = [[stock_name, year-1] + [np.nan] * len(indicators)]
        data = [[stock_name, year-2] + [np.nan] * len(indicators)]
        return pd.DataFrame(data, columns = ["기업", "연도"] + indicators)
    
    else:
        report = report[report['account_nm'].isin(indicators)] # 관련 지표로 필터링
        if sum(report['fs_nm'] == "연결재무제표") > 0:
            # 연결재무제표 데이터가 있으면 연결재무제표를 사용
            report = report.loc[report['fs_nm'] == "연결재무제표"]
            
        else:
            # 연결재무제표 데이터가 없으면 일반재무제표를 사용
            report = report.loc[report['fs_nm'] == "재무제표"]
        
        data = []
        for y, c in zip([year, year-1, year-2], ['thstrm_amount', 'frmtrm_amount', 'bfefrmtrm_amount']):
            record = [stock_name, y]
            for indic in indicators:
                # account_nm이 indic인 행의 c 컬럼 값을 가져옴
                if sum(report['account_nm'] == indic) > 0: # 해당 지표가 있다면 추가 (지표가 없는 경우도 있음)
                    value = report.loc[report['account_nm'] == indic, c].iloc[0]
                else:
                    value = np.nan
                
                record.append(value)
            
            data.append(record)
        return pd.DataFrame(data, columns = ["기업", "연도"] + indicators)
    
import time 
from datetime import datetime
indicators = ['자산총계', '부채총계', '자본총계', '매출액', '영업이익', '당기순이익']
data = pd.DataFrame() # 이 데이터프레임에 각각의 데이터를 추가할 예정

now = datetime.now()
current_year=now.year

for idx, stock_name in enumerate(stock_name_list):
    print(idx+1, "/", len(stock_name_list)) # 현재까지 진행된 상황 출력
    for year in [current_year-4, current_year-1, current_year]:
        try: # 재무제표 데이터가 잘 불러와지지 않는 경우가 있어, try - except으로 넘김
            result = find_financial_indicators(stock_name, year, indicators) # 재무지표 데이터 
            data = pd.concat([data, result], axis = 0, ignore_index = True) # data에 부착
        except:
            pass
        time.sleep(0.5)
# 중복된 행 제거
data.drop_duplicates(inplace = True)

# 숫자로 모두 변환
def str_to_float(value):
    if type(value) == float: # nan의 경우에는 문자가 아니라, 이미 float 취급됨
        return value
    elif value == '-': # -로 되어 있으면 0으로 변환
        return 0
    else:
        return float(value.replace(',', ''))

for indc in indicators:
    data[indc] = data[indc].apply(str_to_float)

data.sort_values(by = ['기업', '연도'], inplace = True) # 기업과 연도를 기준으로 정렬
data['영업이익증가율'] = data['영업이익'].diff() / data['영업이익']
data.loc[data['연도'] == 2018, '영업이익증가율'] = np.nan
# 2018년도에는 전기 정보가 없으므로 계산 불가 (계산된 것들은 다른 종목이랑 섞인 것)

data.sort_values(by = ['기업', '연도'], inplace = True) # 기업과 연도를 기준으로 정렬
data['매출액증가율'] = data['매출액'].diff() / data['매출액']
data.loc[data['연도'] == 2018, '매출액증가율'] = np.nan
# 2018년도에는 전기 정보가 없으므로 계산 불가 (계산된 것들은 다른 종목이랑 섞인 것)
data.sort_values(by = ['기업', '연도'], inplace = True) # 기업과 연도를 기준으로 정렬
data['당기순이익증가율'] = data['당기순이익'].diff() / data['당기순이익']
data.loc[data['연도'] == 2018, '당기순이익증가율'] = np.nan
# 2018년도에는 전기 정보가 없으므로 계산 불가 (계산된 것들은 다른 종목이랑 섞인 것)
data['매출액_상태'] = np.nan # 상태를 모두 결측으로 초기화

# iloc[1:]: 현재, iloc[:-1]: 과거
data['매출액_이전'] = data['매출액'].shift(1)

# 상태 열 초기화
data['매출액_상태'] = None

# 상태를 설정
data.loc[(data['매출액'] > 0) & (data['매출액_이전'] > 0), '매출액_상태'] = "흑자지속"
data.loc[(data['매출액'] <= 0) & (data['매출액_이전'] <= 0), '매출액_상태'] = "적자지속"
data.loc[(data['매출액'] > 0) & (data['매출액_이전'] <= 0), '매출액_상태'] = "흑자전환"
data.loc[(data['매출액'] <= 0) & (data['매출액_이전'] > 0), '매출액_상태'] = "적자전환"

# 이전 값 열 제거 (필요에 따라)
data.drop(columns=['매출액_이전'], inplace=True)

data.loc[data['연도'] == 2018, '매출액_상태'] = np.nan
data['ROA'] = data['당기순이익'] / data['자산총계']
average_equity = data['자본총계'].rolling(2).mean() # 평균 자기 자본
data['ROE'] = data['당기순이익'] / average_equity

data.to_excel("주요재무지표.xlsx", index = False)