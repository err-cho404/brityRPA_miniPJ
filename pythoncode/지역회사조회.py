import requests
from io import BytesIO
import zipfile
import xmltodict
import json
#import dart_fss as dart
import OpenDartReader
import pandas as pd
import numpy as np
from xml.etree.ElementTree import parse
import time
import datetime
local=input('출력하고자하는 기업들의 지역을 입력하시오 : ')
corp_detail=pd.read_excel('회사상세정보.xlsx')
corp_local=corp_detail[corp_detail['adres'].str.contains(local)]
corp_local.to_excel(local+'재무제표결과.xlsx')
data = pd.read_excel('주요재무지표.xlsx')

merged_df = pd.merge(corp_local, data,left_on='corp_name', right_on='기업')

columns_to_extract = [
    '기업', '연도', '자산총계', '부채총계', '자본총계', '매출액',
    '영업이익', '당기순이익', '부채비율', '영업이익증가율',
    '매출액증가율', '당기순이익증가율', '매출액_상태', 'ROA', 'ROE'
]

# 추출할 열이 모두 존재하는지 확인하고 필터링
columns_to_extract = [col for col in columns_to_extract if col in merged_df.columns]
result = merged_df[columns_to_extract]

# 결과를 엑셀 파일로 저장 (선택 사항)
result.to_excel(local+'재무제표결과.xlsx', index=False,sheet_name = local)

print(result)

# 비교할 엑셀 파일들을 불러옵니다.
df_detail= pd.read_excel('회사상세정보.xlsx') 
df_localresult = pd.read_excel(local+'재무제표결과.xlsx')
df_localresult['산업분류코드']='0'
for i in range(len(df_localresult['기업'])) :
    for j in range(len(df_detail['stock_name'])):
        if df_localresult['기업'][i] in df_detail['stock_name'][j]:
            df_localresult['산업분류코드'][i]=df_detail['induty_code'][j]

df_localresult.to_excel(local+'재무제표 induty_code 추가 결과.xlsx', index=False,sheet_name=local)
#기존 가져오는 값에 업종명이 없고 업종 코드만 존재
#rpa로 업종명과 업종코드를 매칭한 자료를 만들고 rpa에서 산업분류코드와 업종명을 매칭해 작업할 것임ㅍ ㅍ