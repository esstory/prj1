# 3개월 이평 상승 시 매수, 매도 시 전량 매도 하는 경우

import pandas as pd
import numpy as np
import os
from pandas import Series, DataFrame
import copy
import sys
import os
from PyQt5.QtWidgets import *


class changExcel:
    def ExcelToBuffet(self, file):
        # 엑셀을 읽어 판다스로 저장
        # skiprows=2: 행 2개 버린다.
        # drop('Unnamed: 0', axis=1) -drop 라벨, axis = 1은 컬럼
        print('시작')
        df = pd.read_excel(file, sheet_name='퀀트데이타', skiprows=2).drop('Unnamed: 0', axis=1)

        print('all data')
        print(df.head())

        print ('읽기 끝')

        # 필요한 데이터만 끄집어 낸다
        # set_index (컬럼명) --> 주어진 컬럼을 인덱스로 쓰게 됨
        df = df[['회사명',
                 '시가총액\n(억)', '발표\nPBR', '과거\nPER', '과거\nPSR', '시가\n배당률\n(%)',  # 저평가
                 '과거\nROE\n(%)', '14년->17년\n3년간 YOY', '단순\n부채비율\n(%)', '배당\n성향\n(%)',  # 우량주
                 '주가\n변동성',  # 변동성
                 ]].set_index('회사명')
        print('필요한 것만 출여서 새로 만듦')
        print(df.head())

        #########################################
        # 1. 저평가 종목 구하기
        # 데이터 필터링
        # pbr, per, psr 이 0 인 데이터를 버린다.  (대박,, 이렇게 간단히 빼다니)
        df = df[df['발표\nPBR'] != 0]
        df = df[df['과거\nPER'] != 0]
        df = df[df['과거\nPSR'] != 0]

        # 새로운 컬럼도 추가
        df['1/PBR'] = 1 / df['발표\nPBR']
        df['1/PER'] = 1 / df['과거\nPER']
        df['1/PSR'] = 1 / df['과거\nPSR']


        # 순위 배당수익 - 시가배당률 컬럼 데이터의 순위를 매김
        # 순위 pbr - 1/pbr 즉 pbr 이 작은게 좋지만 값이 큰걸 랭킹 점수 매기기 좋아서, 그걸로 순위를 처리
        # 순위 per 도 마찬가지
        # 순위 psr 도 마찬가지
        # 통합 저평가 지표는 위 4가지 값의 평균을 구한다.
        df['순위 배당수익'] = df['시가\n배당률\n(%)'].rank(ascending=False)
        df['순위 PBR'] = df['1/PBR'].rank(ascending=False)
        df['순위 PER'] = df['1/PER'].rank(ascending=False)
        df['순위 PSR'] = df['1/PSR'].rank(ascending=False)
        df['통합 저평가 지표'] = df[['순위 배당수익', '순위 PBR', '순위 PER', '순위 PSR']].mean(axis=1)

        print(df.sort_values(by='통합 저평가 지표').head())


        #########################################
        # 2. 우량주 순위 구하기
        # roe/yoy/부채비율/배당성향 을 각각 순위 점수를 매긴다.
        # 부채비율은 작은게 좋은 것임으로 작은 순으로 랭킹을 매긴다 나머지는 클수록 좋은 거니 큰순으로 랭킹
        df['ROE 순위'] = df['과거\nROE\n(%)'].rank(ascending=False)
        df['영업이익 성장 순위'] = df['14년->17년\n3년간 YOY'].rank(ascending=False)
        df['부채비율 순위'] = df['단순\n부채비율\n(%)'].rank(ascending=True)
        df['배당성향 순위'] = df['배당\n성향\n(%)'].rank(ascending=False)
        df['통합 우량주 지표'] = df[['ROE 순위', '영업이익 성장 순위', '부채비율 순위', '배당성향 순위']].mean(axis=1)

        df.sort_values(by='통합 우량주 지표').head()


        #########################################
        # 3. 변동성 순위
        # 주가 변동성이 낮은 게 좋다고 보고, 그걸로 순위를 매긴다.
        df['변동성 순위'] = df['주가\n변동성'].rank(ascending=True)
        df.sort_values(by='변동성 순위').head()

        #########################################
        # 4. 통합 순위 구하기기
        df['통합순위'] = df['통합 저평가 지표'] * 0.4 + df['통합 우량주 지표'] * 0.4 + df['변동성 순위'] * 0.2
        pd.set_option('display.max_columns', None)
        pd.set_option('display.expand_frame_repr', False)
        pd.set_option('max_colwidth', -1)
        df = df.sort_values(by='통합순위')


        outFile = os.path.dirname(file)
        ofilename = 'result_저평가'  + '.xlsx'
        outFile += '\\' + ofilename

        writer = pd.ExcelWriter(outFile, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='spread')

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(outFile)

