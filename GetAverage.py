################################################
# 이동 평균 구하기
def GetAverage(stock_code):
    objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    if objStockChart.GetDibStatus() == 0:
        #print(objStockChart.GetDibStatus())

        # 조회할 주식 변수 할당
        objStockChart.SetInputValue(0, stock_code)          # 종목코드
        objStockChart.SetInputValue(1, ord('2'))            # 1 : 기간, 2 : 개수
        objStockChart.SetInputValue(4, 15)                  # 최근 15개
        objStockChart.SetInputValue(5, (5))                 # 0 : 날짜, 1 : 시간, 5 : 종가, 8 : 거래량, 9 : 거래대금
        objStockChart.SetInputValue(6, ord('m'))            # m : 분, D : 일, W : 주, M : 월
        objStockChart.SetInputValue(9, ord('1'))            # 1 : 수정주가, 0 :  무수정주가

        # BlockRequest
        objStockChart.BlockRequest()

        # GetHeaderValue
        numData = objStockChart.GetHeaderValue(3)
        numField = objStockChart.GetHeaderValue(1)

        close_price_list = []

        # GetDataValue
        for i in range(numData):
            for j in range(numField):
                close_price_list.append(objStockChart.GetDataValue(j, i))
                #print(objStockChart.GetDataValue(j, i))

        avg_min_5 = sum(close_price_list[:5]) / 5
        avg_min_15 = sum(close_price_list[:15]) / 15
        #print("5분 평균 : ", avg_min_5)
        #print("15분 평균 : ", avg_min_15)
        return avg_min_5, avg_min_15
