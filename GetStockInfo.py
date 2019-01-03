################################################
# 주식 현재 가격 리턴
def GetStockInfo(stock_code):
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, stock_code)  # 종목 코드 
    objStockMst.BlockRequest()

    # 현재가 통신 및 통신 에러 처리
    if objStockMst.GetDibStatus() == 0:
        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)        # 종목코드
        name = objStockMst.GetHeaderValue(1)        # 종목명
        time = objStockMst.GetHeaderValue(4)        # 시간
        cprice = objStockMst.GetHeaderValue(11)     # 종가
        diff = objStockMst.GetHeaderValue(12)       # 대비
        open = objStockMst.GetHeaderValue(13)       # 시가
        high = objStockMst.GetHeaderValue(14)       # 고가
        low = objStockMst.GetHeaderValue(15)        # 저가
        offer = objStockMst.GetHeaderValue(16)      # 매도호가
        bid = objStockMst.GetHeaderValue(17)        # 매수호가
        vol = objStockMst.GetHeaderValue(18)        # 거래량
        vol_value = objStockMst.GetHeaderValue(19)  # 거래대금

        #print("코드", code)
        #print("이름", name)
        #print("시간", time)
        #print("종가", cprice)
        #print("대비", diff)
        #print("시가", open)
        #print("고가", high)
        #print("저가", low)
        #print("매도호가", offer)
        #print("매수호가", bid)
        #print("거래량", vol)
        #print("거래대금", vol_value)
        #print("------------------------------------------------------------------------------------")
        #print("[ ", code, " : ", name, " ] 시간 : ", time, " / 종가 : ",cprice)

        return name, cprice
