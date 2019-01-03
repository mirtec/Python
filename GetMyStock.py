################################################
# 주식 잔고 조회
def GetMyStock(stock_code):
    account = objCpTrade.AccountNumber[0]           # 계좌번호
    accountFlag = objCpTrade.GoodsList(account, 1)  # 주식상품 구분
    #print("계정 번호 : ", account, "계좌 Flag : ", accountFlag[0])

    objMyStock = win32com.client.Dispatch('CpTrade.CpTd6033')
    objMyStock.SetInputValue(0, account)            # 계좌번호
    objMyStock.SetInputValue(1, accountFlag[0])     # 상품구분 : 주식
    objMyStock.SetInputValue(2, 50)                 # 요청건스 (최대 50개)

    # BlockRequest
    objMyStock.BlockRequest()

    if objMyStock.GetDibStatus() == 0:
        cnt = objMyStock.GetHeaderValue(7)
        print("cnt : ", cnt)

        if (cnt == 0 or cnt > 100):
            return "None", "None"
        else:
            for i in range(cnt):
                code = objMyStock.GetDataValue(12, i)       # 종목코드
                name = objMyStock.GetDataValue(0, i)        # 종목명
                amount = objMyStock.GetDataValue(7, i)      # 체결잔고수량
                sell = objMyStock.GetDataValue(15, i)       # 매도가능수량
                buyPrice = objMyStock.GetDataValue(17, i)   # 체결장부단가

                print("결재장보단가 : ", objMyStock.GetDataValue(18, i) )
                print("[", code, "] ", name,  ":", "체결잔고수량 : " , amount, " 매도가능수량 : " , sell, " 체결장부단가 : " , buyPrice)
                return sell, buyPrice

                #if stock_code == code:
                #    print(" 존재함 [", code, "] ", name, ":", "체결잔고수량 : ", amount, " 매도가능수량 : ", sell, " 체결장부단가 : ", buyPrice)
                #    return sell, buyPrice
    else:
        return "None", "None"
