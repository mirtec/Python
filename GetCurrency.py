################################################
# 현금 잔고 찾기
def GetCurrency():
    account = objCpTrade.AccountNumber[0]           # 계좌번호
    accountFlag = objCpTrade.GoodsList(account, 1)  # 주식상품 구분
    #print("계정 번호 : ", account, "계좌 Flag : ", accountFlag[0])

    objStockCurrency = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
    objStockCurrency.SetInputValue(0, account)          # 계좌번호 : objCpTrade.AccountNumber[0]
    objStockCurrency.SetInputValue(1, accountFlag[0])   # 상품구분 : 주식
    #objStockCurrency.SetInputValue(2, 50)  # 요청건수 : 최대 50개

    # BlockRequest
    objStockCurrency.BlockRequest()

    if objStockCurrency.GetDibStatus() == 0:
        #print("10번 : ", objStockCurrency.GetHeaderValue(10))
        #print("45번 : ", objStockCurrency.GetHeaderValue(45))
        #print("47번 : ", objStockCurrency.GetHeaderValue(47))
        return objStockCurrency.GetHeaderValue(47)
