################################################
# 주식 Code 리턴
def CodeList(code_num):     # code_num : 1 = 코스피, 2 = 코스닥
    codelist = objCpCodeMgr.GetStockListByMarket(code_num)

    for i, code in enumerate(codelist):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        print(i, code, secondCode, stdPrice, name)
