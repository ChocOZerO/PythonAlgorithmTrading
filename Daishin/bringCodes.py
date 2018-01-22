import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = instCpCodeMgr.GetStockListByMarket(1)
print(codeList)

kospi = {}
for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    kospi[code] = name

f = open('d:\\IT\\ProgrammingStudy\\PythonAlgorithmTrading\\Daishin\\kospi.csv', 'w')
for key, value in kospi.items():
    f.write("%s,%s\n" % (key, value))
f.close()

f = open('d:\\IT\\ProgrammingStudy\\PythonAlgorithmTrading\\Daishin\\kospi_detail.csv', 'w')
for i, code in enumerate(codeList):
    secondCode = instCpCodeMgr.GetStockSectionKind(code)
    name = instCpCodeMgr.CodeToName(code)
    f.write("%s,%s,%s,%s\n" % (i, code, secondCode, name))
f.close()
