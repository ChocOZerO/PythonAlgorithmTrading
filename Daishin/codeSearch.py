import win32com.client

instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
print(instCpStockCode.GetCount())
print('//////////////////////////////')
print(instCpStockCode.GetData(1, 0))
print('//////////////////////////////')
print(instCpStockCode.GetData(0, 0))
print('//////////////////////////////')
for i in range(0, 10):
    print(instCpStockCode.GetData(1, i))

print('//////////////////////////////')
stockNum = instCpStockCode.GetCount()
for i in range(stockNum):
    if instCpStockCode.GetData(1, i) == 'NAVER':
        print(instCpStockCode.GetData(0, i))
        print(instCpStockCode.GetData(1, i))
        print(i)
print('//////////////////////////////')
naverCode = instCpStockCode.NameToCode('NAVER')
naverIndex = instCpStockCode.CodeToIndex(naverCode)
print(naverCode)
print(naverIndex)
