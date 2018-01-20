class CpStockCode:
    def __init__(self):
        self.stocks = {'유한양행':'A000100'}

    def GetCount(self):
        return len(self.stocks)

    def NameToCode(self, name):
        return self.stocks[name]


instCpStockCode = CpStockCode()

print(instCpStockCode.GetCount())
print(instCpStockCode.NameToCode('유한양행'))
