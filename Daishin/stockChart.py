import win32com.client
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

instStockChart.SetInputValue(0, "A003540")
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(4, 10)
instStockChart.SetInputValue(5, 5)
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

numData = instStockChart.GetHeaderValue(3)
for i in range(numData):
    print(instStockChart.GetDataValue(0, i))

print("////////////////////")

instStockChart.SetInputValue(0, "A003540")
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(4, 10)
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8))
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)

for i in range(numData):
    for j in range(numField):
        print(instStockChart.GetDataValue(j, i), end=" ")
    print("")

print("////////////////////")

instStockChart.SetInputValue(0, "A003540")
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, 20161031)
instStockChart.SetInputValue(3, 20161020)
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8))
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)

for i in range(numData):
    for j in range(numField):
        print(instStockChart.GetDataValue(j, i), end=" ")
    print("")
