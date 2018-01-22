import win32com.client
import time

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

instStockChart.SetInputValue(0, "A003540")
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(4, 60)
instStockChart.SetInputValue(5, 8)
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

volumes = []
numData = instStockChart.GetHeaderValue(3)
for i in range(numData):
    volume = instStockChart.GetDataValue(0, i)
    volumes.append(volume)
print(volumes)

averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) - 1)

if volumes[0] > averageVolume * 10:
    print("대박주")
else:
    print("일반주", volumes[0] / averageVolume)

print("/////////////////////")


def CheckVolumn(instStockChart, code):
    instStockChart.SetInputValue(0, code)
    instStockChart.SetInputValue(1, ord('2'))
    instStockChart.SetInputValue(4, 60)
    instStockChart.SetInputValue(5, 8)
    instStockChart.SetInputValue(6, ord('D'))
    instStockChart.SetInputValue(9, ord('1'))

    instStockChart.BlockRequest()

    volumes = []
    numData = instStockChart.GetHeaderValue(3)
    for i in range(numData):
        volume = instStockChart.GetDataValue(0, i)
        volumes.append(volume)

    averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) - 1)

    if volumes[0] > averageVolume * 10:
        return 1
    else:
        return 0


if __name__ == "__main__":
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)
    buyList = []
    for code in codeList:
        if CheckVolumn(instStockChart, code) == 1:
            buyList.append(code)
            print(code)
        time.sleep(1)
