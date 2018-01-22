import win32com.client

instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

instMarketEye.setInputValue(0, (4, 67, 70, 111))
instMarketEye.SetInputValue(1, 'A003540')

instMarketEye.BlockRequest()

print("현재가: ", instMarketEye.GetDataValue(0, 0))
print("PER: ", instMarketEye.GetDataValue(1, 0))
print("EPS: ", instMarketEye.GetDataValue(2, 0))
print("최근분기년월: ", instMarketEye.GetDataValue(3, 0))
