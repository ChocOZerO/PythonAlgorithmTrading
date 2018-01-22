import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
industryCodeList = instCpCodeMgr.GetIndustryList()

for industryCode in industryCodeList:
    print(industryCode, instCpCodeMgr.GetIndustryName(industryCode))

print("/////////////////")

targetCodeList = instCpCodeMgr.GetGroupCodeList(5)

for code in targetCodeList:
    print(code, instCpCodeMgr.CodeToName(code))

print("/////////////////")

instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
instMarketEye.SetInputValue(0, 67)
instMarketEye.SetInputValue(1, targetCodeList)

instMarketEye.BlockRequest()

numStock = instMarketEye.GetHeaderValue(2)

sumPer = 0
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(0, i)

print("Average PER: ", sumPer / numStock)
