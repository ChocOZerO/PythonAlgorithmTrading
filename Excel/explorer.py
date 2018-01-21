import win32com.client

explore = win32com.client.Dispatch("InternetExplorer.Application")
explore.Visible = True
