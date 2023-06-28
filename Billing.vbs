Set Billing = CreateObject("Excel.Application")
Billing.DisplayAlerts = False
Billing.AlertBeforeOverwriting = False
Billing.Workbooks.Open(Replace(WScript.ScriptFullName,"vbs","xlsm"))
Billing.Quit