FilePath = Replace(WScript.ScriptFullName, ".vbs", ".xlsm")
Const FunctionName = "main"

With WScript.CreateObject("Excel.Application")
    .Visible = True 'TrueならExcelの画面を表示
    .Workbooks.Open FilePath
    .Application.Run FunctionName
    .Quit
End With
