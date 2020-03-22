# call_vba

vbsファイルをダブルクリックすることで、同名のxlsmファイルのVBAを実行する。<br/>

```VBScript
FilePath = Replace(WScript.ScriptFullName, ".vbs", ".xlsm")
Const FunctionName = "main"

With WScript.CreateObject("Excel.Application")
    .Visible = True 'Trueなら実行時にExcelの画面を表示
    .Workbooks.Open FilePath
    .Application.Run FunctionName
    .Quit
End With
```
例：book.xlsmのmainプロシージャor関数を呼び出す。<br/>
このvbsファイルを実行したいxlsmファイルと同じフォルダに、xlsmファイルと同じ名前で保存(この例では book.vbs )する。<br/>
このvbsファイルをダブルクリックすれば対象のプロシージャor関数が呼び出される。
<br/>

補足："WScript.ScriptFullName"で自分(book.vbs)のファイルパスを取得し.vbsを.xlsmに置換することで、呼び出すxlsmファイルのパスを作成している。
