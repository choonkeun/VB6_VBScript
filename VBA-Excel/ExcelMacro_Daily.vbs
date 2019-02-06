'---C:\> cscript.exe "C:\Daily\ExcelMacro_Daily.vbs"

Option Explicit
On Error Resume Next
    Dim sheetName

    ExcelMacro_Daily
    WScript.Echo "Daily is Finished"
    '--WScript.Echo sheetName

Sub ExcelMacro_Daily() 

    Dim xlApp 
    Dim xlBook 

    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Open("C:\Daily\Portal Counts.xlsm", 0, True) 
    xlApp.Run "Module1.BeginStep1"

    sheetName = xlBook.ActiveSheet.Name
    xlbook.SaveAs "C:\Daily\Portal Counts " + sheetName + ".xlsm"
    xlBook.Close False
    
    xlApp.Quit 
    Set xlBook = Nothing 
    Set xlApp = Nothing 
End Sub


