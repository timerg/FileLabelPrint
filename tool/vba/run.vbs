Option Explicit

ExcelMacroExample

Sub ExcelMacroExample() 

  Dim xlApp 
  Dim xlBook 
  Dim objShell
  Dim strPath
  
  Set xlApp = CreateObject("Excel.Application") 
  xlApp.Visible = True
  
  Set objShell = CreateObject("Wscript.Shell")
  strPath = objShell.CurrentDirectory
  Set xlBook = xlApp.Workbooks.Open(strPath & "\tool\vba\Main.xls", 0, False) 
  xlApp.Run "Clear"
  xlApp.Run "ParseContentTxt"
  xlApp.Run "Main"

  'xlBook.Close False
  'xlApp.Quit
  Set xlBook = Nothing 
  Set xlApp = Nothing 
  Set ObjShell = Nothing
End Sub 