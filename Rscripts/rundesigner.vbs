
'First write everything from a vbscript
Dim Arg, DesPath, GeoPath, SetupPath, LLDir, LLName, SetupLang, LLLang

Set Arg = WScript.Arguments

DesPath    = Arg(0)
GeoPath    = Arg(1)
SetupPath  = Arg(2)
LLDir      = Arg(3)
LLName     = Arg(4)
SetupLang  = Arg(5)
LLLang = Arg(6)

Set xlsApp = CreateObject("Excel.Application")
xlsApp.visible = True
xlsApp.DisplayAlerts = False
xlsApp.ScreenUpdating = False

Set Wkb = xlsApp.Workbooks.Open(DesPath)

'Setting up parameters for the designer
Wkb.Worksheets("Main").Range("RNG_PathDico").value  = SetupPath
Wkb.Worksheets("Main").Range("RNG_PathGeo").value   = GeoPath
Wkb.Worksheets("Main").Range("RNG_LLDir").value     = LLDir
Wkb.Worksheets("Main").Range("RNG_LLName").value    = LLName
Wkb.Worksheets("Main").Range("RNG_LangSetup").value = SetupLang
Wkb.Worksheets("Main").Range("RNG_LLForm").value = LLLang

'Generate linelist data
xlsApp.Run  Wkb.Name & "!" & "GenerateData"

Wkb.Close False
xlsApp.Quit

Set Wkb = Nothing
Set xlsApp = Nothing

'Set Arg and xlsapp to Nothing
Set Arg = Nothing