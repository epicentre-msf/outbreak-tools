Attribute VB_Name = "DesignerTests"
Option Explicit



'Test of validations formulas
Sub TestValidation()

    Dim sFormula As String
    Dim iSheetStartLine As Integer
    Dim VarNameData As New BetterArray
    Dim ColumnIndexData As New BetterArray
    Dim IsValidation As Boolean
    Dim FormulaData As New BetterArray
    Dim SpecCharData As New BetterArray

    FormulaData.FromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange SheetFormulas.ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False

    VarNameData.Push "date_notification", "var2", "deceased"
    ColumnIndexData.Push 5, 5, 3
    iSheetStartLine = 1
    sFormula = "IF(ISBLANK(date_notification)," & Chr(34) & Chr(34) & ",EPIWEEK(date_notification))"
    IsValidation = False

    Debug.Print ValidationFormula(sFormula, VarNameData, ColumnIndexData, FormulaData, SpecCharData)

End Sub


'Test of Opening the Control for Generate with different scenario

Sub test()
 LoadFile ("*.xlsx")
End Sub

'Test for the
