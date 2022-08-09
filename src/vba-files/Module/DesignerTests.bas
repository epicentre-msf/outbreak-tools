Attribute VB_Name = "DesignerTests"
Option Explicit
Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub

'Test for the Formulas

Sub TestBuildFormula()

    'Testing without filters
    Debug.Print BuildVariateFormula("table1", "outcome")
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="univariate")
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="univariate", sFirstCondVar:="outcome", sFirstCondVal:=Cells(1, 1).Address)
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="bivariate", sFirstCondVar:="outcome", sFirstCondVal:="dead")
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="bivariate", sFirstCondVar:="outcome", sFirstCondVal:="dead", sSecondCondVar:="sex", sSecondCondVal:="Male")

    'Testing with filters
    Debug.Print BuildVariateFormula("table1", "outcome", isFiltered:=True)
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="univariate", isFiltered:=True)
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="univariate", sFirstCondVar:="outcome", sFirstCondVal:="dead", isFiltered:=True)
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="bivariate", sFirstCondVar:="outcome", sFirstCondVal:="dead", isFiltered:=True)
    Debug.Print BuildVariateFormula("table1", "outcome", sVariate:="bivariate", sFirstCondVar:="outcome", sFirstCondVal:="dead", _
                                sSecondCondVar:="sex", sSecondCondVal:="Male", isFiltered:=True)

End Sub



Sub TestAnalysisFormula()

    'Analysis Formula

    Debug.Print AnalysisFormula("COUNT(outcome)", ThisWorkbook, isFiltered:=True, sVariate:="univariate", sFirstCondVar:="outcome", sFirstCondVal:="dead")



End Sub


'Testing the case_wen

Sub TestCaseWhen()

    Debug.Print ParseCaseWhen(ThisWorkbook.Worksheets("Test").Range("A1").value)

End Sub

Sub TestFindFistDay()
Debug.Print Format(FindFirstDay("Quarter", 44780), "Long Date")

End Sub


Sub TestAutoFill()
  Dim DictHeaders As BetterArray
  Dim sForm As String
  Dim sTimeVar As String
  
  Set DictHeaders = GetHeaders(ThisWorkbook, C_sParamSheetDict, 1)
  sForm = "COUNTA(case_id)"
  sTimeVar = "date_notification"
  
  
  Debug.Print TimeSeriesCount(ThisWorkbook, DictHeaders, sTimeVar, "2022", "2023", _
                        isFiltered:=True)
End Sub





