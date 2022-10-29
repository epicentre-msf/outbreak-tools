Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("E15").Select
    Range("K1").Value = 1
    Range("M1").Value = 2
    With Range("E15").Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=K1", Formula2:="=M1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = "Error"
        .InputMessage = ""
        .errorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("K1").Value = ""
    Range("M1").Value = ""
End Sub
