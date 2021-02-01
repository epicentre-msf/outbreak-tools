Attribute VB_Name = "Helpers"
Option Explicit

Dim ProcessNbr As Integer


Sub DisplayErr(Err As ErrObject)
    If Err.Number <> 0 Then MsgBox Err.Description, , Err.Source + " Error", Err.HelpFile, Err.HelpContext
End Sub

Sub ProcessReset()
   Application.Cursor = xlDefault
   Application.Interactive = True
   Application.DisplayAlerts = True
   Application.CutCopyMode = False
   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.Calculate
   Application.Calculation = xlCalculationAutomatic
   ProcessNbr = 0
End Sub

Sub ProcessBegin()
    On Error GoTo ErrorHandler
    If ProcessNbr = 0 Then
        
        #If Not InDevelopment Then
            Application.Cursor = xlWait
        #End If
        
        Application.Calculation = xlCalculationManual
        Application.CutCopyMode = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
    End If
    ProcessNbr = ProcessNbr + 1
    Exit Sub
ErrorHandler:
    DisplayErr Err
End Sub

Sub ProcessEnding()
    On Error GoTo ErrorHandler
    If ProcessNbr = 1 Then
        
        #If Not InDevelopment Then
            Application.Cursor = xlDefault
        #End If
        
        Application.Interactive = True
        Application.DisplayAlerts = True
        Application.CutCopyMode = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        'Application.Calculate
        'Application.Calculation = xlCalculationAutomatic
    End If
    ProcessNbr = ProcessNbr - 1
    Exit Sub
ErrorHandler:
    DisplayErr Err
End Sub
