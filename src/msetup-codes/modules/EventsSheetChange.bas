Attribute VB_Name = "EventsSheetChange"
Attribute VB_Description = "Events for changes in a worksheet"
Option Explicit

'@ModuleDescription("Events for changes in a worksheet")
'@IgnoreModule UndeclaredVariable, UnassignedVariableUsage, ProcedureNotUsed, VariableNotAssigned
'@Folder("Events")



'speed app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.Cursor = xlNorthWestArrow

    BusyApp

    If Me.Cells(2, 4).Value = "DISSHEET" Then
        EventsGlobal.UpdateDiseaseSheet Me, Target
    ElseIf Me.Name <> "__compRep" Then
        EventsGlobal.checkUpdateStatus Me, Target
    Else
       EventsGlobal.ComparationSheet Target
    End If

    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub


