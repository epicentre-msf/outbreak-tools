Attribute VB_Name = "Helpers"
Option Private Module

'Basic Helper functions used in the creation of the linelist and other stuffs
'Most of them are explicit functions. Contains all the ancillary sub/
'Functions used when creating the linelist and also in the linelist
'itself

'@Obsolete, will be removed in future releases.

Option Explicit

'Speed up before a work
Public Sub BeginWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = False
    xlsapp.DisplayAlerts = False
    xlsapp.Calculation = xlCalculationManual
    xlsapp.EnableAnimations = False
End Sub

'Return previous state
Public Sub EndWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = True
    xlsapp.DisplayAlerts = True
    xlsapp.EnableAnimations = True
    xlsapp.DisplayStatusBar = bstatusbar
End Sub


'Safely delete databodyrange of a listobject
Public Sub DeleteLoDataBodyRange(Lo As ListObject)
    If Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
End Sub


'Test if a listobject exists
Public Function ListObjectExists(Wksh As Worksheet, sListObjectName As String) As Boolean
    ListObjectExists = False
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = Wksh.ListObjects(sListObjectName)
    ListObjectExists = (Not Lo Is Nothing)
    On Error GoTo 0
End Function


'Unique of a betteray sorted
Function GetUniqueBA(BA As BetterArray) As BetterArray
    Dim sVal As String
    Dim i As Long
    Dim Outable As BetterArray

    BA.Sort

    Set Outable = New BetterArray
    Outable.LowerBound = 1

    sVal = Application.WorksheetFunction.Trim(BA.Item(BA.LowerBound))

    If sVal <> vbNullString Then
        Outable.Push sVal
    End If

    If BA.Length > 0 Then
        For i = BA.LowerBound To BA.UpperBound
            If sVal <> Application.WorksheetFunction.Trim(BA.Item(i)) And Application.WorksheetFunction.Trim(BA.Item(i)) <> vbNullString Then
                sVal = Application.WorksheetFunction.Trim(BA.Item(i))
                Outable.Push sVal
            End If
        Next
    End If

    Set GetUniqueBA = Outable.Clone()

End Function
