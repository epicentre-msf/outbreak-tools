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
Public Sub DeleteLoDataBodyRange(Lo As listObject)
    If Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
End Sub


'Test if a listobject exists
Public Function ListObjectExists(Wksh As Worksheet, sListObjectName As String) As Boolean
    ListObjectExists = False
    Dim Lo As listObject
    On Error Resume Next
    Set Lo = Wksh.ListObjects(sListObjectName)
    ListObjectExists = (Not Lo Is Nothing)
    On Error GoTo 0
End Function


'Unique Values of a BetterArray
Function GetUniqueBA(inputTable As BetterArray, _
                     Optional ByVal Sort As Boolean = False) As BetterArray
    Dim tableValue As String
    Dim counter As Long
    Dim Outable As BetterArray

    Set Outable = New BetterArray
    Outable.LowerBound = 1

    If inputTable.Length > 0 Then
        For counter = inputTable.LowerBound To inputTable.UpperBound

            tableValue = Application.WorksheetFunction.Trim( _
                                inputTable.Item(counter) _
                        )

            If (tableValue <> vbNullString) And _
               (Not Outable.Includes(tableValue)) Then _
                Outable.Push tableValue
        Next
    End If

    'sort
    If Sort Then Outable.Sort
    Set GetUniqueBA = Outable.Clone()
End Function
