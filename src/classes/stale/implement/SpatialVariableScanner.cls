VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpatialVariableScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Parses spatial configuration worksheets to discover variables"

'@Folder("Implement")
'@ModuleDescription("Scans HList worksheets to identify spatial variables without mutations")

Option Explicit

'@section State
'===============================================================================

Private Type TSpatialVariableScanner
    configSheet As Worksheet
End Type

Private this As TSpatialVariableScanner
Private Const CLASSNAME As String = "SpatialVariableScanner"

'@section Factory
'===============================================================================

Public Function Create(ByVal configSheet As Worksheet) As SpatialVariableScanner
    Dim scanner As SpatialVariableScanner

    If configSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Configuration worksheet is required"
    End If

    Set scanner = New SpatialVariableScanner
    Set scanner.ConfigSheet = configSheet
    Set Create = scanner
End Function

'@section Dependencies
'===============================================================================

Public Property Set ConfigSheet(ByVal sh As Worksheet)
    Set this.configSheet = sh
End Property

Public Property Get ConfigSheet() As Worksheet
    Set ConfigSheet = this.configSheet
End Property

'@section Public API
'===============================================================================

Public Function ScanVariables() As BetterArray
    Dim variables As BetterArray
    Dim tableObject As ListObject
    Dim dataRange As Range
    Dim rowRange As Range
    Dim valueText As String

    Set variables = New BetterArray
    variables.LowerBound = 1

    On Error Resume Next
        Set tableObject = ConfigSheet.ListObjects("listofgeovars")
    On Error GoTo 0

    If tableObject Is Nothing Then
        Set ScanVariables = variables.Clone
        Exit Function
    End If

    Set dataRange = tableObject.DataBodyRange
    If dataRange Is Nothing Then
        Set ScanVariables = variables.Clone
        Exit Function
    End If

    For Each rowRange In dataRange.Rows
        valueText = Trim$(CStr(rowRange.Cells(1, 1).Value))
        If LenB(valueText) > 0 Then
            If Not variables.Includes(valueText) Then
                variables.Push valueText
            End If
        End If
    Next rowRange

    Set ScanVariables = variables.Clone
End Function

End Class
