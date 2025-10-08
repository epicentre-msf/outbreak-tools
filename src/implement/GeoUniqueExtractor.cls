VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoUniqueExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Transforms ranges into BetterArray collections of unique values"

'@Folder("Implement")
'@ModuleDescription("Extracts unique values from ranges with trimming and blank control")

Option Explicit

Public Function Extract(ByVal sourceRange As Range, Optional ByVal ignoreBlanks As Boolean = True) As BetterArray
    Dim result As BetterArray
    Dim cell As Range
    Dim valueText As String

    If sourceRange Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, "GeoUniqueExtractor", "Source range is required"
    End If

    Set result = New BetterArray
    result.LowerBound = 1

    For Each cell In sourceRange
        valueText = Trim$(CStr(cell.Value))
        If (ignoreBlanks = False) Or LenB(valueText) > 0 Then
            If Not result.Includes(valueText) Then
                result.Push valueText
            End If
        End If
    Next cell

    Set Extract = result.Clone
End Function

End Class
