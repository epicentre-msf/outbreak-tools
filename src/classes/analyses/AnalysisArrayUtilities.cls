VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisArrayUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Shared helpers for cloning and creating BetterArray buffers used across analysis outputs")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Const CLASSNAME As String = "AnalysisArrayUtilities"

'@section Buffer helpers
'===============================================================================
'Create a BetterArray configured with a LowerBound of 1 so callers can push items safely.
'@return BetterArray   A fresh buffer instance.
Public Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateBuffer = buffer
End Function

'Clone every element from the source array without creating object copies (shallow copy).
'@param source    Source BetterArray that may be Nothing.
'@return BetterArray containing the same references/values, or an empty buffer when source is Nothing/empty.
Public Function CloneItems(ByVal source As BetterArray) As BetterArray
    Set CloneItems = CloneInternal(source, False)
End Function

'Clone every element from the source array by invoking a Clone method on each object.
'@param source    Source BetterArray that may be Nothing.
'@return BetterArray populated with cloned objects; empty buffer when source is Nothing/empty.
'@throws ProjectError.ErrorUnexpectedState when an item does not expose a Clone method.
Public Function CloneCloneable(ByVal source As BetterArray) As BetterArray
    Set CloneCloneable = CloneInternal(source, True)
End Function

'@section Internal helpers
'===============================================================================
Private Function CloneInternal(ByVal source As BetterArray, ByVal requireClone As Boolean) As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim item As Variant
    Dim objectItem As Object

    Set result = CreateBuffer()

    If source Is Nothing Then
        Set CloneInternal = result
        Exit Function
    End If

    If source.Length = 0 Then
        Set CloneInternal = result
        Exit Function
    End If

    For idx = source.LowerBound To source.UpperBound
        item = source.Item(idx)

        If IsObject(item) Then
            Set objectItem = item

            If objectItem Is Nothing Then
                If Not requireClone Then
                    result.Push objectItem
                End If
            ElseIf requireClone Then
                result.Push CloneObject(objectItem)
            Else
                result.Push objectItem
            End If
        Else
            result.Push item
        End If

        Set objectItem = Nothing
    Next idx

    Set CloneInternal = result
End Function

Private Function CloneObject(ByVal candidate As Object) As Object
    On Error GoTo MissingClone
        Set CloneObject = CallByName(candidate, "Clone", VbMethod)
        Exit Function

MissingClone:
    RaiseError CLng(ProjectError.ErrorUnexpectedState), _
              "Array utilities expected items to expose a Clone method but one was missing for type '" & TypeName(candidate) & "'"
End Function

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub
