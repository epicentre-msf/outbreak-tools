Attribute VB_Name = "TableSpecsPolicyHelpers"
Option Explicit

'@Folder("Analysis")
'@ModuleDescription("Shared utilities and constants for TableSpecs policy strategy objects")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Public Const TABLE_TYPE_GLOBAL_SUMMARY As Byte = 1
Public Const TABLE_TYPE_UNIVARIATE As Byte = 2
Public Const TABLE_TYPE_BIVARIATE As Byte = 3
Public Const TABLE_TYPE_TIME_SERIES As Byte = 4
Public Const TABLE_TYPE_SPATIAL As Byte = 5
Public Const TABLE_TYPE_SPATIO_TEMPORAL As Byte = 6

'@section Normalisation helpers
'===============================================================================

'Normalise a value for case-insensitive comparisons.
Public Function NormalizeValue(ByVal valueText As String) As String
    NormalizeValue = LCase$(Trim$(valueText))
End Function

'Compare two values after normalising their casing and spacing.
Public Function ValueEquals(ByVal leftValue As String, ByVal rightValue As String) As Boolean
    ValueEquals = (NormalizeValue(leftValue) = NormalizeValue(rightValue))
End Function

'Determine whether a value matches any candidate when compared case-insensitively.
Public Function ValueInList(ByVal target As String, ParamArray candidates() As Variant) As Boolean
    Dim idx As Long
    Dim candidate As Variant

    ValueInList = False

    For idx = LBound(candidates) To UBound(candidates)
        candidate = candidates(idx)
        If ValueEquals(target, CStr(candidate)) Then
            ValueInList = True
            Exit Function
        End If
    Next idx
End Function

'Guard against empty strings when a value must be present.
Public Function HasText(ByVal valueText As String) As Boolean
    HasText = (NormalizeValue(valueText) <> vbNullString)
End Function
