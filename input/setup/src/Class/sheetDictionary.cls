VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim vPage As Variant
    
    If bUpdate Then Exit Sub
    
    If Not Intersect(Target, sheetDictionary.Range("Tab_Dictionary")) Is Nothing Then
    
        'note the changes in the columns to be translated
        For Each vPage In ActiveWorkbook.Names
            If Right(vPage.Name, 3) = Left(ActiveSheet.Name, 3) Then
                If Not Intersect(Target, sheetDictionary.Range(vPage.Name)) Is Nothing Then
                    bUpdate = True
                    Exit Sub
                End If
            End If
        Next

    End If

End Sub
