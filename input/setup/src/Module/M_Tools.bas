Attribute VB_Name = "M_Tools"
Option Explicit

Sub ClicCmdAddRows()

    Dim oLstobj As Object
    Dim iRow1 As Integer, iRow2 As Integer, iRowHearder As Integer

    Application.EnableEvents = False
    
    ActiveSheet.Unprotect C_sPassword
    
    For Each oLstobj In ActiveSheet.ListObjects
        iRowHearder = oLstobj.DataBodyRange.Row - 1
        oLstobj.Resize Range(Cells(iRowHearder, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + iRowHearder, Cells(iRowHearder, 1).End(xlToRight).Column))
        iRow1 = oLstobj.DataBodyRange.Rows.Count - C_iNbLinesLLData + iRowHearder + 1

        iRow2 = iRow1 + C_iNbLinesLLData - 1
    Next
    
    Rows(iRow1 - 1).Copy

    Rows(iRow1 & ":" & iRow2).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Cells(iRow1, 1).Select
    
    Call ProtectSheet
    
    Application.EnableEvents = True
    
End Sub


Sub ProtectSheet()
ActiveSheet.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True

End Sub

