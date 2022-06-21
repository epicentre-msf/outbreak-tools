VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Dim vPage As Variant
    
    Reponse = 0
    
    On Error GoTo ManageError

    If bUpdate Then Call Translate_Manage("Close")

    If Reponse = vbNo Then
        Cancel = True 'cancel closing the workbook
        Exit Sub
    End If
    
'removal of ranges for columns to be translated
    For Each vPage In ActiveWorkbook.Names
        If Left(vPage.Name, 8) = "T_Trans_" Then ActiveWorkbook.Names(vPage.Name).Delete
    Next
    
    ActiveWorkbook.Save
    
ManageError:
    Resume Next
    
End Sub

Private Sub Workbook_Open()

    Dim iCol As Integer, iLastRow As Integer, iRowStart As Integer, i As Integer
    Dim iCptCol As Integer, iCptSheet As Integer
    Dim arrColumn() As String
    Dim SheetActive As Worksheet
    
    Application.ScreenUpdating = False
    
    iLastRow = sheetTranslation.[Tab_Translations].Rows.Count + sheetTranslation.[T_Translate_Start].Row
    
    i = 1
    
    For iCptSheet = 1 To 3

        Select Case iCptSheet
            Case 1
                arrColumn = Split(sCstColDictionary, "|")
                Set SheetActive = sheetDictionary
                iRowStart = SheetActive.[T_Dictionary_Start].Row + 1
                iLastRow = SheetActive.[Tab_Dictionary].Rows.Count + 2
            Case 2
                arrColumn = Split(sCstColChoices, "|")
                Set SheetActive = SheetChoice
                iRowStart = SheetActive.[T_Choices_Start].Row + 1
                iLastRow = SheetActive.[Tab_Choices].Rows.Count + 1
            Case 3
                arrColumn = Split(sCstColExport, "|")
                Set SheetActive = sheetExport
                iRowStart = SheetActive.[T_Export_Start].Row + 1
                iLastRow = SheetActive.[Tab_Export].Rows.Count + 1
        End Select

        SheetActive.Select
        
    'level column
        For iCptCol = LBound(arrColumn, 1) To UBound(arrColumn, 1)
        
            If Not SheetActive.Rows(iRowStart - 1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole) Is Nothing Then _
            iCol = SheetActive.Rows(iRowStart - 1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole).Column
            
            SheetActive.Range(Cells(iRowStart, iCol), Cells(iLastRow, iCol)).Name = "T_Trans_" & i & Left(ActiveSheet.Name, 3)
            
            i = i + 1
            
        Next iCptCol

    Next iCptSheet

    sheetDictionary.Select
    sheetDictionary.Range("A3").Activate

    Application.ScreenUpdating = True
    
End Sub
