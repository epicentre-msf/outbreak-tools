Attribute VB_Name = "DesignerTranslation"
Option Explicit

'Translate one shape using informations on languages
Sub TranslateShape(oShape As Object, sValue As String)
    Dim bVis As Integer                          'visibility of the shape
    Dim sFont As String                          'actual font of the shape
    bVis = oShape.Visible
    'be sure the shape is visible before updating its text
    oShape.Visible = msoTrue
    With SheetMain.Shapes(oShape.Name)
        'keeping the previous font selected
        sFont = .TextFrame.Characters.Font.Name
        .TextFrame.Characters.Text = sValue
        .TextFrame.Characters.Font.Name = sFont
    End With
    oShape.Visible = bVis

End Sub

Sub TranslateRange(rngCode As String, rngValue As String)
    SheetMain.Range(rngCode).value = rngValue
End Sub

Sub TranslateDesigner()

    Dim oShape As Object
    Dim i As Integer                             'variable used for indexation
    Dim k As Integer
    Dim T_data As BetterArray                    'temporary array to get translation of values in shapes and Messages and languages
    Dim T_values As BetterArray                  'Array of headings or values of languages)
    Dim T_codes As BetterArray                   'array of values of languages, shapes codes or ranges codes
    Dim sString As String                        'A string used as a temporary variable

    'Initializing
    Set T_data = New BetterArray
    T_data.LowerBound = 1
    Set T_values = New BetterArray
    T_values.LowerBound = 1
    Set T_codes = New BetterArray
    T_codes.LowerBound = 1

    'First Get the language code
    sString = GetLanguageCode([RNG_LangDesigner].value)

    T_codes.Clear
    T_values.Clear
    T_data.Clear

    'Now check if the language code is not empty before moving foward
    If sString <> "" Then
        Application.ScreenUpdating = False

        'First Get done the Translation of Shapes ----------------------------------------
        T_data.FromExcelRange SheetDesTranslation.ListObjects("T_tradShape").DataBodyRange
        'language codes
        T_codes.FromExcelRange SheetDesTranslation.ListObjects("T_tradShape").HeaderRowRange
        i = T_codes.IndexOf(sString)             'Column index of the language code
        T_codes.Clear

        'Be sure the index is of the language positif because if the value is not found, it returns -9999
        If (i > 0) Then
            'Shapes codes
            T_codes.Items = T_data.ExtractSegment(ColumnIndex:=1)
            'Shapes text values
            T_values.Items = T_data.ExtractSegment(ColumnIndex:=i)
            T_data.Clear

            For Each oShape In SheetMain.Shapes
                If T_codes.Includes(oShape.Name) Then
                    k = T_codes.IndexOf(oShape.Name)
                    TranslateShape oShape, T_values.Item(k)
                End If
            Next
        Else
            MsgBox "Update values for the current language " & sString & "for shapes in designer-translation Sheet"
            Exit Sub
        End If

        'Translation of Labels in Ranges ------------------------------------------------
        T_codes.Clear
        T_values.Clear
        'pour les range
        T_data.FromExcelRange SheetDesTranslation.ListObjects("T_tradRange").DataBodyRange
        'language codes
        T_codes.FromExcelRange SheetDesTranslation.ListObjects("T_tradRange").HeaderRowRange
        i = T_codes.IndexOf(sString)             'Column index of the language code
        T_codes.Clear
        'Be sure the index is of the language positif because if the value is not found, it returns -9999
        If (i > 0) Then
            'Ranges codes
            T_codes.Items = T_data.ExtractSegment(ColumnIndex:=1)
            'Ranges text values
            T_values.Items = T_data.ExtractSegment(ColumnIndex:=i)
            T_data.Clear

            On Error Resume Next                 'In case the range is not found or not set, just resume
            For k = 1 To T_values.UpperBound
                TranslateRange T_codes.Item(k), T_values.Item(k)
            Next
            On Error GoTo 0
        Else
            MsgBox "Update values for the current language " & sString & "for ranges in designer-translation Sheet"
            Exit Sub
        End If
        T_values.Clear
        T_codes.Clear
        SheetMain.Range("RNG_Edition").value = TranslateMsg("MSG_Traduit")

        SheetMain.Range("RNG_LLName").NoteText TranslateMsg("NoteText_Forbidden_Caracteres")

        Application.ScreenUpdating = True
    End If
End Sub

Function TranslateMsg(sMsgId As String) As String
    'Translating the message of for displays

    Dim T_data As BetterArray                    'Array of messages and languages data
    Dim T_values As BetterArray                  'value of the message translated
    Dim T_codes As BetterArray                   'Array of messages and languages codes
    Dim slCod As String                          'String for the code of the language
    Dim i As Integer                             'value index for the column
    Dim k As Integer

    'Set values here
    Set T_data = New BetterArray
    T_data.LowerBound = 1
    Set T_values = New BetterArray
    T_values.LowerBound = 1
    Set T_codes = New BetterArray
    T_codes.LowerBound = 1

    TranslateMsg = ""
    T_data.FromExcelRange SheetDesTranslation.ListObjects("T_tradMsg").DataBodyRange
    T_codes.FromExcelRange SheetDesTranslation.ListObjects("T_tradMsg").HeaderRowRange
    slCod = GetLanguageCode(SheetMain.Range(C_sRngLangDes).value)

    If slCod <> "" Then
        i = T_codes.IndexOf(slCod)
        T_codes.Clear
        If (i > 0) Then                          'if the index is found then continue
            T_codes.Items = T_data.ExtractSegment(ColumnIndex:=1)
            T_values.Items = T_data.ExtractSegment(ColumnIndex:=i)
            T_data.Clear
            If (T_codes.Includes(sMsgId)) Then   'be sure the message exists
                k = T_codes.IndexOf(sMsgId)
                TranslateMsg = T_values.Item(k)
            End If
        End If
    End If
End Function

Sub TranslateHeadGeo()
    'translation of column headers in the GEO tab

    Dim sIsoCountry As String, sCountry As String, sSubCounty As String, sWard As String, sPlace As String, sFacility As String

    sIsoCountry = GetLanguageCode(SheetMain.Range(C_sRngLLFormLang).value)

    'Get the isoCode for the linelist
    SheetLLTranslation.Range(C_sRngLLLanguageCode).value = sIsoCountry

    sCountry = Application.WorksheetFunction.HLookup(sIsoCountry, SheetGeo.ListObjects(C_sTabNames).Range, 2, False)
    sSubCounty = Application.WorksheetFunction.HLookup(sIsoCountry, SheetGeo.ListObjects(C_sTabNames).Range, 3, False)
    sWard = Application.WorksheetFunction.HLookup(sIsoCountry, SheetGeo.ListObjects(C_sTabNames).Range, 4, False)
    sPlace = Application.WorksheetFunction.HLookup(sIsoCountry, SheetGeo.ListObjects(C_sTabNames).Range, 5, False)
    sFacility = Application.WorksheetFunction.HLookup(sIsoCountry, SheetGeo.ListObjects(C_sTabNames).Range, 6, False)

    SheetGeo.Range("A1,E1,J1,P1,Z1").value = sCountry
    SheetGeo.Range("F1,K1,Q1,Y1").value = sSubCounty
    SheetGeo.Range("L1,R1,X1").value = sWard
    SheetGeo.Range("S1").value = sPlace
    SheetGeo.Range("W1").value = sFacility

    SheetLLTranslation.Range(C_sRngLLLanguage).value = SheetMain.Range(C_sRngLLFormLang) 'check Language of linelist's forms

End Sub

'TRANSLATIONS OF THE LINELIST, DONE IN THE DESIGNER =============================================================================================================

'Import languages for linelist translation
Sub ImportLang()
    'Import languages from the setup file and sheets Translation and Analysis

    Dim Wkb As Workbook
    Dim sAdr1 As String
    Dim sAdr2 As String
    Dim src As Range
    Dim dest As Range

    With SheetDesTranslation
        .Range(.Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").Column), _
               .Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").End(xlToRight).Column)).ClearContents
    End With

    SheetSetTranslation.Cells.Clear

    BeginWork xlsapp:=Application
    Application.EnableEvents = False
    Application.EnableAnimations = False

    Set Wkb = Workbooks.Open(FileName:=SheetMain.Range(C_sRngPathDic).value)

    SheetSetTranslation.Cells.Clear

    'Copy the languages
    Set src = Wkb.Worksheets(C_sParamSheetTranslation).ListObjects(C_sTabTranslation).Range
    With SheetSetTranslation
        Set dest = .Range(.Cells(C_eStartLinesTransdata, 1), .Cells(C_eStartLinesTransdata + src.Rows.Count, src.Columns.Count))
        dest.value = src.value
        .ListObjects.Add(xlSrcRange, dest, , xlYes).Name = C_sTabTranslation
    End With
    Set src = Wkb.Worksheets(C_sParamSheetTranslation).ListObjects(C_sTabTranslation).HeaderRowRange
    src.Copy SheetDesTranslation.Range("T_Lst_Lang")

    sAdr1 = SheetDesTranslation.Range("T_Lst_Lang").Address
    sAdr2 = SheetDesTranslation.Range("T_Lst_Lang").End(xlToRight).Address

    Wkb.Close

    'Set Validation, 1 is Error
    Call Helpers.SetValidation(SheetMain.Range(C_sRngLangSetup), "='" & SheetDesTranslation.Name & "'!" & sAdr1 & ":" & sAdr2, 1)


    SheetMain.Range(C_sRngLangSetup).value = SheetSetTranslation.Cells(C_eStartLinesTransdata, 1).value
    EndWork xlsapp:=Application
    Application.EnableEvents = True
    Application.EnableAnimations = True
End Sub

'--------------- Writing functions to translate the dictionary and other parts -----------------------------------------

'A function to translate on column in one sheet

Function GetTranslatedValue(ByVal sText As String) As String

    GetTranslatedValue = vbNullString

    Dim iColLang As Integer
    Dim rngTrans As Range
    Dim sLangSetup As String
    Dim iRow As Integer

    'search in linelist language
    sLangSetup = SheetMain.Range(C_sRngLangSetup).value
    iColLang = IIF(sLangSetup <> "", SheetSetTranslation.Rows(C_eStartLinesTransdata).Find(What:=sLangSetup, LookAt:=xlWhole).Column, C_eStartcolumntransdata)

    With DesignerWorkbook.Worksheets(C_sParamSheetTranslation)
        Set rngTrans = .ListObjects(C_sTabTranslation).DataBodyRange
        If rngTrans Is Nothing Then Exit Function
    End With

    On Error Resume Next
    iRow = rngTrans.Find(What:=Application.WorksheetFunction.Trim(sText), LookAt:=xlWhole).Row
    GetTranslatedValue = SheetSetTranslation.Cells(iRow, iColLang).value
    On Error GoTo 0

End Function

Sub TranslateColumn(iCol As Integer, sSheetName As String, iLastRow As Long, Optional iStartRow As Long = 2)
    Dim Wksh As Worksheet
    Dim i As Long
    Dim sText As String

    If iCol > 0 Then                             'Be sure the column exists
        Set Wksh = DesignerWorkbook.Worksheets(sSheetName)

        i = iStartRow

        Do While i <= iLastRow
            If Wksh.Cells(i, iCol).value <> vbNullString Then
                sText = Wksh.Cells(i, iCol).value
                sText = GetTranslatedValue(sText)
                If sText <> vbNullString Then
                    Wksh.Cells(i, iCol).value = sText
                End If
            End If
            i = i + 1
        Loop
    End If
End Sub

Function TranslateCellFormula(ByVal sFormText As String) As String
    Dim j As Integer
    Dim iStart As Integer
    Dim sText As String

    Dim sFormula As String
    Dim sLabelTranlate As String

    TranslateCellFormula = vbNullString
    sText = sFormText

    iStart = 0

    sFormula = Replace(sText, Chr(34) & Chr(34), vbNullString)

    If InStr(1, sFormula, Chr(34), 1) > 0 Then
        For j = 1 To Len(sFormula)
            If Mid(sFormula, j, 1) = Chr(34) Then
                If iStart = 0 Then
                    iStart = j + 1
                Else
                    sLabelTranlate = GetTranslatedValue(Mid(sFormula, iStart, j - iStart))
                    If sLabelTranlate <> vbNullString Then
                        sText = Replace(sText, Mid(sFormula, iStart, j - iStart), sLabelTranlate)
                    End If
                    iStart = 0
                End If
            End If
        Next
    End If

    If sText <> vbNullString Then
        TranslateCellFormula = sText
    End If
End Function

Sub TranslateColumnFormula(iCol As Integer, sSheetName As String, iLastRow As Long, Optional iStartRow As Long = 2)

    Dim i As Integer
    Dim sText As String
    Dim Wksh As Worksheet

    Set Wksh = DesignerWorkbook.Worksheets(sSheetName)

    i = iStartRow

    Do While i <= iLastRow
        sText = Wksh.Cells(i, iCol).value
        sText = TranslateCellFormula(sText)
        If sText <> vbNullString Then Wksh.Cells(i, iCol).value = sText
        i = i + 1
    Loop

End Sub

'A Function to translate the dictionary

'Translation of the dictionary

Sub TranslateDictionary()

    'List of columns to Translate
    Dim DictHeaders As BetterArray
    Dim iCol As Integer
    Dim iLastRow As Long

    Set DictHeaders = New BetterArray
    Set DictHeaders = GetHeaders(DesignerWorkbook, C_sParamSheetDict, 1)

    iLastRow = DesignerWorkbook.Worksheets(C_sParamSheetDict).Cells(Rows.Count, 1).End(xlUp).Row

    'Translate different columns

    'Main label
    iCol = DictHeaders.IndexOf(C_sDictHeaderMainLab)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Sub-label
    iCol = DictHeaders.IndexOf(C_sDictHeaderSubLab)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Note
    iCol = DictHeaders.IndexOf(C_sDictHeaderNote)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Sheet Name
    iCol = DictHeaders.IndexOf(C_sDictHeaderSheetName)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Main Section
    iCol = DictHeaders.IndexOf(C_sDictHeaderMainSec)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Sub Section
    iCol = DictHeaders.IndexOf(C_sDictHeaderSubSec)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)
    'Message
    iCol = DictHeaders.IndexOf(C_sDictHeaderMessage)
    Call TranslateColumn(iCol, C_sParamSheetDict, iLastRow)

    'Formula
    iCol = DictHeaders.IndexOf(C_sDictHeaderFormula)
    Call TranslateColumnFormula(iCol, C_sParamSheetDict, iLastRow)

End Sub

'Translation of the choices
Sub TranslateChoices()

    Dim ChoiceHeaders As BetterArray
    Dim iCol As Integer
    Dim iLastRow As Long

    iLastRow = DesignerWorkbook.Worksheets(C_sParamSheetChoices).Cells(Rows.Count, 1).End(xlUp).Row

    Set ChoiceHeaders = New BetterArray
    Set ChoiceHeaders = GetHeaders(DesignerWorkbook, C_sParamSheetChoices, 1)

    'Label Short
    iCol = ChoiceHeaders.IndexOf(C_sChoiHeaderLabShort)
    Call TranslateColumn(iCol, C_sParamSheetChoices, iLastRow)

    'Label
    iCol = ChoiceHeaders.IndexOf(C_sChoiHeaderLab)
    Call TranslateColumn(iCol, C_sParamSheetChoices, iLastRow)

End Sub

'Translation of the Exports

Sub TranslateExports()
    Dim iLastRow As Long

    iLastRow = DesignerWorkbook.Worksheets(C_sParamSheetExport).Cells(Rows.Count, 1).End(xlUp).Row

    'Second column is for label button (I hope)
    Call TranslateColumn(2, C_sParamSheetExport, iLastRow)
End Sub

Sub TranslateAnalysis()

    Dim iLast As Long
    Dim iCol As Integer
    Dim Wksh As Worksheet
    Dim Headers As BetterArray
    Dim iStartLine As Long
    Dim iStartColumn As Long

    Set Headers = New BetterArray

    Set Wksh = DesignerWorkbook.Worksheets(C_sParamSheetAnalysis)



    'GLOBAL SUMMARY ============================================================

    With Wksh.ListObjects(C_sTabGS)
        iStartLine = .Range.Row
        iLast = .DataBodyRange.Rows.Count + iStartLine
        iStartColumn = .Range.Column
        Set Headers = GetHeaders(DesignerWorkbook, C_sParamSheetAnalysis, iStartLine, iStartColumn)
    End With


    'Translate the column of label
    iCol = Headers.IndexOf(C_sAnaSumLabel)
    If iCol < 0 Then Exit Sub
    Call TranslateColumn(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'Translate the column of formulas
    iCol = Headers.IndexOf(C_sAnaSumFunction)
    If iCol < 0 Then Exit Sub
    Call TranslateColumnFormula(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'UNIVARIATE ANALYSIS =======================================================
    With Wksh.ListObjects(C_sTabUA)
        iStartLine = .Range.Row
        iLast = .DataBodyRange.Rows.Count + iStartLine
        iStartColumn = .Range.Column
        Set Headers = GetHeaders(DesignerWorkbook, C_sParamSheetAnalysis, iStartLine, iStartColumn)
    End With

    'Translate the column of label
    iCol = Headers.IndexOf(C_sAnaSumLabel)
    If iCol < 0 Then Exit Sub
    Call TranslateColumn(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'Translate the column of formulas
    iCol = Headers.IndexOf(C_sAnaSumFunction)
    If iCol < 0 Then Exit Sub
    Call TranslateColumnFormula(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'Translate the column of section
    iCol = Headers.IndexOf(C_sAnaSection)
    If iCol < 0 Then Exit Sub
    Call TranslateColumnFormula(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'BIVARIATE ANALYSIS ========================================================
    With Wksh.ListObjects(C_sTabBA)
        iStartLine = .Range.Row
        iLast = .DataBodyRange.Rows.Count + iStartLine
        iStartColumn = .Range.Column
        Set Headers = GetHeaders(DesignerWorkbook, C_sParamSheetAnalysis, iStartLine, iStartColumn)
    End With

    'Translate the column of label
    iCol = Headers.IndexOf(C_sAnaSumLabel)
    If iCol < 0 Then Exit Sub
    Call TranslateColumn(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'Translate the column of formulas
    iCol = Headers.IndexOf(C_sAnaSumFunction)
    If iCol < 0 Then Exit Sub
    Call TranslateColumnFormula(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

    'Translate the column of section
    iCol = Headers.IndexOf(C_sAnaSection)
    If iCol < 0 Then Exit Sub
    Call TranslateColumnFormula(iCol, C_sParamSheetAnalysis, iLast, iStartLine)

End Sub

Sub TranslateLinelistData()
    'translation of the Export, Dictionary and Choice sheets for the linelist

    BeginWork xlsapp:=Application

    'Dictionary
    Call TranslateDictionary
    'Choices
    Call TranslateChoices
    'Exports
    Call TranslateExports
    'Analysis...
    Call TranslateAnalysis


    EndWork xlsapp:=Application

End Sub


