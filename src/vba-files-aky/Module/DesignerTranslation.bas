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


