Attribute VB_Name = "M_traduction"
Option Explicit

Function GetLanguageCode(sString As String) As String
    Dim T_data As BetterArray                    'array of languages
    Dim T_codes As BetterArray                   'array of languages codes
    Dim T_values As BetterArray                  'values of languages
    
    Set T_data = New BetterArray
    T_data.LowerBound = 1
    Set T_codes = New BetterArray
    T_codes.LowerBound = 1
    Set T_values = New BetterArray
    T_values.LowerBound = 1
    
    GetLanguageCode = ""
    
    T_data.FromExcelRange SheetDesTranslation.ListObjects(C_sTabLang).DataBodyRange 'Language table
    T_values.Items = T_data.ExtractSegment(ColumnIndex:=1) 'language values
    T_codes.Items = T_data.ExtractSegment(ColumnIndex:=2) 'language codes
    
    If T_values.Includes(sString) Then
        'The Language code
        GetLanguageCode = T_codes.Item(T_values.IndexOf(sString))
    End If
    
End Function

'Translate one shape using informations on languages
Sub TranslateShape(oShape As Object, sValue As String)
    Dim bVis As Integer                          'visibility of the shape
    Dim sFont As String                          'actual font of the shape
    bVis = oShape.Visible
    'be sure the shape is visible before updating its text
    oShape.Visible = -1
    With Sheets("MAIN").Shapes(oShape.Name)
        'keeping the previous font selected
        sFont = .TextFrame.Characters.Font.Name
        .TextFrame.Characters.Text = sValue
        .TextFrame.Characters.Font.Name = sFont
    End With
    oShape.Visible = bVis

End Sub

Sub TranslateRange(rngCode As String, rngValue As String)
    Sheets("MAIN").Range(rngCode).value = rngValue
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
                   
            For Each oShape In Sheets("MAIN").Shapes
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
        Sheets("Main").Range("RNG_Edition").value = TranslateMsg("MSG_Traduit")
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

'This is just to translate the translation shape and headers to help for translation
Public Sub StartTranslate()
    
    Dim slCod As String                          'Language code
    Dim T_data As BetterArray                    'data of the Shapes
    Dim T_codes As BetterArray                   'codes
    Dim T_values As BetterArray                  'values
    Dim indexShapeCode As Integer                'index of the shape
    Dim indexRangeCode As Integer                'index of the range
    Dim indexLangCod As Integer                  'index of the language

    Set T_data = New BetterArray
    Set T_codes = New BetterArray
    Set T_values = New BetterArray
        
    Application.ScreenUpdating = False
    slCod = GetLanguageCode(SheetMain.Range("RNG_LangDesigner").value)
    If slCod <> "" Then
        T_codes.FromExcelRange SheetDesTranslation.ListObjects("T_tradShape").HeaderRowRange
        indexLangCod = T_codes.IndexOf(slCod)    'index of the language code
        T_codes.Clear
        'updating one shape: Translation shape
        If (indexLangCod > 0) Then               'we are sure the index column is present
            T_data.FromExcelRange SheetDesTranslation.ListObjects("T_tradShape").DataBodyRange
            T_values.Items = T_data.ExtractSegment(ColumnIndex:=indexLangCod)
            T_codes.Items = T_data.ExtractSegment(ColumnIndex:=1) 'index of all the shapes codes
            'where the shape of the code is
            indexShapeCode = T_codes.IndexOf("SHP_Trad")
            If (indexShapeCode > 0) Then
                TranslateShape Sheets("Main").Shapes("SHP_Trad"), T_values.Item(indexShapeCode)
            End If
            'Then do the same for the range above the RNG_Designer
            T_data.Clear
            T_values.Clear
            T_codes.Clear
            T_data.FromExcelRange SheetDesTranslation.ListObjects("T_tradRange").DataBodyRange
            T_values.Items = T_data.ExtractSegment(ColumnIndex:=indexLangCod)
            'index of all the Ranges codes is 1
            T_codes.Items = T_data.ExtractSegment(ColumnIndex:=1)
            indexRangeCode = T_codes.IndexOf("RNG_LabLangDesigner")
            If (indexRangeCode > 0) Then
                TranslateRange "RNG_LabLangDesigner", T_values.Item(indexRangeCode)
            End If
            Sheets("Main").Range("RNG_LangSetup").value = ""
            Sheets("Main").Range("RNG_LangGeo").value = ""
        End If
    End If
    Set T_data = Nothing
    Set T_values = Nothing
    Set T_codes = Nothing
    Application.ScreenUpdating = True
End Sub

Sub TranslateForm(sNameForm As String)

    Dim i As Integer
    Dim T_data
    Dim D_data As Scripting.Dictionary

    T_data = ThisWorkbook.Sheets("translation").[T_tradForm]

    i = 1
    While i <= UBound(T_data, 1) And sNameForm <> T_data(0, i)
        i = i + 1
    Wend

    If sNameForm = T_data(0, i) Then
        While sNameForm = T_data(0, i)

            ThisWorkbook.VBProject.VBComponents(sNameForm).Controls(T_data(1, i)).value = T_data(2, i)
            i = i + 1
        Wend
    End If

End Sub

