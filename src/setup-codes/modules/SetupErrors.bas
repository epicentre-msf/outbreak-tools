Attribute VB_Name = "SetupErrors"
Option Explicit

'@Folder("Checks")

'Module for checkings in the Setup file
'This is a long long long module.

Private Const DICTSHEETNAME As String = "Dictionary"
Private Const CHOICESHEETNAME As String = "Choices"
Private Const EXPORTSHEETNAME As String = "Exports"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const TRANSLATIONSHEETNAME As String = "Translations"

Private checkTables As BetterArray
Private wb As Workbook
Private errTab As ICustomTable 'Custom table for Error Messages
Private pass As IPasswords
Private dict As ILLdictionary
Private formData As IFormulaData
Private choi As ILLchoice

Private Sub Initialize()
    Dim shform As Worksheet
    BusyApp

    'Initialize formula
    Set wb = ThisWorkbook
    Set shform = wb.Worksheets("__formula")
    Set formData = FormulaData.Create(shform)

    'Initialize the checking
    Set checkTables = New BetterArray
    Set errTab = CustomTable.Create(shform.ListObjects("Tab_Error_Messages"), idCol:="Key")
    Set pass = Passwords.Create(wb.Worksheets("__pass"))
    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEETNAME), 5, 1)
    Set choi = LLchoice.Create(wb.Worksheets(CHOICESHEETNAME), 4, 1)
End Sub

Private Function FormulaMessage(ByVal formValue As String, _
                                ByVal keyName As String, _
                                Optional ByVal value_one As String = vbNullString, _
                                Optional ByVal value_two As String = vbNullString, _
                                Optional ByVal formulaType As String = "linelist") As String
    Dim setupForm As IFormulas
    Set setupForm = Formulas.Create(dict, formData, formValue)
    If Not setupForm.Valid(formulaType:=formulaType) Then _
    FormulaMessage = ConvertedMessage(keyName, value_one, value_two, _
                                     setupForm.reason())

End Function

Private Function ConvertedMessage(ByVal keyName As String, _
                                  Optional ByVal value_one As String = vbNullString, _
                                  Optional ByVal value_two As String = vbNullString, _
                                  Optional ByVal value_three As String = vbNullString, _
                                  Optional ByVal value_four As String = vbNullString) As String
    Dim infoMessage As String

    infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
    infoMessage = Replace(infoMessage, "{$}", value_one)
    infoMessage = Replace(infoMessage, "{$$}", value_two)
    infoMessage = Replace(infoMessage, "{$$$}", value_three)
    infoMessage = Replace(infoMessage, "{$$$$}", value_four)

    ConvertedMessage = infoMessage
End Function

Private Sub CheckDictionary()

    Dim check As IChecking
    Dim csTab As ICustomTable
    Dim expTab As ICustomTable
    Dim varRng As Range
    Dim sheetRng As Range
    Dim FUN As WorksheetFunction
    Dim varValue As String
    Dim sheetValue As String
    Dim shdict As Worksheet
    Dim shexp As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim cellRng As Range
    Dim controlDetailsValue As String
    Dim controlValue As String
    Dim setupForm As Object
    Dim checkingCounter As Long 'Counter As 0 for each variable
    Dim choiCategories As BetterArray
    Dim formCategories As BetterArray
    Dim controlsList As BetterArray
    Dim choiName As String
    Dim tabCounter As Long
    Dim catValue As String 'category value when dealing with choice formulas
    Dim expCounter As Long 'Counter for each of the exports
    Dim expRng As Range    'Range for exports
    Dim expStatusRng As Range
    Dim minValue As String
    Dim maxValue As String
    Dim mainVarRng As Range
    Dim mainLabValue As String
    Dim uniqueValue As String
    Dim typeValue As String 'Type of the variable
    Dim formatValue As String 'Format of the variable

    BusyApp

    Set shdict = wb.Worksheets(DICTSHEETNAME)
    Set shexp = wb.Worksheets(EXPORTSHEETNAME)
    Set check = Checking.Create(titleName:="Dictionary incoherences Type--Where?--Details")
    Set csTab = CustomTable.Create(shdict.ListObjects(1), idCol:="Variable Name")
    Set expTab = CustomTable.Create(shexp.ListObjects(1), idCol:="Export Number")
    Set FUN = Application.WorksheetFunction

    Set choiCategories = New BetterArray
    Set formCategories = New BetterArray
    Set controlsList = New BetterArray

    ' Some preparation steps: Resize the dictionary table, sort on sheetNames
    pass.UnProtect DICTSHEETNAME
    pass.UnProtect EXPORTSHEETNAME

    csTab.RemoveRows
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name")
    Set sheetRng = csTab.DataRange("Sheet Name")
    Set mainVarRng = csTab.DataRange("Main Label")
    Set cellRng = varRng.Cells(varRng.Rows.Count, 1)
    controlsList.Push "choice_manual", "choice_formula", "formula", _
                      "geo", "hf", "custom", "list_auto", "case_when", _ 
                      "choice_custom", "choice_multiple"

    'Errors on columns
    Do While cellRng.Row >= varRng.Row
        checkingCounter = 0 'checkingCounter is just an id for errors and checkings
        varValue = FUN.Trim(cellRng.Value)

        'Duplicates variable names
        If FUN.COUNTIF(varRng, varValue) > 1 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-unique"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Variabel lenths < 4
        If Len(varValue) < 4 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-length"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Empty variable name
        mainLabValue = shdict.Cells(cellRng.Row, mainVarRng.Column)

        If (mainLabValue = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-main-lab"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        sheetValue = shdict.Cells(cellRng.Row, sheetRng.Column)

        'Empty sheet names
        If sheetValue = vbNullString Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-empty-sheet"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        controlValue = csTab.Value("Control", varValue)
        controlDetailsValue = csTab.Value("Control Details", varValue)

        'Unkown control
        If (Not controlsList.Includes(controlValue)) And _ 
           (controlValue <> vbNullString) And _ 
           (Not (InStr(1, controlValue, "choice_multiple") = 1)) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unknown-control"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, controlValue, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Incoherences in choice formula
        If (controlValue = "choice_formula") And (InStr(1, controlDetailsValue, "CHOICE_FORMULA") = 1) Then
            'choice formula
            Set setupForm = ChoiceFormula.Create(controlDetailsValue)

            'Test if the choice_name exists
            choiName = setupForm.choiceName()
            If Not choi.ChoiceExists(choiName) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choiform-empty"
                infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            Else
                'Test if categories in choice_formula exists
                keyName = "dict-cat-notfound"
                Set choiCategories = choi.Categories(choiName)
                Set formCategories = setupForm.Categories()

                For tabCounter = formCategories.LowerBound To formCategories.UpperBound
                    catValue = formCategories.Item(tabCounter)
                    If Not choiCategories.Includes(catValue) Then
                        checkingCounter = checkingCounter + 1
                        infoMessage = ConvertedMessage(keyName, cellRng.Row, catValue, choiName)
                        check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingInfo
                    End If
                Next
            End If
        End If

        'Incorrect formulas (should include tests case_when and choice_formula)
        If (controlValue = "formula" Or controlValue = "case_when" Or controlValue = "choice_formula") Then
            keyName = "dict-incor-form"
            infoMessage = FormulaMessage(controlDetailsValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        'Choices not present in choice sheet
        'Could be choice manual, choice_custom or a choice multiple. For choice_custom, important
        'To precise that the controlDetails shoud be non empty.
        If (controlValue = "choice_manual") Or _ 
           ((controlValue = "choice_custom") And (controlDetailsValue <> vbNullString)) Or _ 
           (InStr(1, controlValue, "choice_multiple") = 1) Then
            If Not choi.ChoiceExists(controlDetailsValue) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choi-empty"
                infoMessage = ConvertedMessage(keyName, cellRng.Row, controlDetailsValue)
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        'Incorrect Min/Max formulas
        minValue = csTab.Value("Min", varValue)
        maxValue = csTab.Value("Max", varValue)

        If (minValue <> vbNullString) Then
            keyName = "dict-incor-min"
            infoMessage = FormulaMessage(minValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (maxValue <> vbNullString) Then
            keyName = "dict-incor-max"
            infoMessage = FormulaMessage(maxValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        'If there is validation and the type is not precised
        typeValue = csTab.Value("Variable Type", varValue)

        If ((minValue <> vbNullString)  Or (maxValue <> vbNullString)) And (typeValue = vbNullString) Then
            keyName = "dict-valid-control"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
            End If
        End If

        'If the variable format is precised without the type
        formatValue = csTab.Value("Variable Format", varValue)
        If (formatValue <> vbNullString) And (typeValue = vbNullString) Then
            keyName = "dict-format-control"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If
        

        'Inform for validation on unique values
        uniqueValue = csTab.Value("Unique", varValue)

        If (uniqueValue = "yes") Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unique-val"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingInfo
        End If


        Set cellRng = cellRng.Offset(-1)
    Loop

    'Exports Range
    expTab.Sort "Export Number"

    For expCounter = 1 To 5
    
        Set expRng = csTab.DataRange("Export " & expCounter)
        Set expStatusRng = expTab.DataRange("Status")

        If (Not (expRng Is Nothing)) And (Not (expStatusRng Is Nothing)) Then
            If (FUN.CountBlank(expRng) <> expRng.Rows.Count) And (expStatusRng.Cells(expCounter, 1).Value <> "active") Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-export-na"
                infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
                infoMessage = Replace(infoMessage, "{$}", expCounter)
                check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingNote
            End If
        End If
    Next

    checkTables.Push check
    pass.Protect DICTSHEETNAME
    pass.Protect EXPORTSHEETNAME
End Sub


Private Sub CheckChoice()
    Dim check As IChecking
    Dim shchoi As Worksheet
    Dim shdict As Worksheet
    Dim choiTab As ICustomTable
    Dim dictTab As ICustomTable
    Dim choiLst As BetterArray
    Dim counter As Long
    Dim checkingCounter As Long
    Dim choiName As String
    Dim infoMessage As String
    Dim cellRng As Range
    Dim choiNameRng As Range
    Dim sortValue As String
    Dim choiLabValue As String
    Dim keyName As String
    Dim usedChoicesLst As BetterArray
    Dim setupForm As Object
    Dim cntrlRng As Range
    Dim cntrlDetRng As Range
    Dim actualControl As String

    BusyApp

    Set shchoi = wb.Worksheets(CHOICESHEETNAME)
    Set shdict = wb.Worksheets(DICTSHEETNAME)
    Set choiTab = CustomTable.Create(shchoi.ListObjects(1))
    Set usedChoicesLst = New BetterArray

    pass.UnProtect CHOICESHEETNAME

    'Sort the choices in choice sheet
    choi.Sort
    choiTab.RemoveRows

    Set check = Checking.Create(titleName:="Choices incoherences Type--Where?--Details")
    Set dictTab = CustomTable.Create(shdict.ListObjects(1))

    'List of all choice names
    Set choiLst = choi.AllChoices()
    Set choiNameRng = choiTab.DataRange("List Name")
    Set cntrlRng = dictTab.DataRange("Control")
    Set cntrlDetRng = dictTab.DataRange("Control Details")

    'Initialize the number of checkings to do.
    checkingCounter = 0

    'List of choices_formulas
    For counter = 1 To cntrlDetRng.Rows.Count

        actualControl = cntrlRng.Cells(counter, 1).Value
        'add choices to the list of used choices
        If (actualControl = "choice_manual") Or _ 
            (actualControl = "choice_custom") Or _ 
            (InStr(1, actualControl, "choice_multiple") = 1) Then

           If (cntrlDetRng.Cells(counter, 1).Value <> vbNullString) Then _ 
              usedChoicesLst.Push cntrlDetRng.Cells(counter, 1).Value

        ' add choice formulas to the list of used choices
        ElseIf (actualControl = "choice_formula") Then

            If (cntrlDetRng.Cells(counter, 1).Value <> vbNullString)  Then 
                Set setupForm = ChoiceFormula.Create(cntrlDetRng.Cells(counter, 1).Value)
                'avoid sending empty strings to the list of used choices, so test for that before
                choiName = setupForm.ChoiceName()
                If choiName <> vbNullString Then usedChoicesLst.Push choiName
            End If

        End If
    Next

    'choices not used
    For counter = choiLst.LowerBound To choiLst.UpperBound
        choiName = choiLst.Item(counter)
        If Not usedChoicesLst.Includes(choiName) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-unfound-choi"
            infoMessage = ConvertedMessage(keyName, choiName)

            check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If
    Next

    'Going through the choice sheet for eventual incoherences
    'cellRng in the last cell of the choices.
    Set cellRng = choiNameRng.Cells(choiNameRng.Rows.Count, 1)

    Do While cellRng.Row >= choiNameRng.Row

        choiName = cellRng.Value
        sortValue = cellRng.Offset(, 1).Value
        choiLabValue = cellRng.Offset(, 2).Value

        'Labels without choice name
        If (choiLabValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-emptychoi-lab"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiLabValue)

            check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Sort without choice name
        If (sortValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-emptychoi-order"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, sortValue)

            check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Sort not filled
        If (sortValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-empty-order"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)

            check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If

        'missing label for choice name (info)
        If (choiLabValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-mis-lab"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)

            check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
        End If
        Set cellRng = cellRng.Offset(-1)
    Loop


    checkTables.Push check
    pass.Protect CHOICESHEETNAME
End Sub

'Checking on exports
Private Sub CheckExports()

    Dim expTab As ICustomTable
    Dim counter As Long
    Dim shexp As Worksheet
    Dim check As IChecking
    Dim keyName As String
    Dim checkingCounter As Long
    Dim expStatus As String
    Dim cellRng As Range
    Dim infoMessage As String
    Dim keysLst As BetterArray
    Dim headersLst As BetterArray
    Dim headerCounter As Long
    Dim exportRng As Range
    Dim statusRng As Range
    Dim FUN As WorksheetFunction
    Dim fileNameLst As BetterArray
    Dim actFileName As String
    Dim fileCounter As Long
    Dim fileNameChunk As String
    Dim vars As ILLVariables
    Dim numberOfExports As Long
    Dim pwd As String
    Dim expId As String

    BusyApp

    Set shexp = wb.Worksheets(EXPORTSHEETNAME)
    Set expTab = CustomTable.Create(shexp.ListObjects(1), idCol:="Export Number")
    Set keysLst = New BetterArray
    Set headersLst = New BetterArray
    Set fileNameLst = New BetterArray
    Set statusRng = expTab.DataRange("Status")

    Set check = Checking.Create(titleName:="Export incoherences type--Where?--Details")
    Set FUN = Application.WorksheetFunction
    Set vars = LLVariables.Create(dict)

    headersLst.Push "Label button", "Password", "Export metadata sheets", _
                    "File format", "File name", "Header format"
    keysLst.Push "exp-mis-lab", "exp-mis-pass", "exp-mis-meta",  _
                 "exp-mis-form", "exp-mis-name", "exp-mis-head"

    checkingCounter = 0
    numberOfExports = statusRng.Rows.Count

    For counter = 1 To numberOfExports

        
        'This cellRng is used not only for the status, but also for identifying the 
        'row of the checking.
        Set cellRng = expTab.CellRange("Status", counter + statusRng.Row - 1)
        expStatus = cellRng.Value
        pwd = expTab.Value(colName:="Password", keyName:=CStr(counter))
        expId = expTab.Value(colName:="Include personal identifiers", keyName:=CStr(counter))
        
        Debug.Print "Export: " & expId
        
        For headerCounter = keysLst.LowerBound To keysLst.UpperBound
            'Empty label, password, metadata, translation file format or file name, file header
            'The check is done for each of the export.

            If IsEmpty(expTab.CellRange(headersLst.Item(headerCounter), cellRng.Row)) And (expStatus = "active") Then
                checkingCounter = checkingCounter + 1
                keyName = keysLst.Item(headerCounter)
                infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)

                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
            End If
        Next

        'Active export not filled in the dictionary
        On Error Resume Next
        Set exportRng = Nothing
        Set exportRng = dict.DataRange("Export " & counter, includeHeaders:=False)
        On Error GoTo 0

        If (expStatus = "active") And (exportRng Is Nothing) Then
            
            checkingCounter = checkingCounter + 1
            keyName = "exp-unfound-dictcolumn"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning

        ElseIf (expStatus = "active") And (FUN.CountBlank(exportRng) = exportRng.Rows.Count) Then
            
            checkingCounter = checkingCounter + 1
            keyName = "exp-act-empty"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        End If

        'Exports with identifiers without a password
        If (pwd <> "yes") And (expId = "yes") Then

            checkingCounter = checkingCounter + 1
            keyName = "exp-id-passwd"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        
        End If

        'Variable in File name not found:
        'Get the file name, split on + and loop through elements
        actFileName = expTab.Value(colName:="File Name", keyName:=counter)
        fileNameLst.Clear
        fileNameLst.Items = Split(actFileName, "+")

        'Loop through each file name
        For fileCounter = fileNameLst.LowerBound To fileNameLst.UpperBound

            fileNameChunk = fileNameLst.Item(fileCounter)

            If Not (InStr(1, fileNameChunk, Chr(34)) > 0) Then

                'Test for linelist variables presence
                If Not vars.Contains(fileNameChunk) Then

                    checkingCounter = checkingCounter + 1
                    keyName = "exp-filename-varunfound"
                    infoMessage = ConvertedMessage(keyName, cellRng.Row, counter, fileNameChunk)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning

                    'Test for variables in vlist1D
                ElseIf (vars.Value(colName :="Sheet Type", varName:=fileNameChunk) <> "vlist1D") Then

                    checkingCounter = checkingCounter + 1
                    keyName = "exp-filename-varnotvlist"
                    infoMessage = ConvertedMessage(keyName, cellRng.Row, counter, fileNameChunk)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning

                End If

            End If
        Next
    Next

    'Active export not filled in the dictionary
    checkTables.Push check
End Sub


'Checking on Translations
Private Sub CheckTranslations()
    Dim Lo As ListObject
    Dim shTrans As Worksheet
    Dim hRng As Range
    Dim messageMissing As String
    Dim nbMissing As Long
    Dim langName As String
    Dim check As IChecking
    Dim counter As Long
    Dim colRng As Range

    BusyApp

    Set shTrans = wb.Worksheets(TRANSLATIONSHEETNAME)
    Set Lo = shTrans.ListObjects(1)
    Set hRng = Lo.HeaderRowRange
    Set check = Checking.Create(titleName:="Translation incoherences--Where?--Details")
    If (Not Lo.DataBodyRange Is Nothing) Then
        For counter = 1 To hRng.Columns.Count
            langName = hRng.Cells(1, counter).Value
            Set colRng = Lo.ListColumns(langName).DataBodyRange
            nbMissing = Application.WorksheetFunction.CountBlank(colRng)
            If nbMissing > 0 Then
                messageMissing = "Translations Sheet--" & nbMissing & _
                            " labels are missing for column " & _
                            langName & "."
                'Add the message to checkings
                check.Add "trads-mis-labs-" & counter, messageMissing, checkingInfo
            End If
        Next
    End If

    checkTables.Push check
End Sub

'adding checks for analysis
Private Sub checkTable(ByVal partName As String)

    Const TABGS As String = "Tab_global_summary"
    Const TABUA As String = "Tab_Univariate_Analysis"
    Const TABBA As String = "Tab_Bivariate_Analysis"
    Const TABTS As String = "Tab_TimeSeries_Analysis"
    Const TABSP As String = "Tab_Spatial_Analysis"
    Const TABSPTEMP As String = "Tab_SpatioTemporal_Analysis"
    Const TABTSGRAPHS As String = "Tab_Graph_TimeSeries"

    Dim tabLo As ListObject 'ListObject for each of the tables in analysis
    Dim loname As String 'ListObject Name
    Dim tabHeaderRng As Range 'HeaderRange for table specs
    Dim tabSpecsRng As Range 'Specification range of one table
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim check As IChecking
    Dim nbLines As Long
    Dim checkingCounter As Long
    Dim keyName As String
    Dim infoMessage As String
    Dim anaForm As String 'analysis formula
    Dim FUN As WorksheetFunction
    Dim tableScope As Byte

    BusyApp

    'The listObject is related to the name of part given
    loname = Switch( _
        partName = "Global summary", TABGS, _
        partName = "Univariate analysis", TABUA, _
        partName = "Bivariate analysis", TABBA, _
        partName = "Time series analysis", TABTS, _
        partName = "Spatial analysis", TABSP, _
        partName = "Spatio-temporal analysis", TABSPTEMP, _
        partName = "Time series graphs", TABTSGRAPHS _ 
    )

    Set sh = wb.Worksheets(ANALYSISSHEETNAME)
    Set FUN = Application.WorksheetFunction
    checkingCounter = 0
    'If the list Object does not exists, exit the checkings
    On Error Resume Next
    Set tabLo = sh.ListObjects(loname)
    On Error GoTo 0
    If tabLo Is Nothing Then Exit Sub
    Set check = Checking.Create(titleName:="Analysis incoherences----", _
                                subtitleName:=partName & "--Where?--Details")

    Set tabHeaderRng = tabLo.HeaderRowRange

    For nbLines = 1 To tabLo.ListRows.Count
        'Non valid table
        Set tabSpecsRng = tabLo.ListRows(nbLines).Range()
        'If the specs range is empty, move to next one with a warning
        If ((FUN.CountBlank(tabSpecsRng) = tabSpecsRng.Columns.Count) And _
           (loname <> TABTS)) Or (FUN.CountBlank(tabSpecsRng) >= tabSpecsRng.Columns.Count - 1) Then
            keyName = "ana-empty-tab"
            checkingCounter = checkingCounter + 1
            infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
        Else

            Set specs = TablesSpecs.Create(tabHeaderRng, tabSpecsRng, dict, choi)
            tableScope = specs.TableType

            'Invalid table
            If (Not specs.ValidTable()) Then
                checkingCounter = checkingCounter + 1
                keyName = "ana-inv-tab"
                infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row, specs.ValidityReason())
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError

                'on new section on time series add another Error
                If (tableScope = TypeTimeSeries) And (specs.isNewSection()) _
                    And (FUN.CountBlank(tabSpecsRng) < tabSpecsRng.Columns.Count - 1) Then
                    keyName = "ana-ts-newsec"
                    infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
                End If
            End If 

            'No Title
            If (specs.Value("title") = vbNullString) And _ 
               (tableScope <> TypeGlobalSummary) And _ 
               (tableScope <> TypeTimeSeriesGraph) Then

                checkingCounter = checkingCounter + 1
                keyName = "ana-empty-title"
                infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row)

                check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
            End If

            'flip coordinates = yes on univariate analysis table
            If (tableScope = TypeUnivariate) And _
               specs.HasGraph() And _
               (specs.Value("flip") = "yes") And _
               (specs.Value("percentage") = "yes") Then

                checkingCounter = checkingCounter + 1
                keyName = "ana-uaflip-perc"
                infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row)

                check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
            End If

            'add graph = no and flip coordinates = "yes" (on tables expect time series)
            If (tableScope <> TypeTimeSeries) And (tableScope <> TypeTimeSeriesGraph) Then
                If (Not specs.HasGraph()) And (specs.Value("flip") = "yes")  Then

                    checkingCounter = checkingCounter + 1
                    keyName = "ana-adgr-flip"
                    infoMessage = ConvertedMessage(keyName, tabSpecsRng.Row)

                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
                End If
            End If

            'Retrieve the formula of one specs and test them
            If (tableScope <> TypeTimeSeriesGraph) Then
                anaForm = specs.Value("function")
                If (anaForm <> vbNullString) Then
                    keyName = "ana-incor-form"
                    infoMessage = FormulaMessage(anaForm, keyName, tabSpecsRng.Row, _
                                                formulaType:="analysis")
                    If (infoMessage <> vbNullString) Then
                        checkingCounter = checkingCounter + 1
                        check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                    End If
                End If
            End If
        End If
    Next

    checkTables.Push check
End Sub

Private Sub CheckAnalysis()

    'Resize all the tables on the analysis sheet
    EventsRibbon.ManageRows sheetName:="Analysis", del:=True, allAnalysis:=True
    'This busyapp is important because after managerows, enablevents = True
    BusyApp
    'Check all the tables progressively
    checkTable "Global summary"
    checkTable "Univariate analysis"
    checkTable "Bivariate analysis"
    checkTable "Time series analysis"
    checkTable "Time series graphs"
    checkTable "Spatial analysis"
    checkTable "Spatio-temporal analysis"
End Sub

Private Sub PrintReport()
    Const CHECKSHEETNAME As String = "__checkRep"
    Const DROPSHEETNAME As String = "__variables"
    'Initilialize the dropdown array and list

    Dim checKout As ICheckingOutput
    Dim sh As Worksheet
    Dim drop As IDropdownLists
    Dim formatRng As Range

    BusyApp
    Set sh = wb.Worksheets(CHECKSHEETNAME)
    Set checKout = CheckingOutput.Create(sh)
    Set drop = DropdownLists.Create(wb.Worksheets(DROPSHEETNAME))
    sh.Cells.EntireRow.Hidden = False

    checKout.PrintOutput checkTables
    'Set validation for filtering
    drop.SetValidation sh.Range("RNG_CheckingFilter"), "__checking_types", "error"
    With sh
        Set formatRng = .Range(.Cells(1, 2), .Cells(1, 3))
        formatRng.Font.color = RGB(21, 133, 255)
        formatRng.Font.Bold = True

        With .Range("RNG_CheckingFilter")
            .Interior.Color = RGB(221, 235, 247)
            .HorizontalAlignment = xlHAlignCenter
            .Value = "All"
        End With
    End With
End Sub


Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
  End Sub

Public Sub CheckTheSetup()
    BusyApp
    Initialize
    CheckDictionary
    CheckChoice
    CheckExports
    CheckAnalysis
    CheckTranslations
    PrintReport
    Application.EnableEvents = True
End Sub
