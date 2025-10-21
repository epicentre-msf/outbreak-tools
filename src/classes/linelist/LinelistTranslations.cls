VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Simplified linelist translations manager backed by worksheet names and tables"

'@Folder("Linelist")
'@ModuleDescription("Simplified linelist translations manager backed by worksheet names and tables")
'@depends BetterArray, CustomTable, ICustomTable, TranslationObject, ITranslationObject, ProjectError
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName

Option Explicit
Implements ILinelistTranslations

Private Type TLinelistTranslations
    sh As Worksheet               'All translation tables and named values live on this sheet
End Type

Private this As TLinelistTranslations
Private internalTableMappings As BetterArray
Private internalValueMappings As BetterArray

Private Const CLASSNAME As String = "LinelistTranslations"
Private Const DICTIONARY_TABLE As String = "Tab_Translations"
Private Const DICTIONARY_NAME As String = "LLDictionaryTranslation"
Private Const DICTIONARY_EXPORT_SHEET As String = "Translation"
Private Const SETUP_SHEET_NAME_ID As String = "RNG_SetupDictionarySheet"


'@section Factory
'===============================================================================
'@details Entry point returning the improved translations manager.
'@param sh Worksheet hosting all linelist translations.
'@return ILinelistTranslations configured instance.
Public Function Create(ByVal sh As Worksheet) As ILinelistTranslations

    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation worksheet is required"
    End If

    Dim instance As LinelistTranslations
    Set instance = New LinelistTranslations
    instance.Initialise sh
    Set Create = instance
End Function


'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal sh As Worksheet)
    Set this.sh = sh
    ValidateContext
End Sub

Private Sub ValidateContext()
    EnsureTranslationTables
    EnsureDictionaryName
    EnsureValueNames
    EnsureSetupSheetName
End Sub

Private Sub EnsureTranslationTables()
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant

    Set mappings = TableMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        EnsureListObjectExists this.sh, CStr(mapping(2))
    Next idx
End Sub

Private Sub EnsureDictionaryName()
    Dim dictLo As ListObject
    Set dictLo = FindListObject(this.sh, DICTIONARY_TABLE)

    If dictLo Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Table '" & DICTIONARY_TABLE & "' does not exist on worksheet '" & this.sh.Name & "'"
    End If

    AssignDictionaryName this.sh, dictLo
End Sub

Private Sub EnsureValueNames()
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant

    Set mappings = ValueMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        EnsureNameExists CStr(mapping(2))
    Next idx
End Sub

Private Sub EnsureNameExists(ByVal name As String)
    If TrySheetNameDefinition(this.sh, name) Is Nothing Then
        AddOrUpdateNameConstant this.sh, name, vbNullString
    End If
End Sub

Private Sub EnsureSetupSheetName()
    Dim definition As Name
    Set definition = TrySheetNameDefinition(this.sh, SETUP_SHEET_NAME_ID)

    If definition Is Nothing Then
        AddOrUpdateNameConstant this.sh, SETUP_SHEET_NAME_ID, DICTIONARY_EXPORT_SHEET
    ElseIf LenB(NameDefinitionValue(definition)) = 0 Then
        SetNameDefinition definition, DICTIONARY_EXPORT_SHEET
    End If
End Sub


'@section Public API
'===============================================================================
'@label:TransObject
'@sub-title Provide a translation object for the requested scope.
'@param scope Byte describing the table (defaults to messages).
'@return ITranslationObject wrapping the requested ListObject.
Public Function TransObject(Optional ByVal scope As Byte = translationScopeMessages) As ITranslationObject
    Dim mapping As Variant
    Dim tableName As String
    Dim languageNameDef As String
    Dim sourceTable As ListObject
    Dim languageValue As String

    mapping = ResolveTableMapping(scope)
    tableName = CStr(mapping(1))
    languageNameDef = CStr(mapping(2))

    Set sourceTable = FindListObject(this.sh, tableName)
    If sourceTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Table '" & tableName & "' not found on worksheet '" & this.sh.Name & "'"
    End If

    languageValue = NameValue(languageNameDef)

    Set TransObject = TranslationObject.Create(sourceTable, languageValue)
End Function

'@label:Value
'@sub-title Retrieve a named translation value stored as a worksheet-level name.
'@param tagName Lowercase tag identifying the name.
'@return String value stored in the name definition.
Public Property Get Value(ByVal tagName As String) As String
    Value = NameValue(TagToName(tagName))
End Property

'@label:SetValue
'@sub-title Update a named translation value stored as a worksheet-level name.
'@param tagName Lowercase tag identifying the name.
'@param newValue String value to assign.
Public Sub SetValue(ByVal tagName As String, ByVal newValue As String)
    SetNameValue TagToName(tagName), newValue
End Sub

'@label:ExportToWorkbook
'@sub-title Export translation sheet and dictionary table to the supplied workbook.
'@param target Workbook that will receive cloned content.
'@param visibility Optional XlSheetVisibility applied to exported sheets.
Public Sub ExportToWorkbook(ByVal target As Workbook, _
                            Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)

    EnsureWorkbook target

    Dim translationTarget As Worksheet

    Set translationTarget = PrepareOutputSheet(target, this.sh.Name, visibility)
    CopyTranslationContent translationTarget
End Sub

'@label:ImportFromWorkbook
'@sub-title Import translation tables, dictionary table, and named values from the supplied workbook.
Public Sub ImportFromWorkbook(ByVal source As Workbook)
    EnsureWorkbook source

    Dim sourceTranslation As Worksheet
    Set sourceTranslation = ResolveWorksheet(source, this.sh.Name)
    If sourceTranslation Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & this.sh.Name & "' not found in workbook '" & source.Name & "'"
    End If

    ImportTranslationTables sourceTranslation
    ImportDictionaryFromWorkbook source, sourceTranslation
    ImportNamedValues sourceTranslation
End Sub
'@label:ExportSetup
'@sub-title Export dictionary translations for setup integration using the last imported name.
Public Sub ExportSetup(ByVal target As Workbook, _
                       Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)

    EnsureWorkbook target

    Dim sheetName As String
    sheetName = SetupSheetName()

    Dim setupSheet As Worksheet
    Set setupSheet = PrepareOutputSheet(target, sheetName, visibility)
    CopyDictionaryContent setupSheet
End Sub

'@label:ImportFromSetup
'@sub-title Import dictionary translations from a setup worksheet containing a single table.
Public Sub ImportFromSetup(ByVal sourceSheet As Worksheet)
    If sourceSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Setup worksheet is required for import"
    End If

    Dim sourceTable As ListObject
    Set sourceTable = LookupDictionaryListObject(sourceSheet)

    If sourceTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "No listobject found on setup worksheet '" & sourceSheet.Name & "'"
    End If

    ImportDictionaryTable sourceTable
    SetNameValue SETUP_SHEET_NAME_ID, sourceSheet.Name
End Sub


'@section Mapping helpers
'===============================================================================
Private Function TableMappings() As BetterArray
    If internalTableMappings Is Nothing Then
        Set internalTableMappings = New BetterArray
        internalTableMappings.LowerBound = 1
        internalTableMappings.Push Array(CByte(translationScopeDictionary), DICTIONARY_TABLE, "RNG_DictionaryLanguage")
        internalTableMappings.Push Array(CByte(translationScopeShapes), "T_TradLLShapes", "RNG_LLLanguageCode")
        internalTableMappings.Push Array(CByte(translationScopeForms), "T_TradLLForms", "RNG_LLLanguageCode")
        internalTableMappings.Push Array(CByte(translationScopeMessages), "T_TradLLMsg", "RNG_LLLanguageCode")
        internalTableMappings.Push Array(CByte(translationScopeRibbon), "T_TradLLRibbon", "RNG_LLLanguageCode")
    End If

    Set TableMappings = internalTableMappings
End Function

Private Function ValueMappings() As BetterArray
    If internalValueMappings Is Nothing Then
        Set internalValueMappings = New BetterArray
        internalValueMappings.LowerBound = 1
        internalValueMappings.Push Array("lllanguage", "RNG_LLLanguage")
        internalValueMappings.Push Array("lllanguagecode", "RNG_LLLanguageCode")
        internalValueMappings.Push Array("gotosection", "RNG_GoToSection")
        internalValueMappings.Push Array("anaperiod", "RNG_AnaPeriod")
        internalValueMappings.Push Array("gotoheader", "RNG_GoToHeader")
        internalValueMappings.Push Array("dictionarylanguage", "RNG_DictionaryLanguage")
        internalValueMappings.Push Array("nodevide", "RNG_NoDevide")
        internalValueMappings.Push Array("devide", "RNG_Devide")
        internalValueMappings.Push Array("gotograph", "RNG_GoToGraph")
        internalValueMappings.Push Array("onfiltered", "RNG_OnFiltered")
        internalValueMappings.Push Array("customdropdown", "RNG_CustomDrop")
        internalValueMappings.Push Array("uasheet", "RNG_UASheet")
        internalValueMappings.Push Array("tssheet", "RNG_TSSheet")
        internalValueMappings.Push Array("spsheet", "RNG_SPSheet")
        internalValueMappings.Push Array("sptsheet", "RNG_SPTSheet")
        internalValueMappings.Push Array("customtable", "RNG_CustomPivot")
        internalValueMappings.Push Array("week", "RNG_Week")
        internalValueMappings.Push Array("quarter", "RNG_Quarter")
        internalValueMappings.Push Array("infostart", "RNG_InfoStart")
        internalValueMappings.Push Array("infoend", "RNG_InfoEnd")
    End If

    Set ValueMappings = internalValueMappings
End Function

Private Function ResolveTableMapping(ByVal scope As Byte) As Variant
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant

    Set mappings = TableMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        If mapping(1) = scope Then
            ResolveTableMapping = mapping
            Exit Function
        End If
    Next idx

    ThrowError ProjectError.InvalidArgument, "Unsupported translation scope: " & CStr(scope)
End Function

Private Function TagToName(ByVal tagName As String) As String
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant
    Dim lowered As String

    lowered = LCase$(tagName)
    Set mappings = ValueMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        If mapping(1) = lowered Then
            TagToName = CStr(mapping(2))
            Exit Function
        End If
    Next idx

    ThrowError ProjectError.ElementNotFound, "Unknown translation value tag '" & tagName & "'"
End Function


'@section Workbook Exports
'===============================================================================
Private Function PrepareOutputSheet(ByVal target As Workbook, _
                                    ByVal sheetName As String, _
                                    ByVal visibility As XlSheetVisibility) As Worksheet

    Dim existing As Worksheet
    Set existing = ResolveWorksheet(target, sheetName)
    If Not existing Is Nothing Then
        ThrowError ProjectError.ElementShouldNotExists, _
                  "Sheet '" & sheetName & "' already exists in workbook '" & target.Name & "'"
    End If

    Dim newSheet As Worksheet
    Set newSheet = target.Worksheets.Add(After:=target.Worksheets(target.Worksheets.Count))
    newSheet.Name = sheetName
    newSheet.Visible = visibility

    Set PrepareOutputSheet = newSheet
End Function

Private Sub CopyTranslationContent(ByVal targetSheet As Worksheet)
    Dim sourceTable As ListObject

    For Each sourceTable In this.sh.ListObjects
        CopyTable sourceTable, targetSheet
    Next sourceTable

    CopyNamedValues targetSheet
End Sub

Private Sub CopyDictionaryContent(ByVal targetSheet As Worksheet)
    Dim sourceTable As ListObject
    Set sourceTable = FindListObject(this.sh, DICTIONARY_TABLE)

    If sourceTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Table '" & DICTIONARY_TABLE & "' not found on worksheet '" & this.sh.Name & "'"
    End If

    CopyTable sourceTable, targetSheet
    AssignDictionaryName targetSheet, FindListObject(targetSheet, sourceTable.Name)
    CopyNamedValues targetSheet
End Sub

Private Sub CopyTable(ByVal sourceTable As ListObject, ByVal targetSheet As Worksheet)
    Dim targetRange As Range
    Dim newTable As ListObject

    Set targetRange = targetSheet.Range(sourceTable.Range.Address)
    targetRange.Value = sourceTable.Range.Value
    targetRange.Formula = sourceTable.Range.Formula

    Set newTable = targetSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
    newTable.Name = sourceTable.Name
    ApplyTableStyle newTable, sourceTable.TableStyle
End Sub

Private Sub ApplyTableStyle(ByVal targetTable As ListObject, ByVal styleName As String)
    If LenB(styleName) = 0 Then Exit Sub

    On Error Resume Next
        targetTable.TableStyle = styleName
    On Error GoTo 0
End Sub

Private Sub CopyNamedValues(ByVal targetSheet As Worksheet)
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant
    Dim nameId As String

    Set mappings = ValueMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        nameId = CStr(mapping(2))
        AddOrUpdateNameConstant targetSheet, nameId, NameValue(nameId)
    Next idx

    AddOrUpdateNameConstant targetSheet, SETUP_SHEET_NAME_ID, NameValue(SETUP_SHEET_NAME_ID)
End Sub


'@section Workbook Imports
'===============================================================================
Private Sub ImportTranslationTables(ByVal sourceSheet As Worksheet)
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant
    Dim tableName As String
    Dim sourceTable As ListObject
    Dim targetTable As ListObject

    Set mappings = TableMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        tableName = CStr(mapping(2))

        If StrComp(tableName, DICTIONARY_TABLE, vbTextCompare) = 0 Then
            GoTo NextMapping
        End If

        Set sourceTable = FindListObject(sourceSheet, tableName)
        If sourceTable Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "Table '" & tableName & "' not found on worksheet '" & sourceSheet.Name & "'"
        End If

        Set targetTable = FindListObject(this.sh, tableName)
        ImportTable sourceTable, targetTable
NextMapping:
    Next idx
End Sub

Private Sub ImportDictionaryTable(ByVal sourceTable As ListObject)
    Dim targetTable As ListObject
    Set targetTable = FindListObject(this.sh, DICTIONARY_TABLE)

    ImportTable sourceTable, targetTable
    AssignDictionaryName this.sh, targetTable
End Sub

Private Sub ImportDictionaryFromWorkbook(ByVal sourceBook As Workbook, _
                                         ByVal translationSheet As Worksheet)
    Dim sourceTable As ListObject
    Set sourceTable = ResolveDictionaryTable(sourceBook, translationSheet)

    If sourceTable Is Nothing Then
        ClearDictionaryTable
    Else
        ImportDictionaryTable sourceTable
    End If
End Sub

Private Sub ClearDictionaryTable()
    Dim targetTable As ListObject
    Dim data As Range

    Set targetTable = FindListObject(this.sh, DICTIONARY_TABLE)
    If targetTable Is Nothing Then Exit Sub

    Set data = targetTable.DataBodyRange
    If Not data Is Nothing Then
        data.ClearContents
    End If
End Sub

Private Sub ImportTable(ByVal sourceTable As ListObject, ByVal targetTable As ListObject)
    Dim sourceCustom As ICustomTable
    Dim targetCustom As ICustomTable

    Set sourceCustom = CustomTable.Create(sourceTable)
    Set targetCustom = CustomTable.Create(targetTable)

    targetCustom.ImportAll sourceCustom
End Sub

Private Sub ImportNamedValues(ByVal sourceSheet As Worksheet)
    Dim mappings As BetterArray
    Dim idx As Long
    Dim mapping As Variant
    Dim nameId As String
    Dim sourceName As Name

    Set mappings = ValueMappings()

    For idx = mappings.LowerBound To mappings.UpperBound
        mapping = mappings.Item(idx)
        nameId = CStr(mapping(2))

        Set sourceName = TrySheetNameDefinition(sourceSheet, nameId)
        If Not sourceName Is Nothing Then
            AddOrUpdateNameConstant this.sh, nameId, NameDefinitionValue(sourceName)
        Else
            AddOrUpdateNameConstant this.sh, nameId, vbNullString
        End If
    Next idx

    Dim setupSource As Name
    Set setupSource = TrySheetNameDefinition(sourceSheet, SETUP_SHEET_NAME_ID)
    If Not setupSource Is Nothing Then
        AddOrUpdateNameConstant this.sh, SETUP_SHEET_NAME_ID, NameDefinitionValue(setupSource)
    Else
        EnsureSetupSheetName
    End If
End Sub

Private Function ResolveDictionaryTable(ByVal book As Workbook, _
                                        ByVal translationSheet As Worksheet) As ListObject
    Dim exportSheet As Worksheet
    Set exportSheet = ResolveWorksheet(book, DICTIONARY_EXPORT_SHEET)

    If Not exportSheet Is Nothing Then
        Set ResolveDictionaryTable = ResolveDictionaryTableOnSheet(exportSheet)
        If Not ResolveDictionaryTable Is Nothing Then Exit Function
    End If

    Set ResolveDictionaryTable = ResolveDictionaryTableOnSheet(translationSheet)
End Function

Private Function ResolveDictionaryTableOnSheet(ByVal sheet As Worksheet) As ListObject
    If sheet Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = FindListObject(sheet, DICTIONARY_TABLE)
    If Not lo Is Nothing Then
        Set ResolveDictionaryTableOnSheet = lo
        Exit Function
    End If

    Dim namedRange As Range
    Set namedRange = TrySheetNameRange(sheet, DICTIONARY_NAME)
    If Not namedRange Is Nothing Then
        Set ResolveDictionaryTableOnSheet = ListObjectByRange(sheet, namedRange)
    End If
End Function

Private Function LookupDictionaryListObject(ByVal sheet As Worksheet) As ListObject
    Dim namedRange As Range
    Set namedRange = TrySheetNameRange(sheet, DICTIONARY_NAME)

    If Not namedRange Is Nothing Then
        Set LookupDictionaryListObject = ListObjectByRange(sheet, namedRange)
        If Not LookupDictionaryListObject Is Nothing Then Exit Function
    End If

    Set LookupDictionaryListObject = FirstListObject(sheet)
End Function


'@section Name helpers
'===============================================================================
Private Function NameValue(ByVal nameId As String) As String
    Dim definition As Name
    Set definition = GetSheetNameDefinition(this.sh, nameId)
    NameValue = NameDefinitionValue(definition)
End Function

Private Sub SetNameValue(ByVal nameId As String, ByVal newValue As String)
    Dim definition As Name
    Set definition = GetSheetNameDefinition(this.sh, nameId)
    SetNameDefinition definition, newValue
End Sub

Private Function NameDefinitionValue(ByVal definition As Name) As String
    Dim targetRange As Range

    On Error Resume Next
        Set targetRange = definition.RefersToRange
    On Error GoTo 0

    If Not targetRange Is Nothing Then
        NameDefinitionValue = CStr(targetRange.Value)
    Else
        NameDefinitionValue = DecodeNamedValue(definition.RefersTo)
    End If
End Function

Private Sub SetNameDefinition(ByVal definition As Name, ByVal newValue As String)
    definition.RefersTo = EncodeNamedValue(newValue)
End Sub

Private Sub AddOrUpdateNameConstant(ByVal sheet As Worksheet, _
                                    ByVal nameId As String, _
                                    ByVal value As String)
    RemoveSheetName sheet, nameId
    sheet.Names.Add Name:=nameId, RefersTo:=EncodeNamedValue(value)
End Sub

Private Sub AssignDictionaryName(ByVal sheet As Worksheet, ByVal table As ListObject)
    If table Is Nothing Then Exit Sub
    RemoveSheetName sheet, DICTIONARY_NAME
    sheet.Names.Add Name:=DICTIONARY_NAME, RefersTo:=table.Range
End Sub

Private Sub RemoveSheetName(ByVal sheet As Worksheet, ByVal nameId As String)
    Dim idx As Long

    For idx = sheet.Names.Count To 1 Step -1
        If CompareName(sheet.Names(idx), nameId) Then
            sheet.Names(idx).Delete
        End If
    Next idx
End Sub

Private Function EncodeNamedValue(ByVal value As String) As String
    EncodeNamedValue = "=""" & Replace(value, """", """""") & """"
End Function

Private Function DecodeNamedValue(ByVal refersTo As String) As String
    Dim text As String
    text = refersTo

    If Len(text) > 0 And Left$(text, 1) = "=" Then
        text = Mid$(text, 2)
    End If

    If Len(text) >= 2 Then
        If Left$(text, 1) = """" And Right$(text, 1) = """" Then
            text = Mid$(text, 2, Len(text) - 2)
        End If
    End If

    DecodeNamedValue = Replace(text, """""", """")
End Function

Private Function TrySheetNameDefinition(ByVal sheet As Worksheet, ByVal nameId As String) As Name
    Dim definition As Name

    For Each definition In sheet.Names
        If CompareName(definition, nameId) Then
            Set TrySheetNameDefinition = definition
            Exit Function
        End If
    Next definition
End Function

Private Function GetSheetNameDefinition(ByVal sheet As Worksheet, ByVal nameId As String) As Name
    Dim definition As Name
    Set definition = TrySheetNameDefinition(sheet, nameId)

    If definition Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Name '" & nameId & "' not found on worksheet '" & sheet.Name & "'"
    End If

    Set GetSheetNameDefinition = definition
End Function


'@section Setup helpers
'===============================================================================
Private Function SetupSheetName() As String
    Dim stored As String

    stored = NameValue(SETUP_SHEET_NAME_ID)

    If LenB(stored) = 0 Then
        SetupSheetName = DICTIONARY_EXPORT_SHEET
    Else
        SetupSheetName = stored
    End If
End Function


'@section Resolution helpers
'===============================================================================
Private Function ResolveWorksheet(ByVal book As Workbook, ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In book.Worksheets
        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then
            Set ResolveWorksheet = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function FindListObject(ByVal sheet As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject

    For Each lo In sheet.ListObjects
        If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
            Set FindListObject = lo
            Exit Function
        End If
    Next lo
End Function

Private Function ListObjectByRange(ByVal sheet As Worksheet, ByVal targetRange As Range) As ListObject
    Dim lo As ListObject
    Dim address As String

    address = targetRange.Address(True, True, xlA1, False)

    For Each lo In sheet.ListObjects
        If StrComp(lo.Range.Address(True, True, xlA1, False), address, vbTextCompare) = 0 Then
            Set ListObjectByRange = lo
            Exit Function
        End If
    Next lo
End Function

Private Function FirstListObject(ByVal sheet As Worksheet) As ListObject
    If sheet.ListObjects.Count = 0 Then Exit Function
    Set FirstListObject = sheet.ListObjects(1)
End Function

Private Function TrySheetNameRange(ByVal sheet As Worksheet, ByVal nameId As String) As Range
    Dim definition As Name
    Dim cleanName As String

    For Each definition In sheet.Names
        cleanName = ExtractSimpleName(definition.Name)
        If StrComp(cleanName, nameId, vbTextCompare) = 0 Then
            On Error Resume Next
                Set TrySheetNameRange = definition.RefersToRange
            On Error GoTo 0
            Exit Function
        End If
    Next definition
End Function

Private Function ExtractSimpleName(ByVal qualifiedName As String) As String
    Dim exclPos As Long
    exclPos = InStr(qualifiedName, "!")
    If exclPos = 0 Then
        ExtractSimpleName = qualifiedName
    Else
        ExtractSimpleName = Mid$(qualifiedName, exclPos + 1)
    End If
End Function

Private Function CompareName(ByVal definition As Name, ByVal expected As String) As Boolean
    CompareName = (StrComp(ExtractSimpleName(definition.Name), expected, vbTextCompare) = 0)
End Function


'@section Validation helpers
'===============================================================================
Private Sub EnsureWorkbook(ByVal book As Workbook)
    If book Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference is missing"
    End If
End Sub

Private Sub EnsureListObjectExists(ByVal sheet As Worksheet, ByVal tableName As String)
    If FindListObject(sheet, tableName) Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Table '" & tableName & "' does not exist on worksheet '" & sheet.Name & "'"
    End If
End Sub


'@section Error handling
'===============================================================================
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, "Error: " & message
End Sub


'@section Interface
'===============================================================================
Private Property Get ILinelistTranslations_TransObject(Optional ByVal scope As Byte = translationScopeMessages) As ITranslationObject
    Set ILinelistTranslations_TransObject = TransObject(scope)
End Property

Private Property Get ILinelistTranslations_Value(ByVal tagName As String) As String
    ILinelistTranslations_Value = Value(tagName)
End Property

Private Sub ILinelistTranslations_SetValue(ByVal tagName As String, ByVal newValue As String)
    SetValue tagName, newValue
End Sub

Private Sub ILinelistTranslations_ExportToWorkbook(ByVal target As Workbook, _
                                                   Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)
    ExportToWorkbook target, visibility
End Sub

Private Sub ILinelistTranslations_ImportFromWorkbook(ByVal source As Workbook)
    ImportFromWorkbook source
End Sub

Private Sub ILinelistTranslations_ExportSetup(ByVal target As Workbook, _
                                              Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)
    ExportSetup target, visibility
End Sub

Private Sub ILinelistTranslations_ImportFromSetup(ByVal sourceSheet As Worksheet)
    ImportFromSetup sourceSheet
End Sub

