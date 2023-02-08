VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Formulas")
'@ModuleDescription("Formula Object")
'@IgnoreModule

Option Explicit
Implements IFormulas

Private Type TFormulas
    Data As IFormulaData
    dict As ILLdictionary
    sheets As ILLSheets
    vars As ILLVariables
    setupform As String
    testDone As Boolean
    validity As Boolean
    literals As Boolean
    formulavariables As BetterArray
    formulaTable As BetterArray
End Type

Private This As TFormulas
Private customFormulas As BetterArray
Private initialized As Boolean

'Instanciation
Public Function Create(ByVal dict As ILLdictionary, ByVal formData As IFormulaData, ByVal setupform As String) As IFormulas
    With New Formulas
        Set .Dictionary = dict
        Set .Data = formData
        .SetupFormula = setupform
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IFormulas
    Set Self = Me
End Property

'Properties
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = This.dict
End Property

Public Property Get Data() As IFormulaData
    Set Data = This.Data
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set This.dict = dict
    Set This.vars = LLVariables.Create(dict)
    Set This.sheets = LLSheets.Create(dict)
    Set This.formulavariables = New BetterArray
    Set This.formulaTable = New BetterArray
    This.formulaTable.LowerBound = 1
    This.formulavariables.LowerBound = 1
    This.testDone = False
    initialized = False
End Property

Public Property Set Data(ByVal formData As IFormulaData)
    Set This.Data = formData
End Property

Public Property Get SetupFormula() As String
    SetupFormula = This.setupform
End Property

Public Property Let SetupFormula(setupform As String)
    This.setupform = setupform
End Property

Private Property Get Variables() As ILLVariables
    Set Variables = This.vars
End Property

Private Property Get FormulaSheets() As ILLSheets
    Set FormulaSheets = This.sheets
End Property

Private Property Get FormulaVariablesList() As BetterArray
    Set FormulaVariablesList = This.formulavariables.Clone()
End Property

Private Property Set FormulaVariablesList(ByVal baTable As BetterArray)
    Set This.formulavariables = baTable.Clone()
End Property

Private Property Get FormulaTableList() As BetterArray
    Set FormulaTableList = This.formulaTable.Clone()
End Property

Private Property Set FormulaTableList(ByVal baTable As BetterArray)
    Set This.formulaTable = baTable.Clone()
End Property

Private Property Get ValidityTestDone() As Boolean
    ValidityTestDone = This.testDone
End Property

Private Property Let ValidityTestDone(test As Boolean)
    This.testDone = test
End Property

Private Property Get FormulaValidity() As Boolean
    FormulaValidity = This.validity
End Property

Private Property Let FormulaValidity(test As Boolean)
    This.validity = test
End Property

Private Property Get ConvertedSetupFormula() As String
    Dim setupform As String
    Dim casewhenObject As ICaseWhen
    Dim choiformObj As IChoiceFormula

    setupform = SetupFormula()
    If (InStr(1, setupform, "CASE_WHEN") = 1) Then
        'Case when custom formula, need to be parsed
        Set casewhenObject = CaseWhen.Create(setupform)
        setupform = casewhenObject.parsedFormula()

        'choice formula
    ElseIf (InStr(1, setupform, "CHOICE_FORMULA") = 1) Then
        Set choiformObj = ChoiceFormula.Create(setupform)
        setupform = choiformObj.parsedFormula()
    End If

    ConvertedSetupFormula = setupform
End Property

'Initialise custom formulas
Private Sub Initialize()
    Set customFormulas = New BetterArray
    customFormulas.Push "MEAN", "N", "N()"
    initialized = True
End Sub

'Has literals
Public Property Get HasLiterals() As Boolean
    HasLiterals = This.literals
End Property

Private Property Let HasLiterals(ByVal lit As Boolean)
    This.literals = lit
End Property

'Clean the formula and remove all eventual trailing spaces
Private Function Clean(chunk As String) As String
    Dim cleanedChunk As String
    cleanedChunk = Application.WorksheetFunction.SUBSTITUTE(chunk, chr(160), " ")
    cleanedChunk = Application.WorksheetFunction.Clean(cleanedChunk)
    Clean = Application.WorksheetFunction.Trim(cleanedChunk)
End Function

'Test if a chunk is a parenthesis, or a quote
Private Function isQuoted(ByVal chunk As String) As Boolean
    isQuoted = (Mid(chunk, 1, 1) = chr(34))
End Function

'Test if a chunk is an opened parenthesis
Private Function isOpenedParenthesis(ByVal chunk As String) As Boolean
    isOpenedParenthesis = (chunk = chr(40))
End Function

Private Function isClosedParenthesis(ByVal chunk As String) As Boolean
    isClosedParenthesis = (chunk = chr(41))
End Function

Private Property Get FormulaSheet() As String
    Dim dict As ILLdictionary
    Dim sheetLists As BetterArray
    Dim setupform As String

    Set dict = Dictionary()
    setupform = SetupFormula()

    FormulaSheet = vbNullString

    Set sheetLists = dict.Data.FilterData("control details", setupform, "sheet name")
    
    If sheetLists.Length > 1 Then Debug.Print setupform; " is present on more than 1 sheet "
    If sheetLists.Length = 1 Then FormulaSheet = sheetLists.Item(sheetLists.LowerBound)
End Property

'Test if a formula is valid
Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    Dim setupform As String
    Dim chunk As String
    Dim midChar As String
    Dim vars As ILLVariables
    Dim formData As IFormulaData
    Dim counter As Long
    Dim nbParenthesis As Long
    Dim QuotedCharacter As Boolean
    Dim previousBreak As Long
    Dim formulaTable As BetterArray
    Dim VariablesTable As BetterArray

    setupform = ConvertedSetupFormula()
    nbParenthesis = 0
    QuotedCharacter = False
    previousBreak = 1

    setupform = Clean(setupform)
    Set vars = Variables()
    Set formData = Data()

    Set formulaTable = New BetterArray
    Set VariablesTable = New BetterArray

    formulaTable.LowerBound = 1
    VariablesTable.LowerBound = 1

    If (Not initialized) Then Initialize

    'If the test is already done return the value
    If ValidityTestDone Then
        Valid = FormulaValidity
        Exit Property
    End If
    
    'Initialize the literals (testing if the formula contains literals)
    HasLiterals = False
    
    'if it is a variable and we are on analysis, it is not valid, we need aggregation
    'If the formula is empty validity is false
    If setupform = vbNullString Then
        GoTo FormulaIsNotValid
    ElseIf (vars.Contains(setupform) And formulaType = "analysis") Then
        HasLiterals = True
        GoTo FormulaIsNotValid
    ElseIf formulaType = "analysis" And customFormulas.Includes(setupform) Then
        HasLiterals = True
        formulaTable.Push setupform
        GoTo FormulaIsValid
    ElseIf (vars.Contains(setupform)) Then       'simple type formula, we can have only varname
        VariablesTable.Push setupform
        formulaTable.Push setupform
        HasLiterals = True
        GoTo FormulaIsValid
    End If

    'Now Loop to the formula
    counter = 1
    Do While (counter <= Len(setupform))
        midChar = Mid(setupform, counter, 1)

        If isQuoted(midChar) Then QuotedCharacter = (Not QuotedCharacter)
        If isOpenedParenthesis(midChar) And (Not QuotedCharacter) Then nbParenthesis = nbParenthesis + 1
        If isClosedParenthesis(midChar) And (Not QuotedCharacter) Then nbParenthesis = nbParenthesis - 1

        If (formData.SpecialCharacterIncludes(midChar) And (Not QuotedCharacter)) Or (counter = Len(setupform)) Then
            
            chunk = Clean(Mid(setupform, previousBreak, counter - previousBreak))
            'Now you test for variables
            If (Not vars.Contains(chunk)) And _
                                          (Not formData.ExcelFormulasIncludes(UCase(chunk))) And _
                                          (chunk <> vbNullString) And _
                                          (Not IsNumeric(chunk)) And _
                                          (Not isQuoted(chunk)) And _
                                          (Not (formulaType = "analysis" And customFormulas.Includes(chunk))) Then
                'Write here to say why
                'Debug.Print chunk & " is unknown in formula " & setupform
                'Error because we don't know the meaning of chunk
                
                GoTo FormulaIsNotValid
            ElseIf vars.Contains(chunk) Then
                HasLiterals = True
                VariablesTable.Push chunk
                formulaTable.Push chunk, midChar
            Else
                If formData.ExcelFormulasIncludes(chunk) Then HasLiterals = True
                formulaTable.Push chunk, midChar
            End If

            previousBreak = counter + 1
        End If
        'Test for the number of parenthesis
        If nbParenthesis < 0 Then GoTo FormulaIsNotValid
        counter = counter + 1
    Loop

    If nbParenthesis > 0 Then GoTo FormulaIsNotValid

FormulaIsValid:
    ValidityTestDone = True
    FormulaValidity = True
    Set FormulaTableList = formulaTable
    Set FormulaVariablesList = VariablesTable
    Valid = FormulaValidity
    Exit Property

FormulaIsNotValid:
    ValidityTestDone = True
    FormulaValidity = False
    Valid = FormulaValidity
End Property

'List of formulas
Public Property Get Varlists(Optional ByVal formulaType As String = "analysis") As BetterArray
    Dim varTable As BetterArray
    Set varTable = New BetterArray
    varTable.LowerBound = 1
    If Valid(formulaType) Then
        Set varTable = This.formulavariables.Clone()
    End If
    Set Varlists = varTable.Clone()
End Property

'Parse the formula for linelist
Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
    Dim vars As ILLVariables
    Dim sheets As ILLSheets
    Dim counter As Long
    Dim formulaTable As BetterArray
    Dim chunk As String
    Dim formshName As String

    Set formulaTable = New BetterArray
    Set vars = Variables()
    Set sheets = FormulaSheets()

    ParsedLinelistFormula = vbNullString

    If Not Valid(formulaType:="simple") Then Exit Property

    Set formulaTable = FormulaTableList
    For counter = formulaTable.LowerBound To formulaTable.UpperBound
        chunk = formulaTable.Item(counter)
        If vars.Contains(chunk) Then
            If useTableName Then
                formulaTable.Item(counter) = tablePrefix & vars.Value(colName:="table name", varName:=chunk) & "[" & chunk & "]"
            Else
                formshName = FormulaSheet()
                formulaTable.Item(counter) = sheets.VariableAddress(chunk, formshName)
            End If
        ElseIf UCase(chunk) = "MEAN" Then
            formulaTable.Item(counter) = "AVERAGE"
        End If
    Next

    ParsedLinelistFormula = formulaTable.ToString(Separator:="", OpeningDelimiter:="", _
                                                  ClosingDelimiter:="", QuoteStrings:=False)
End Property

'Parse the formula for the analysis
Public Property Get ParsedAnalysisFormula(formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
    Dim vars As ILLVariables
    Dim sheets As ILLSheets
    Dim dict As ILLdictionary
    Dim tableName As String
    Dim formulaTable As BetterArray
    Dim counter As Long
    Dim tableLength As Long
    Dim parsedCustom As String
    Dim chunk As String


    Set vars = Variables()
    Set sheets = FormulaSheets()
    Set dict = Dictionary()

    If Valid(formulaType:="analysis") Then
        Set formulaTable = FormulaTableList
        counter = formulaTable.LowerBound
        tableLength = formulaTable.UpperBound

        Do While counter <= tableLength
            chunk = formulaTable.Item(counter)

            If vars.Contains(chunk) Then

                tableName = vars.Value(varName:=chunk, colName:="table name")
        
                If Not formCond.Valid(dict, tableName) Then Exit Property 'Trace
                formulaTable.Item(counter) = formCond.ConditionString(varName:=chunk, _
                                                                      tableName:=tablePrefix & tableName, Connector:=Connector)

            ElseIf customFormulas.Includes(chunk) Then
                parsedCustom = ParsedCustomFormula(customFunction:=chunk, _
                                                   formCond:=formCond, tableName:=tableName, tablePrefix:=tablePrefix)
                If parsedCustom = vbNullString Then Exit Property 'Trace
                formulaTable.Item(counter) = parsedCustom

            End If
            counter = counter + 1
        Loop
    End If

    ParsedAnalysisFormula = formulaTable.ToString(Separator:="", OpeningDelimiter:="", _
                                                  ClosingDelimiter:="", QuoteStrings:=False)
End Property

Private Function Quoted(ByVal val As String)
    Quoted = chr(34) & val & chr(34)
End Function

Private Function ClearCountIf(ByVal val As String)
    Dim cleared As String
    cleared = Replace(val, Quoted("=") & "&", vbNullString)
    cleared = Replace(val, "=", vbNullString)
    cleared = Replace(cleared, chr(34) & chr(34) & chr(34) & chr(34), chr(34) & chr(34))
    cleared = Replace(cleared, "<>", "######")
    cleared = IIf(InStr(1, cleared, "<") > 0, Replace(cleared, "<", Quoted("<=") & " & "), cleared)
    cleared = IIf(InStr(1, cleared, ">") > 0, Replace(cleared, ">", Quoted(">=") & " & "), cleared)
    cleared = Replace(cleared, "######", "<>")
    ClearCountIf = cleared
End Function

Private Function ParsedCustomFormula(ByVal customFunction As String, ByVal formCond As IFormulaCondition, _
                                     Optional ByVal tableName As String = vbNullString, _
                                     Optional ByVal tablePrefix As String = vbNullString) As String
    Dim cond As BetterArray
    Dim condVars As BetterArray
    Dim parsedFormula As String
    Dim customtableName As String
    Dim dict As ILLdictionary
    Dim counter As Long

    Set dict = Dictionary()
    customtableName = tableName
    If tableName = vbNullString Then customtableName = formCond.VariablesTable(dict)
    customtableName = tablePrefix & customtableName
  

    Select Case customFunction

    Case "N", "N()"
        Set cond = formCond.Conditions()
        Set condVars = formCond.Variables()
        counter = condVars.LowerBound
        parsedFormula = customtableName & "[" & condVars.Item(counter) & "]" & ", " & ClearCountIf(cond.Item(counter))

        For counter = condVars.LowerBound + 1 To condVars.UpperBound
            parsedFormula = parsedFormula & "," & customtableName & "[" & condVars.Item(counter) & "]" & "," & ClearCountIf(cond.Item(counter))
        Next
        parsedFormula = "COUNTIFS(" & parsedFormula & ")"

    Case "MEAN"
        parsedFormula = "AVERAGE"
    Case Else
        parsedFormula = customFunction
    End Select

    ParsedCustomFormula = parsedFormula
End Function

'Interfaces
Public Property Get IFormulas_Varlists() As BetterArray
    Set IFormulas_Varlists = Varlists()
End Property

Public Property Get IFormulas_Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    IFormulas_Valid = Valid(formulaType)
End Property

Private Property Get IFormulas_FormulaTableList() As BetterArray
    Set IFormulas_FormulaTableList = FormulaTableList
End Property

Public Property Get IFormulas_ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                                    Optional ByVal tablePrefix As String = vbNullString) As String
    IFormulas_ParsedLinelistFormula = ParsedLinelistFormula(useTableName, tablePrefix)
End Property

Public Property Get IFormulas_ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, Optional ByVal tablePrefix As String = vbNullString, _
                                                    Optional ByVal Connector As String = "*") As String
    IFormulas_ParsedAnalysisFormula = ParsedAnalysisFormula(formCond, tablePrefix, Connector)
End Property

Public Property Get IFormulas_HasLiterals() As Boolean
    IFormulas_HasLiterals = HasLiterals()
End Property


