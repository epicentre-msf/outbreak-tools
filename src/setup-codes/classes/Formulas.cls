VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Convert formulas to excel"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Convert formulas to excel")
'@IgnoreModule IIfSideEffect, ExcelMemberMayReturnNothing

Option Explicit
Implements IFormulas

Private Type TFormulas
    Data As IFormulaData
    dict As ILLdictionary
    sheets As ILLSheets
    vars As ILLVariables
    setupForm As String
    testDone As Boolean
    validity As Boolean
    literals As Boolean
    formulavariables As BetterArray
    formulaTable As BetterArray
    ValidityReason As String
End Type

Private this As TFormulas
Private customFormulas As BetterArray
Private initialized As Boolean

'Instanciation
Public Function Create(ByVal dict As ILLdictionary, ByVal formData As IFormulaData, ByVal setupForm As String) As IFormulas
    With New Formulas
        Set .Dictionary = dict
        Set .Data = formData
        .SetupFormula = setupForm
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IFormulas
    Set Self = Me
End Property

'Properties
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Get Data() As IFormulaData
    Set Data = this.Data
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    Set this.vars = LLVariables.Create(dict)
    Set this.sheets = LLSheets.Create(dict)
    Set this.formulavariables = New BetterArray
    Set this.formulaTable = New BetterArray
    this.formulaTable.LowerBound = 1
    this.formulavariables.LowerBound = 1
    this.testDone = False
    initialized = False
End Property

Public Property Set Data(ByVal formData As IFormulaData)
    Set this.Data = formData
End Property

Public Property Get SetupFormula() As String
    SetupFormula = this.setupForm
End Property

Public Property Let SetupFormula(setupForm As String)
    this.setupForm = setupForm
End Property

Private Property Get Variables() As ILLVariables
    Set Variables = this.vars
End Property

Private Property Get FormulaSheets() As ILLSheets
    Set FormulaSheets = this.sheets
End Property

Private Property Get FormulaVariablesList() As BetterArray
    Set FormulaVariablesList = this.formulavariables.Clone()
End Property

Private Property Set FormulaVariablesList(ByVal baTable As BetterArray)
    Set this.formulavariables = baTable.Clone()
End Property

Private Property Get FormulaTableList() As BetterArray
    Set FormulaTableList = this.formulaTable.Clone()
End Property

Private Property Set FormulaTableList(ByVal baTable As BetterArray)
    Set this.formulaTable = baTable.Clone()
End Property

Private Property Get ValidityTestDone() As Boolean
    ValidityTestDone = this.testDone
End Property

Private Property Let ValidityTestDone(test As Boolean)
    this.testDone = test
End Property

Private Property Get FormulaValidity() As Boolean
    FormulaValidity = this.validity
End Property

Private Property Let FormulaValidity(test As Boolean)
    this.validity = test
End Property

Private Property Get reason(Optional ByVal formulaType As String = "analysis") As String
    If Valid(formulaType) Then this.ValidityReason = "The formula seems correct"

    reason = this.ValidityReason
End Property

Private Property Get ConvertedSetupFormula() As String
    Dim setupForm As String
    Dim casewhenObject As ICaseWhen
    Dim choiformObj As IChoiceFormula

    setupForm = SetupFormula()
    If (InStr(1, setupForm, "CASE_WHEN") = 1) Then
        'Case when custom formula, need to be parsed
        Set casewhenObject = CaseWhen.Create(setupForm)
        setupForm = casewhenObject.parsedFormula()

        'choice formula
    ElseIf (InStr(1, setupForm, "CHOICE_FORMULA") = 1) Then
        Set choiformObj = ChoiceFormula.Create(setupForm)
        setupForm = choiformObj.parsedFormula()
    End If

    ConvertedSetupFormula = setupForm
End Property

'Initialise custom formulas
Private Sub Initialize()
    Set customFormulas = New BetterArray
    customFormulas.Push "MEAN", "N", "N()"
    initialized = True
End Sub

'Has literals
Public Property Get HasLiterals() As Boolean
    HasLiterals = this.literals
End Property

Private Property Let HasLiterals(ByVal lit As Boolean)
    this.literals = lit
End Property

'Clean the formula and remove all eventual trailing spaces
Private Function Clean(chunk As String) As String
    Dim cleanedChunk As String
    cleanedChunk = Application.WorksheetFunction.SUBSTITUTE(chunk, Chr(160), " ")
    cleanedChunk = Application.WorksheetFunction.Clean(cleanedChunk)
    Clean = Application.WorksheetFunction.Trim(cleanedChunk)
End Function

'Test if a chunk is a parenthesis, or a quote
Private Function isQuoted(ByVal chunk As String) As Boolean
    isQuoted = (Mid(chunk, 1, 1) = Chr(34))
End Function

'Test if a chunk is an opened parenthesis
Private Function isOpenedParenthesis(ByVal chunk As String) As Boolean
    isOpenedParenthesis = (chunk = Chr(40))
End Function

Private Function isClosedParenthesis(ByVal chunk As String) As Boolean
    isClosedParenthesis = (chunk = Chr(41))
End Function

Private Property Get FormulaSheet() As String
    Dim dict As ILLdictionary
    Dim sheetLists As BetterArray
    Dim setupForm As String

    Set dict = Dictionary()
    setupForm = SetupFormula()

    FormulaSheet = vbNullString

    Set sheetLists = dict.Data.FilterData("control details", setupForm, "sheet name")
    
    If sheetLists.Length > 1 Then Debug.Print setupForm; " is present on more than 1 sheet "
    If sheetLists.Length = 1 Then FormulaSheet = sheetLists.Item(sheetLists.LowerBound)
End Property

'Test if a formula is valid
Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    Dim setupForm As String
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

    setupForm = ConvertedSetupFormula()
    nbParenthesis = 0
    QuotedCharacter = False
    previousBreak = 1

    setupForm = Clean(setupForm)
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
    If setupForm = vbNullString Then
        this.ValidityReason = "The formula is empty. No formula were found"
        GoTo FormulaIsNotValid
    ElseIf (vars.Contains(setupForm) And formulaType = "analysis") Then
        HasLiterals = True
        this.ValidityReason = "Analysis formula can not consist of only one variable, you should use aggregation function"
        GoTo FormulaIsNotValid
    ElseIf formulaType = "analysis" And customFormulas.Includes(setupForm) Then
        HasLiterals = True
        formulaTable.Push setupForm
        GoTo FormulaIsValid
    ElseIf (vars.Contains(setupForm)) Then       'simple type formula, we can have only varname
        VariablesTable.Push setupForm
        formulaTable.Push setupForm
        HasLiterals = True
        GoTo FormulaIsValid
    End If

    'Now Loop to the formula
    counter = 1
    Do While (counter <= Len(setupForm))
        midChar = Mid(setupForm, counter, 1)

        If isQuoted(midChar) Then QuotedCharacter = (Not QuotedCharacter)
        If isOpenedParenthesis(midChar) And (Not QuotedCharacter) Then nbParenthesis = nbParenthesis + 1
        If isClosedParenthesis(midChar) And (Not QuotedCharacter) Then nbParenthesis = nbParenthesis - 1

        If (formData.SpecialCharacterIncludes(midChar) And (Not QuotedCharacter)) Or (counter = Len(setupForm)) Then
            
            chunk = Clean(Mid(setupForm, previousBreak, counter - previousBreak))
            'Now you test for variables
            If (Not vars.Contains(chunk)) And _
                                          (Not formData.ExcelFormulasIncludes(UCase(chunk))) And _
                                          (chunk <> vbNullString) And _
                                          (Not IsNumeric(chunk)) And _
                                          (Not isQuoted(chunk)) And _
                                          (Not (formulaType = "analysis" And customFormulas.Includes(chunk))) Then
                'Write here to say why
                this.ValidityReason = "The formula parsing stoped at " & Mid(setupForm, 1, counter) & ": " & _
                                        chunk & " is unknown in formula " & setupForm
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
        Set varTable = this.formulavariables.Clone()
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

    ParsedLinelistFormula = formulaTable.ToString(Separator:=vbNullString, OpeningDelimiter:=vbNullString, _
                                                  ClosingDelimiter:=vbNullString, QuoteStrings:=False)
End Property

'Parse the formula for the analysis
Public Property Get ParsedAnalysisFormula(formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim TableName As String
    Dim formulaTable As BetterArray
    Dim counter As Long
    Dim tableLength As Long
    Dim parsedCustom As String
    Dim chunk As String


    Set vars = Variables()
    Set dict = Dictionary()

    If Valid(formulaType:="analysis") Then
        Set formulaTable = FormulaTableList
        counter = formulaTable.LowerBound
        tableLength = formulaTable.UpperBound

        Do While counter <= tableLength
            chunk = formulaTable.Item(counter)

            If vars.Contains(chunk) Then

                TableName = vars.Value(varName:=chunk, colName:="table name")
        
                If Not formCond.Valid(dict, TableName) Then Exit Property 'Trace
                formulaTable.Item(counter) = formCond.ConditionString(varName:=chunk, _
                                                                      TableName:=tablePrefix & TableName, Connector:=Connector)

            ElseIf customFormulas.Includes(chunk) Then
                parsedCustom = ParsedCustomFormula(customFunction:=chunk, _
                                                   formCond:=formCond, TableName:=TableName, tablePrefix:=tablePrefix)
                If parsedCustom = vbNullString Then Exit Property 'Trace
                formulaTable.Item(counter) = parsedCustom

            End If
            counter = counter + 1
        Loop
    End If

    ParsedAnalysisFormula = formulaTable.ToString(Separator:=vbNullString, OpeningDelimiter:=vbNullString, _
                                                  ClosingDelimiter:=vbNullString, QuoteStrings:=False)
End Property

Private Function Quoted(ByVal val As String)
    Quoted = Chr(34) & val & Chr(34)
End Function

Private Function ClearCountIf(ByVal val As String)
    Dim cleared As String
    cleared = Replace(val, Quoted("=") & "&", vbNullString)
    cleared = Replace(val, "=", vbNullString)
    cleared = Replace(cleared, Chr(34) & Chr(34) & Chr(34) & Chr(34), Chr(34) & Chr(34))
    cleared = Replace(cleared, "<>", "######")
    cleared = IIf(InStr(1, cleared, "<") > 0, Replace(cleared, "<", Quoted("<=") & " & "), cleared)
    cleared = IIf(InStr(1, cleared, ">") > 0, Replace(cleared, ">", Quoted(">=") & " & "), cleared)
    cleared = Replace(cleared, "######", "<>")
    ClearCountIf = cleared
End Function

Private Function ParsedCustomFormula(ByVal customFunction As String, ByVal formCond As IFormulaCondition, _
                                     Optional ByVal TableName As String = vbNullString, _
                                     Optional ByVal tablePrefix As String = vbNullString) As String
    Dim cond As BetterArray
    Dim condVars As BetterArray
    Dim parsedFormula As String
    Dim customtableName As String
    Dim dict As ILLdictionary
    Dim counter As Long

    Set dict = Dictionary()
    customtableName = TableName
    If TableName = vbNullString Then customtableName = formCond.VariablesTable(dict)
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
Private Property Get IFormulas_Varlists() As BetterArray
    Set IFormulas_Varlists = Varlists()
End Property

Private Property Get IFormulas_Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    IFormulas_Valid = Valid(formulaType)
End Property

Private Property Get IFormulas_FormulaTableList() As BetterArray
    Set IFormulas_FormulaTableList = FormulaTableList
End Property

Private Property Get IFormulas_ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                                    Optional ByVal tablePrefix As String = vbNullString) As String
    IFormulas_ParsedLinelistFormula = ParsedLinelistFormula(useTableName, tablePrefix)
End Property

Private Property Get IFormulas_ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, Optional ByVal tablePrefix As String = vbNullString, _
                                                    Optional ByVal Connector As String = "*") As String
    IFormulas_ParsedAnalysisFormula = ParsedAnalysisFormula(formCond, tablePrefix, Connector)
End Property

Private Property Get IFormulas_HasLiterals() As Boolean
    IFormulas_HasLiterals = HasLiterals()
End Property

Private Property Get IFormulas_Reason(Optional ByVal formulaType As String = "analysis") As String
    IFormulas_Reason = reason(formulaType)
End Property


