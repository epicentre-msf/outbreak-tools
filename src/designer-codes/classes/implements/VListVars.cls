VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VListVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Vlist sheet")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Implements IVListVars

Private Type TVListVars
    ll As ILinelist
    varName As String
    inputRange As Range
End Type

Private this As TVListVars
Private Const CLASSNAME As String = "VListVars"
Private Const STARTCOL As Long = 5               'Where the variables should start in the output worksheet

Public Function Create(ByVal varName As String, ByVal ll As ILinelist) As IVListVars

    CheckRequirements varName, ll

    With New VListVars
        Set .LinelistObject = ll
        .ActualVariable = varName
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IVListVars
    Set Self = Me
End Property

Public Property Get LinelistObject() As ILinelist
    Set LinelistObject = this.ll
End Property

Public Property Set LinelistObject(ByVal ll As ILinelist)
    Set this.ll = ll
End Property

Public Property Get ActualVariable() As String
    ActualVariable = this.varName
End Property

Public Property Let ActualVariable(ByVal varName As String)
    this.varName = varName
End Property

'The linelist Data
Private Property Get LinelistData() As ILinelistSpecs
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set LinelistData = ll.LinelistData()
End Property

'Get the sheets object
Private Property Get SheetsObject() As ILLSheets
    Dim dict As ILLdictionary
    Set dict = Dictionary()
    Set SheetsObject = LLSheets.Create(dict)
End Property

'Get the worksheet Object where the variable should be written
Public Property Get Wksh() As Worksheet
    Dim ll As ILinelist
    Dim sheetName As String

    Set ll = LinelistObject()
    sheetName = ValueOf("sheet name")
    Set Wksh = ll.Wksh(sheetName)
End Property

'The dropdown object as part of the linelist
Public Property Get Dropdown() As IDropdownLists
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set Dropdown = ll.Dropdown()
End Property

Private Property Get CustomDropdown() As IDropdownLists
    Dim ll As ILinelist

    Set ll = LinelistObject()
    Set CustomDropdown = ll.Dropdown(scope:=2)
End Property

Private Property Get Dictionary() As ILLdictionary
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set Dictionary = ll.Dictionary
End Property

Private Property Get vars() As ILLVariables
    Dim dict As ILLdictionary
    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)
End Property

Private Property Get ValueOf(ByVal colName As String) As String
    Dim varObject As ILLVariables
    Set varObject = vars()
    ValueOf = varObject.Value(colName:=colName, varName:=ActualVariable)
End Property

Private Property Get VarRange() As Range         'Where to input informations on the variable
    Dim StartRow As Long
    Dim sh As Worksheet
    Dim rng As Range

    If this.inputRange Is Nothing Then
        Set sh = Wksh()
        With sh
            StartRow = ValueOf("column index")
            Set this.inputRange = .Cells(StartRow, STARTCOL)
        End With
    End If

    Set VarRange = this.inputRange
End Property

'Type using type and format of the variable
Private Sub AddType()

    Dim varTypes As String
    Dim varFormat As String
    Dim cellRng As Range
    Dim nbDecimals As Integer
    Dim zerosDecimals As String
    Dim formatText As String

    Set cellRng = VarRange()

    varTypes = LCase(Application.WorksheetFunction.Trim(ValueOf("variable type")))
    varFormat = LCase(Application.WorksheetFunction.Trim(ValueOf("variable format")))
    nbDecimals = 1

    'Find the format depending on the Types and format
    Select Case varTypes

    Case "text"

        cellRng.NumberFormat = "@"

    Case "date"

        'Apply default type before changing
        cellRng.NumberFormat = "d-mmm-yyyy"

        On Error Resume Next
        If varFormat <> vbNullString Then cellRng.NumberFormat = varFormat
        On Error GoTo 0

    Case "decimal", "integer"

        If varFormat = vbNullString Then

            'Default format if there is nothing
            formatText = IIf(varTypes = "decimal", "#,##0.00;-#,##0.00;0.00;@", _
                             "###0;-###0;0;@")

        'Other formats
        ElseIf varFormat = "integer" Then

            formatText = "###0;-###0;0;@"

        ElseIf varFormat = "round0" Then

            formatText = "#,##0;-#,##0;0;@"

        ElseIf varFormat = "percentage0" Then

            formatText = "0%"

        ElseIf (InStr(1, varFormat, "round") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "round", vbNullString))
            On Error GoTo 0

            zerosDecimals = String(nbDecimals, "0")
            formatText = "#,###0." & zerosDecimals & ";-#,###0." & _
                         zerosDecimals & ";0." & zerosDecimals & ";@"

        ElseIf (InStr(1, varFormat, "percentage") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "percentage", vbNullString))
            On Error GoTo 0
            zerosDecimals = String(nbDecimals, "0")
            formatText = "0." & zerosDecimals & "%"

        ElseIf varFormat = "euros" Then

            formatText = "#,##0.00 [$€-x-euro1];-#,##0.00 [$€-x-euro1];0.00 [$€-x-euro1];@"

        ElseIf varFormat = "dollars" Then

            formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"

        ElseIf varFormat <> "text" Then

            formatText = "custom"

        End If

        'Define the format
        If formatText <> "custom" Then
            cellRng.NumberFormat = formatText
        Else
            On Error Resume Next
            cellRng.NumberFormat = varFormat
            On Error GoTo 0
        End If

    End Select
End Sub

'Failed formula
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim sh As Worksheet
    Dim prevAlert As Boolean

    Set sh = Wksh()
    With sh
        .Cells(1, 1).Clear
        On Error GoTo Fail
        prevAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False

        'Test the formula on one cell
        .Cells(1, 1).formula = "= " & frmName

        .Cells(1, 1).Clear
        FailedFormula = False
        Exit Function
    End With

    Application.DisplayAlerts = prevAlert
Fail:
    FailedFormula = True
    Application.DisplayAlerts = prevAlert
End Function

'Formula
Private Sub AddFormula()
    Dim varForm As String
    Dim excelForm As String
    Dim cellRng As Range
    Dim lData As ILinelistSpecs
    Dim varFormData As IFormulaData
    Dim dict As ILLdictionary
    Dim varFormObject As IFormulas
    Dim DesignFormat As ILLFormat
    Dim varControl As String

    varControl = ValueOf("control")
    If (varControl <> "formula") And (varControl <> "case_when") And _
     (varControl <> "choice_formula") Then Exit Sub

    Set lData = LinelistData()
    Set dict = lData.Dictionary()
    Set varFormData = lData.FormulaDataObject()
    Set DesignFormat = lData.DesignFormat()
    varForm = ValueOf("control details")
    Set cellRng = VarRange()

    Set varFormObject = Formulas.Create(dict, varFormData, varForm)

    If varFormObject.Valid(formulaType:="linelist") Then
        excelForm = varFormObject.ParsedLinelistFormula()
        'Test if the formula is correct and parse it
        If FailedFormula(excelForm) Then
            excelForm = chr(34) & "##== formula parsing failed :( ==##" & chr(34)
        End If
        cellRng.formula = "= " & excelForm
        cellRng.Locked = True
    End If
    'Apply format of the calculated range
    DesignFormat.ApplyFormat cellRng, scope:=HListCalculatedFormulaCell
End Sub

'Conditional formatting on variables
Private Sub AddConditionalFormatting()
    Dim varValid As String 'variable (dummies for validation, should contain 1 Or 0)
    Dim cellRng As Range 'Range of the variable
    Dim formatRng As Range 'Range with values defined for the validation format
    Dim varObject As ILLVariables 'List of variables
    Dim shObj As ILLSheets
    Dim condAdd As String 'condition variable address

    varValid = ValueOf("formatting condition")
    Set varObject = vars()

    If (varValid = vbNullString) Then Exit Sub
    If Not varObject.Contains(varValid) Then Exit Sub
    Set cellRng = VarRange()
    Set formatRng = varObject.CellRange(colName:="formatting values", _
                                         varName:=ActualVariable)
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid)

    cellRng.FormatConditions.Add Type:=xlExpression, Formula1:="=(" & condAdd & "= 1)"
    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.color = formatRng.Interior.color
        .Font.color = formatRng.Font.color
        .Font.Bold = formatRng.Font.Bold
        .Font.Italic = formatRng.Font.Italic
        'Stop if this condition is met (This is the first priority)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'Main Label
Private Sub AddLabel(Optional ByVal notevisibility As Boolean = False)
    Dim varMainLab As String
    Dim varIndex As Long
    Dim varName As String
    Dim varNote As String
    Dim varSublab As String
    Dim varStatus As String
    Dim varListAutoOrigin As String
    Dim varControl As String
    Dim cellRng As Range
    Dim trans As ITranslation
    Dim lData As ILinelistSpecs
    Dim calculated As String
    Dim mandatory As String
    Dim mainLabLength As Long
    Dim design As ILLFormat
    Dim varEditLab As String


    Set cellRng = VarRange()
    Set lData = LinelistData()
    Set trans = lData.TransObject()
    Set design = lData.DesignFormat()

    varMainLab = ValueOf("main label")
    varName = ValueOf("variable name")
    varNote = ValueOf("note")
    varSublab = ValueOf("sub label")
    varControl = ValueOf("control")
    varStatus = ValueOf("status")
    varIndex = CLng(ValueOf("column index"))
    varEditLab = ValueOf("editable label")

    calculated = trans.TranslatedValue("MSG_Calculated")
    mandatory = trans.TranslatedValue("MSG_Mandatory")

    'Add the table start name
    If varIndex = 4 Then cellRng.Name = ValueOf("table name") & "_START"

    'sublabel
    If (varControl = "formula") Or (varControl = "case_when") Or _
      (varControl = "choice_formula") Then
        varSublab = varSublab & chr(10) & calculated
    End If

    mainLabLength = Len(varMainLab)

    'Add subLabel to main label
    If (varSublab <> vbNullString) Then varMainLab = varMainLab & chr(10) & varSublab

    'Status to add to Notes
    If varStatus = "mandatory" Then varNote = mandatory & chr(10) & varNote

    'variable main label and sub lable
    cellRng.Offset(, -1).Value = varMainLab

    If varControl <> "choice_custom" Then
        design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=1, Length:=mainLabLength), _
        scope:=VListMainLab                      'format of the main label
        design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=mainLabLength + 1), _
        scope:=VListSublab                       'format of the sub label
    End If

    'variable name
    cellRng.Name = varName


    'variable note
    If varNote <> vbNullString Then
        cellRng.AddComment
        cellRng.Comment.Text Text:=varNote
        cellRng.Comment.Visible = notevisibility
    End If

    'Hidden Row
    If (varStatus = "hidden") Or (varStatus = "optional, hidden") Then 
        cellRng.EntireRow.HIDDEN = True
    End If
    
    cellRng.Locked = False
    'On Editable label = yes, unlock the cell
    If (varEditLab = "yes") Then cellRng.Offset(, -1).Locked = False
    cellRng.Offset(, -2).Locked = True
    cellRng.Offset(, -3).Locked = True

    'Adding variable control for future usage
    cellRng.Offset(, 1).Value = varControl
    design.ApplyFormat cellRng.Offset(, 1), scope:=LinelistHiddenCell
End Sub

'Add the choices
Private Sub AddChoices()
    Dim varControl As String
    Dim varName As String
    Dim varAlert As String
    Dim varMessage As String

    Dim lData As ILinelistSpecs
    Dim cat As BetterArray
    Dim drop As IDropdownLists
    Dim cellRng As Range
    Dim choicust As String
    Dim trans As ITranslation
    Dim varMainLab As String
    Dim varSubLab As String
    Dim sh As Worksheet
    Dim labRng As Range
    Dim targetAddress As String
    Dim mainLabLength As Long
    Dim design As ILLFormat

    
    varControl = ValueOf("control")

    If (varControl <> "choice_manual") And _
        (varControl <> "choice_custom") And (varControl <> "choice_multiple") _
        Then Exit Sub
    
    varName = ActualVariable()
    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")

    Set lData = LinelistData()
    Set cat = New BetterArray
    Set cellRng = VarRange()

    Select Case varControl

    Case "choice_manual", "choice_multiple"

        Set drop = Dropdown()
        Set cat = lData.Categories(varName)
        'Add dropdown
        drop.Add cat, varName

    Case "choice_custom"

        Set drop = CustomDropdown()
        Set cat = lData.Categories(varName)
        Set trans = lData.TransObject()
        Set design = lData.DesignFormat()
        Set sh = cellRng.Parent
        
        varSubLab = ValueOf("sub label")
        varMainLab = ValueOf("main label")
        mainLabLength = Len(varMainLab)
        
        If cat.Length = 0 Then cat.Push vbNullString, vbNullString

        choicust = trans.TranslatedValue("MSG_CustomChoice")
        drop.Add cat, varName, addLabel:=True, counterPrefix:=choicust
    
        'For choice customs, special treatment to add the main label and sub label after
        'adding a link
        'If the labRng is Nothing, just continue
        On Error Resume Next
            Set labRng = drop.LabelRange(varName)
            targetAddress = "'" & labRng.Parent.Name & "'" & "!" & labRng.Address
            'Adding hyperlinks remove the writings (sub label and so on, so need to set it back)
            With sh
                .Hyperlinks.Add _ 
                Anchor:=cellRng.Offset(, -1), _ 
                Address:="", _
                SubAddress:= targetAddress, _
                TextToDisplay:= ""
            End With
        On Error GoTo 0

        cellRng.Offset(, -1).Font.UnderLine = xlUnderlineStyleNone

        'Add the custom dropdown label value (including the number) to sub label
        varSubLab = varSubLab & chr(10) & "(" & labRng.Value & ")"
        varMainLab = varMainLab & chr(10) & varSublab

        cellRng.Offset(, -1).Value = varMainLab
        'Set the style back for main label and sub label (could be removed)
        design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=1, Length:=mainLabLength), _
        scope:=VListMainLab                      'format of the main label
        design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=mainLabLength + 1), _
        scope:=VListSublab

        'Add a return link to the vlist var worksheet
        drop.AddReturnLink listName:=varName, returnRng:=cellRng.Offset(, -1)
    End Select

    'Set dropdown validations
    drop.SetValidation cellRng:=cellRng, listName:=varName, _
                       alertType:=varAlert, message:=varMessage, _ 
                       showError:= (varControl <> "choice_multiple")
End Sub

'Add validation for VList var
Private Sub AddValidation()

    Dim varMin As String
    Dim varMax As String
    Dim varMessage As String
    Dim varAlert As String
    Dim varTypes As String                       'Something weird but we can't name a variable varType, so the s is important
    Dim validationTypes As Long
    Dim validationAlerts As Long
    Dim excelForm As Variant
    Dim excelOSForm As Variant
    Dim excelForm2 As Variant
    Dim excelOSForm2 As Variant
    Dim cellRng As Range

    varTypes = ValueOf("variable type")
    varMin = ValueOf("min")
    varMax = ValueOf("max")
    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")

    If (varMin = vbNullString And varMax = vbNullString) Or (varTypes = "text") Then Exit Sub

    Set cellRng = VarRange()

    With cellRng.validation
        .Delete

        validationTypes = Switch(varTypes = "integer", xlValidateWholeNumber, _
                                 varTypes = "date", xlValidateDate, _
                                 varTypes = "decimal", _
                                 xlValidateDecimal)

        validationAlerts = Switch(varAlert = "error", xlValidAlertStop, _
                                  varAlert = "warning", xlValidAlertWarning, _
                                  True, xlValidAlertInformation)

        If (varMin <> vbNullString) And (varMax = vbNullString) Then

            excelForm = ExcelFormula(varMin, validationTypes)
            excelOSForm = ExcelFormula(varMin, validationTypes, useOS:=True)

            If IsEmpty(excelForm) Then Exit Sub

            On Error Resume Next
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlGreaterEqual, Formula1:=excelOSForm

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlGreaterEqual, Formula1:=excelForm
            On Error GoTo 0

        ElseIf (varMin = vbNullString) And (varMax <> vbNullString) Then
            
            'Validation on Maximum
            excelForm = ExcelFormula(varMax, validationTypes)
            excelOSForm = ExcelFormula(varMin, validationTypes, useOS:=True)

            If IsEmpty(excelForm) Then Exit Sub

            On Error Resume Next

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
            Operator:=xlLessEqual, Formula1:=excelOSForm

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
            Operator:=xlLessEqual, Formula1:=excelForm

            On Error GoTo 0

        ElseIf (varMin <> vbNullString) And (varMax <> vbNullString) Then

            'Validation on Both minimum and maximum
            excelForm = ExcelFormula(varMin, validationTypes)
            excelOSForm = ExcelFormula(varMin, validationTypes, useOS:=True)
            excelForm2 = ExcelFormula(varMax, validationTypes)
            excelOSForm2 = ExcelFormula(varMax, validationTypes, useOS:=True)

            Debug.Print "ExcelForm: " & excelForm
            Debug.Print "ExcelOSForm: " & excelOSForm

            Debug.Print "ExcelForm2: " & excelForm2
            Debug.Print "ExcelOSForm2: " & excelOSForm2


            If IsEmpty(excelForm) Or IsEmpty(excelForm2) Then Exit Sub

            On Error Resume Next

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlBetween, _
                 Formula1:=excelOSForm, Formula2:=excelOSForm2

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlBetween, _
                 Formula1:=excelForm, Formula2:=excelForm2

            On Error GoTo 0

        End If

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .errorTitle = vbNullString
        .InputMessage = vbNullString
        .errorMessage = varMessage
        .ShowInput = True
        .ShowError = True

    End With
    Exit Sub

End Sub

'Excel formula uses the OS to convert to local formula
Private Function ExcelFormula(ByVal formVal As String, _ 
                              ByVal validationTypes As Long, _
                              Optional ByVal useOS As Boolean = False) As Variant

    Dim lData As ILinelistSpecs
    Dim dict As ILLdictionary
    Dim varFormData As IFormulaData
    Dim varFormObject As IFormulas
    Dim excelForm As String
    Dim convForm As Variant


    'Elements to build the formula
    Set lData = LinelistData()
    Set dict = lData.Dictionary()
    Set varFormData = lData.FormulaDataObject()
    Set varFormObject = Formulas.Create(dict, varFormData, formVal)

    'If the formula is valid, test it before setting the validation
    If Not varFormObject.Valid(formulaType:="linelist") Then Exit Function
    excelForm = varFormObject.ParsedLinelistFormula()
    If FailedFormula(excelForm) Then Exit Function

    'Removed conversion to OS formula on french desktops
    If useOS Then
        excelForm = OSFormula(excelForm)
    Else
        excelForm = "= " & excelForm
    End If

    If Not varFormObject.HasLiterals Then

        'Return the required type
        excelForm = Application.WorksheetFunction.Trim(Replace(excelForm, "=", ""))

        Select Case validationTypes
        Case xlValidateWholeNumber
            convForm = CInt(excelForm)
        Case xlValidateDate
            convForm = CDate(excelForm)
        Case xlValidateDecimal
            convForm = CLng(excelForm)
        End Select
        ExcelFormula = convForm

    Else
        'We have a formula, we just return the formula
        ExcelFormula = excelForm
    End If

End Function

'Operating system formula (on Mac, the formula is not in formula Local)
Private Function OSFormula(ByVal formVal As String) As String

    OSFormula = vbNullString
    Dim rng As Range
    Dim sh As Worksheet
    Dim prevForm As String                       'Previous formula


    If (Application.OperatingSystem Like "*Mac*") Then
        OSFormula = formVal
    Else
        'On windows, validations are done using the
        'formula Local and not the formula itself
        Set sh = Wksh()
        prevForm = sh.Cells(1, 1).formula
        sh.Cells(1, 1).formula = "=" & formVal
        OSFormula = sh.Cells(1, 1).FormulaLocal
        'restore previous formula in cell
        sh.Cells(1, 1).formula = prevForm
    End If
End Function

'Write one variable information in the corresponding cell.

Private Sub WriteInfo()

    'Add the label
    AddLabel

    'Formulas
    AddFormula

    'Types
    AddType

    'Choices
    AddChoices

    'Add conditional formatting
    AddConditionalFormatting

    'Addvalidation
    On Error Resume Next
    AddValidation
    On Error GoTo 0

End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub CheckRequirements(ByVal varName As String, ll As ILinelist)
    Dim dict As ILLdictionary
    Dim varObject As ILLVariables
    Dim sheetName As String

    If ll Is Nothing Then ThrowError "No linelist object found"
    Set dict = ll.Dictionary()
    Set varObject = LLVariables.Create(dict)
    If Not dict.VariableExists(varName) Then _
    ThrowError "Variable" & varName & "Does not exists!"
    sheetName = varObject.Value(varName:=varName, colName:="sheet name")
    If Not ll.sheetExists(sheetName) Then _
    ThrowError varName & "worksheet named" & sheetName & "is not created"

End Sub

'Exposed methods
Public Sub IVListVars_WriteInfo()
    WriteInfo
End Sub

Public Property Get IVListVars_ValueOf(ByVal colName As String) As String
    IVListVars_ValueOf = ValueOf(colName)
End Property
