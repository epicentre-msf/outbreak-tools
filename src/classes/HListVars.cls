VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HListVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("HList variables")
'@IgnoreModule
Option Explicit
Implements IHListVars

Private Type THListVars
    ll As ILinelist
    varName As String
    inputRange As Range
End Type

Private this As THListVars
Private Const CLASSNAME As String = "HListVars"
Private Const STARTROW As Long = 8               'Where the variables should start in the output worksheet

Public Function Create(ByVal varName As String, ByVal ll As ILinelist) As IHListVars

    CheckRequirements varName, ll

    With New HListVars
        Set .LinelistObject = ll
        .ActualVariable = varName
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IHListVars
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
    Dim drop As IDropdownLists

    Set ll = LinelistObject()
    Set drop = ll.Dropdown()

    Set Dropdown = drop
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

Private Property Get SheetsObject() As ILLSheets
    Dim dict As ILLdictionary
    Set dict = Dictionary()
    Set SheetsObject = LLSheets.Create(dict)
End Property

Private Property Get ValueOf(ByVal colName As String) As String
    Dim varObject As ILLVariables
    Set varObject = vars()
    ValueOf = varObject.Value(colName:=colName, varName:=ActualVariable)
End Property

Private Property Get VarRange() As Range         'Where to input informations on the variable
    Dim Startcol As Long
    Dim sh As Worksheet
    Dim rng As Range

    If this.inputRange Is Nothing Then
        Set sh = Wksh()
        With sh
            Startcol = ValueOf("column index")
            Set this.inputRange = .Cells(STARTROW + 1, Startcol)
        End With
    End If

    Set VarRange = this.inputRange
End Property

'Type
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
            formatText = IIf(varTypes = "decimal", "#,##0.00;-#,##0.00;0.00;@", "###0;-###0;;@")
        'Other formats
        ElseIf varFormat = "integer" Then

            formatText = "###0;-###0;;@"

        ElseIf varFormat = "round0" Then

            formatText = "#,##0;-#,##0;0;@"

        ElseIf varFormat = "percentage0" Then

            formatText = "0%"

        ElseIf (InStr(1, varFormat, "round") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "round", vbNullString))
            On Error GoTo 0
            zerosDecimals = String(nbDecimals, "0")
            formatText = "#,###0." & zerosDecimals & ";-#,###0." & zerosDecimals & ";0." & zerosDecimals & ";@"

        ElseIf (InStr(1, varFormat, "percentage") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "percentage", vbNullString))
            On Error GoTo 0
            zerosDecimals = String(nbDecimals, "0")
            formatText = "0." & zerosDecimals & "%"

        ElseIf varFormat = "euros" Then

            formatText = "#,##0.00 $;-#,##0.00 $;0.00 $;@"

        ElseIf varFormat = "dollars" Then

            formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"

        ElseIf varFormat <> "text" Then

            formatText = "custom"
        End If

        'Define the format, if your custom format is not valid, it will be ignored
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

        'Lock the formula cell
        cellRng.Locked = True

    End If

    'Apply format of the calculated range
    DesignFormat.ApplyFormat cellRng, scope:=HListCalculatedFormulaCell 'Cell
    DesignFormat.ApplyFormat cellRng.Offset(-2), scope:=HListCalculatedFormulaHeader 'Header
End Sub

'Conditional formatting
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
        .Interior.Color = formatRng.Interior.Color
        .Font.Color = formatRng.Font.Color
        .Font.Bold = formatRng.Font.Bold
        .Font.Italic = formatRng.Font.Italic
        'Stop if this condition is met (This is the first priority)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'Conditional formatting on unique values
Private Sub AddUniqueConditonalFormatting()
    Dim varUnique As String
    Dim cellRng As Range
    varUnique = ValueOf("unique")
    If (varUnique <> "yes") Then Exit Sub
    Set cellRng = VarRange()
    cellRng.FormatConditions.AddUniqueValues
    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .DupeUnique = xlDuplicate
        .Interior.Color = vbRed
    End With
End Sub

'Main Label
Private Sub AddLabel(Optional ByVal notevisibility As Boolean = False)
    Dim varMainLab As String
    Dim varName As String
    Dim varNote As String
    Dim varSublab As String
    Dim varStatus As String
    Dim varListAutoOrigin As String
    Dim varControl As String
    Dim varIndex As Long
    Dim cellRng As Range
    Dim trans As ITranslation
    Dim lData As ILinelistSpecs
    Dim calculated As String
    Dim mandatory As String
    Dim mainLabLength As Long
    Dim design As ILLFormat

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
    varListAutoOrigin = ValueOf("list auto")
    varIndex = CLng(ValueOf("column index"))
    calculated = trans.TranslatedValue("MSG_Calculated")
    mandatory = trans.TranslatedValue("MSG_Mandatory")

    'add an anchor to where the table starts
    If varIndex = 1 Then cellRng.Name = ValueOf("table name") & "_" & "START"

    'sublabel
    If (varControl = "formula") Or (varControl = "case_when") Or _
       (varControl = "choice_formula") Then
        varSublab = varSublab & chr(10) & calculated
    End If

    mainLabLength = Len(varMainLab)

    'Add subLabel to main label
    If (varSublab <> vbNullString) Then
        varMainLab = varMainLab & chr(10) & varSublab
    End If

    'Status to add to Notes
    If varStatus = "mandatory" Then varNote = mandatory & chr(10) & varNote

    'variable main label and sub lable
    cellRng.Offset(-2).Value = varMainLab
    design.ApplyFormat cellRng.Offset(-2).Characters(Start:=1, Length:=mainLabLength), _
        scope:=HListMainLab                      'format of the main label
    design.ApplyFormat cellRng.Offset(-2).Characters(Start:=mainLabLength + 1), _
        scope:=HListSublab                       'format of the sub label

    'variable name
    cellRng.Offset(-1).Value = varName
    'Name the cell range
    cellRng.Offset(-2).Name = varName

    'variable note
    If varNote <> vbNullString Then
        cellRng.Offset(-2).AddComment
        cellRng.Offset(-2).Comment.Text Text:=varNote
        cellRng.Offset(-2).Comment.Visible = notevisibility
    End If

    'Control
    cellRng.Offset(-5).Value = varControl
    design.ApplyFormat cellRng.Offset(-5), scope:=LinelistHiddenCell

    'List auto origin
    cellRng.Offset(-6).Value = varListAutoOrigin
    design.ApplyFormat cellRng.Offset(-6), scope:=LinelistHiddenCell

    'Geo labels
    If (InStr(1, varControl, "geo") > 0) Or (varControl = "hf") Or _
        (InStr(1, varControl, "pcode") > 0) Then

        'Put geo color
        design.ApplyFormat cellRng.Offset(-2), scope:=HListGeoHeader
        design.ApplyFormat cellRng, scope:=HListGeo
    End If

    'Custom columns, user can modify them
    If (varControl = "custom") Then  cellRng.Offset(-2).Locked = False
    'Hidden columns
    If varStatus = "hidden" Then  cellRng.EntireColumn.Hidden = True

    'The cell header itself
    design.ApplyFormat cellRng.Offset(-2), scope:=HListSimpleHeaderCell
    cellRng.Locked = False
End Sub

'Add the choices
Private Sub AddChoices()

    Dim varControl As String
    Dim dropName As String                       'drop abreviation for dropdown
    Dim varAlert As String
    Dim varMessage As String
    Dim lData As ILinelistSpecs
    Dim cat As BetterArray
    Dim drop As IDropdownLists
    Dim cellRng As Range
    Dim geo As ILLGeo

    varControl = ValueOf("control")
    If (varControl <> "choice_manual") And (varControl <> "list_auto") And _
     (varControl <> "geo1") Then Exit Sub

    Set lData = LinelistData()
    Set drop = Dropdown()
    Set cat = New BetterArray
    Set cellRng = VarRange()                     'variable range where to set the validation

    Select Case varControl
        'dropName is the name of the dropdown, not the variable. Tt is the variable name only if
        'the control is choices
        'Dropdown list are named dropdown_[dropName] in the workbook
    Case "choice_manual"

        dropName = ActualVariable()
        Set cat = lData.Categories(dropName)
        drop.Add cat, dropName
        varAlert = ValueOf("alert")
        varMessage = ValueOf("message")

    Case "list_auto"

        dropName = ValueOf("control details")
        'The category is empty It will create an empty table
        cat.Push ""
        drop.Add cat, dropName
    Case "geo1"
        dropName = "admin1"                      'The dropdown is only on admin 1 and the logic is done in the worksheet
        'Add the three dropdowns for the four admins
        Set geo = lData.geoObject()
        Set cat = geo.GeoLevel(LevelAdmin1)
        drop.Add cat, "admin1"                   'Admin 1 Levels are used for the dropdown
        cat.Clear
        cat.Push ""

        'Dropdowns on next admin levels
        drop.Add cat, "admin2"
        drop.SetValidation cellRng:=cellRng.Offset(, 1), listName:="admin2", _
        alertType:=varAlert, message:=varMessage


        drop.Add cat, "admin3"

        drop.SetValidation cellRng:=cellRng.Offset(, 2), listName:="admin3", _
        alertType:=varAlert, message:=varMessage

        drop.Add cat, "admin4"
        drop.SetValidation cellRng:=cellRng.Offset(, 3), listName:="admin4", _
        alertType:=varAlert, message:=varMessage
    End Select

    'Set dropdown validations
    drop.SetValidation cellRng:=cellRng, listName:=dropName, _
                       alertType:=varAlert, message:=varMessage
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
    Dim excelForm2 As Variant
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
            If IsEmpty(excelForm) Then Exit Sub

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlGreaterEqual, Formula1:=excelForm

        ElseIf (varMin = vbNullString) And (varMax <> vbNullString) Then
            'Validation on Maximum
            excelForm = ExcelFormula(varMax, validationTypes)
            If IsEmpty(excelForm) Then Exit Sub

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
            Operator:=xlLessEqual, Formula1:=excelForm

        ElseIf (varMin <> vbNullString) And (varMax <> vbNullString) Then
            'Validation on Both minimum and maximum
            excelForm = ExcelFormula(varMin, validationTypes)
            excelForm2 = ExcelFormula(varMax, validationTypes)

            If IsEmpty(excelForm) Or IsEmpty(excelForm2) Then Exit Sub

            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlBetween, _
                 Formula1:=excelForm, Formula2:=excelForm2

        End If

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .errorMessage = varMessage
        .ShowInput = True
        .ShowError = True

    End With

End Sub

'Excel formula uses the OS to convert to local formula
Private Function ExcelFormula(ByVal formVal As String, ByVal validationTypes As Long) As Variant

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
    'excelForm = OSFormula(excelForm)

    excelForm = "= " & excelForm

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
        'On windows, validations are done using the formula Local and not the formula itself
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

    'Add the validations
    On Error Resume Next
    AddValidation
    On Error GoTo 0

    'Conditional formatting
    AddConditionalFormatting

    'Conditional formatting on unique values
    AddUniqueConditonalFormatting
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
Public Sub IHListVars_WriteInfo()
    WriteInfo
End Sub

Public Property Get IHListVars_ValueOf(ByVal colName As String) As String
    IHListVars_ValueOf = ValueOf(colName)
End Property