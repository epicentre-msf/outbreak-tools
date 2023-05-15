Attribute VB_Name = "EventsSpatialAnalysis"

Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events related to spatial analysis tables")

'spatial analyses Sheet
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const PASSSHEET As String = "__pass"

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Cursor = xlDefault
End Sub

'@Description("Update all spatial tables in the spatial sheet")
'@EntryPoint
Public Sub UpdateSpTables()
    Dim sp As ILLSpatial
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(SPATIALSHEET)
    Set sp = LLSpatial.Create(sh)

    UpdateFilterTables calculate:=False

    'Update all the spatial tables
    'For each geo variable in analyses, update computations
    'This sub circles throughout all the worksheets.
    'Updating filters returns the application to normal state, so need to
    'add a BusyApp here before proceeding.
    BusyApp

    sp.Update

    DoEvents
    ActiveSheet.calculate
    ActiveSheet.UsedRange.calculate
    ActiveSheet.Columns("A:E").calculate

    NotBusyApp
End Sub


'@Description("Update all values in a table when the user changes the admin level")
'@EntryPoint
Public Sub UpdateSingleSpTable(ByVal rngName As String)

    'rngName is the name of the range where we have the admin level

    Dim tabId As String
    Dim adminName As String
    Dim selectedAdmin As String
    Dim formulaValue As String
    Dim prevAdmValue As String
    Dim cellRng As Range
    Dim rng As Range
    Dim sh As Worksheet
    Dim geo As ILLGeo
    Dim hasFormula As Boolean
    Dim pass As ILLPasswords

    BusyApp

    Set pass = LLPasswords.Create(ThisWorkbook.Worksheets("__pass"))
    pass.UnProtect "_active"

    'Spatial analysis worksheet
    Set sh = ActiveSheet

    'selected admin level
    selectedAdmin = sh.Range(rngName).Value

    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    adminName = geo.AdminCode(selectedAdmin)

    'remove the admdropdown to get the table id
    tabId = Replace(rngName, "ADM_DROPDOWN_", "")
    prevAdmValue = sh.Range("PREVIOUS_ADM_" & tabId).Value

    'Interior table range, including missing row and total column ranges
    Set rng = sh.Range("OUTER_VALUES_" & tabId)

    For Each cellRng In rng

        hasFormula = False
        formulaValue = cellRng.FormulaArray

        If formulaValue = vbNullString Then
            formulaValue = cellRng.formula
            hasFormula = True
        End If

        If (InStr(1, formulaValue, "concat_" & prevAdmValue) > 0) Then

            formulaValue = Replace(formulaValue, "concat_" & prevAdmValue, "concat_" & adminName)

            'some cells have formula, others have formulaArray
            If (hasFormula) Then
                cellRng.formula = formulaValue
            Else
                cellRng.FormulaArray = formulaValue
            End If
        End If
    Next

    'change the previous admin
    sh.Range("PREVIOUS_ADM_" & tabId).Value = adminName

    'Calculate the outer range
    rng.calculate

    pass.Protect "_active", True
    EndWork xlsapp:=Application
End Sub



'@Description("Devide all computed Values by the population")
Public Sub DevideByPopulation(ByVal rngName As String, Optional ByVal revertBack As Boolean = False)

    Dim sh As Worksheet
    Dim hasFormula As Boolean
    Dim factorMult As Long
    Dim prevFact As Long
    Dim rng As Range
    Dim cellRng As Range
    Dim formulaValue As String
    Dim adminName As String
    Dim AdminCode As Byte
    Dim rowRng As Range
    Dim geo As ILLGeo
    Dim selectedAdmin As String
    Dim popValue As String
    Dim tabId As String
    Dim pass As ILLPasswords

    BusyApp

    Set sh = ActiveSheet
    Set pass = LLPasswords.Create(ThisWorkbook.Worksheets(PASSSHEET))
    pass.UnProtect "_active"

    tabId = Replace(rngName, "POPFACT_", vbNullString)
    prevFact = sh.Range("POPPREVFACT_" & tabId).Value

    factorMult = 100
    On Error Resume Next
    factorMult = CLng(Application.WorksheetFunction.Trim(sh.Range(rngName).Value))
    On Error GoTo Errkz

    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    selectedAdmin = sh.Range("ADM_DROPDOWN_" & tabId).Value
    adminName = geo.AdminCode(selectedAdmin)

    Set rng = sh.Range("OUTER_VALUES_" & tabId)
    Set rowRng = sh.Range("ROW_CATEGORIES_" & tabId)

    For Each cellRng In rng

        hasFormula = False
        formulaValue = cellRng.FormulaArray

        If formulaValue = vbNullString Then
            formulaValue = cellRng.formula
            hasFormula = True
        End If

        If (InStr(1, formulaValue, "concat_" & adminName) > 0) And (cellRng.Column > rowRng.Column) Then

            popValue = sh.Cells(cellRng.Row, rowRng.Column - 1).Address

            If (Not revertBack) And (prevFact = 0) Then
                formulaValue = Replace(formulaValue, "=", vbNullString)
                formulaValue = "= " & factorMult & "*" & formulaValue & "/" & popValue
            ElseIf (Not revertBack) And (prevFact <> 0) Then
                formulaValue = Replace(formulaValue, prevFact, factorMult)
            ElseIf (prevFact <> 0) Then 'If the previous factor is 0, then no need to revert Back
                'Remove the factor
                formulaValue = Replace(formulaValue, prevFact & "*", vbNullString)
                'Remove the denominator
                formulaValue = Replace(formulaValue, "/" & popValue, vbNullString)
            End If

            'some cells have formula, others have formulaArray
            If (hasFormula) Then
                cellRng.formula = formulaValue
            Else
                cellRng.FormulaArray = formulaValue
            End If
        End If
    Next

    'Update the previous Factor
    If revertBack Then
        sh.Range("POPPREVFACT_" & tabId).Value = 0
    Else
        sh.Range("POPPREVFACT_" & tabId).Value = factorMult
    End If
    rng.calculate

Err:
    pass.Protect "_active", True
    NotBusyApp
End Sub


'
