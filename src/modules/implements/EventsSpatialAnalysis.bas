Attribute VB_Name = "EventsSpatialAnalysis"
Attribute VB_Description = "Events related to spatial analysis tables"

Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events related to spatial analysis tables")

'spatial analyses Sheet
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const PASSSHEET As String = "__pass"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

Private pass As ILLPasswords
Private lltrads As ILLTranslations

'Initialize trads and passwords
Private Sub Initialize()

    Dim dicttranssh As Worksheet
    Dim psh As Worksheet
    Dim lltranssh As Worksheet

    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set psh = ThisWorkbook.Worksheets(PASSSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set pass = LLPasswords.Create(psh)
End Sub

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

'@Description("Update all spatial tables in the spatial sheet")
'@EntryPoint
Public Sub UpdateSpTables()
    Attribute UpdateSpTables.VB_Description = "Update all spatial tables in the spatial sheet"

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

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
Public Sub UpdateSingleSpTable(ByVal rngName As String)
    Attribute UpdateSingleSpTable.VB_Description = "Update all values in a table when the user changes the admin level"

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

    BusyApp cursor:=xlNorthwestArrow

    'initialize passwords, translations etc.
    Initialize
    pass.UnProtect "_active"

    'Spatial analysis worksheet
    Set sh = ActiveSheet

    'selected admin level
    selectedAdmin = sh.Range(rngName).Value

    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    adminName = geo.AdminCode(selectedAdmin)

    'remove the admdropdown to get the table id
    tabId = Replace(rngName, "ADM_DROPDOWN_", vbNullString)
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

            formulaValue = Replace(formulaValue, "concat_" & prevAdmValue, _
                                   "concat_" & adminName)

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
    sh.Columns("C").EntireColumn.AutoFit

    'Calculate the outer range
    rng.calculate
    pass.Protect "_active", True
    NotBusyApp
End Sub

'@Description("Devide all computed Values by the population")
Public Sub DevideByPopulation(ByVal rngName As String, _
                             Optional ByVal revertBack As Boolean = False)
    Attribute DevideByPopulation.VB_Description = "Devide all computed Values by the population"

    Dim sh As Worksheet
    Dim hasFormula As Boolean
    Dim factorMult As Long
    Dim prevFact As Long
    Dim rng As Range
    Dim cellRng As Range
    Dim formulaValue As String
    Dim adminName As String
    Dim rowRng As Range
    Dim geo As ILLGeo
    Dim selectedAdmin As String
    Dim popValue As String
    Dim tabId As String
    Dim sp As ILLSpatial

    BusyApp cursor:=xlNorthwestArrow
    Initialize

    Set sh = ActiveSheet
    pass.UnProtect "_active"

    tabId = Replace(rngName, "POPFACT_", vbNullString)
    prevFact = sh.Range("POPPREVFACT_" & tabId).Value

    factorMult = 100
    On Error Resume Next
    factorMult = CLng(Application.WorksheetFunction.Trim(sh.Range(rngName).Value))
    On Error GoTo Err

    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    selectedAdmin = sh.Range("ADM_DROPDOWN_" & tabId).Value
    adminName = geo.AdminCode(selectedAdmin)

    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))
    Set rng = sh.Range("OUTER_VALUES_" & tabId)
    Set rowRng = sh.Range("ROW_CATEGORIES_" & tabId)

    'Sort the spatial tables on either attack rate or values
    sp.Sort tabId:=tabId, onAR:=(Not revertBack)

    For Each cellRng In rng

        hasFormula = False
        formulaValue = cellRng.FormulaArray

        If formulaValue = vbNullString Then
            formulaValue = cellRng.formula
            hasFormula = True
        End If

        If (InStr(1, formulaValue, "concat_" & adminName) > 0) And _
           (cellRng.Column > rowRng.Column) Then

            popValue = sh.Cells(cellRng.Row, rowRng.Column - 1).Address

            If (Not revertBack) And (prevFact = 0) Then
                formulaValue = Replace(formulaValue, "=", vbNullString)
                formulaValue = "= " & factorMult & "*" & formulaValue & "/" & popValue
            ElseIf (Not revertBack) And (prevFact <> 0) Then
                formulaValue = Replace(formulaValue, prevFact, factorMult)
            'If the previous factor is 0, then no need to revert Back
            ElseIf (prevFact <> 0) Then
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
    pass.Protect "_active"

Err:
    NotBusyApp
End Sub

'@Description("Format the devide by population")
Public Sub FormatDevidePop(ByVal rngName As String)
    Attribute FormatDevidePop.VB_Description = "Format the devide by population"

    Dim sh As Worksheet
    Dim tabId As String

    Set sh = ActiveSheet

    Initialize
    BusyApp cursor:=xlNorthwestArrow

    pass.UnProtect "_active"
    tabId = Replace(rngName, "DEVIDEPOP_", vbNullString)

    'lltrads is the linelist translation object in the Initialize sub
    If sh.Range(rngName).Value = lltrads.Value("nodevide") Then

        'Do not devide
        sh.Range("POPFACT_" & tabId).Font.color = vbWhite
        sh.Range("POPFACT_" & tabId).Locked = True
        sh.Range("POPFACTLABEL_" & tabId).Font.color = vbWhite
        sh.Range("POPFACTLABEL_" & tabId).Locked = True
        sh.Range("POPFACT_" & tabId).FormulaHidden = True

        DevideByPopulation rngName:="POPFACT_" & tabId, revertBack:=True

    ElseIf sh.Range(rngName).Value = lltrads.Value("devide") Then

        'Devide by the population
        sh.Range("POPFACT_" & tabId).Font.color = vbBlack
        sh.Range("POPFACT_" & tabId).Locked = False
        sh.Range("POPFACT_" & tabId).FormulaHidden = False
        sh.Range("POPFACTLABEL_" & tabId).Font.color = vbBlack
        sh.Range("POPFACTLABEL_" & tabId).Locked = False

        DevideByPopulation "POPFACT_" & tabId

    End If

    NotBusyApp
    pass.Protect "_active", True
End Sub

'@Description("Update formulas in spatio-temporal tables")
'@EntryPoint
Public Sub UpdateSpatioTemporalFormulas(ByVal rngName As String, ByVal actAdm As Long)
    Attribute UpdateSpatioTemporalFormulas.VB_Description = "Update formulas in spatio-temporal tables"

    Dim tabId As String
    Dim prevAdm As Long
    Dim sh As Worksheet
    Dim counter As Long
    Dim headerRng As Range
    Dim cellRng As Range
    Dim valuesRng As Range
    Dim headerFormula As String
    Dim valuesFormula As String
    Dim headerCellName As String
    Dim hasFormula As Boolean
    

    BusyApp cursor:=xlNorthwestArrow
    Initialize
    
    On Error GoTo Err

    Set sh = ActiveSheet
    tabId = "SPT_" & Split(rngName, "_")(3)
    Set headerRng = sh.Range("SPT_FORMULA_COLUMN_" & tabId)
    prevAdm = sh.Range(rngName).Offset(, 1).Value

    pass.UnProtect "_active"

    For counter = 1 To headerRng.Columns.Count
        headerFormula = Replace(headerRng.Cells(1, counter).Formula, "=", vbNullString)
        headerFormula = Application.WorksheetFunction.Trim(headerFormula)

        'Change the formula for only columns where headers are the selected input
        If (headerFormula = rngName) Then
            
            Set valuesRng = Nothing

            On Error Resume Next
            headerCellName = headerRng.Cells(1, counter).Name.Name
            'replace LABEL with VALUES to get the column range of values
            Set valuesRng = sh.Range(Replace(headerCellName, "LABEL", "VALUES"))
            On Error GoTo Err

            If (Not valuesRng Is Nothing) Then
                'Shift values Rng to take in account total and missing lines
                Set valuesRng = sh.Range(valuesRng.Cells(1, 1), _ 
                                         valuesRng.Cells(valuesRng.Rows.Count + 2, 1))

                For Each cellRng In valuesRng

                    hasFormula = False
                    valuesFormula = cellRng.FormulaArray

                    If valuesFormula = vbNullString Then
                        valuesFormula = cellRng.formula
                        hasFormula = True
                    End If

                    If (InStr(1, valuesFormula, "concat_adm" & prevAdm) > 0) Then

                        valuesFormula = Replace( _ 
                                            valuesFormula, _ 
                                            "concat_adm" & prevAdm, _
                                            "concat_adm" & actAdm)

                        'some cells have formula, others have formulaArray
                        If (hasFormula) Then
                            cellRng.formula = valuesFormula
                        Else
                            cellRng.FormulaArray = valuesFormula
                        End If
                    End If
                Next
            End If
        End If
    Next

    'Change previous admin values
    sh.Range(rngName).Offset(, 1).Value = actAdm
    sh.UsedRange.Calculate

Err:
    pass.Protect sh, True
    NotBusyApp
End Sub