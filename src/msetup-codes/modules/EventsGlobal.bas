
Attribute VB_Name = "EventsGlobal"
Option Explicit

'@Folder("Events")
'This modules contains events fired at the workbook level

Private Const PASSSHEETNAME As String = "__pass"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const TRANSLATIONSHEET As String = "Translations"
Private Const RIBBONTRADSHEET As String = "__ribbonTranslation"

'speed app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub


Public Sub SetAllUpdatedTo(ByVal Value As String)
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim rng As Range
    Dim counter As Long

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(UPDATEDSHEETNAME)

    For Each Lo In sh.ListObjects
        Set rng = Lo.ListColumns("updated").Range
        For counter = 2 To rng.Rows.Count
            rng.Cells(counter, 1).Value = Value
        Next
    Next
End Sub

'Check update status when something changes in a range on a worksheet
'@EntryPoint
Public Sub checkUpdateStatus(ByVal sh As Worksheet, ByVal Target As Range)
    Dim upsh As Worksheet
    Dim upObj As IUpdatedValues
    Dim upId As String
    Dim Lo As ListObject
    Dim wb As Workbook
    Dim choiObj As IChoices
    Dim pass As IPasswords
    Dim ribTrads As ITranslationObject
    Dim ribSh As Worksheet

    BusyApp

    Set wb = ThisWorkbook

    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)
    upId = LCase(sh.Name)
    
    If sh.Name = "Variables" Then
        Set Lo = sh.ListObjects(1)
        If (Target.Column =  Lo.ListColumns(3).Column) Then
            'Test if Default choices has been modified
            Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
            Set ribSh = wb.Worksheets(RIBBONTRADSHEET)
            Set ribTrads = Translation.Create(ribSh.ListObjects(1), ribSh.Range("RNG_FileLang").Value)
            Set choiObj = LLChoices.Create(wb.Worksheets("Choices"), 4, 1, ribTrads)
            pass.UnProtect sh
            Target.Offset(, 1).Value = choiObj.PasteCategories(Target.Value)
            pass.Protect sh
        End If
    End If
    
    Set upObj = UpdatedValues.Create(upsh, upId)
    upObj.CheckUpdate sh, Target
End Sub

'Fire this event when the workbook is opened
'@EntryPoint
Public Sub OpenedWorkbook()
    Dim wb As Workbook
    Dim rng As Range
    
    Set wb = ThisWorkbook
    BusyApp
    On Error Resume Next
    Set rng = wb.Worksheets(TRANSLATIONSHEET).Range("RNG_NbTimesTrans")
    rng.Value = 0
    On Error GoTo 0
    'Set all update ranges to no
    SetAllUpdatedTo "yes"
    Application.CalculateBeforeSave = False
End Sub


'Fire Event to update a disease worksheet

Public Sub UpdateDiseaseSheet(ByVal disSh As Worksheet, ByVal Target As Range)

    Dim Lo As ListObject
    Dim tradLo As ListObject
    Dim ribTrads As ITranslationObject
    Dim ribSh As Worksheet
    Dim varLo As ListObject
    Dim trads As ITranslationObject
    Dim wb As Workbook
    Dim newLang As String
    Dim tradRng As Range
    Dim counter As Long
    Dim varTab As ICustomTable
    Dim varValue As String
    Dim labValue As String
    Dim choiValue As String
    Dim statusValue As String
    Dim choiObj As IChoices
    Dim pass As IPasswords
    Dim tradRib As ITranslationObject

    Set wb = ThisWorkbook
    Set ribSh = wb.Worksheets("__ribbonTranslation")
    Set ribTrads = Translation.Create(ribSh.ListObjects(1), ribSh.Range("RNG_FileLang").Value)
    Set Lo = dissh.ListObjects(1)
    Set tradLo = wb.Worksheets("Translations").ListObjects(1)
    Set choiObj = LLChoices.Create(wb.Worksheets("Choices"), 4, 1, ribTrads)
    newLang = disSh.Cells(2, 2).Value
    Set trads = Translation.Create(tradLo, newLang)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))

    pass.UnProtect dissh

    'Translate the elements in the worksheet to another language
    If Not (Intersect(Target, disSh.Cells(2, 2)) Is Nothing) Then Set tradRng = Lo.DataBodyRange

    'Update Variable label, status, choice and Default choice
    If Not (Intersect(Target, Lo.ListColumns(3).Range) Is Nothing) Then Set tradRng = Target


    'Update translation elements
    If Not (tradRng Is Nothing) Then
        'Variable Name has changed, update the columns
        Set varLo = wb.Worksheets("Variables").ListObjects(1)
        Set varTab = CustomTable.Create(varLo, idCol:=varLo.Range.Cells(1, 1).Value)

        'Get the value of the variable label
        For counter = 1 To tradRng.Rows.Count
            varValue = tradRng.Cells(counter, 1).Value
            
            'label
            labValue = varTab.Value(colName:=varLo.Range.Cells(1, 2).Value, keyName:=varValue) 
            tradRng.Cells(counter, 2).Value = trads.TranslatedValue(labValue)

            'choice
            choiValue = varTab.Value(colName:=varLo.Range.Cells(1, 3).Value, keyName:=varValue)
            tradRng.Cells(counter, 3).Value = choiValue
            
            'Choices elements (values)
            tradRng.Cells(counter, 4).Value = choiObj.PasteTranslatedCategories(choiValue, trads)
        Next
    End If

    'Update Categories
    If Not (Intersect(Target, Lo.ListColumns(5).Range) Is Nothing) Then
        For counter = 1 To Target.Rows.Count
            choiValue = Target.Value
            Target.Cells(counter, 2).Value = choiObj.PasteTranslatedCategories(choiValue, trads)
        Next
    End If


    'For core variables, update the visibility to always visible if possible
    ' If Not (Intersect(Target, Lo.ListColumns(5).Range) Is Nothing) Then
    '     For counter = 1 To Target.Rows.Count    
    '         If Target.Cells(counter, 1).Value = ribTrads.TranslatedValue("coreVar") Then
    '             On Error Resume Next
    '                 Target.Cells(counter, 2).Value = ribTrads.TranslatedValue("alwaysVis")
    '                 Target.Cells(counter, 2).Font.Color = RGB(59, 8, 145)
    '                 Target.Cells(counter, 2).Font.Italic = True
    '             On Error GoTo 0
    '         Else
    '             On Error Resume Next
    '                 Target.Cells(counter, 2).ClearContents
    '                 Target.Cells(counter, 2).Font.Color = vbBlack
    '                 Target.Cells(counter, 2).Font.Italic = False
    '             On Error GoTo 0
    '         End If
    '     Next
    ' End If

    pass.Protect dissh
End Sub


'Filter the comparation report worksheet

Public Sub ComparationSheet(ByVal Target As Range)
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filterValue As String
    Dim sheetRng As Range 'used range in the worksheet
    Dim cellRng As Range 'corresponding cellRange
    
    On Error Resume Next

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets("__compRep")
    Set sheetRng = sh.UsedRange
    If Intersect(Target, sh.Range("RNG_CheckingFilter")) Is Nothing Then Exit Sub
    filterValue = Target.Value

    If filterValue = vbNullString Then Exit Sub
    sh.Cells.EntireRow.Hidden = False

    If filterValue <> "All" Then
        Set cellRng = sheetRng.Cells(sheetRng.Rows.Count, 1)
        Do While cellRng.Row > sheetRng.Row
            'Hide cells with values corresponding to those selected (keeping headers)
            If (cellRng.Value <> filterValue) And (Not cellRng.Font.Size = 14) And (Not cellRng.Font.Size = 12) _
             And (Not cellRng.Value = vbNullString) Then cellRng.EntireRow.Hidden = True
            Set cellRng = cellRng.Offset(-1)
        Loop
    End If

    On Error GoTo 0
End Sub