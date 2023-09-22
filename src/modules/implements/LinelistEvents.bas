Attribute VB_Name = "LinelistEvents"

Option Explicit
Option Private Module

Private Const GOTOSECCODE As String = "go_to_section" 'Go To section constant
Private Const DROPDOWNSHEET As String = "dropdown_lists__"
Private Const GEOSHEET As String = "Geo"
Private Const DICTSHEET As String = "Dictionary"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const UPSHEET As String = "updates__"  'worksheet for updated values

Private tradmess As ITranslation              'Translation of messages object
Private lltrads As ILLTranslations
Private wb As Workbook

'Speed up before a work
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Return previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableAnimations = True
End Sub


'Initialize translation
Private Sub InitializeTrads()
    Set wb = ThisWorkbook
    Set lltrads = LLTranslations.Create( _
                                        wb.Worksheets(LLSHEET), _
                                        wb.Worksheets(TRADSHEET) _
                                        )
    Set tradmess = lltrads.TransObject()
End Sub

'Unique Values of a BetterArray
Private Function GetUniqueBA(inputTable As BetterArray, _
                     Optional ByVal Sort As Boolean = False) As BetterArray
    Dim tableValue As String
    Dim counter As Long
    Dim Outable As BetterArray

    Set Outable = New BetterArray
    Outable.LowerBound = 1

    If inputTable.Length > 0 Then
        For counter = inputTable.LowerBound To inputTable.UpperBound

            tableValue = Application.WorksheetFunction.Trim( _
                                inputTable.Item(counter) _
                        )

            If (tableValue <> vbNullString) And _
               (Not Outable.Includes(tableValue)) Then _
                Outable.Push tableValue
        Next
    End If

    'sort
    If Sort Then Outable.Sort
    Set GetUniqueBA = Outable.Clone()
End Function


'Trigerring event when the linelist sheet has some values within                                                          -                                                      -
Public Sub EventValueChangeLinelist(Target As Range)

    Dim adminTable As BetterArray              'Table with admin names (input by user)
    Dim adminNames As BetterArray              'Table with admin names (extracted from geobase)
    Dim sh As Worksheet                        'Active sheet where the event fires
    
    Dim rng As Range
    Dim hRng As Range                          'Header Row Range of the listObject
    Dim calcRng As Range                       'Calculate range
    Dim cellRng As Range
    
    Dim varControl As String                   'Control of variable
    Dim varLabel As String                     'Updated variable labels
    Dim varName As String                      'Variable name
    Dim varSubLabel As String                  'variable sub-label to remove
    Dim tablename As String                    'Name of the listObject on Hlist
    Dim varEditable As String                  'Test if a variable is editable or not
    Dim sectionName As String                  'Section name for GoTo Section
    
    Dim targetColumn As Long                   'Column of the target range
    Dim startLine As Long                      'The row of the anchor range for table start
    Dim nbOffset As Long                       'Number of offset from the headerrow range
    
    Dim drop As IDropdownLists                 'Dropdown Object for updating geolevels
    Dim geo As ILLGeo
    Dim upobj As IUpVal
    Dim dict As ILLdictionary
    Dim vars As ILLVariables

    Dim choiSep As String                      'Choice separator for multiple choices selection

    On Error GoTo ErrHand
    
    Set sh = ActiveSheet
    Set hRng = sh.ListObjects(1).HeaderRowRange
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1                  'This is mandatory for geolevel function, lowerbound = 1
    Set adminTable = New BetterArray

    'Initialize translations
    InitializeTrads

    tablename = sh.Cells(1, 4).Value
    targetColumn = Target.Column
    startLine = sh.Range(tablename & "_START").Row
    varControl = sh.Cells(startLine - 5, targetColumn).Value

    If (Target.Row >= startLine) Then

        nbOffset = Target.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)
        
        calcRng.calculate

        'Geo variables
        If (varControl = "geo1") Or _
           (varControl = "geo2") Or _
           (varControl = "geo3") Or _
           (varControl = "geo4") Then

            If (Target.Value = vbNullString) Then Exit Sub
                        
            Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))
            Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))

            Select Case varControl

            Case "geo1"
                'adm1 has been modified, we will correct and set validation to adm2-4
                BusyApp
                
                'Clear admin2, admin3 and admin4, update admin2 dropdown
                drop.ClearList "admin2"
                Target.Offset(, 1).Value = vbNullString
                drop.ClearList "admin3"
                Target.Offset(, 2).Value = vbNullString
                drop.ClearList "admin4"
                Target.Offset(, 3).Value = vbNullString

                'Filter the geobase on admin1
                Set adminTable = geo.GeoLevel(LevelAdmin2, _
                                              GeoScopeAdmin, _
                                              Target.Value)

                'Update the validation list for admin2
                drop.Update adminTable, "admin2"

                NotBusyApp

            Case "geo2"

                'Adm2 has been modified, we will correct and filter adm3 and 4
                BusyApp
                    
                'Clear admin3 and admin4, update admin 3 dropdown
                drop.ClearList "admin3"
                Target.Offset(, 1).Value = vbNullString
                drop.ClearList "admin4"
                Target.Offset(, 2).Value = vbNullString
                adminNames.Push Target.Offset(, -1).Value, Target.Value
                Set adminTable = geo.GeoLevel(LevelAdmin3, _
                                              GeoScopeAdmin, _
                                              adminNames)

                'Update the validation list for admin3
                drop.Update adminTable, "admin3"

                NotBusyApp

            Case "geo3"
                
                'Adm 3 has been modified, correct and filter adm4
                BusyApp

                drop.ClearList "admin4"
                Target.Offset(, 1).Value = vbNullString

                adminNames.Push Target.Offset(, -2).Value, _
                                    Target.Offset(, -1).Value, _
                                    Target.Value
                'Take the adm4 table
                Set adminTable = geo.GeoLevel(LevelAdmin4, _
                                              GeoScopeAdmin, _
                                              adminNames)
                drop.Update adminTable, "admin4"

                NotBusyApp

            End Select
        
        'Exit as soon as geo variables are updated
        Exit Sub
        End If
    End If

    'Update variables with editable column
    If (Target.Row = startLine - 2) Then

        Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
        Set vars = LLVariables.Create(dict)

        varName = sh.Cells(startLine - 1, targetColumn).Value
        varEditable = vars.Value(varName:=varName, colName:="editable label")
        
        If (varEditable <> "yes") Then Exit Sub

            
        'The name of custom variables has been updated, update the dictionary
        varSubLabel = vars.Value(varName:=varName, colName:="sub label")

        varLabel = Replace(Target.Value, varSubLabel, vbNullString)
        varLabel = Replace(varLabel, chr(10), vbNullString)

        vars.SetValue varName:=varName, colName:="main label", _
                      newValue:=varLabel

        Exit Sub
    End If

    'Update the list auto
    If (Target.Row >= startLine) And _
       (sh.Cells(startLine - 6, targetColumn).Value = "list_auto_origin") Then
        
        Set upobj = UpVal.Create(wb.Worksheets(UPSHEET))
        upobj.SetValue "RNG_UpdateListAuto", "yes"

        Exit Sub
    End If

    'GoTo section
    Set rng = sh.Range(tablename & "_" & GOTOSECCODE)

    If Not (Intersect(Target, rng) Is Nothing) Then

        sectionName = Replace(Target.Value, _
                         lltrads.Value("gotosection") & ": ", _
                         vbNullString)

        Set hRng = hRng.Offset(-3)
        Set cellRng = hRng.Find(What:=sectionName, lookAt:=xlWhole, MatchCase:=True)
        If Not cellRng Is Nothing Then cellRng.Activate

        Exit Sub
    End If
    
    'Avoid modifying headers of the table (with variable names)
    If (Target.Row = startLine - 1) Then
        
        Target.Value = Target.Offset(-1).Name.Name
        MsgBox tradmess.TranslatedValue("MSG_NotModify"), _
                vbOKOnly + vbCritical, _
                tradmess.TranslatedValue("MSG_Error")

        Exit Sub
    End If

    'Update multiple choices dropdown
    If Instr(1, varControl, "choice_multiple") = 1 Then

        On Error Resume Next
            choiSep = Split(varControl, "(" & Chr(34))(1)
            choiSep = Replace(choiSep,  Chr(34) & ")", vbNullString)
        On Error GoTo 0
        
        If (choiSep = vbNullString) Or _ 
           (Instr(1, choiSep, "choice_multiple") = 1) Then _ 
            choiSep = ", "

        BusyApp
        UpdateMultipleChoice Target, choiSep
        NotBusyApp
        Exit Sub
    End If

ErrHand:
    NotBusyApp
End Sub


'Event to update the list_auto when a sheet containing a list_auto is desactivated
Public Sub EventDesactivateLinelist(ByVal prevSheetName As String)

    Dim prevsh As Worksheet
    Dim upobj As IUpVal

    On Error GoTo ErrHand

    InitializeTrads

    Set upobj = UpVal.Create(wb.Worksheets(UPSHEET))

    'Update the listAuto only and only if update list auto is yes
    If upobj.Value("RNG_UpdateListAuto") <> "yes" Then Exit Sub
    
    Set prevsh = wb.Worksheets(prevSheetName)

    BusyApp
    UpdateListAuto prevsh
    upobj.SetValue "RNG_UpdateListAuto", "no"
    NotBusyApp
    Exit Sub

ErrHand:
   NotBusyApp
End Sub

'Update the list Auto of one Sheet
Public Sub UpdateListAuto(ByVal sh As Worksheet)

    Dim varName As String
    Dim drop As IDropdownLists
    Dim arrTable As BetterArray
    Dim cellRng As Range
    Dim tablename As String

    Set arrTable = New BetterArray
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))

    tablename = sh.Cells(1, 4).Value
    Set cellRng = sh.Range(tablename & "_START").Offset(-2)

    Do While (Not IsEmpty(cellRng))

        If cellRng.Offset(-4).Value = "list_auto_origin" Then

            varName = cellRng.Offset(1).Value
                
            If drop.Exists(varName) Then
                
                'Unique values (removing the spaces and the Null strings and keeping the case
                '(The remove duplicates doesn't do that))
                arrTable.FromExcelRange cellRng.Offset(2), _
                                        DetectLastColumn:=False, _
                                        DetectLastRow:=True

                Set arrTable = GetUniqueBA(arrTable)

                drop.ClearList varName
                drop.Update arrTable, varName
                drop.Sort varName, xlDescending

            End If

        End If
            
        Set cellRng = cellRng.Offset(, 1)
    Loop
End Sub


Public Sub EventValueChangeVList(Target As Range)

   
    Dim sh As Worksheet
    Dim rng As Range
    Dim rngLook As Range
    Dim varLabel As String
    Dim tablename As String

    On Error GoTo Err

    InitializeTrads

    Set sh = ActiveSheet
    tablename = sh.Cells(1, 4).Value

    'Calculate the range where the values are entered
    Set rng = sh.Range(tablename & "_" & "PLAGEVALUES")
    rng.calculate

    Set rng = sh.Range(tablename & "_" & GOTOSECCODE)

    If Not Intersect(Target, rng) Is Nothing Then
        varLabel = Replace(Target.Value, _
                           lltrads.Value("gotosection") & ": ", _
                           vbNullString)
        Set rngLook = sh.Cells.Find(What:=varLabel, lookAt:=xlWhole, MatchCase:=True)
        If Not rngLook Is Nothing Then rngLook.Activate
    End If

    Exit Sub
Err:
End Sub



'Selection change Event for updating geo dropdowns
Public Sub EventSelectionLinelist(ByVal Target As Range)

    Dim targetColumn As Long
    Dim sh As Worksheet
    Dim nbOffset As Long
    Dim hRng As Range
    Dim calcRng As Range
    Dim startLine As Long
    Dim varControl As String
    Dim tablename As String
    Dim adminTable As BetterArray
    Dim geo As ILLGeo
    Dim adminNames As BetterArray
    Dim drop As IDropdownLists


    On Error GoTo ErrHand

    Set sh = ActiveSheet
    tablename = sh.Cells(1, 4).Value
    startLine = sh.Range(tablename & "_START").Row

    'First test if we are on a good line
    If Target.Row < startLine Then Exit Sub
    
    'Calculate the line
    Set hRng = sh.ListObjects(1).HeaderRowRange
    nbOffset = Target.Row - hRng.Row
    Set calcRng = hRng.Offset(nbOffset)
    calcRng.calculate

    'Test for geo control (Exit if not the case)
    targetColumn = Target.Column
    varControl = sh.Cells(startLine - 5, targetColumn).Value
    
    If (varControl <> "geo2") And _
       (varControl <> "geo3") And _
       (varControl <> "geo4") Then _
        Exit Sub

    InitializeTrads

    Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
    Set adminNames = New BetterArray
    'This is mandatory for geolevel function, lowerbound should be 1
    adminNames.LowerBound = 1
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))

    Select Case varControl

    Case "geo2"
        
        BusyApp
        
        'Get admin 2 list for the corresponding admin1
        drop.ClearList "admin2"
        Set adminTable = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, Target.Offset(, -1).Value)
        'Build the validation list for adm2
        drop.Update adminTable, "admin2"
        NotBusyApp

    Case "geo3"
        'Adm3 has been selected, we will update the corresponding dropdown
        ' using admin1, and admin2
        BusyApp
        drop.ClearList "admin3"
        adminNames.Push Target.Offset(, -2).Value, Target.Offset(, -1).Value
        Set adminTable = geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, adminNames)
        drop.Update adminTable, "admin3"
        NotBusyApp

    Case "geo4"

        'Adm 4 has been selected, will update corresponding dropdown using admin1-3
        BusyApp
        drop.ClearList "admin4"
        'Take the adm4 table
        adminNames.Push Target.Offset(, -3).Value, Target.Offset(, -2).Value, Target.Offset(, -1).Value
        Set adminTable = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)
        drop.Update adminTable, "admin4"
        NotBusyApp

    End Select
ErrHand:
End Sub



Private Sub UpdateMultipleChoice(ByVal Target As Range, ByVal choiSep As String)

    Dim actualValue As String
    Dim prevValue As String
    Dim prevTab As BetterArray
    Dim actTab As BetterArray
    
    Set prevTab = New BetterArray
    Set actTab = New BetterArray

    If IsEmpty(Target) Then Exit Sub

    actualValue = Target.Value

    On Error Resume Next
    Application.Undo
    On Error GoTo Err

    If IsEmpty(Target) Then GoTo KeepValue

    prevValue = Target.Value
    prevTab.Items = Split(prevValue, choiSep)
    actTab.Items = Split(actualValue, choiSep)

    'Length reduction, keep values
    If (actTab.Length > 1) And (actTab.Length < prevTab.Length) Then GoTo KeepValue

    'There is no length reduction. Test for presence of actual element
    If prevTab.Includes(actualValue) Then
        Target.Value = prevValue
        Exit Sub
    End If

    'Length of actual Tab is One, previous Tab does not includes newvalue
    Target.Value = prevValue & choiSep & actualValue            
    Exit Sub

KeepValue:
    Target.Value = actualValue

Err:
End Sub