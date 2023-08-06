Attribute VB_Name = "LinelistEvents"

Option Explicit
Option Private Module

Public DebugMode As Boolean


'Trigerring event when the linelist sheet has some values within                                                          -                                                      -
Sub EventValueChangeLinelist(Target As Range)

    Const GOTOSECCODE As String = "go_to_section" 'Go To section constant

    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim varControl As String                   'Control type
    Dim sLabel As String
    Dim varName As String
    Dim varSubLabel As String
    Dim targetColumn As Long 'column of the target range
    Dim rng As Range
    Dim loAdm2 As ListObject
    Dim loAdm3 As ListObject
    Dim loAdm4 As ListObject
    Dim tableName As String
    Dim adminNames As BetterArray
    Dim sh As Worksheet 'Active sheet where the event fires
    Dim geo As ILLGeo
    Dim cellRng As Range
    Dim hRng As Range 'Header Row Range of the listObject
    Dim goToSection As String
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim startLine As Long
    Dim calcRng As Range 'calculate range
    Dim nbOffset As Long 'number of offset from the headerrow range

    On Error GoTo errHand
    Set sh = ActiveSheet
    tableName = sh.Cells(1, 4).Value
    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    Set hRng = sh.ListObjects(1).HeaderRowRange
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1

    targetColumn = Target.Column
    startLine = sh.Range(tableName & "_START").Row
    varControl = sh.Cells(startLine - 5, targetColumn).Value

    If Target.Row >= startLine Then

        nbOffset = Target.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)
        calcRng.calculate

        If (varControl = "geo1") Or (varControl = "geo2") Or (varControl = "geo3") Or (varControl = "geo4") Then

            Set loAdm2 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin2")
            Set loAdm3 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin3")
            Set loAdm4 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin4")


            Select Case varControl

            Case "geo1"
                'adm1 has been modified, we will correct and set validation to adm2

                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm2
                Target.Offset(, 1).Value = vbNullString
                DeleteLoDataBodyRange loAdm3
                Target.Offset(, 2).Value = vbNullString
                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 3).Value = vbNullString

                If Target.Value <> vbNullString Then

                    'Filter on adm1
                    Set T_geo = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, Target.Value)
                    'Build the validation list for adm2
                    T_geo.ToExcelRange loAdm2.Range.Cells(2, 1)
                    T_geo.Clear
                End If


                EndWork xlsapp:=Application

            Case "geo2"

                'Adm2 has been modified, we will correct and filter adm3
                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm3
                Target.Offset(, 1).Value = vbNullString
                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 2).Value = vbNullString

                If Target.Value <> vbNullString Then
                    adminNames.Push Target.Offset(, -1).Value, Target.Value
                    Set T_geo = geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, adminNames)
                    T_geo.ToExcelRange loAdm3.Range.Cells(2, 1)
                    T_geo.Clear
                End If

                EndWork xlsapp:=Application

            Case "geo3"
                'Adm 3 has been modified, correct and filter adm4
                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 1).Value = vbNullString

                If Target.Value <> vbNullString Then

                    adminNames.Push Target.Offset(, -2).Value, Target.Offset(, -1).Value, Target.Value
                    'Take the adm4 table
                    Set T_geo = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)
                    T_geo.ToExcelRange loAdm4.Range.Cells(2, 1)
                    T_geo.Clear
                End If

                EndWork xlsapp:=Application

            End Select
        End If

    End If

    'Update the custom control
    If (Target.Row = startLine - 2) And (varControl = "custom") Then
        Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 1, 1)
        Set vars = LLVariables.Create(dict)
        'The name of custom variables has been updated, update the dictionary
        varName = sh.Cells(startLine - 1, targetColumn).Value
        varSubLabel = vars.Value(varName:=varName, colName:="sub label")

        sLabel = Replace(Target.Value, varSubLabel, vbNullString)
        sLabel = Replace(sLabel, chr(10), vbNullString)

        vars.SetValue varName:=varName, colName:="main label", newValue:=sLabel

    End If

    'Update the list auto
    If Target.Row >= startLine And _
       sh.Cells(startLine - 6, targetColumn).Value = "list_auto_origin" And _
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value <> "list_auto_change_yes" Then
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_yes"
    End If


    'GoTo section
    If Not Intersect(Target, rng) Is Nothing Then
        goToSection = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToSection").Value

        sLabel = Replace(Target.Value, goToSection & ": ", vbNullString)
        Set hRng = sh.ListObjects(1).HeaderRowRange
        Set hRng = hRng.Offset(-3)

        Set cellRng = hRng.Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=True)

        If Not cellRng Is Nothing Then cellRng.Activate
    End If

    If Target.Row = startLine - 1 Then
        Target.Value = Target.Offset(-1).Name.Name
        MsgBox "You can't modify this header."
    End If

errHand:

End Sub


'Event to update the list_auto when a sheet containing a list_auto is desactivated
Public Sub EventDesactivateLinelist(ByVal sSheetName As String)

    Dim PrevWksh As Worksheet

    On Error GoTo errHand

    If ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_yes" Then

        Set PrevWksh = ThisWorkbook.Worksheets(sSheetName)
        BeginWork xlsapp:=Application

        UpdateListAuto PrevWksh
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_no"

        EndWork xlsapp:=Application
        Exit Sub

    End If
errHand:
    EndWork xlsapp:=Application
End Sub

'Update the list Auto of one Sheet

Public Sub UpdateListAuto(Wksh As Worksheet)

    Dim iChoiceCol As Integer
    Dim choiceLo As ListObject
    Dim sVarName As String
    Dim iRow As Long
    Dim i As Long
    Dim arrTable As BetterArray
    Dim listAutoSheet As Worksheet

    Dim rng As Range

    Set arrTable = New BetterArray
    i = 1

    Set listAutoSheet = ThisWorkbook.Worksheets(C_sSheetChoiceAuto)
    With Wksh
        .calculate
        Do While (.Cells(C_eStartLinesLLData, i) <> vbNullString)
            Select Case .Cells(C_eStartLinesLLMainSec - 2, i).Value
            Case C_sDictControlChoiceAuto & "_origin"
                sVarName = .Cells(C_eStartLinesLLData + 1, i).Value
                If ListObjectExists(listAutoSheet, "list_" & sVarName) Then
                    arrTable.FromExcelRange .Cells(C_eStartLinesLLData + 2, i), DetectLastColumn:=False, DetectLastRow:=True
                    'Unique values (removing the spaces and the Null strings and keeping the case (The remove duplicates doesn't do that))
                    Set arrTable = GetUniqueBA(arrTable)
                    With listAutoSheet
                        Set choiceLo = .ListObjects("list_" & sVarName)
                        iChoiceCol = choiceLo.Range.Column
                        If Not choiceLo.DataBodyRange Is Nothing Then choiceLo.DataBodyRange.Delete
                        arrTable.ToExcelRange .Cells(C_eStartlinesListAuto + 1, iChoiceCol)
                        iRow = .Cells(.Rows.Count, iChoiceCol).End(xlUp).Row
                        choiceLo.Resize .Range(.Cells(C_eStartlinesListAuto, iChoiceCol), .Cells(iRow, iChoiceCol))
                        'Sort in descending order
                        Set rng = choiceLo.ListColumns(1).Range
                        With choiceLo.Sort
                            .SortFields.Clear
                            .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, ORDER:=xlDescending
                            .Header = xlYes
                            .Apply
                        End With
                    End With
                End If
            Case Else
            End Select
            i = i + 1
        Loop
    End With

End Sub


Sub EventValueChangeVList(Target As Range)

    Const GOTOSECCODE As String = "go_to_section" 'Go To section constant

    Dim rng As Range
    Dim RngLook As Range
    Dim sLabel As String
    Dim sh As Worksheet
    Dim tableName As String
    Dim goToSection As String


    On Error GoTo Err
    Set sh = ActiveSheet
    tableName = sh.Cells(1, 4).Value

    'Calculate the range where the values are entered
    Set rng = sh.Range(tableName & "_" & "PLAGEVALUES")
    rng.calculate

    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)
    goToSection = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToSection").Value

    If Not Intersect(Target, rng) Is Nothing Then
        sLabel = Replace(Target.Value, goToSection & ": ", vbNullString)
        Set RngLook = sh.Cells.Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=True)
        If Not RngLook Is Nothing Then RngLook.Activate
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
    Dim tableName As String
    Dim loAdm2 As ListObject
    Dim loAdm3 As ListObject
    Dim loAdm4 As ListObject
    Dim T_geo As BetterArray
    Dim geo As ILLGeo
    Dim adminNames As BetterArray


    On Error GoTo errHand
    Set sh = ActiveSheet
    tableName = sh.Cells(1, 4).Value
    Set hRng = sh.ListObjects(1).HeaderRowRange

    targetColumn = Target.Column
    startLine = sh.Range(tableName & "_START").Row
    varControl = sh.Cells(startLine - 5, targetColumn).Value
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1

    If Target.Row < startLine Then Exit Sub

    nbOffset = Target.Row - hRng.Row
    Set calcRng = hRng.Offset(nbOffset)
    calcRng.calculate

    If (varControl <> "geo2") And _
     (varControl <> "geo3") And (varControl <> "geo4") Then Exit Sub

    Set loAdm2 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin2")
    Set loAdm3 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin3")
    Set loAdm4 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin4")

    Select Case varControl
       Case "geo2"
        'adm1 has been modified, we will correct and set validation to adm2
        BeginWork xlsapp:=Application
        If Target.Value <> vbNullString Then
            DeleteLoDataBodyRange loAdm2
          'Filter on adm1
           Set T_geo = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, Target.Offset(, -1).Value)
           'Build the validation list for adm2
            T_geo.ToExcelRange loAdm2.Range.Cells(2, 1)
            T_geo.Clear
        End If
        EndWork xlsapp:=Application

       Case "geo3"
        'Adm2 has been modified, we will correct and filter adm3
         BeginWork xlsapp:=Application
         If Target.Value <> vbNullString Then
            DeleteLoDataBodyRange loAdm3
            adminNames.Push Target.Offset(, -2).Value, Target.Offset(, -1).Value
            Set T_geo = geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, adminNames)
            T_geo.ToExcelRange loAdm3.Range.Cells(2, 1)
            T_geo.Clear
         End If
         EndWork xlsapp:=Application

       Case "geo4"
        'Adm 3 has been modified, correct and filter adm4
         BeginWork xlsapp:=Application

         If Target.Value <> vbNullString Then
            DeleteLoDataBodyRange loAdm4
            adminNames.Push Target.Offset(, -3).Value, Target.Offset(, -2).Value, Target.Offset(, -1).Value
            'Take the adm4 table
             Set T_geo = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)
             T_geo.ToExcelRange loAdm4.Range.Cells(2, 1)
             T_geo.Clear
         End If
        EndWork xlsapp:=Application

       End Select
errHand:
End Sub
