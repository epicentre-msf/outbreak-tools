Attribute VB_Name = "M_CreationFeuille"
Option Explicit

Const C_StartLineTitle1 As Byte = 3
Const C_StartLineTitle2 As Byte = 4
Const C_TitleLine As Byte = 5
Const C_ligneDeb As Byte = 6

Const C_CmdWidht As Byte = 60
Const C_PWD As String = "1234"

'Building the linelist from the different input data
'D_TitleDic: The headers of the dictionnary sheet
'T_dataDic: Dictionnary data
'D_Choices: The headers of the Choices sheet
'T_Choices: The choices data
'T_Export: The export data

Sub BuildList(D_TitleDic As Scripting.Dictionary, T_dataDic, D_Choices As Scripting.Dictionary, T_Choices, T_Export, sPath As String)
    'c'est parti pour le fichier de sortie !

    Dim xlsapp As New Excel.Application          'will contain the linelist
    Dim sExitPath As String
    Dim i As Integer                             'cpt result
    Dim j As Integer                             'cpt source
    Dim l As Integer                             'cpt nbcolonne
    Dim T_NbCol                                  'tab nb colonne dans une feuille determin�e
    Dim sPrevSheetName As String

    Dim oCell As Object                          'pour colors of titles
    Dim iPrevColS1 As Integer
    Dim iPrevColS2 As Integer
    Dim sTitle1 As String
    Dim sTitle2 As String

    Dim iDecNb As Integer                        'pour le calcul de dcimal
    Dim k As Integer
    Dim sNbDeci As String

    Dim sValidationList As String                'For validationList

    Dim bCmdGeoExist As Boolean                  'Check for existence of a geo button before creation
    Dim bCmdVisibleNameExist As Boolean
    Dim bCmdAddLine As Boolean
    Dim bCmdExport As Boolean
    Dim bSheetEvent As Boolean
    Dim bCmdExportMigration As Boolean

    Dim oKey As Variant

    Dim iPrevStartS1 As Integer

    Dim T_RowToHide
    Dim m As Byte
    Dim n As Byte

    Dim T_Formula
    Dim sFormula As String
    Dim O As Byte
    Dim sSheetname As String
    Dim sFormulaMin As String
    Dim sFormulaMax As String

    Dim p As Integer                             'cpt Admin

    Dim sPrevSheetNameSHP As String

    With xlsapp
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Visible = False
        .AutoCorrect.DisplayAutoCorrectOptions = False
        .Workbooks.Add
        .ActiveWorkbook.VBProject.References.AddFromFile ("C:\windows\system32\scrrun.dll") 'Using the scripting dictionnary dll
    
        DoEvents
        On Error Resume Next
        MkDir ("C:\LineListeApp\")               'create a folder for sending all the data from designer
        On Error GoTo 0
        DoEvents
        
        'TransferCode is for sending modules  (M) or classes (C) from actual excel workbook to another excel workbook
        Call TransferCode(xlsapp, "M_LineList", "M")
        
        'Transfert form is for sending forms from the actual excel workbook to another
        Call TransferForm(xlsapp, "F_Geo")
        Call TransferForm(xlsapp, "F_NomVisible")
        Call TransferForm(xlsapp, "F_Export")
    
        Call TransferCode(xlsapp, "M_Geo", "M")
        Call TransferCode(xlsapp, "M_NomVisible", "M")
        Call TransferCode(xlsapp, "M_FonctionsTransf", "M")
        Call TransferCode(xlsapp, "M_Export", "M")
        Call TransferCode(xlsapp, "M_Traduction", "M")
        Call TransferCode(xlsapp, "M_Migration", "M")
        Call TransferCode(xlsapp, "BetterArray", "C")
    
        DoEvents
    
        'TransfertSheet is for sending worksheets from the actual workbook to another
        Call TransfertSheet(xlsapp, "GEO")
        Call TransfertSheet(xlsapp, "PASSWORD")
        'on a besoin de la table ascii
        Call TransfertSheet(xlsapp, "ControleFormule")
    
        DoEvents
        On Error Resume Next
        RmDir ("C:\LineListeApp\")
        On Error GoTo 0
        
        '-------------- Creating the dictionnary sheet
        .Sheets.Add.Name = "Dico"
        i = 1
        For Each oKey In D_TitleDic.Keys
            .Sheets("dico").Cells(1, i).value = oKey
            i = i + 1
        Next oKey
        .Sheets("dico").Range("A2").Resize(UBound(T_dataDic, 2) + 1, D_TitleDic.Count) = .WorksheetFunction.Transpose(T_dataDic)
        .Sheets("dico").Visible = False
    
        '-------------- Creating the export sheet
        .Sheets.Add.Name = "Export"
        .Sheets("Export").Cells(1, 1).value = "ID"
        .Sheets("Export").Cells(1, 2).value = "Lbl"
        .Sheets("Export").Cells(1, 3).value = "Pwd"
        .Sheets("Export").Cells(1, 4).value = "Actif"
        .Sheets("Export").Cells(1, 5).value = "FileName"
        .Sheets("Export").Range("A2").Resize(UBound(T_Export, 2) + 1, UBound(T_Export, 1) + 1) = .WorksheetFunction.Transpose(T_Export)
        .Sheets("Export").Visible = False
    
        '--------------- adding the other sheets in the dictionary
        i = 1
        j = 0
        ReDim T_NbCol(j)
        sPrevSheetName = ""
        While i <= UBound(T_dataDic, 2)
            If LCase(T_dataDic(D_TitleDic("Sheet") - 1, i)) = "admin" Then
                .Sheets(1).Name = "Admin"
            Else
                If sPrevSheetName <> T_dataDic(D_TitleDic("Sheet") - 1, i) Then
                    j = j + 1
                    'T_NbCol is a table with number columns for each sheet T(1) the number of columns
                    'of first sheet (linelist-patient).
                    ReDim Preserve T_NbCol(j)
                    T_NbCol(j) = 1
                    .Worksheets.Add.Name = T_dataDic(D_TitleDic("Sheet") - 1, i)
                    .Worksheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Rows("1:2").RowHeight = 25
                
                    .ActiveWindow.DisplayZeros = False
                    .ActiveWindow.SplitColumn = 2
                    .ActiveWindow.SplitRow = 5
                    .ActiveWindow.FreezePanes = True
                
                    sPrevSheetName = T_dataDic(D_TitleDic("Sheet") - 1, i)
                    ReDim Preserve T_NbCol(UBound(T_NbCol) + 1)
                Else
                    T_NbCol(j) = T_NbCol(j) + 1
                End If
            End If
            'i will hep move from one values of dictionnary data to another But here
            'we focus mainly on the Sheet column
            i = i + 1
        Wend
    End With

    sPrevSheetName = ""                          'Checking if we moved from one sheet to another
    sTitle1 = ""                                 'First title on the linelist-patient sheet (or on every other sheet)
    sTitle2 = ""                                 'Second title on the linelist-patient sheet (or on every other sheet)

    iPrevColS1 = 1                               'Counter for previous columns for the mainlabel
    iPrevColS2 = 1                               'Counter for previous columns for the sub-section

    ReDim T_RowToHide(0)                         'Table for the number of columns to hide

    j = 1                                        'cpt result
    i = 0                                        'Dictionnary iterator (columns of the dictionnary)
    l = 0                                        'iterators for the number of colums in one sheet
    p = C_TitleLine
    While i <= UBound(T_dataDic, 2)
        With xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i))
            If LCase(T_dataDic(D_TitleDic("Sheet") - 1, i)) <> "admin" Then
                
                If sPrevSheetName <> T_dataDic(D_TitleDic("Sheet") - 1, i) Then 'only on new sheet
                    If sPrevSheetName <> "" Then 'only if the sheet name is not empty
                        With xlsapp.Sheets(sPrevSheetName)
                        
                            'Title 1
                            'C_StartLineTitle1: Line for the first title
                            'iPrevColS1: previous column for the first title
                            'iPreColS2: previous column for the second title
                            'j: is absolutely greater than 1 in this section
                            
                            .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1, j - 1)).Merge
                            .Cells(C_StartLineTitle1, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                            .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1 + 1, j - 1)).Interior.Color = LetColor("DarkBlueTitle")
                            Call WriteBorderLines(.Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle2, j - 1)))
                            
                            'Title 2
                            If .Cells(C_StartLineTitle2, iPrevColS2) <> "" Then
                                .Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)).Merge
                                'why 1 and not 2 on iprevcol?
                                .Cells(C_StartLineTitle2, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                                .Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)).Interior.Color = LetColor("LightBlueTitle")
                                Call WriteBorderLines(.Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)))
                            End If
                        End With
                    End If
                
                    'Creation of the list object for the first time on one Sheet
                    j = 1
                    l = l + 1
                    sPrevSheetName = T_dataDic(D_TitleDic("Sheet") - 1, i)
                    .ListObjects.Add(xlSrcRange, .Range(.Cells(C_TitleLine, 1), .Cells(C_TitleLine, T_NbCol(l))), , xlYes).Name = "o" & T_dataDic(D_TitleDic("Sheet") - 1, i)
                    .ListObjects("o" & T_dataDic(D_TitleDic("Sheet") - 1, i)).TableStyle = "TableStyleLight16"
                
                    .Cells.Font.Size = 9
                            
                    iPrevStartS1 = 1
                    
                    'First title
                    sTitle1 = T_dataDic(D_TitleDic("Main section") - 1, i)
                    .Cells(C_StartLineTitle1, j).value = T_dataDic(D_TitleDic("Main section") - 1, i)
                
                    bCmdGeoExist = False
                    bCmdVisibleNameExist = False
                    bSheetEvent = False
                
                End If
            
                'Headers
                .Cells(C_TitleLine, j).Name = Replace(T_dataDic(D_TitleDic("Variable name") - 1, i), " ", "_")
                .Cells(C_TitleLine, j).value = LetWordingWithSpace(xlsapp, CStr(T_dataDic(D_TitleDic("Main label") - 1, i)), CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)))
                .Cells(C_TitleLine, j).VerticalAlignment = xlTop
                
                'Adding the sub-label if needed Chr(10) is the return to line character the sublabel is in gray
                If T_dataDic(D_TitleDic("Sub-label") - 1, i) <> "" Then
                    .Cells(C_TitleLine, j).value = .Cells(C_TitleLine, j).value & Chr(10) & T_dataDic(D_TitleDic("Sub-label") - 1, i)
                    
                    'Changing the fontsize of the sublabels
                    .Cells(C_TitleLine, j).Characters(Start:=Len(T_dataDic(D_TitleDic("Main label") - 1, i)) + 1, Length:=Len(T_dataDic(D_TitleDic("Sub-label") - 1, i)) + 1).Font.Size = 8
                    .Cells(C_TitleLine, j).Characters(Start:=Len(T_dataDic(D_TitleDic("Main label") - 1, i)) + 1, Length:=Len(T_dataDic(D_TitleDic("Sub-label") - 1, i)) + 1).Font.Color = LetColor("Grey")
                End If
            
                'Adding the notes as comment
                If T_dataDic(D_TitleDic("Note") - 1, i) <> "" Then
                    .Cells(C_TitleLine, j).AddComment
                    .Cells(C_TitleLine, j).Comment.Text Text:=T_dataDic(D_TitleDic("Note") - 1, i)
                    .Cells(C_TitleLine, j).Comment.Visible = False
                End If
            
                'Geo Titles or Customs
                Select Case LCase(T_dataDic(D_TitleDic("Control") - 1, i))
                Case "geo"
                    If T_dataDic(D_TitleDic("Sub-section") - 1, i) = "" Then
                        T_dataDic(D_TitleDic("Sub-section") - 1, i) = T_dataDic(D_TitleDic("Main label") - 1, i)
                    End If
                Case "custom"
                    .Cells(C_TitleLine, j).Locked = False
                End Select
                
                'Now the sections
                If sTitle1 <> T_dataDic(D_TitleDic("Main section") - 1, i) Then
                
                    'Merge previous cells if the title changes
                    .Cells(C_StartLineTitle1, j).value = T_dataDic(D_TitleDic("Main section") - 1, i)
                    sTitle1 = T_dataDic(D_TitleDic("Main section") - 1, i)
                
                    .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1, j - 1)).Merge
                    .Cells(C_StartLineTitle1, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                    .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1, j - 1)).Interior.Color = LetColor("DarkBlueTitle")
    
                    For Each oCell In .Range(.Cells(C_StartLineTitle2, iPrevColS1), .Cells(C_StartLineTitle2, j - 1))
                        If oCell.value = "" Then
                            oCell.Interior.Color = LetColor("DarkBlueTitle")
                        End If
                    Next
                    Set oCell = Nothing
                    Call WriteBorderLines(.Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle2, j - 1)))
                
                    iPrevColS1 = j
                Else
                    If i = UBound(T_dataDic, 2) Then 'Derniere case
                        .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1, j)).Merge
                        .Cells(C_StartLineTitle1, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                        .Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle1, j)).Interior.Color = LetColor("DarkBlueTitle")
                        For Each oCell In .Range(.Cells(C_StartLineTitle2, iPrevColS1), .Cells(C_StartLineTitle2, j)) 'coloriage
                            If oCell.value = "" Then
                                oCell.Interior.Color = LetColor("DarkBlueTitle")
                            End If
                        Next
                        Set oCell = Nothing
                        Call WriteBorderLines(.Range(.Cells(C_StartLineTitle1, iPrevColS1), .Cells(C_StartLineTitle2, j)))
                    End If
                End If
            
                If sTitle2 <> T_dataDic(D_TitleDic("Sub-section") - 1, i) Then
                    'si le titre change, on fusionne les prec cellules
                    .Cells(C_StartLineTitle2, j).value = T_dataDic(D_TitleDic("Sub-section") - 1, i)
                
                    sTitle2 = T_dataDic(D_TitleDic("Sub-section") - 1, i)
                    If j > 1 Then
                        If .Cells(C_StartLineTitle2, iPrevColS2) <> "" Then
                            .Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)).Merge
                            .Cells(C_StartLineTitle2, iPrevColS2).MergeArea.HorizontalAlignment = xlCenter
                            .Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)).Interior.Color = LetColor("LightBlueTitle")
                            Call WriteBorderLines(.Range(.Cells(C_StartLineTitle2, iPrevColS2), .Cells(C_StartLineTitle2, j - 1)))
                        End If
                    Else
                        If .Cells(C_StartLineTitle2, iPrevColS2) <> "" Then
                            .Cells(C_StartLineTitle2, iPrevColS2).HorizontalAlignment = xlCenter
                            .Cells(C_StartLineTitle2, iPrevColS2).Interior.Color = LetColor("LightBlueTitle")
                            Call WriteBorderLines(.Cells(C_StartLineTitle2, iPrevColS2))
                        End If
                    End If
                    iPrevColS2 = j
                End If
            
                .Columns(j).EntireColumn.AutoFit
            
                'Status champ obligatoire
                Select Case LCase(T_dataDic(D_TitleDic("Status") - 1, i))
                Case "mandatory"
                    If T_dataDic(D_TitleDic("Note") - 1, i) <> "" Then
                        .Cells(C_TitleLine, j).Comment.Text Text:="Mandatory data" & Chr(10) & T_dataDic(D_TitleDic("Note") - 1, i)
                    Else
                        .Cells(C_TitleLine, j).AddComment
                        .Cells(C_TitleLine, j).Comment.Text Text:="Mandatory data"
                        .Cells(C_TitleLine, j).Comment.Visible = False
                    End If
                Case "optional"
            
                Case "hidden"
                    .Columns(j).EntireColumn.Hidden = True
                
                End Select
            
                'Protection of the following cell (normally C_TitleLine+1)?
                .Cells(6, j).Locked = False
            
                'typage
                If T_dataDic(D_TitleDic("Type") - 1, i) <> "" Then
                    Select Case LCase(T_dataDic(D_TitleDic("Type") - 1, i))
                    Case "text"
                        .Cells(6, j).NumberFormat = "@"
                    Case "date"
                        .Cells(6, j).NumberFormat = "d-mmm-yyyy"
                    Case "integer"
                        .Cells(6, j).NumberFormat = "0"
                    Case Else
                        If InStr(1, LCase(T_dataDic(D_TitleDic("Type") - 1, i)), "decimal") > 0 Then 'decimal
                            iDecNb = Right(T_dataDic(D_TitleDic("Type") - 1, i), 1)
                            k = 0
                            While k < iDecNb
                                k = k + 1
                            Wend
                            'Only the last character is extracted, so you can have up to 9 digits maximum
                            .Cells(6, j).NumberFormat = "0." & LetDecString(Right(T_dataDic(D_TitleDic("Type") - 1, i), 1))
                        End If
                    End Select
                End If
            
                'Choices / geo et HF
                If T_dataDic(D_TitleDic("Control") - 1, i) <> "" Then
                    Select Case LCase(T_dataDic(D_TitleDic("Control") - 1, i))
                    Case "choices"
    
                        If T_dataDic(D_TitleDic("Choices") - 1, i) <> "" Then
                            sValidationList = GetValidationName(T_Choices, D_Choices, CStr(T_dataDic(D_TitleDic("Choices") - 1, i)))
                            If sValidationList <> "" Then
                                Call LetValidationList(.Cells(6, j), sValidationList, LetValidationLockType(CStr(T_dataDic(D_TitleDic("Alert") - 1, i))), CStr(T_dataDic(D_TitleDic("Message") - 1, i)))
                            End If
                        End If
        
                    Case "geo", "hf"
                        'ajouter colonnes  pour geo
                        .Cells(C_TitleLine, j).Interior.Color = LetColor("Orange")
                        'update the columns only for the geo and do nothing for health facility
                        If LCase(T_dataDic(D_TitleDic("Control") - 1, i)) = "geo" Then
                            Call Add4GeoCol(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), CStr(T_dataDic(D_TitleDic("Main label") - 1, i)), Replace(T_dataDic(D_TitleDic("Variable name") - 1, i), " ", "_"), j, CStr(T_dataDic(D_TitleDic("Message") - 1, i)))
                            j = j + 3
                        End If
                    
                        'adding the button command for the geo only one time
                        If Not bCmdGeoExist Then
                            Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(1, 1).Left, .Cells(1, 1).Top, "SHP_GeoApps", "Geo", C_CmdWidht, 20)
                            With .Shapes("SHP_GeoApps").Fill
                                .Visible = msoTrue
                                .ForeColor.RGB = LetColor("Orange")
                                .BackColor.RGB = LetColor("Orange")
                                '.TwoColorGradient msoGradientHorizontal, 1
                            End With
                            .Shapes("SHP_GeoApps").OnAction = "ClicCmdGeoApps"
                            bCmdGeoExist = True
                        End If
                    Case Else
                
                    End Select
            
                End If
        
                'min max simple / Les complexes sont en dessous
                If T_dataDic(D_TitleDic("Min") - 1, i) <> "" And T_dataDic(D_TitleDic("Max") - 1, i) <> "" Then
                    If IsNumeric(T_dataDic(D_TitleDic("Min") - 1, i)) And IsNumeric(T_dataDic(D_TitleDic("Max") - 1, i)) Then
                        Call BuildValidationMinMax(.Cells(6, j), CStr(T_dataDic(D_TitleDic("Min") - 1, i)), CStr(T_dataDic(D_TitleDic("Max") - 1, i)), LetValidationLockType(CStr(T_dataDic(D_TitleDic("Alert") - 1, i))), CStr(T_dataDic(D_TitleDic("Type") - 1, i)), CStr(T_dataDic(D_TitleDic("Message") - 1, i)))
                        .Cells(6, j).Locked = False
                    End If
                End If
                    
                'buttons for show/hide
                If Not bCmdVisibleNameExist Then
                    Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(2, 1).Left, .Cells(2, 1).Top, "SHP_NomVisibleApps", "Show/Hide", C_CmdWidht, 20)
                    .Shapes("SHP_NomVisibleApps").Fill.ForeColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_NomVisibleApps").Fill.BackColor.RGB = LetColor("DarkBlueTitle")
                    '.Shapes("SHP_NomVisibleApps").Fill.TwoColorGradient msoGradientHorizontal, 1
                    .Shapes("SHP_NomVisibleApps").OnAction = "ClicCmdVisibleName"
                    bCmdVisibleNameExist = True
                End If
                'buttons for adding 200 lines
                If Not bCmdAddLine Then
                    Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(1, 1).Left + C_CmdWidht + 10, .Cells(1, 2).Top, "SHP_Ajout200L", "Add rows", C_CmdWidht, 20)
                    .Shapes("SHP_Ajout200L").Fill.ForeColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_Ajout200L").Fill.BackColor.RGB = LetColor("DarkBlueTitle")
                    '.Shapes("SHP_Ajout200L").Fill.TwoColorGradient msoGradientHorizontal, 1
                    .Shapes("SHP_Ajout200L").OnAction = "clicAdd200L"
                    bCmdAddLine = True
                End If
                'button for export
                If Not bCmdExport Then
                    Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(2, 1).Left + C_CmdWidht + 10, .Cells(2, 2).Top, "SHP_Export", "Export", C_CmdWidht, 20)
                    .Shapes("SHP_Export").Fill.ForeColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_Export").Fill.BackColor.RGB = LetColor("DarkBlueTitle")
                    '.Shapes("SHP_Export").Fill.TwoColorGradient msoGradientHorizontal, 1
                    .Shapes("SHP_Export").OnAction = "clicExport"
                    bCmdExport = False
                End If
        
                j = j + 1
        
            Else                                 'on admin sheet
                'This part should be removed?
                If sPrevSheetName <> "" Then     'on conclue les titres de la feuille prec
                    xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1, iPrevColS1), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1, j - 1)).Merge 'titre1
                    xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                    xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1, iPrevColS1), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1 + 1, j - 1)).Interior.Color = LetColor("DarkBlueTitle")
                    Call WriteBorderLines(xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle1, iPrevColS1), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, j - 1)))
                
                    If xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, iPrevColS2) <> "" Then
                        xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, iPrevColS2), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, j - 1)).Merge 'titre2
                        xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, iPrevColS1).MergeArea.HorizontalAlignment = xlCenter
                        xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, iPrevColS2), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, j - 1)).Interior.Color = LetColor("LightBlueTitle")
                        Call WriteBorderLines(xlsapp.Sheets(sPrevSheetName).Range(xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, iPrevColS2), xlsapp.Sheets(sPrevSheetName).Cells(C_StartLineTitle2, j - 1)))
                    End If
                End If

                'creating admin layout
                .Cells(p, 2).value = T_dataDic(D_TitleDic("Main label") - 1, i)
                .Cells(p, 2).Interior.Color = LetColor("LightBlueTitle")
                .Cells(p, 3).Name = T_dataDic(D_TitleDic("Variable name") - 1, i)
                Call WriteBorderLines(.Cells(p, 3))
            
                If LCase(T_dataDic(D_TitleDic("Control") - 1, i)) = "choices" Then
                    If T_dataDic(D_TitleDic("Choices") - 1, i) <> "" Then
                        sValidationList = GetValidationName(T_Choices, D_Choices, CStr(T_dataDic(D_TitleDic("Choices") - 1, i)))
                        If sValidationList <> "" Then
                            Call LetValidationList(.Cells(p, 3), sValidationList, LetValidationLockType(CStr(T_dataDic(D_TitleDic("Alert") - 1, i))), CStr(T_dataDic(D_TitleDic("Message") - 1, i)))
                        End If
                    End If
                End If
            
                If Not bCmdExportMigration Then
                    Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(1, 5).Left + 10, .Cells(2, 1).Top, "SHP_ExportMig", "Export for" & Chr(10) & "migration", C_CmdWidht + 10, 30)
                    .Shapes("SHP_ExportMig").Fill.ForeColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_ExportMig").Fill.BackColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_ExportMig").OnAction = "clicExportMigration"
                
                    Call AddCmd(xlsapp, CStr(T_dataDic(D_TitleDic("Sheet") - 1, i)), .Cells(1, 5).Left + 20 + .Shapes("SHP_ExportMig").Width, .Cells(2, 1).Top, "SHP_ImportMig", "Import from" & Chr(10) & "migration", C_CmdWidht + 10, 30)
                    .Shapes("SHP_ImportMig").Fill.ForeColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_ImportMig").Fill.BackColor.RGB = LetColor("DarkBlueTitle")
                    .Shapes("SHP_ImportMig").OnAction = "clicImportMigration"
                
                    'pour le logo
                    sPrevSheetNameSHP = xlsapp.ActiveSheet.Name
                    Sheets("Main").Shapes("SHP_Logo").Copy
                    xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Select
                    xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Range("A1").Select
                    xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Paste
                    xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Range("C5").Select
                    xlsapp.Sheets(sPrevSheetName).Select
                    Sheets("Main").Range("a1").Select
                    bCmdExportMigration = True
                End If
            
                p = p + 2
            
            End If
            i = i + 1

        End With
    Wend

    sPrevSheetName = ""

    i = 0
    While i <= UBound(T_dataDic, 2)
        If LCase(T_dataDic(D_TitleDic("Control") - 1, i)) = "formula" Then 'pav� pour le controle de formule
            If T_dataDic(D_TitleDic("Formula") - 1, i) <> "" Then
                'sFormula = UCase(Replace(T_dataDic(D_TitleDic("Formula") - 1, i), " ", ""))
                sFormula = T_dataDic(D_TitleDic("Formula") - 1, i)
                T_Formula = ControlValidationFormula(sFormula, T_dataDic, D_TitleDic, False)
                If Not IsEmptyTable(T_Formula) Then
                    With xlsapp
                        If T_Formula(0) <> "" Then
                            sSheetname = T_dataDic(D_TitleDic("Sheet") - 1, i)
                            j = 0                'on transcrit la formule
                            While j <= UBound(T_Formula)
                                If InStr(1, UCase(sFormula), T_Formula(j)) > 0 Then
                                    sFormula = Replace(UCase(sFormula), UCase(T_Formula(j)), Split(.Cells(, LetColNumberByDataName(xlsapp, CStr(T_Formula(j)), sSheetname)).Address, "$")(1) & C_ligneDeb)
                                End If
                                j = j + 1
                            Wend
                            'on ecrit la formule a la bonne place
                            j = 1
                            While j <= .Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Cells(C_TitleLine, 1).End(xlToRight).Column _
        And .Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Cells(C_TitleLine, j).Name.Name <> T_dataDic(D_TitleDic("Variable name") - 1, i)
                                j = j + 1
                            Wend
                            If .Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Cells(C_TitleLine, j).Name.Name = T_dataDic(D_TitleDic("Variable name") - 1, i) Then
                                .Sheets(sSheetname).Cells(6, j).NumberFormat = "General"
                                .Sheets(sSheetname).Cells(6, j).Formula = "=" & sFormula
                                On Error Resume Next
                                .Sheets(sSheetname).Cells(6, j).Formula2 = "=" & sFormula 'etrange... facultatif sur certaines machines
                                On Error GoTo 0
                                .Sheets(sSheetname).Cells(6, j).Locked = True
                            End If
                        Else
                            MsgBox "Invalid formula will be ignored : " & sFormula 'MSG_InvalidFormula
                        End If
                    End With
                Else
                    MsgBox "Invalid formula will be ignored : " & sFormula 'MSG_InvalidFormula
            
                End If
            End If
        End If
        'min / max en formule
        If T_dataDic(D_TitleDic("Min") - 1, i) <> "" And T_dataDic(D_TitleDic("Max") - 1, i) <> "" Then
            If Not IsNumeric(T_dataDic(D_TitleDic("Min") - 1, i)) And Not IsNumeric(T_dataDic(D_TitleDic("Max") - 1, i)) Then
                'sFormulaMin = UCase(Replace(T_dataDic(D_TitleDic("Min") - 1, i), " ", "")) 'min
                sFormulaMin = T_dataDic(D_TitleDic("Min") - 1, i)
                If IsAFunction(Replace(sFormulaMin, "()", "")) Then
                
                    sFormulaMin = LetInternationalFormula(sFormulaMin)
                
                    If Right(sFormulaMin, 2) <> "()" Then
                        sFormulaMin = sFormulaMin & "()"
                    End If
                Else
                    T_Formula = ControlValidationFormula(sFormulaMin, T_dataDic, D_TitleDic, True)
                    If Not IsEmptyTable(T_Formula) Then
                        sSheetname = T_dataDic(D_TitleDic("Sheet") - 1, i)
                        j = 0
                        While j <= UBound(T_Formula)
                            If T_Formula(j) <> "" Then
                                If InStr(1, T_Formula(j), Chr(124)) Then
                                    sFormulaMin = Replace(UCase(sFormulaMin), Split(T_Formula(j), Chr(124))(0), Split(T_Formula(j), Chr(124))(1)) 's'il y a un pipe (alt 6) : c'est forcement une formule. On remplace donc l'ancienne par la fonction propre au systeme
                                ElseIf InStr(1, UCase(sFormulaMin), T_Formula(j)) > 0 And Not IsAFunction(CStr(T_Formula(j))) Then
                                    sFormulaMin = Replace(UCase(sFormulaMin), UCase(T_Formula(j)), Split(xlsapp.Cells(, LetColNumberByDataName(xlsapp, CStr(T_Formula(j)), sSheetname)).Address, "$")(1) & C_ligneDeb) 'sans pipe, c'est un nom de variable, on recupere uniquement la colonne
                                End If
                            End If
                            j = j + 1
                        Wend
                    End If
                End If
        
                'sFormulaMax = UCase(Replace(T_dataDic(D_TitleDic("Max") - 1, i), " ", "")) 'max
                sFormulaMax = T_dataDic(D_TitleDic("Max") - 1, i)
                If IsAFunction(Replace(sFormulaMax, "()", "")) Then
                    sFormulaMax = LetInternationalFormula(Replace(sFormulaMax, "()", ""))
                
                    If Right(sFormulaMax, 2) <> "()" Then
                        sFormulaMax = sFormulaMax & "()"
                    End If
                Else
                    T_Formula = ControlValidationFormula(sFormulaMax, T_dataDic, D_TitleDic, True)
                    If Not IsEmptyTable(T_Formula) Then
                        sSheetname = T_dataDic(D_TitleDic("Sheet") - 1, i)
                        j = 0
                        While j <= UBound(T_Formula)
                            If InStr(1, T_Formula(j), Chr(124)) Then
                                sFormulaMax = Replace(UCase(sFormulaMax), Split(T_Formula(j), Chr(124))(0), Split(T_Formula(j), Chr(124))(1)) 's'il y a un pipe (alt 6) : c'est forcement une formule. On remplace donc l'ancienne par la fonction propre au systeme
                            ElseIf InStr(1, UCase(sFormulaMax), T_Formula(j)) > 0 And Not IsAFunction(CStr(T_Formula(j))) Then
                                sFormulaMax = Replace(UCase(sFormulaMax), UCase(T_Formula(j)), Split(xlsapp.Cells(, LetColNumberByDataName(xlsapp, CStr(T_Formula(j)), sSheetname)).Address, "$")(1) & C_ligneDeb) 'sans pipe, c'est un nom de variable, on recupere uniquement la colonne
                            End If
                            j = j + 1
                        Wend
                    End If
                End If
            
                'pour ecrire la validation min/max, on recherche la position de champ dans le fichier final
                If sFormulaMin <> "" And sFormulaMax <> "" Then
                    With xlsapp
                        j = 1
                        While j <= .Sheets(sSheetname).Cells(C_TitleLine, 1).End(xlToRight).Column And T_dataDic(D_TitleDic("Variable name") - 1, i) <> .Sheets(sSheetname).Cells(C_TitleLine, j).Name.Name
                            j = j + 1
                        Wend
                        If T_dataDic(D_TitleDic("Variable name") - 1, i) = .Sheets(sSheetname).Cells(C_TitleLine, j).Name.Name Then
                            Call BuildValidationMinMax(.Sheets(sSheetname).Cells(6, j), "=" & sFormulaMin, "=" & sFormulaMax, LetValidationLockType(CStr(T_dataDic(D_TitleDic("Alert") - 1, i))), CStr(T_dataDic(D_TitleDic("Type") - 1, i)), CStr(T_dataDic(D_TitleDic("Message") - 1, i)))
                            '.Sheets(sSheetName).Cells(6, j).Locked = True  'verouille les dates ?
                        End If
                    End With
                End If
            End If
        End If
    
        i = i + 1
    Wend

    Call Add200Lines(xlsapp)

    'on (presque) conclue !
   'Application.ActiveWindow.WindowState = xlMinimized
    With xlsapp
        .Sheets("admin").Columns(2).EntireColumn.AutoFit
        .Sheets(6).Select
        .Sheets(6).Range("A1").Select
        .DisplayAlerts = False
        .ScreenUpdating = False
        '.Visible = True
        .ActiveWindow.DisplayZeros = True
        '.ActiveWindow.WindowState = xlMaximized
    End With
    DoEvents

    sPrevSheetName = ""
    i = 0
    While i <= UBound(T_dataDic, 2)
        If sPrevSheetName <> T_dataDic(D_TitleDic("Sheet") - 1, i) Then
            If LCase(T_dataDic(D_TitleDic("Sheet") - 1, i)) <> "admin" Then
                xlsapp.Sheets(T_dataDic(D_TitleDic("Sheet") - 1, i)).Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                                                                                , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
            End If
            sPrevSheetName = T_dataDic(D_TitleDic("Sheet") - 1, i)
        End If
        i = i + 1
    Wend

    'ecriture de l'evenement "change" dans la feuille de resultat
    Call TransferCodeWks(xlsapp, "linelist-patient", "linelist_sheet_change")
    
    xlsapp.ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, ConflictResolution:=xlLocalSessionChanges
    xlsapp.Quit
    Set xlsapp = Nothing
End Sub

Private Sub Add200Lines(xlsapp As Excel.Application)

    Dim oKey As Variant                          'cpt dico attention typage special !
    Dim oLstobj As Object
    Dim oSheet As Object

    With xlsapp
        For Each oSheet In .ActiveWorkbook.Sheets 'on se cr�e les 200 premieres lignes
            If oSheet.Name <> "GEO" And oSheet.Name <> "TRANSLATION" And oSheet.Name <> "Dico" And oSheet.Name <> "Password" And oSheet.Name <> "ControleFormule" Then
                For Each oLstobj In oSheet.ListObjects
                    'maybe check here to be sure the listobject is a table first?
                    oLstobj.Resize oSheet.Range(oSheet.Cells(C_TitleLine, 1), oSheet.Cells(200 + C_TitleLine, oSheet.Cells(C_TitleLine, 1).End(xlToRight).Column))
                Next
            End If
        Next oSheet
    End With

End Sub

'Set the type of a validation of a cell giving the name of the validation in the
'dictionary
Function LetValidationLockType(sValidationLockType As String) As Byte

    LetValidationLockType = 3                    'liste de validation info, warning ou erreur
    If sValidationLockType <> "" Then
        Select Case LCase(sValidationLockType)
        Case "warning"
            LetValidationLockType = 2
        Case "error"
            LetValidationLockType = 1
        End Select
    End If
    
End Function

'adding a validation list in an excel range
Sub LetValidationList(oRange As Range, sValidList As String, sAlertType As Byte, sMessage As String)
    With oRange.Validation
        .Delete
        Select Case sAlertType
        Case 1                                   '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sValidList
        Case 2                                   '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=sValidList
        Case Else                                'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=sValidList
        End Select
        
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .ErrorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'Setting the min and the max validation in
Sub BuildValidationMinMax(oRange As Range, iMin As String, iMax As String, iAlertType As Byte, sTypeValidation As String, sMessage As String)

    With oRange.Validation
        .Delete
        Select Case LCase(sTypeValidation)
        Case "integer"                           'numerique
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case "date"                              'date
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case Else                                'decimal
            If InStr(1, LCase(sTypeValidation), "decimal") > 0 Then
                Select Case iAlertType
                Case 1                           '"error"
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                Case 2                           '"warning"
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                Case Else
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                End Select
            End If
        End Select
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .ErrorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'Write the borders for one range
Sub WriteBorderLines(oRange As Range)

    Dim i As Integer
    For i = 7 To 10
        With oRange.Borders(i)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next

End Sub

Sub AddCmd(xlsapp As Excel.Application, sSheet As String, iLeft As Integer, iTop As Integer, sName As String, sText As String, iCmdWidth As Integer, iCmdHeight As Integer)

    Dim oShape As Object
    Dim bShapeExist As Boolean

    bShapeExist = False
    For Each oShape In xlsapp.Sheets(sSheet).Shapes
        If oShape.Name = sName Then
            bShapeExist = True
            Exit For
        End If
    Next

    If Not bShapeExist Then
        With xlsapp.Sheets(sSheet)
            .Shapes.AddShape(msoShapeRectangle, iLeft + 3, iTop + 3, iCmdWidth, iCmdHeight).Name = sName
            .Shapes(sName).Placement = xlFreeFloating
            .Shapes(sName).TextFrame2.TextRange.Characters.Text = sText
            .Shapes(sName).TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .Shapes(sName).TextFrame2.VerticalAnchor = msoAnchorMiddle
            .Shapes(sName).TextFrame2.WordWrap = msoFalse
            .Shapes(sName).TextFrame2.TextRange.Font.Size = 9
            xlsapp.Sheets(sSheet).Shapes(sName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbBlack
            '.Shapes(sName).ShapeStyle = msoShapeStylePreset30
        End With
    End If

End Sub

'The purpose of this procedure is to create the geo columns using the geo data  (its also adds the first dropdowns)
' we shift the columns to the right until we reached the number of columns required
Sub Add4GeoCol(xlsapp As Excel.Application, sSheetname As String, sLib As String, sNameCell As String, iCol As Integer, sMessage As String)

    'sSheetName: Sheet name
    'sNameCell: Name of the cell
    'iCol: Column to start shifting
    'sMessage: message in case of error
    'sLib: header message
    
    Dim i As Byte
    Dim j As Byte
    Dim sTemp As String

    With xlsapp.Sheets(sSheetname)
        i = 4
        While i > 1
            .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            .Cells(C_TitleLine, iCol + 1).value = LetWordingWithSpace(xlsapp, Sheets("GEO").ListObjects("T_ADM" & i).HeaderRowRange.Item(i).value, CStr(sSheetname))
            .Cells(C_TitleLine, iCol + 1).Name = "adm" & i & "_" & sNameCell
            .Cells(C_TitleLine, iCol + 1).Interior.Color = vbWhite
            .Cells(C_TitleLine, iCol + 1).Locked = False
            i = i - 1
        Wend
        .Cells(C_TitleLine, iCol).value = LetWordingWithSpace(xlsapp, Sheets("GEO").ListObjects("T_ADM" & i).HeaderRowRange.Item(1).value, CStr(sSheetname))
        .Range(.Cells(C_StartLineTitle2, iCol), .Cells(C_StartLineTitle2, iCol + 3)).Merge
    
        'ajout des formules de validation
        .Cells(C_TitleLine + 1, iCol).Validation.Delete

        .Cells(C_TitleLine + 1, iCol).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, _
                                                     Formula1:="=GEO!" & xlsapp.Sheets("GEO").Range("T_ADM1").Columns(1).Address
     
        .Cells(C_TitleLine + 1, iCol).Validation.IgnoreBlank = True
        .Cells(C_TitleLine + 1, iCol).Validation.InCellDropdown = True
        .Cells(C_TitleLine + 1, iCol).Validation.InputTitle = ""
        .Cells(C_TitleLine + 1, iCol).Validation.errorTitle = ""
        .Cells(C_TitleLine + 1, iCol).Validation.InputMessage = ""
        .Cells(C_TitleLine + 1, iCol).Validation.ErrorMessage = sMessage
        .Cells(C_TitleLine + 1, iCol).Validation.ShowInput = True
        .Cells(C_TitleLine + 1, iCol).Validation.ShowError = True
    End With

End Sub

Private Sub TransfertSheet(xlsapp As Object, sSheetname As String)
    
    'Since We can't move worksheet from one instance to another
    'we need to save as a temporary file and then move it to another app
    ThisWorkbook.Sheets(sSheetname).Copy
    DoEvents
    
    ActiveWorkbook.SaveAs "C:\LineListeApp\tampon.xlsx"
    ActiveWorkbook.Close
    DoEvents

    With xlsapp
        .Workbooks.Open Filename:="C:\LineListeApp\tampon.xlsx", UpdateLinks:=False
        
        .Sheets(sSheetname).Select
        .Sheets(sSheetname).Copy after:=.Workbooks(1).Sheets(1)
        
        'Sheets PASSWORD and CONTROLEFORMULE have some hidden properties to update
        Select Case UCase(sSheetname)
        Case "PASSWORD"
            .Sheets(sSheetname).Visible = xlSheetVeryHidden
        Case "CONTROLEFORMULE"
            .Sheets(sSheetname).Visible = False
        End Select
        
        DoEvents
        .Workbooks("tampon.xlsx").Close
    End With
    DoEvents
    
    Kill "C:\LineListeApp\tampon.xlsx"

End Sub

Private Function LetColNumberByDataName(xlsapp As Excel.Application, sDataName As String, sSheetname As String) As Integer

    Dim i As Integer

    'T_dataDic(D_TitleDic("Choices") - 1, i)
    With xlsapp
        i = 1
        While i <= .Sheets(sSheetname).Cells(C_TitleLine, 1).End(xlToRight).Column And UCase(.Sheets(sSheetname).Cells(C_TitleLine, i).Name.Name) <> sDataName
            i = i + 1
        Wend
        If UCase(.Sheets(sSheetname).Cells(C_TitleLine, i).Name.Name) = sDataName Then
            LetColNumberByDataName = i
        End If
    End With

End Function

Private Function LetWordingWithSpace(xlsapp As Excel.Application, sDataWording As String, sSheetname As String)
    'The goal of this function is to add space to duplicates labels so that excels does not force a unique name with number at the end
    Dim i As Integer

    LetWordingWithSpace = ""
    With xlsapp
        i = 1
        While i <= .Sheets(sSheetname).Cells(C_TitleLine, 1).End(xlToRight).Column And Replace(UCase(.Sheets(sSheetname).Cells(C_TitleLine, i).value), " ", "") <> Replace(UCase(sDataWording), " ", "")
            i = i + 1
        Wend
        
        If Replace(UCase(xlsapp.Sheets(sSheetname).Cells(C_TitleLine, i).value), " ", "") = Replace(UCase(sDataWording), " ", "") Then
            LetWordingWithSpace = xlsapp.Sheets(sSheetname).Cells(C_TitleLine, i).value & " "
        Else
            LetWordingWithSpace = sDataWording
        End If
    End With

End Function


