Attribute VB_Name = "DesignerBuildListHelpers"
Option Explicit
Option Private Module
'-------
'Transfert Codes and forms to the designer

Public Sub TransferDesignerCodes(wkb As Workbook)


    'Transfert form is for sending forms from the actual excel workbook to another
    TransferForm wkb, C_sFormGeo
    TransferForm wkb, C_sFormShowHide
    TransferForm wkb, C_sFormExport
    TransferForm wkb, C_sFormExportMig
    TransferForm wkb, C_sFormAdvanced
    TransferForm wkb, C_sFormImportRep

    'TransferCode is for sending modules  (Modules) or classes (Classes) from actual excel workbook to another excel workbook
    TransferCode wkb, C_sModLinelist, "Module"
    TransferCode wkb, C_sModLLGeo, "Module"
    TransferCode wkb, C_sModLLShowHide, "Module"
    TransferCode wkb, C_sModHelpers, "Module"
    TransferCode wkb, C_sModLLMigration, "Module"
    TransferCode wkb, C_sModLLConstants, "Module"
    TransferCode wkb, C_sModEsthConstants, "Module"
    TransferCode wkb, C_sModLLExport, "Module"
    TransferCode wkb, C_sModLLTrans, "Module"
    TransferCode wkb, C_sModLLDict, "Module"
    TransferCode wkb, C_sModLLCustFunc, "Module"

    TransferCode wkb, C_sClaBA, "Class"
    'Formula classes?

End Sub

'Transfert code from one module to a worksheet to trigger some events
'@sSheetName the sheet name we want to transfer to
'@sNameModule the name of the module we want to copy code from

Public Sub TransferCodeWksh(wkb As Workbook, sSheetName As String, _
                           sNameModule As String)

    Dim sNouvCode As String                      'a string to contain code to add
    Dim sheetComp As String
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object

    'save the code module in the string sNouvCode
    With DesignerWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With

    With wkb
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.sheets(sSheetName).CodeName)
        Set codeMod = vbComp.CodeModule
    End With

    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With
End Sub

'-----------------
'Transfert a Worksheet from the current designer to another
'Excel workbook. (in one application)

'@Wkb: a workbook
'@sSheetName: the name of the Sheet in the designer we want to move

Public Sub TransferSheet(wkb As Workbook, sSheetName As String, sPrevSheetName As String)
    DesignerWorkbook.Worksheets(sSheetName).Copy after:=wkb.Worksheets(sPrevSheetName)
End Sub

'-----
'Transfert a form the actual Designer to the linelist's Workbook
'@Wkb : A workbook
'@sFormName: The name of the form to transfert

Private Sub TransferForm(wkb As Workbook, sFormName As String)

    'The form is sent to the LinelisteApp folder
    On Error Resume Next
    Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frm"
    Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frx"
    On Error GoTo 0

    DoEvents
    DesignerWorkbook.VBProject.VBComponents(sFormName).Export SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frm"
    wkb.VBProject.VBComponents.Import SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frm"
    DoEvents

    On Error Resume Next
    Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frm"
    Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "CopieUsf.frx"
    On Error GoTo 0
End Sub

'---
'The Goal is to transfer one Code Module/Class from the designer to the
'linelist sheet
'@wkb a workbook
'@sType the type of the code to transfer (Module or Class)
'@sModule: The Name of the module to transfer

Private Sub TransferCode(wkb As Workbook, sNameModule As String, sType As String)

    Dim oNouvM As Object                         'New module name
    Dim sNouvCode As String                      'New module code

    'get all the values within the actual module to transfer
    With DesignerWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With

    'create to code or module if needed
    Select Case sType
    Case "Module"
        Set oNouvM = wkb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "Class"
        Set oNouvM = wkb.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select

    'keep the name and add the codes
    oNouvM.Name = sNameModule
    With wkb.VBProject.VBComponents(oNouvM.Name).CodeModule
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With

End Sub

'----
'When you have same headers name in a listobject table, excels add a number (1)
'for example to the following header.

'The goal of this function is to add space to duplicates labels so that excels
'does not force a unique name with number at the end in listcolumn header

'@Wkb a Workbook
'sHeader the String we want to add space (in case) to
'sSheetName: The concernec SheetName
'iStartLine: Integer, the line where the table listobject starts

Public Function AddSpaceToHeaders(wkb As Workbook, _
                                  sHeader As String, _
                                  sSheetName As String, iStartLine As Integer)
    Dim i As Integer

    AddSpaceToHeaders = ""
    With wkb
        i = 1
        Do While i <= .Worksheets(sSheetName).Cells(iStartLine, Columns.Count).End(xlToLeft).Column And Replace(UCase(.sheets(sSheetName).Cells(iStartLine, i).Value), " ", "") <> Replace(UCase(sHeader), " ", "")
            i = i + 1
        Loop
        If Replace(UCase(wkb.Worksheets(sSheetName).Cells(iStartLine, i).Value), " ", "") = Replace(UCase(sHeader), " ", "") Then
            AddSpaceToHeaders = wkb.Worksheets(sSheetName).Cells(iStartLine, i).Value & " "
        Else
            AddSpaceToHeaders = sHeader
        End If
    End With

End Function

'----


'Add a Button command to a Sheet (create the button and addit)
'@Wkb: a Workbook
'@sSheet: The Sheet we want to add the button
'@sShpName: The name we want to give to the shape (Shape Name)
'@sText: The text to put on the button
'@iCmdWidth: The command with
'@iCmdHeight: The command height
'@sCommand: The binding command on the Shape
'@sShpColor: The color of the Shape
'@sShpTextColor: color of the text for each of the shapes

Sub AddCmd(wkb As Workbook, sSheetName As String, iLeft As Integer, iTop As Integer, _
           sShpName As String, sText As String, iCmdWidth As Integer, iCmdHeight As Integer, _
           sCommand As String, Optional sShpColor As String = "MainSecBlue", _
           Optional sShpTextColor As String = "White", Optional iTextFontSize As Integer = 10)


    sText = LineListTranslatedValue(sShpName, C_sTabTradLLShapes)

    With wkb.Worksheets(sSheetName)
        .Shapes.AddShape(msoShapeRectangle, iLeft + 3, iTop + 3, iCmdWidth, iCmdHeight).Name = sShpName
        .Shapes(sShpName).Placement = xlFreeFloating
        .Shapes(sShpName).TextFrame2.TextRange.Characters.Text = sText
        .Shapes(sShpName).TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Shapes(sShpName).TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shapes(sShpName).TextFrame2.WordWrap = msoTrue
        .Shapes(sShpName).TextFrame2.TextRange.Font.Size = iTextFontSize
        .Shapes(sShpName).TextFrame2.TextRange.Font.Bold = False
        .Shapes(sShpName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Helpers.GetColor(sShpTextColor)
        .Shapes(sShpName).Fill.ForeColor.RGB = Helpers.GetColor(sShpColor)
        .Shapes(sShpName).Fill.BackColor.RGB = Helpers.GetColor(sShpColor)
        .Shapes(sShpName).Shadow.Type = msoShadow24
        .Shapes(sShpName).OnAction = sCommand
    End With

End Sub

'Little Subs used when working with the Creation of the data Entry for a sheet of type Linelist
'Add the Sub Label
Sub AddSubLab(Wksh As Worksheet, iSheetStartLine As Integer, _
              iCol As Integer, sMainLab As String, sSubLab As String, _
              Optional sSubLabColor As String = "SubLabBlue")
    With Wksh
        .Cells(iSheetStartLine, iCol).Value = _
        .Cells(iSheetStartLine, iCol).Value & Chr(10) & sSubLab

        'Changing the fontsize of the sublabels
        .Cells(iSheetStartLine, iCol).Characters(Start:=Len(sMainLab) + 1, _
            Length:=Len(sSubLab) + 1).Font.Size = C_iLLSheetFontSize - 2
        .Cells(iSheetStartLine, iCol).Characters(Start:=Len(sMainLab) + 1, _
            Length:=Len(sSubLab) + 1).Font.Color = Helpers.GetColor(sSubLabColor)
    End With

End Sub

'Add the notes
Sub AddNotes(Wksh As Worksheet, iSheetStartLine As Integer, _
             iCol As Integer, sNote As String, _
             Optional bNoteVisibility As Boolean = False)
    With Wksh

        .Cells(iSheetStartLine, iCol).AddComment
        .Cells(iSheetStartLine, iCol).Comment.Text Text:=sNote
        .Cells(iSheetStartLine, iCol).Comment.Visible = bNoteVisibility

    End With

End Sub

'Add the status to notes

Sub AddStatus(Wksh As Worksheet, iSheetStartLine As Integer, _
              iCol As Integer, sNote As String, sStatus As String, _
              Optional sMandatory As String = "Mandatory data", _
              Optional bNoteVisibility As Boolean = False)
    With Wksh
        Select Case sStatus
        Case C_sDictStatusMan
            If sNote <> "" Then
                'Update the notes to add the Status
                .Cells(iSheetStartLine, iCol).Comment.Text Text:=sMandatory & Chr(10) & sNote
            Else
                'or  Add comment on status
                Call AddNotes(Wksh, _
                              iSheetStartLine, _
                              iCol, sMandatory)
            End If
        Case C_sDictStatusHid
            'Hidden, hid the actual column
            .Columns(iCol).EntireColumn.Hidden = True
        Case C_sDictStatusOpt
            'Do nothing for the moment for optional status
        End Select
    End With
End Sub

'Add the type
Sub AddType(Wksh As Worksheet, iSheetStartLine As Integer, _
            iCol As Integer, sType As String)

    Dim iDecType As Integer                      'Just to get the decimal number at the end of decimal
    'Dim iDecNb As Integer
    Dim i As Integer
    Dim sNbDeci As String                        'Number of decimals


    'Check to be sure that the actual type contains decimal
    With Wksh
        If InStr(1, sType, C_sDictTypeDec) > 0 Then
            iDecType = CInt(Replace(sType, C_sDictTypeDec, ""))
            sType = C_sDictTypeDec
            i = 0
            sNbDeci = ""
            Do While i < iDecType
                sNbDeci = "0" & sNbDeci
                i = i + 1
            Loop
        End If

        Select Case sType
            'Text Type
        Case C_sDictTypeText
            .Cells(iSheetStartLine + 2, iCol).NumberFormat = "@"
            'Integer
        Case C_sDictTypeInt
            .Cells(iSheetStartLine + 2, iCol).NumberFormat = "0"
            'Date Type
        Case C_sDictTypeDate
            .Cells(iSheetStartLine + 2, iCol).NumberFormat = "d-mmm-yyy"
            'Decimal
        Case C_sDictTypeDec
            .Cells(iSheetStartLine + 2, iCol).NumberFormat = "0." & sNbDeci
        Case Else
            'If I don't know the type, put in text
            .Cells(iSheetStartLine + 2, iCol).NumberFormat = "@"
        End Select
    End With
End Sub

'Add the choices
Sub AddChoices(wkb As Workbook, sSheetName As String, iSheetStartLine As Integer, iCol As Integer, _
               ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
               sChoice As String, sAlert As String, sMessage As String)

    Dim ValidationList As BetterArray
    Dim iChoiceCol As Long
    Dim iChoiceRow As Long
    Dim sListObjectName As String
    Dim WkshChoice As Worksheet
    Dim WkshLL As Worksheet
    Dim LoRng As Range

    'Set the differents worksheets
    Set WkshChoice = wkb.Worksheets(C_sSheetChoiceAuto)
    Set WkshLL = wkb.Worksheets(sSheetName)

    sListObjectName = "lo" & "_" & sChoice

    'Add the list object if it does not exists
    If Not ListObjectExists(WkshChoice, sListObjectName) Then
        With WkshChoice
            iChoiceCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 2
            .Cells(C_eStartlinesListAuto, iChoiceCol).Value = sChoice

            'Get the validation list
            Set ValidationList = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sChoice)

            'Write it to the worksheet and free memory
            ValidationList.ToExcelRange .Cells(C_eStartlinesListAuto + 1, iChoiceCol)

            'Add the list object to the worksheet
            iChoiceRow = .Cells(.Rows.Count, iChoiceCol).End(xlUp).Row

            Set LoRng = .Range(.Cells(C_eStartlinesListAuto, iChoiceCol), .Cells(iChoiceRow, iChoiceCol))

            'Add the list object here
            .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = sListObjectName
        End With

        'Now Add dynamic name for the choice
        wkb.Names.Add Name:=sChoice, RefersToR1C1:="=" & sListObjectName & "[" & sChoice & "]"
    End If

    'Add the validation
    With WkshLL
        Call Helpers.SetValidation(.Cells(iSheetStartLine, iCol), _
                                   "=" & sChoice, Helpers.GetValidationType(sAlert), _
                                   sMessage)
    End With

End Sub

'Add The GoTo Section for a worksheet

'Add Geo
Sub AddGeo(wkb As Workbook, DictData As BetterArray, DictHeaders As BetterArray, sSheetName As String, iSheetStartLine As Integer, iCol As Integer, _
           iSheetSubSecStartLine As Integer, iDictLine As Integer, sVarName As String, sMessage As String, iNbshifted As Integer)

    With wkb.Worksheets(sSheetName)
        .Cells(iSheetStartLine, iCol).Interior.Color = GetColor("Orange")
        'update the columns only for the geo
        Call Add4GeoCol(wkb, DictData, DictHeaders, sSheetName, sVarName, iSheetStartLine, _
                        iCol, sMessage, _
                        iSheetSubSecStartLine, iDictLine, iNbshifted)

    End With
End Sub

'For Columns of Type Geo, we need to insert the 4 admin levels So when working with these type of column,
'This functions's purpose is to add the three other remaining columns of the geo.

'The purpose of this procedure is to create the geo columns using the geo data  (its also adds the first dropdowns)
' we shift the columns to the right until we reached the number of columns required

'@Wksh: Excel Worksheet where to add the Geo columns
'@sVarName: The Name of the cell of type geo, which is the variable name
'@iCol: the column where we want to insert the geo
'@sMessage: the validation message in case the user enters different information
'@iStartLine: Starting line of Data in the Linelist
'@iStartLineSubLab: Starting line of the Sub label

Sub Add4GeoCol(wkb As Workbook, DictData As BetterArray, DictHeaders As BetterArray, _
               sSheetName As String, sVarName As String, iStartLine As Integer, iCol As Integer, _
               sMessage As String, iStartLineSubLab As Integer, iDictLine As Integer, iNbshifted As Integer)


    Dim sLab As String                           'Temporary variable, label of the Admin level
    Dim LineValues As BetterArray
    Dim iRow As Integer

    Set LineValues = New BetterArray
    LineValues.LowerBound = 1

    iRow = iDictLine + iNbshifted

    With wkb.Worksheets(sSheetName)

        'Admin 4
        sLab = SheetGeo.ListObjects(C_sTabAdm4).HeaderRowRange.Item(4).Value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).Value = AddSpaceToHeaders(wkb, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "4" & "_" & sVarName

        'Add the type
        .Cells(C_eStartLinesLLMainSec - 1, iCol + 1).Value = C_sDictControlGeo & "4"
        'Put in bold
        .Range(.Cells(iStartLine, iCol + 1), .Cells(iStartLine + 1, iCol + 1)).Font.Bold = True
        .Cells(iStartLine + 2, iCol + 1).Locked = False

        Call Helpers.SetValidation(.Cells(iStartLine + 2, iCol + 1), "=" & C_sAdmName & "_4_" & "dropdown", 2, sMessage)

        'Admin 3
        sLab = SheetGeo.ListObjects(C_sTabAdm3).HeaderRowRange.Item(3).Value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).Value = AddSpaceToHeaders(wkb, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "3" & "_" & sVarName
        .Cells(iStartLine + 1, iCol + 1).Value = C_sAdmName & "3" & "_" & sVarName

        Call Helpers.WriteBorderLines(.Range(.Cells(iStartLine, iCol + 1), .Cells(iStartLine + 1, iCol + 1)))


        .Cells(C_eStartLinesLLMainSec - 1, iCol + 1).Value = C_sDictControlGeo & "3"
        .Cells(iStartLine + 2, iCol + 1).Locked = False
        'Set validation for admin 3
        Call Helpers.SetValidation(.Cells(iStartLine + 2, iCol + 1), "=" & C_sAdmName & "_3_" & "dropdown", 2, sMessage)


        'Admin 2
        sLab = SheetGeo.ListObjects(C_sTabAdm2).HeaderRowRange.Item(2).Value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).Value = AddSpaceToHeaders(wkb, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "2" & "_" & sVarName
        .Cells(iStartLine + 1, iCol + 1).Value = C_sAdmName & "2" & "_" & sVarName

        Call Helpers.WriteBorderLines(.Range(.Cells(iStartLine, iCol + 1), .Cells(iStartLine + 1, iCol + 1)))
        .Range(.Cells(iStartLine, iCol + 1), .Cells(iStartLine + 1, iCol + 1)).Font.Bold = True
        .Cells(C_eStartLinesLLMainSec - 1, iCol + 1).Value = C_sDictControlGeo & "2"
        .Cells(iStartLine + 2, iCol + 1).Locked = False

        'Set validation for admin 2
        Call Helpers.SetValidation(.Cells(iStartLine + 2, iCol + 1), "=" & C_sAdmName & "_2_" & "dropdown", 2, sMessage)

        'Admin 1
        sLab = SheetGeo.ListObjects(C_sTabadm1).HeaderRowRange.Item(1).Value
        .Cells(iStartLine, iCol).Value = AddSpaceToHeaders(wkb, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol).Name = C_sAdmName & "1" & "_" & sVarName
        .Cells(iStartLine, iCol).Interior.Color = GetColor("Orange")
        .Cells(iStartLine + 1, iCol).Value = C_sAdmName & "1" & "_" & sVarName


        Call Helpers.WriteBorderLines(.Range(.Cells(iStartLine, iCol), .Cells(iStartLine + 1, iCol)))
        .Cells(iStartLine + 2, iCol).Locked = False

        'ajout des formules de validation
        .Cells(iStartLine + 2, iCol).Validation.Delete
        'Add name and reference for adm1 (in case someone adds one adm1)
        wkb.Names.Add Name:=C_sAdmName & "1" & "_column", RefersToR1C1:="=" & C_sTabadm1 & "[" & SheetGeo.Cells(1, 1).Value & "]"

        Call Helpers.SetValidation(.Cells(iStartLine + 2, iCol), "=" & C_sAdmName & 1 & "_column", 2, sMessage)
        Call Helpers.WriteBorderLines(.Range(.Cells(iStartLine, iCol), .Cells(iStartLine + 1, iCol)))
    End With

    'Updating the Dictionary for future uses
    With wkb.Worksheets(C_sParamSheetDict)
        'Admin 4
        LineValues.Items = DictData.ExtractSegment(RowIndex:=iDictLine)
        LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "4"
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).Value = ""
        .Cells(iRow + 2, DictHeaders.Length + 1).Value = .Cells(iRow + 1, DictHeaders.Length + 1).Value + 3
        'Admin 3
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "3"
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).Value = ""
        .Cells(iRow + 2, DictHeaders.Length + 1).Value = .Cells(iRow + 1, DictHeaders.Length + 1).Value + 2
        'Admin 2
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "2"
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).Value = ""
        .Cells(iRow + 2, DictHeaders.Length + 1).Value = .Cells(iRow + 1, DictHeaders.Length + 1).Value + 1

    End With
End Sub

Sub BuildGotoArea(wkb As Workbook, sTableName As String, sSheetName As String, iGoToCol As Long, _
                  Optional iCol As Long = C_eSectionsLookupColumns, _
                  Optional iFontSize As Long = 10)

    Dim iChoiceRow As Long
    Dim LoRng As Range
    Dim sListObjectName As String
    Dim sGoToSourceName As String

    With wkb.Worksheets(sSheetName)
        'Where to write the GoTo section in the worksheet
        .Cells(1, iCol).Locked = False
        .Cells(1, iCol).Value = TranslateLLMsg("MSG_SelectSection")
        .Cells(1, iCol).Name = sTableName & "_" & C_sGotoSection
        .Cells(1, iCol).Font.Size = iFontSize
        .Cells(1, iCol).HorizontalAlignment = xlHAlignCenter
        .Cells(1, iCol).Interior.Color = Helpers.GetColor("MainSecBlue")
        .Cells(1, iCol).Font.Color = vbWhite
        .Cells(1, iCol).Font.Bold = True
        .Cells(1, iCol).VerticalAlignment = xlVAlignCenter
        .Cells(1, iCol).FormulaHidden = True
        .Cells(1, iCol).WrapText = True
    End With

    sListObjectName = "lo" & "_" & sTableName & "_" & C_sGotoSection
    sGoToSourceName = sTableName & "_" & C_sGotoSection & "_" & "source"

    With wkb.Worksheets(C_sSheetChoiceAuto)

        .Cells(C_eStartlinesListAuto, iGoToCol).Value = sGoToSourceName
        'Add the list object to the worksheet
        iChoiceRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).Row
        Set LoRng = .Range(.Cells(C_eStartlinesListAuto, iGoToCol), .Cells(iChoiceRow, iGoToCol))

        'Add the list object here
        .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = sListObjectName
    End With

    'Add the validation
    wkb.Names.Add Name:=sGoToSourceName, RefersToR1C1:="=" & sListObjectName & "[" & sGoToSourceName & "]"

    'Add the validation
    With wkb.Worksheets(sSheetName)
        Call Helpers.SetValidation(.Cells(1, iCol), "=" & sGoToSourceName, 1, TranslateLLMsg("MSG_SectionNotExist"))
    End With
End Sub

'Merging Areas on Sheets of type LL and Adms ___________________________________________________________________________________________________________________________________________________________

Sub MergeArea(Wksh As Worksheet, iLineFrom As Integer, iLineTo As Integer, iColFrom As Integer, iColTo As Integer)
    With Wksh
        .Range(.Cells(iLineFrom, iColFrom), .Cells(iLineTo, iColTo)).Merge
        .Cells(iLineFrom, iColFrom).MergeArea.HorizontalAlignment = xlCenter
        .Cells(iLineFrom, iColFrom).MergeArea.VerticalAlignment = xlCenter
    End With
End Sub

'Main Section Vertical Merge
Sub BuildMainSectionVMerge(Wksh As Worksheet, iLineFrom As Integer, iLineTo As Integer, _
                           iColumnFrom As Integer)

    Dim oCell As Object
    Dim iColumnTo As Integer

    iColumnTo = iColumnFrom + 1

    'Merge
    MergeArea Wksh, iLineFrom:=iLineFrom, iLineTo:=iLineTo - 1, iColFrom:=iColumnFrom, iColTo:=iColumnFrom

    With Wksh
        With .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineTo - 1, iColumnFrom))
            .Interior.Color = Helpers.GetColor("MainSecBlue")
            .Font.Color = Helpers.GetColor("White")
            .Font.Bold = True
            .Font.Size = 12
        End With

        For Each oCell In .Range(.Cells(iLineFrom, iColumnTo), .Cells(iLineTo - 1, iColumnTo))
            If oCell.Value = vbNullString Then oCell.Interior.Color = vbWhite
        Next

        'Write the borders line
        WriteBorderLines .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineTo - 1, iColumnTo)), iWeight:=xlHairline, sColor:="DarkBlue"
        WriteBorderLines .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineTo - 1, iColumnFrom)), iWeight:=xlHairline, sColor:="DarkBlue"
        'Entire main section
        WriteBorderLines .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineTo - 1, iColumnFrom + 3)), iWeight:=xlThin, sColor:="DarkBlue"
    End With

End Sub

'Sub Section Vertical Merge
Sub BuildSubSectionVMerge(Wksh As Worksheet, iColumn As Integer, iLineFrom As Integer, _
                          iLineTo As Integer)

    Dim iLastLine As Integer
    'Last Column can be 1, in that case move to the first column
    iLastLine = IIf(iLineTo <= 1, 1, iLineTo - 1)

    'Merge Area for Worksheet
    MergeArea Wksh, iLineFrom:=iLineFrom, iLineTo:=iLastLine, iColFrom:=iColumn, iColTo:=iColumn

    With Wksh
        With .Range(.Cells(iLineFrom, iColumn), .Cells(iLastLine, iColumn))
            .Interior.Color = Helpers.GetColor("SubSecBlue")
            .Font.Color = Helpers.GetColor("BlueButton")
            .Font.Bold = True
            .Font.Size = 10
        End With

        'Draw borders
        WriteBorderLines .Range(.Cells(iLineFrom, iColumn), .Cells(iLastLine, iColumn)), iWeight:=xlHairline, sColor:="DarkBlue"
        WriteBorderLines .Range(.Cells(iLineFrom, iColumn), .Cells(iLastLine, iColumn + 2)), iWeight:=xlHairline, sColor:="DarkBlue"
    End With
End Sub

' Build sections horizontal merge area
Sub BuildMainSectionHMerge(Wksh As Worksheet, iLineFrom As Integer, iLineTo As Integer, _
                           iColumnFrom As Integer, iColumnTo As Integer, _
                           Optional sColorMainSec As String = "MainSecBlue")

    Dim oCell As Object

    'Merge
    MergeArea Wksh, iLineFrom:=iLineFrom, iLineTo:=iLineFrom, iColFrom:=iColumnFrom, iColTo:=iColumnTo - 1

    With Wksh
        With .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineFrom, iColumnTo - 1))
            .Interior.Color = Helpers.GetColor(sColorMainSec)
            .Font.Color = Helpers.GetColor("White")
            .Font.Bold = True
            .Font.Size = C_iLLMainSecFontSize
        End With

        For Each oCell In .Range(.Cells(iLineTo, iColumnFrom), .Cells(iLineTo, iColumnTo - 1))
            If oCell.Value = vbNullString Then oCell.Interior.Color = vbWhite
        Next

        'Write the borders line
        WriteBorderLines .Range(.Cells(iLineFrom, iColumnFrom), .Cells(iLineTo, iColumnTo - 1))
    End With

End Sub

'Build subsections horizontal merge area

Sub BuildSubSectionHMerge(Wksh As Worksheet, iLine As Integer, iColumnFrom As Integer, _
                          iColumnTo As Integer, Optional sColorSubSec As String = "SubSecBlue", _
                          Optional sColorMainSec As String = "MainSecBlue")

    Dim iLastCol As Integer
    'Last Column can be 1, in that case move to the first column
    iLastCol = IIf(iColumnTo <= 1, 1, iColumnTo - 1)

    'Merge Area for Worksheet
    MergeArea Wksh, iLineFrom:=iLine, iLineTo:=iLine, iColFrom:=iColumnFrom, iColTo:=iLastCol

    With Wksh
        With .Range(.Cells(iLine, iColumnFrom), .Cells(iLine, iLastCol))
            .Interior.Color = Helpers.GetColor(sColorSubSec)
            .Font.Color = Helpers.GetColor(sColorMainSec)
            .Font.Size = C_iLLSubSecFontSize
        End With

        'Draw borders
        WriteBorderLines .Range(.Cells(iLine, iColumnFrom), .Cells(iLine, iLastCol))
    End With
End Sub

'Get the Validation Formulas
'
'Using the formulas Entered by the user, guess the acual formula to use in. This one is almost the same idea as Lionel's one:

'- split the formulas on special characters and put all those things that are not special in a table
'- Pay attention to string in formulas by checking the character " (Chr(34)) in the Abstract Syntax Tree
'- Check to see if non special strings are either in the allowed formulas, or in the varname data
'- replace varnames with the column index.
'- If the validation formula does not succeed, return empty character.

'@sFormula the formula String
'@VarNameData: the data with all the varnames on ONE sheet
'@ColumnIndexData: The data of all the column indexes on ONE Sheet
'@FormulaData: The accepted list of formulas in English
'@SpecCharData: The data with all the special characters


Public Function ValidationFormula(sFormula As String, AllSheetNamesData As BetterArray, VarNameData As BetterArray, _
                                  ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                  SpecCharData As BetterArray, ChoiceAutoVarData As BetterArray, _
                                  sActualVarName As String, Wksh As Worksheet, Optional bLocal As Boolean = True) As String
    'Returns a string of cleared formula

    ValidationFormula = vbNullString
    If sFormula = vbNullString Then Exit Function

    Dim sFormulaATest As String                  'same formula, with all the spaces replaced with
    Dim sAlphaValue As String                    'Alpha numeric values in a formula
    Dim sLetter As String                        'counter for every letter in one formula
    Dim scolAddress As String                    'address of one column used in a formula

    Dim FormulaAlphaData As BetterArray          'Table of alphanumeric data in one formula

    Dim i As Integer
    Dim iPrevBreak As Integer
    Dim iNbParentO As Integer                    'Number of left parenthesis
    Dim iNbParentF As Integer                    'Number of right parenthesis
    Dim icolNumb As Integer                      'Column number on one sheet of one column used in a formual
    Dim sSheetName As String


    Dim isError As Boolean
    Dim OpenedQuotes As Boolean                  'Test if the formula has opened some quotes
    Dim QuotedCharacter As Boolean
    Dim NoErrorAndNoEnd As Boolean
    Set FormulaAlphaData = New BetterArray       'Alphanumeric values of one formula

    FormulaAlphaData.LowerBound = 1

    'squish the formula (removing multiple spaces) to avoid problems related to
    'space collapsing and upper/lower cases


    sFormulaATest = "(" & ClearNonPrintableUnicode(sFormula) & ")"

    iNbParentO = 0                               'Number of open brakets
    iNbParentF = 0                               'Number of closed brackets
    iPrevBreak = 1
    OpenedQuotes = False
    NoErrorAndNoEnd = True
    QuotedCharacter = False
    i = 1

    If VarNameData.Includes(sFormulaATest) Then
        ValidationFormula = sFormulaATest
        Exit Function
    Else
        Do While (i <= Len(sFormulaATest))
            QuotedCharacter = False

            sLetter = Mid(sFormulaATest, i, 1)
            If sLetter = Chr(34) Then
                OpenedQuotes = Not OpenedQuotes
            End If

            If Not OpenedQuotes And SpecCharData.Includes(sLetter) Then 'A special character, not in quotes
                If sLetter = Chr(40) Then
                    iNbParentO = iNbParentO + 1
                End If
                If sLetter = Chr(41) Then
                    iNbParentF = iNbParentF + 1
                End If

                sAlphaValue = Application.WorksheetFunction.Trim(Mid(sFormulaATest, iPrevBreak, i - iPrevBreak))
                If sAlphaValue <> "" Then
                    'It is either a formula or a variable name or a quoted string
                    If Not VarNameData.Includes(LCase(sAlphaValue)) And Not FormulaData.Includes(UCase(sAlphaValue)) And Not IsNumeric(sAlphaValue) Then
                        'Testing if not opened the quotes
                        If Mid(sAlphaValue, 1, 1) <> Chr(34) Then
                            isError = True
                            Exit Do
                        Else
                            QuotedCharacter = True
                        End If
                    End If

                    If Not isError And Not QuotedCharacter Then
                        'It is either a variable name or a formula
                        If VarNameData.Includes(sAlphaValue) Then 'It is a variable name, I will track its column

                            'update the choice auto data
                            If ChoiceAutoVarData.Includes(sActualVarName) Then ChoiceAutoVarData.Push sAlphaValue

                            icolNumb = ColumnIndexData.Item(VarNameData.IndexOf(sAlphaValue))
                            sSheetName = EnsureGoodSheetName(AllSheetNamesData.Item(VarNameData.IndexOf(sAlphaValue)))
                            sAlphaValue = "'" & sSheetName & "'!" & Cells(C_eStartLinesLLData + 2, icolNumb).Address(False, True)
                        ElseIf FormulaData.Includes(UCase(sAlphaValue)) Then 'It is a formula, excel will do the translation for us
                            sAlphaValue = Application.WorksheetFunction.Trim(sAlphaValue)
                        End If
                    End If
                    FormulaAlphaData.Push sAlphaValue, sLetter
                Else
                    'I have a special character, at the value sLetter But nothing between this special character and previous one, just add it
                    FormulaAlphaData.Push sLetter
                End If

                iPrevBreak = i + 1
            End If
            i = i + 1
        Loop
    End If

    If iNbParentO <> iNbParentF Then
        isError = True
    End If

    If Not isError Then
        sAlphaValue = FormulaAlphaData.ToString(Separator:="", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
        'If local, get the local formula
        If (bLocal) And Not isMac() Then
            ValidationFormula = Helpers.GetInternationalFormula(sAlphaValue, Wksh)
        Else
            ValidationFormula = "=" & sAlphaValue
        End If
    End If


End Function

'Setting the min and the max validation
Sub BuildValidationMinMax(oRange As Range, iMin As String, iMax As String, iAlertType As Byte, sTypeValidation As String, sMessage As String)

    On Error Resume Next
    With oRange.Validation
        .Delete
        Select Case LCase(sTypeValidation)
        Case "integer"                           'if the validation should be for integer
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case "date"                              'Date
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case Else                                'Decimals
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
        .errorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

'Update the headers of the choice auto when everything is done.
'For each sheet, write a list_auto_origin on top of the mainsection if there is a list
'auto in the sheet, so that we can use them later on for updating things.
Public Sub UpdateChoiceAutoHeaders(wkb As Workbook, ChoiceAutoVarData As BetterArray, DictHeaders As BetterArray)

    Dim i As Integer
    Dim sVarName As String
    Dim sSheetName As String
    Dim iIndex As Integer
    i = 2
    With wkb
        sVarName = .Worksheets(C_sParamSheetDict).Cells(2, DictHeaders.IndexOf(C_sDictHeaderVarName)).Value
        Do While (sVarName <> vbNullString)
            If ChoiceAutoVarData.Includes(sVarName) Then
                sSheetName = .Worksheets(C_sParamSheetDict).Cells(i, DictHeaders.IndexOf(C_sDictHeaderSheetName)).Value
                iIndex = .Worksheets(C_sParamSheetDict).Cells(i, DictHeaders.Length + 1).Value
                .Worksheets(sSheetName).Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value)
                .Worksheets(sSheetName).Cells(C_eStartLinesLLMainSec - 2, iIndex).Value = C_sDictControlChoiceAuto & "_origin"
                .Worksheets(sSheetName).Cells(C_eStartLinesLLMainSec - 2, iIndex).Font.Color = vbWhite
                .Worksheets(sSheetName).Cells(C_eStartLinesLLMainSec - 2, iIndex).FormulaHidden = True
                .Worksheets(sSheetName).Protect Password:=(ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
            End If
            i = i + 1
            sVarName = .Worksheets(C_sParamSheetDict).Cells(i, DictHeaders.IndexOf(C_sDictHeaderVarName)).Value
        Loop
    End With
End Sub

'Add the metadata Sheet
Public Sub AddMetadataSheet(wkb As Workbook)
    Dim iRow As Integer

    iRow = 0

    With wkb
        'Metadata sheet
        .Worksheets.Add.Name = C_sSheetMetadata

        With SheetGeo.ListObjects(C_sTabGeoMetadata)
            If Not .DataBodyRange Is Nothing Then iRow = .DataBodyRange.Rows.Count
        End With

        'Add metadata of the Geo
        With .Worksheets(C_sSheetMetadata)
            .Cells(1, 1).Value = C_sVariable
            .Cells(1, 2).Value = C_sValue
            If iRow > 0 Then
                .Range(.Cells(2, 1), .Cells(2 + iRow, 2)).Value = SheetGeo.ListObjects(C_sTabGeoMetadata).DataBodyRange.Value
            Else
                iRow = 1
            End If
            'Add other informations to the metadata sheet:

            'language
            .Cells(iRow + 1, 1).Value = C_sLanguage
            .Cells(iRow + 1, 2).Value = SheetLLTranslation.Range(C_sRngLLLanguage).Value

            'linelist creation date
            .Cells(iRow + 2, 1).Value = C_sLLDate
            .Cells(iRow + 2, 2).Value = Format(Now, "yyyy/mm/dd Hh:Nn")

            'linelist version... Other infos will be added

            .Visible = xlSheetHidden
        End With
    End With
End Sub

'Add the temporary sheets for computation and stuffs
Public Sub AddTemporarySheets(wkb As Workbook)

    Const iCmdWidthFactor As Integer = C_iCmdWidth
    Const iCmdHeightFactor As Integer = 30
    Dim LoRng As Range

    With wkb
        'Add Debug sheet
        .Worksheets.Add.Name = C_sSheetDebug
        .Worksheets(C_sSheetDebug).Visible = xlSheetHidden
        'Add de bug command on the debug sheet

        With wkb.Worksheets(C_sSheetDebug)
            'Debug button
            Call AddCmd(wkb, C_sSheetDebug, _
                        .Cells(2, 10).Left + 3 * C_iCmdWidth + 3 * iCmdWidthFactor + 30, _
                        .Cells(2, 1).Top, C_sShpDebug, _
                        "Debug", _
                        C_iCmdWidth + iCmdWidthFactor, C_iCmdHeight + iCmdHeightFactor, _
                        C_sCmdDebug, sShpColor:="Orange", sShpTextColor:="Black", _
                        iTextFontSize:=12)
        End With

        '--------- Adding a temporary sheets for computations
        'temp sheet
        .Worksheets.Add.Name = C_sSheetTemp
        .Worksheets(C_sSheetTemp).Visible = xlSheetVeryHidden
        'temporary sheet for analysis
        .Worksheets.Add.Name = C_sSheetAnalysisTemp
        .Worksheets(C_sSheetAnalysisTemp).Visible = xlSheetVeryHidden
        'temporary sheet for imports report
        .Worksheets.Add.Name = C_sSheetImportTemp
        .Worksheets(C_sSheetImportTemp).Visible = xlSheetVeryHidden
        'Add list auto temporary sheet
        .Worksheets.Add.Name = C_sSheetChoiceAuto
        .Worksheets(C_sSheetChoiceAuto).Visible = xlSheetVeryHidden

        'Add in choice auto list for the Geo dropdowns
        With .Worksheets(C_sSheetChoiceAuto)

            .Cells(1, 1).Value = C_sAdmName & "_2_" & "dropdown"
            .Cells(1, 3).Value = C_sAdmName & "_3_" & "dropdown"
            .Cells(1, 5).Value = C_sAdmName & "_4_" & "dropdown"

            'Add the listObjects

            'Adm 2
            Set LoRng = .Range(.Cells(1, 1), .Cells(2, 1))
            'Add the list object here
            .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = C_sTabAdm2 & "_dropdown"

            'Adm3
            Set LoRng = .Range(.Cells(1, 3), .Cells(2, 3))
            'Add the list object here
            .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = C_sTabAdm3 & "_dropdown"

            'Adm4
            Set LoRng = .Range(.Cells(1, 5), .Cells(2, 5))
            'Add the list object here
            .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = C_sTabAdm4 & "_dropdown"

            'Adm 2
            wkb.Names.Add Name:=C_sAdmName & "_2_" & "dropdown", RefersToR1C1:="=" & C_sTabAdm2 & "_dropdown" & "[" & C_sAdmName & "_2_" & "dropdown" & "]"
            'Adm3
            wkb.Names.Add Name:=C_sAdmName & "_3_" & "dropdown", RefersToR1C1:="=" & C_sTabAdm3 & "_dropdown" & "[" & C_sAdmName & "_3_" & "dropdown" & "]"
            'Adm 4
            wkb.Names.Add Name:=C_sAdmName & "_4_" & "dropdown", RefersToR1C1:="=" & C_sTabAdm4 & "_dropdown" & "[" & C_sAdmName & "_4_" & "dropdown" & "]"
        End With
    End With
End Sub

Public Sub AddAdminSheet(wkb As Workbook)

    Const iCmdWidthFactor As Integer = C_iCmdWidth
    Const iCmdHeightFactor As Integer = 30


    wkb.Worksheets(1).Name = sParamSheetAdmin
    Call RemoveGridLines(wkb.Worksheets(sParamSheetAdmin))

    'ADD BUTTONS

    With wkb.Worksheets(sParamSheetAdmin)

        'Import migration buttons
        Call AddCmd(wkb, sParamSheetAdmin, _
                    .Cells(2, 12).Left, .Cells(2, 1).Top, C_sShpAdvanced, _
                    "Import for Migration", _
                    C_iCmdWidth + iCmdWidthFactor, C_iCmdHeight + iCmdHeightFactor, _
                    C_sCmdImportMigration, iTextFontSize:=12)

        'Export Button
        Call AddCmd(wkb, sParamSheetAdmin, _
                    .Cells(2, 12).Left + C_iCmdWidth + iCmdWidthFactor + 15, _
                    .Cells(2, 1).Top, C_sShpExport, _
                    "Export", _
                    C_iCmdWidth + iCmdWidthFactor, C_iCmdHeight + iCmdHeightFactor, C_sCmdExport, _
                    iTextFontSize:=12)

    End With

End Sub

