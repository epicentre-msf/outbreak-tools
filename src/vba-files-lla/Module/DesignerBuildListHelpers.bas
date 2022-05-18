Attribute VB_Name = "DesignerBuildListHelpers"
Option Explicit
'-------
'Transfert Codes and forms to the designer

Public Sub TransferDesignerCodes(xlsapp As Excel.Application)

    On Error Resume Next
    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp")
    MkDir (Environ("Temp") & Application.PathSeparator & "LinelistApp") 'create a folder for sending all the data from designer
    On Error GoTo 0
    
    Dim Wkb As Workbook
    Set Wkb = xlsapp.ActiveWorkbook
    
    DoEvents
        
    'Transfert form is for sending forms from the actual excel workbook to another
    Call TransferForm(Wkb, C_sFormGeo)
    Call TransferForm(Wkb, C_sFormShowHide)
    Call TransferForm(Wkb, C_sFormExport)
    Call TransferForm(Wkb, C_sFormExportMig)
    Call TransferForm(Wkb, C_sFormImportMig)

    'TransferCode is for sending modules  (Modules) or classes (Classes) from actual excel workbook to another excel workbook
    Call TransferCode(Wkb, C_sModLinelist, "Module")
    Call TransferCode(Wkb, C_sModLLGeo, "Module")
    Call TransferCode(Wkb, C_sModLLShowHide, "Module")
    Call TransferCode(Wkb, C_sModHelpers, "Module")
    Call TransferCode(Wkb, C_sModLLMigration, "Module")
    Call TransferCode(Wkb, C_sModLLConstants, "Module")
    Call TransferCode(Wkb, C_sModEsthConstants, "Module")
    Call TransferCode(Wkb, C_sModLLExport, "Module")
    Call TransferCode(Wkb, C_sModLLTrans, "Module")
    Call TransferCode(Wkb, C_sModLLDict, "Module")
    Call TransferCode(Wkb, C_sClaBA, "Class")
    
    Set Wkb = Nothing
    
End Sub


'Transfert code from one module to a worksheet to trigger some events
'@sSheetName the sheet name we want to transfer to
'@sNameModule the name of the module we want to copy code from

Sub TransferCodeWks(xlsapp As Excel.Application, sSheetName As String, _
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
    
    With xlsapp
        Set vbProj = .ActiveWorkbook.VBProject
        Set vbComp = vbProj.VBComponents(.Sheets(sSheetName).CodeName)
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

'@xlsapp: excel application we want to move
'@sSheetName: the name of the Sheet in the designer we want to move

Public Sub TransferSheet(xlsapp As Excel.Application, sSheetName As String)

    'Since We can't move worksheet from one instance to another
    'we need to save as a temporary file and then move it to another instance.
    'DesignerWorkbook is the actual workbook we want to copy from

    DesignerWorkbook.Worksheets(sSheetName).Copy
    DoEvents

    On Error Resume Next
        Kill Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"
    On Error GoTo 0

    ActiveWorkbook.SaveAs Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"
    ActiveWorkbook.Close

    DoEvents

    With xlsapp
        .Workbooks.Open Filename:=Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx", UpdateLinks:=False
        
        .Sheets(sSheetName).Select
        .Sheets(sSheetName).Copy After:=.Workbooks(1).Sheets(1)
 
        DoEvents
        .Workbooks("Temp.xlsx").Close
    End With
    
    DoEvents
    
    Kill Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"

End Sub

'-----
'Transfert a form the actual Designer to the linelist's Workbook
'@Wkb : A workbook
'@sFormName: The name of the form to transfert

Sub TransferForm(Wkb As Workbook, sFormName As String)
    
    'The form is sent to the LinelisteApp folder
    On Error Resume Next
    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm")
    On Error GoTo 0
    
    DoEvents
    DesignerWorkbook.VBProject.VBComponents(sFormName).Export Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm"
    Wkb.VBProject.VBComponents.Import Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm"
    DoEvents

    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm")
End Sub


'---
'The Goal is to transfer one Code Module/Class from the designer to the
'linelist sheet
'@wkb a workbook
'@sType the type of the code to transfer (Module or Class)
'@sModule: The Name of the module to transfer

Sub TransferCode(Wkb As Workbook, sNameModule As String, sType As String)

    Dim oNouvM As Object 'New module name
    Dim sNouvCode As String 'New module code

    'get all the values within the actual module to transfer
    With DesignerWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    'create to code or module if needed
    Select Case sType
    Case "Module"
        Set oNouvM = Wkb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "Class"
        Set oNouvM = Wkb.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select

    'keep the name and add the codes
    oNouvM.Name = sNameModule
    With Wkb.VBProject.VBComponents(oNouvM.Name).CodeModule
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

'@xlsapp an excel application
'sHeader the String we want to add space (in case) to
'sSheetName: The concernec SheetName
'iStartLine: Integer, the line where the table listobject starts

Public Function AddSpaceToHeaders(xlsapp As Excel.Application, _
                                  sHeader As String, _
                                  sSheetName As String, iStartLine As Integer)
    Dim i As Integer

    AddSpaceToHeaders = ""
    With xlsapp
        i = 1
        While i <= .Sheets(sSheetName).Cells(iStartLine, Columns.Count).End(xlToLeft).Column And Replace(UCase(.Sheets(sSheetName).Cells(iStartLine, i).value), " ", "") <> Replace(UCase(sHeader), " ", "")
            i = i + 1
        Wend
        If Replace(UCase(xlsapp.Sheets(sSheetName).Cells(iStartLine, i).value), " ", "") = Replace(UCase(sHeader), " ", "") Then
            AddSpaceToHeaders = xlsapp.Sheets(sSheetName).Cells(iStartLine, i).value & " "
        Else
            AddSpaceToHeaders = sHeader
        End If
    End With

End Function

'----


'Add a Button command to a Sheet (create the button and addit)
'@xlsapp: excel applicaiton
'@sSheet: The Sheet we want to add the button
'@sShpName: The name we want to give to the shape (Shape Name)
'@sText: The text to put on the button
'@iCmdWidth: The command with
'@iCmdHeight: The command height
'@sCommand: The binding command on the Shape
'@sShpColor: The color of the Shape
'@sShpTextColor: color of the text for each of the shapes

Sub AddCmd(xlsapp As Excel.Application, sSheetName As String, iLeft As Integer, iTop As Integer, _
           sShpName As String, sText As String, iCmdWidth As Integer, iCmdHeight As Integer, _
           sCommand As String, Optional sShpColor As String = "MainSecBlue", _
           Optional sShpTextColor As String = "White")

    
    sText = translate_LineList(sText, Sheets("linelist-translation").[T_tradShapeLL])

    With xlsapp.Sheets(sSheetName)
        .Shapes.AddShape(msoShapeRectangle, iLeft + 3, iTop + 3, iCmdWidth, iCmdHeight).Name = sShpName
        .Shapes(sShpName).Placement = xlFreeFloating
        .Shapes(sShpName).TextFrame2.TextRange.Characters.Text = sText
        .Shapes(sShpName).TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Shapes(sShpName).TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shapes(sShpName).TextFrame2.WordWrap = msoFalse
        .Shapes(sShpName).TextFrame2.TextRange.Font.Size = 9
        .Shapes(sShpName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Helpers.GetColor(sShpTextColor)
        .Shapes(sShpName).Fill.ForeColor.RGB = Helpers.GetColor(sShpColor)
        .Shapes(sShpName).Fill.BackColor.RGB = Helpers.GetColor(sShpColor)
        '.Shapes("SHP_NomVisibleApps").Fill.TwoColorGradient msoGradientHorizontal, 1
        .Shapes(sShpName).OnAction = sCommand
    End With

End Sub


'Little Subs used when working with the Creation of the data Entry for a sheet of type Linelist
'Add the Sub Label
Sub AddSubLab(Wksh As Worksheet, iSheetStartLine As Integer, _
              iCol As Integer, sMainLab As String, sSubLab As String, _
              Optional sSubLabColor As String = "SubLabBlue")
    With Wksh
        .Cells(iSheetStartLine, iCol).value = _
        .Cells(iSheetStartLine, iCol).value & Chr(10) & sSubLab
    
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

    Dim iDecType As Integer 'Just to get the decimal number at the end of decimal
    'Dim iDecNb As Integer
    Dim i As Integer
    Dim sNbDeci As String 'Number of decimals


    'Check to be sure that the actual type contains decimal
    With Wksh
        If InStr(1, sType, C_sDictTypeDec) > 0 Then
            iDecType = CInt(Replace(sType, C_sDictTypeDec, ""))
            sType = C_sDictTypeDec
            i = 0
            sNbDeci = ""
            While i < iDecType
                sNbDeci = "0" & sNbDeci
                i = i + 1
            Wend
        End If

        Select Case sType
            'Text Type
            Case C_sDictTypeText
                .Cells(iSheetStartLine + 1, iCol).NumberFormat = "@"
                'Integer
            Case C_sDictTypeInt
                 .Cells(iSheetStartLine + 1, iCol).NumberFormat = "0"
                'Date Type
            Case C_sDictTypeDate
                 .Cells(iSheetStartLine + 1, iCol).NumberFormat = "d-mmm-yyy"
                'Decimal
            Case C_sDictTypeDec
                 .Cells(iSheetStartLine + 1, iCol).NumberFormat = "0." & sNbDeci
            Case Else
            'If I don't know the type, put in text
             .Cells(iSheetStartLine + 1, iCol).NumberFormat = "@"
        End Select
    End With

End Sub

'Add the choices
Sub AddChoices(Wksh As Worksheet, iSheetStartLine As Integer, iCol As Integer, _
             ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
             sChoice As String, sAlert As String, sMessage As String)
    
    Dim sValidationList As String
    With Wksh
        sValidationList = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sChoice)
        If sValidationList <> "" Then
             Call Helpers.SetValidation(.Cells(iSheetStartLine + 1, iCol), _
                                            sValidationList, _
                                            Helpers.GetValidationType(sAlert), _
                                            sMessage)
        End If
    End With
End Sub


'Add Geo
Sub AddGeo(xlsapp As Excel.Application, DictData As BetterArray, DictHeaders As BetterArray, sSheetName As String, iSheetStartLine As Integer, iCol As Integer, _
          iSheetSubSecStartLine As Integer, iDictLine As Integer, sVarName As String, sMessage As String, iNbshifted As Integer)

    With xlsapp.Worksheets(sSheetName)
        .Cells(iSheetStartLine, iCol).Interior.Color = GetColor("Orange")
                        'update the columns only for the geo
        Call Add4GeoCol(xlsapp, DictData, DictHeaders, sSheetName, sVarName, iSheetStartLine, _
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

Sub Add4GeoCol(xlsapp As Excel.Application, DictData As BetterArray, DictHeaders As BetterArray, _
            sSheetName As String, sVarName As String, iStartLine As Integer, iCol As Integer, _
            sMessage As String, iStartLineSubLab As Integer, iDictLine As Integer, iNbshifted As Integer)


    Dim sLab As String 'Temporary variable, label of the Admin level
    Dim LineValues As BetterArray
    Dim iRow As Integer
      
    Set LineValues = New BetterArray
    LineValues.LowerBound = 1
    
    iRow = iDictLine + iNbshifted
    
    With xlsapp.Worksheets(sSheetName)

        'Admin 4
        sLab = SheetGeo.ListObjects(C_sTabAdm4).HeaderRowRange.Item(4).value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).value = AddSpaceToHeaders(xlsapp, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "4" & "_" & sVarName
        .Cells(iStartLine, iCol + 1).Interior.Color = vbWhite
        .Cells(iStartLine + 1, iCol + 1).Locked = False

        'Admin 3
        sLab = SheetGeo.ListObjects(C_sTabAdm3).HeaderRowRange.Item(3).value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).value = AddSpaceToHeaders(xlsapp, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "3" & "_" & sVarName
        .Cells(iStartLine, iCol + 1).Interior.Color = vbWhite
        .Cells(iStartLine + 1, iCol + 1).Locked = False

        'Admin 2
        sLab = SheetGeo.ListObjects(C_sTabAdm2).HeaderRowRange.Item(2).value
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(iStartLine, iCol + 1).value = AddSpaceToHeaders(xlsapp, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol + 1).Name = C_sAdmName & "2" & "_" & sVarName
        .Cells(iStartLine, iCol + 1).Interior.Color = vbWhite
        .Cells(iStartLine + 1, iCol + 1).Locked = False

        'Admin 1
        sLab = SheetGeo.ListObjects(C_sTabadm1).HeaderRowRange.Item(1).value
        .Cells(iStartLine, iCol).value = AddSpaceToHeaders(xlsapp, sLab, sSheetName, iStartLine)
        .Cells(iStartLine, iCol).Name = C_sAdmName & "1" & "_" & sVarName
        .Cells(iStartLine, iCol).Interior.Color = GetColor("Orange")
        .Cells(iStartLine + 1, iCol).Locked = False

        'ajout des formules de validation
        .Cells(iStartLine + 1, iCol).Validation.Delete

        .Cells(iStartLine + 1, iCol).Validation.Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, _
                         Formula1:="=" & C_sSheetGeo & "!" & SheetGeo.Range(C_sTabadm1).Columns(1).Address

        .Cells(iStartLine + 1, iCol).Validation.IgnoreBlank = True
        .Cells(iStartLine + 1, iCol).Validation.InCellDropdown = True
        .Cells(iStartLine + 1, iCol).Validation.InputTitle = ""
        .Cells(iStartLine + 1, iCol).Validation.errorTitle = ""
        .Cells(iStartLine + 1, iCol).Validation.InputMessage = ""
        .Cells(iStartLine + 1, iCol).Validation.ErrorMessage = sMessage
        .Cells(iStartLine + 1, iCol).Validation.ShowInput = True
        .Cells(iStartLine + 1, iCol).Validation.ShowError = True
    End With
    
    'Updating the Dictionary for future uses
    With xlsapp.Worksheets(C_sParamSheetDict)
        'Admin 4
        LineValues.Items = DictData.ExtractSegment(RowIndex:=iDictLine)
        LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "4"
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).value = ""
        'Admin 3
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "3"
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).value = ""
        'Admin 2
        .Rows(iRow + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         LineValues.Item(DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "2"
        LineValues.ToExcelRange Destination:=.Cells(iRow + 2, 1), TransposeValues:=True
        .Cells(iRow + 2, 1).value = ""
         
         Set LineValues = Nothing
    End With
End Sub




'Build a merge area for subsections and sections
'Wksh the workheet on which we want to build the merge area
Sub BuildMergeArea(Wksh As Worksheet, iStartLineOne As Integer, iPrevColumn As Integer, _
                        Optional iActualColumn As Integer = -1, Optional iStartLineTwo As Integer = -1, _
                        Optional sColorMainSec As String = "MainSecBlue", _
                        Optional sColorSubSec As String = "SubSecBlue")

    Dim oCell As Object

    With Wksh

        If iActualColumn = -1 Then
            .Cells(iStartLineOne, iPrevColumn).HorizontalAlignment = xlCenter
            .Cells(iStartLineOne, iPrevColumn).Interior.Color = Helpers.GetColor(sColorSubSec)
            Call Helpers.WriteBorderLines(.Cells(iStartLineOne, iPrevColumn))
            Exit Sub
        End If

        .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1)).Merge
        .Cells(iStartLineOne, iPrevColumn).MergeArea.HorizontalAlignment = xlCenter

        If (iStartLineTwo <> -1) Then
            With .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1))
                .Interior.Color = Helpers.GetColor(sColorMainSec)
                .Font.Color = Helpers.GetColor("White")
                .Font.Bold = True
                .Font.Size = C_iLLMainSecFontSize
            End With
            'For the sub sections, if nothing is mentionned,
            'just put them in white (or the same color as the main sections)
            For Each oCell In .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineTwo, iActualColumn - 1))
                  If oCell.value = "" Then
                    oCell.Interior.Color = Helpers.GetColor("White")
                  End If
            Next
            Set oCell = Nothing
            'Write borders to the ranges including the subsection
            Call Helpers.WriteBorderLines(.Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineTwo, iActualColumn - 1)))
        Else
            With .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1))
                .Interior.Color = Helpers.GetColor(sColorSubSec)
                .Font.Color = Helpers.GetColor(sColorMainSec)
                .Font.Size = C_iLLSubSecFontSize
            End With
            Call Helpers.WriteBorderLines(.Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1)))
        End If
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


Public Function ValidationFormula(sFormula As String, sSheetName As String, VarNameData As BetterArray, _
                                    ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                    SpecCharData As BetterArray, Optional bLocal As Boolean = True) As String
    'Returns a string of cleared formula

    ValidationFormula = ""

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
   

    Dim isError As Boolean
    Dim OpenedQuotes As Boolean                  'Test if the formula has opened some quotes
    Dim QuotedCharacter As Boolean
    Dim NoErrorAndNoEnd As Boolean
    Set FormulaAlphaData = New BetterArray       'Alphanumeric values of one formula

    FormulaAlphaData.LowerBound = 1

    'squish the formula (removing multiple spaces) to avoid problems related to
    'space collapsing and upper/lower cases
    sFormulaATest = "(" & Application.WorksheetFunction.Trim(sFormula) & ")"

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
                            icolNumb = ColumnIndexData.Item(VarNameData.IndexOf(sAlphaValue))
                            sAlphaValue = "'" & sSheetName & "'!" & Cells(C_eStartLinesLLData + 1, icolNumb).Address(False, True)
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
        If (bLocal) Then
            ValidationFormula = Helpers.GetInternationalFormula(sAlphaValue)
        Else
            ValidationFormula = "=" & sAlphaValue
        End If
    End If
    
    Set FormulaAlphaData = Nothing

End Function





'Setting the min and the max validation
Sub BuildValidationMinMax(oRange As Range, iMin As String, iMax As String, iAlertType As Byte, sTypeValidation As String, sMessage As String)

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
        .ErrorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With
End Sub



'Ensure a sheet name has good name
Public Function EnsureGoodSheetName(ByVal sSheetName As String) As String

    EnsureGoodSheetName = sSheetName

    If sSheetName = C_sSheetGeo Or sSheetName = C_sSheetFormulas Or sSheetName = C_sSheetPassword Or sSheetName = C_sSheetTemp Or sSheetName = C_sSheetLLTranslation Or _
                     sSheetName = C_sParamSheetDict Or sSheetName = C_sParamSheetExport Or sSheetName = C_sParamSheetChoices Or sSheetName = C_sParamSheetTranslation Then
        EnsureGoodSheetName = sSheetName & "_"
    End If

End Function
