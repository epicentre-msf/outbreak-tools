Attribute VB_Name="Misc"

Option Explicit

Private Const TRADSHEETNAME As String = "Translations"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const PASSSHEETNAME As String = "__pass"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const TRADTABLE As String = "TabTransId"
Private Const TRADTABLESHEET As String = "__ribbonTranslation"
Private Const RNG_FileLang As String = "RNG_FileLang"
Private Const DROPSHEET As String = "__dropdowns"

'@IgnoreModule EmptyMethod

Public Sub TransferCodeWksh(ByVal sheetName As String)

   Const CHANGEMODULENAME As String = "EventsSheetChange"
   Const WBMODULENAME As String = "EventsWorkbook"

   Dim codeContent As String                    'a string to contain code to add
   Dim vbProj As Object                         'component, project and modules
   Dim vbComp As Object
   Dim codeMod As Object
   Dim modName As String
   Dim currwb As Workbook

   Set currwb = ThisWorkbook

    modName = IIf(sheetName = "__WorkbookLevel", WBMODULENAME, CHANGEMODULENAME)
    'save the code module in the string sNouvCode
    With currwb.VBProject.VBComponents(modName).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With
    With currwb
        Set vbProj = .VBProject
        'The component could be the workbook code name for workbook related transfers
        If sheetName = "__WorkbookLevel" Then
            Set vbComp = vbProj.VBComponents(.codeName)
        Else
            Set vbComp = vbProj.VBComponents(.sheets(sheetName).codeName)
        End If
        Set codeMod = vbComp.CodeModule
    End With
    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub


Public Sub TranslateWbElmts(ByVal langId As String)

    Dim wb As Workbook
    Dim pass As IPasswords
    Dim drop As IDropdownLists
    Dim sh As Worksheet
    Dim hRng As Range
    Dim trads As ITranslation
    Dim selectValue As String
    
    Set wb = ThisWorkbook
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translation.Create(wb.Worksheets(TRADTABLESHEET).ListObjects(1), langId)
    Set drop = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    selectValue = trads.TranslatedValue("selectValue")

    For Each sh In wb.Worksheets

        'Update elements in the disease worksheet
        If sh.Cells(2, 4).Value = "DISSHEET" Then
            pass.UnProtect sh

            'Change the headers to the corresponding language
            Set hRng = sh.ListObjects(1).HeaderRowRange

            hRng.Cells(1, 1).Value = trads.TranslatedValue("varOrder")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("varSection")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("varName")
            hRng.Cells(1, 4).Value = trads.TranslatedValue("varLabel")
            hRng.Cells(1, 5).Value = trads.TranslatedValue("varChoice")
            hRng.Cells(1, 6).Value = trads.TranslatedValue("choiceVal")
            hRng.Cells(1, 7).Value = trads.TranslatedValue("varStatus")


            'Change the dropdown values for the columns status and visibility
            With sh.ListObjects(1)
                
                'variable status
                drop.SetValidation cellRng:=.ListColumns(5).DataBdoyRange, _
                                   listName:="__var_status_" & LCase(langId), _
                                   alertType:="error", _
                                   message:=selectValue
                
                'variable visibility
                drop.SetValidation cellRng:=.ListColumns(6).DataBodyRange, _
                                   listName:="__var_status_" & LCase(langId), _
                                   alertType:="error", message:=selectValue
            End With

            pass.Protect sh

        'Update columns in the variable worksheet
        
        ElseIf sh.Name = "Variables" Then


            Set hRng = sh.ListObjects(1).HeaderRowRange
            
            pass.UnProtect sh
            
            hRng.Cells(1, 1).Value = trads.TranslatedValue("varOrder")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("varSection")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("varName")
            hRng.Cells(1, 4).Value = trads.TranslatedValue("varLabel")
            hRng.Cells(1, 5).Value = trads.TranslatedValue("defChoice")
            hRng.Cells(1, 6).Value = trads.TranslatedValue("choiceVal")
            hRng.Cells(1, 7).Value = trads.TranslatedValue("defStatus")
            hRng.Cells(1, 8).Value = trads.TranslatedValue("comments")

            'Variable status validation
            pass.Protect sh

        ElseIf sh.Name = "Choices" Then

            Set hRng = sh.ListObjects(1).HeaderRowRange
            
            pass.UnProtect sh

            hRng.Cells(1, 1).Value = trads.TranslatedValue("listName")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("orderingList")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("longLabel")
            hRng.Cells(1, 4).Value = trads.TranslatedValue("shortLabel")
            

            pass.Protect sh
        End If

    Next
End Sub

