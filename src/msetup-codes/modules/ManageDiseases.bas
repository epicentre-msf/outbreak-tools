Attribute VB_Name = "ManageDiseases"
Option Explicit

Private Const TRADTABLE As String = "TabTransId"
Private Const TRADTABLESHEET As String = "__ribbonTranslation"
Private Const DROPSHEET As String = "__dropdowns"
Private Const PASSSHEET As String = "__pass"
Private Const RNG_FileLang As String = "RNG_FileLang"
Private trads As ITranslationObject
Private wb As Workbook
Private pass As IPasswords


Private Sub InitializeTrads()

    Dim Lo As ListObject
    Dim tradTagsh As Worksheet
    Dim fileLang As String

    Set wb = ThisWorkbook
    Set tradTagsh = wb.Worksheets(TRADTABLESHEET)
    Set Lo = tradTagsh.ListObjects(TRADTABLE)
    fileLang = wb.Worksheets(TRADTABLESHEET).Range(RNG_FileLang).Value
    Set trads = Translation.Create(Lo, fileLang)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEET))
End Sub

Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'Clean the formula and remove all eventual trailing spaces
Private Function Clean(chunk As String) As String

    'maximum length for excel sheetname is 30
    Const SHEETNAMEMAXLENGTH As Integer = 30

    Dim cleanedChunk As String

    cleanedChunk = Application.WorksheetFunction.SUBSTITUTE(chunk, chr(160), " ")
    cleanedChunk = Application.WorksheetFunction.Clean(cleanedChunk)
    cleanedChunk = Replace(cleanedChunk, "<", "_")
    cleanedChunk = Replace(cleanedChunk, ">", "_")
    cleanedChunk = Replace(cleanedChunk, ":", "_")
    cleanedChunk = Replace(cleanedChunk, "|", "_")
    cleanedChunk = Replace(cleanedChunk, "?", "_")
    cleanedChunk = Replace(cleanedChunk, "/", "_")
    cleanedChunk = Replace(cleanedChunk, "\", "_")
    cleanedChunk = Replace(cleanedChunk, "*", "_")
    cleanedChunk = Replace(cleanedChunk, ".", "_")
    cleanedChunk = Replace(cleanedChunk, """", "_")

    Clean = Left(Application.WorksheetFunction.Trim(cleanedChunk), SHEETNAMEMAXLENGTH)
End Function


Public Sub AddDisease()

    Dim disObj As IDisease
    Dim disName As String
    Dim dropObj As IDropdownLists
    Dim counter As Long
   
    InitializeTrads

    'Ask for Autorisation
    If MsgBox(trads.TranslatedValue("askConfirmAddDis"), _ 
             vbExclamation + vbYesNo, _ 
              trads.TranslatedValue("askConfirm")) = vbNo _ 
    Then Exit Sub

    counter = 0
    Do While (disName = vbNullString)
        disName = InputBox(trads.TranslatedValue("enterDis"), _
                        trads.TranslatedValue("enterValue"))
        disName = Clean(disName)
        counter = counter + 1
        If (counter > 5) Then Exit Do
    Loop

    If (disName = vbNullString) Then 
        If (counter > 5) Then MsgBox trads.TranslatedValue("errDisName"), vbCritical, trads.TranslatedValue("error")
        Exit Sub
    End If

    BusyApp
    On Error GoTo ExitAdd

    Set dropObj = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    Set disObj = Disease.Create(wb, dropObj)
    pass.UnProtectWkb wb
    disObj.Add disName

    'Transfert codes to a disease worksheet
    Misc.TransferCodeWksh disName

    'Protect.Worksheet
    pass.Protect disName
    pass.ProtectWkb wb

ExitAdd:
    NotBusyApp
End Sub


Public Sub RemoveDisease()

    Dim dropObj As IDropdownLists
    Dim disObj As IDisease
    Dim sh As Worksheet

    InitializeTrads
    Set sh = ActiveSheet

    If sh.Cells(2, 4).Value <> "DISSHEET" Then
        MsgBox trads.TranslatedValue("errDisNotFound"), vbCritical, trads.TranslatedValue("error")
        Exit Sub
    End If

    'Ask for Autorisation
    If MsgBox(trads.TranslatedValue("askConfirmRemDis"), _ 
              vbCritical + vbYesNo, _ 
              trads.TranslatedValue("askConfirm")) = vbNo _ 
    Then Exit Sub

    BusyApp
    On Error GoTo ExitRemove
    pass.UnProtect sh
    pass.UnProtectWkb wb

    Set dropObj = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    Set disObj = Disease.Create(wb, dropObj)
    disObj.Remove sh.Name

    pass.ProtectWkb wb

ExitRemove:
    NotBusyApp
End Sub


Public Sub ClearDiseaseSheet()

    Dim sh As Worksheet

    InitializeTrads
    Set sh = ActiveSheet

    If sh.Cells(2, 4).Value <> "DISSHEET" Then
        MsgBox trads.TranslatedValue("errDisNotFound"), vbCritical, trads.TranslatedValue("error")
        Exit Sub
    End If

     If MsgBox(trads.TranslatedValue("askConfirmClearDis"), _ 
              vbCritical + vbYesNo, _ 
              trads.TranslatedValue("askConfirm")) = vbNo _ 
    Then Exit Sub

    On Error Resume Next
        BusyApp
        pass.UnProtect sh
        sh.ListObjects(1).DataBodyRange.ClearContents
        pass.Protect sh
        NotBusyApp
    On Error GoTo 0
End Sub