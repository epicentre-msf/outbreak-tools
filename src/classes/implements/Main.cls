VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "main interface to linelist specifications paths"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("main interface to linelist specifications paths")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IMain

'Constants for ranges in the main sheet
Private Const RNGLLFORM As String = "RNG_LLForm"
Private Const RNGLANGSETUP As String = "RNG_LangSetup"
Private Const RNGDESIGNLL As String = "RNG_DesignLL"
Private Const RNGPATHDICO As String = "RNG_PathDico"
Private Const RNGPATHGEO As String = "RNG_PathGeo"
Private Const RNGLLDIR As String = "RNG_LLDir"
Private Const RNGLLNAME As String = "RNG_LLName"
Private Const RNGEDITION As String = "RNG_Edition"
Private Const RNGMAINREADY As String = "RNG_MainReady"
Private Const RNGUPDATE As String = "RNG_Update"
Private Const RNGLLPWDOPEN As String = "RNG_LLPwdOpen"
Private Const RNGTEMPPATH As String = "RNG_LLTemp"
Private Const RNGMAINTEMP As String = "RNG_MainTemp"
Private Const RNGASKOPEN As String = "RNG_AskOpen"

Private Const CLASSNAME As String = "Main"


Private Type TMain
    sh As Worksheet
End Type

Private this As TMain

Public Function Create(sh As Worksheet) As IMain

    CheckRequirements sh

    With New Main
        Set .Wksh = sh
        Set Create = .Self()
    End With

End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Self() As IMain
    Set Self = Me
End Property

Private Property Get Ready() As Boolean
    Dim sh As Worksheet
    Set sh = Wksh()
    Ready = (sh.Range(RNGMAINREADY).Value = "ready")
End Property

'Check if a workbook is Opened
Private Function WorkbookIsOpen(wkbName As String) As Boolean
    Dim wb As Workbook                           'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set wb = Application.Workbooks.Item(wkbName)
    WorkbookIsOpen = (Not wb Is Nothing)
    On Error GoTo 0
End Function

Private Property Get LinelistPassWord() As String
    Dim sh As Worksheet
    Set sh = Wksh()
    LinelistPassWord = sh.Range(RNGLLPWDOPEN).Value
End Property

Private Property Get ValueOf(ByVal tagName As String) As String
    Dim rngName As String
    Dim sh As Worksheet
    Dim retVal As String

    Set sh = Wksh()
    If tagName <> "temppath" Then 'special treatment for temp path
        Select Case tagName
        Case "llpassword"
            rngName = RNGLLPWDOPEN
        Case "geopath"
            rngName = RNGPATHGEO
        Case "setuppath"
            rngName = RNGPATHDICO
        Case "lldir"
            rngName = RNGLLDIR
        Case "llname"
            rngName = RNGLLNAME
        Case "setuplang"
            rngName = RNGLANGSETUP
        Case "lllang"
            rngName = RNGLLFORM
        Case "askopen"
            rngName = RNGASKOPEN
        End Select

        On Error Resume Next
            ValueOf = sh.Range(rngName).Value
        On Error GoTo 0
        Exit Property 'once you get the value of exit
    End If

    'for template path, if it is the default, return the _ribbontemplate file
    ' in the desiger misc directory
    On Error Resume Next
        retVal = sh.Range(RNGTEMPPATH).Value
    On Error GoTo 0

    If retVal = "_default" Or retVal = "default" Then
        retVal = ThisWorkbook.Path & Application.PathSeparator & "_ribbontemplate.xlsb"
    End If

    ValueOf = retVal
End Property

'Check if the main requires a template
Private Property Get HasTemplate() As Boolean
    Dim sh As Worksheet
    Set sh = Wksh()
    HasTemplate = (sh.Range(RNGMAINTEMP).Value = "has template")
End Property

'Check if the main sheet is ready for generation process

Public Sub CheckReadiness(ByVal trads As IDesTranslation)
    Dim sh As Worksheet
    Dim wkbName As String                        'Eventual Workbook name to check if a workbook is opened
    Dim redColor As Variant

    Set sh = Wksh()
    redColor = RGB(252, 228, 214)

    sh.Range(RNGMAINREADY).Value = "not ready"
    sh.Range(RNGMAINREADY).Font.color = redColor

    'Be sure the dictionary path is not empty
    If sh.Range(RNGPATHDICO).Value = vbNullString Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_PathDic")
        sh.Range(RNGPATHDICO).Interior.color = redColor
        Exit Sub
    End If

    'Now check if the file exists
    wkbName = Dir(sh.Range(RNGPATHDICO).Value)

    If wkbName = vbNullString Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_PathDic")
        sh.Range(RNGPATHDICO).Interior.color = redColor
        Exit Sub
    End If

    'Be sure the dictionnary is not opened
    If WorkbookIsOpen(wkbName) Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_CloseDic")
        sh.Range(RNGPATHDICO).Interior.color = redColor
        MsgBox trads.TranslationMsg("MSG_AlreadyOpen"), vbExclamation + vbOKOnly, trads.TranslationMsg("MSG_Title_Dictionnary")
        Exit Sub
    End If

    sh.Range(RNGPATHDICO).Interior.color = vbWhite 'if path is OK

    'Checking coherence of the GEO  ----------------------------------------------------------

    'Be sure the geo path is not empty and in that case, test the file.
    'You can create a linelist with an empty geobase though

    If Not IsEmpty(sh.Range(RNGPATHGEO)) Then
        'Now check if the file exists
        If Dir(sh.Range(RNGPATHGEO).Value) = vbNullString Then
            sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_PathGeo")
            sh.Range(RNGPATHGEO).Interior.color = redColor
            MsgBox trads.TranslationMsg("MSG_PathGeo"), vbExclamation + vbOKOnly, trads.TranslationMsg("MSG_TitleGeo")
            Exit Sub
        End If
    End If

    sh.Range(RNGPATHGEO).Interior.color = vbWhite 'if path is OK

    'Checking coherence of the Linelist File ------------------------------------------------------

    'Be sure the linelist directory is not empty
    If sh.Range(RNGLLDIR).Value = vbNullString Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_PathLL")
        sh.Range(RNGLLDIR).Interior.color = redColor
        MsgBox trads.TranslationMsg("MSG_PathLL"), vbExclamation + vbOKOnly, trads.TranslationMsg("MSG_TitleLL")
        Exit Sub
    End If

    'Be sure the linelist is not opened
    wkbName = Dir(sh.Range(RNGLLDIR).Value & Application.PathSeparator & sh.Range(RNGLLNAME).Value & ".xlsb")

    If WorkbookIsOpen(wkbName) Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_CloseOutPut")
        sh.Range(RNGLLDIR).Interior.color = redColor
        MsgBox trads.TranslationMsg("MSG_CloseOutPut"), vbExclamation + vbOKOnly, trads.TranslationMsg("MSG_Title_OutPut")
        Exit Sub
    End If

    'Be sure the directory for the linelist exists
    If Dir(sh.Range(RNGLLDIR).Value & "*", vbDirectory) = vbNullString Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_PathLL")
        sh.Range(RNGLLDIR).Interior.color = redColor
        MsgBox trads.TranslationMsg("MSG_PathLL"), vbExclamation + vbOKOnly, trads.TranslationMsg("MSG_TitleLL")
        Exit Sub
    End If

    sh.Range(RNGLLDIR).Interior.color = vbWhite 'if path is OK

    'Checking coherence of the linelist name ------------------------------------------------------

    'be sure the linelist name is not empty
    If IsEmpty(sh.Range(RNGLLNAME)) Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_LLName")
        sh.Range(RNGLLNAME).Interior.color = redColor
        Exit Sub
    End If

    'Be sure the linelist workbook is not already opened
    If WorkbookIsOpen(sh.Range(RNGLLNAME).Value & ".xlsb") Then
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_CloseLL")
        sh.Range(RNGLLNAME).Interior.color = redColor
        Exit Sub
    End If

    'Be sure the linelist name is well written (or change it)
    sh.Range(RNGLLNAME).Value = CorrectLinelistName(sh.Range(RNGLLNAME).Value)


    'If everything is ready
    sh.Range(RNGMAINREADY).Value = "ready"
    sh.Range(RNGMAINREADY).Font.color = vbWhite
End Sub

Private Sub CheckFileExistence(ByVal trads As IDesTranslation)

    Dim sh As Worksheet

    Set sh = Wksh() 'main worksheet

    If Dir(sh.Range(RNGLLDIR).Value & _
           Application.PathSeparator & _
           sh.Range(RNGLLNAME).Value & ".xlsb") <> vbNullString Then

        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_Correct") & ": " _
        & sh.Range(RNGLLNAME).Value & ".xlsb " _
         & trads.TranslationMsg("MSG_Exists")

        sh.Range(RNGEDITION).Interior.color = RGB(235, 232, 232)

        If MsgBox(sh.Range(RNGLLNAME).Value & ".xlsb " & _
                  trads.TranslationMsg("MSG_Exists") & chr(10) & _
                  trads.TranslationMsg("MSG_Question"), vbYesNo, _
                  trads.TranslationMsg("MSG_Title")) = vbNo Then

            'The user doen't want to change
            sh.Range(RNGLLNAME).Value = vbNullString
            sh.Range(RNGLLNAME).Interior.color = RGB(252, 228, 214)

            'not ready
            sh.Range(RNGMAINREADY).Value = "not ready"
            sh.Range(RNGMAINREADY).Font.color = vbWhite
            Exit Sub
        End If
    Else
        sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_Correct")
    End If
End Sub


'Check Existence of the ribbon file (default to linelist with buttons if not exists)
Private Sub CheckRibbonExistence(ByVal trads As IDesTranslation)
    Dim tempPath As String
    Dim sh As Worksheet
    Dim askRibbon As Boolean

    Set sh = Wksh()
    sh.Range(RNGMAINTEMP).Value = "has template"
    sh.Range(RNGMAINTEMP).Font.color = vbWhite

    tempPath = ValueOf("temppath")
    If tempPath = vbNullString Then
        askRibbon = True
    ElseIf Dir(tempPath) = vbNullString Then
        askRibbon = True
    End If

    If askRibbon Then
        If MsgBox(trads.TranslationMsg("MSG_Tempfile") & " " & _
                  tempPath & " " & trads.TranslationMsg("MSG_NotExists") & " " & _
                  trads.TranslationMsg("MSG_DefaultButton"), vbYesNo, _
                trads.TranslationMsg("MSG_Title")) = vbNo Then

            'The user does not want a linelist with button, he should provide
            'a valid template (If something went wrong with the template, I will default back to buttons.)
            sh.Range(RNGTEMPPATH).Value = vbNullString
            sh.Range(RNGTEMPPATH).Interior.color = RGB(252, 228, 214)
            sh.Range(RNGEDITION).Value = trads.TranslationMsg("MSG_ProvTemp")

            'Not ready
            sh.Range(RNGMAINREADY).Value = "not ready"
            sh.Range(RNGMAINREADY).Font.color = vbWhite
            Exit Sub
        Else
               sh.Range(RNGMAINTEMP).Value = "has button"
               sh.Range(RNGMAINTEMP).Font.color = vbWhite
        End If
    End If
End Sub



Private Sub ClearInputRanges(Optional ByVal clearValues As Boolean = False)
    Dim mainsh As Worksheet
    Set mainsh = Wksh()
    mainsh.Range(RNGPATHGEO).Interior.color = vbWhite
    mainsh.Range(RNGPATHDICO).Interior.color = vbWhite
    mainsh.Range(RNGLLNAME).Interior.color = vbWhite
    mainsh.Range(RNGLLDIR).Interior.color = vbWhite
    mainsh.Range(RNGEDITION).Interior.color = vbWhite
    mainsh.Range(RNGTEMPPATH).Interior.color = vbWhite

    If clearValues Then
        mainsh.Range(RNGPATHGEO).Value = vbNullString
        mainsh.Range(RNGPATHDICO).Value = vbNullString
        mainsh.Range(RNGLLNAME).Value = vbNullString
        mainsh.Range(RNGLLDIR).Value = vbNullString
        mainsh.Range(RNGEDITION).Value = vbNullString
        mainsh.Range(RNGLANGSETUP).Value = vbNullString
        mainsh.Range(RNGLLFORM).Value = vbNullString
    End If
End Sub

'Open the linelist
Private Sub OpenLL()
    Dim mainsh As Worksheet
    Set mainsh = Wksh()
    On Error GoTo ErrOpen
    'Then open it
    Application.DisplayAlerts = False
    Application.Workbooks.Open _
     fileName:=mainsh.Range(RNGLLDIR).Value & Application.PathSeparator & _
               mainsh.Range(RNGLLNAME).Value & ".xlsb"
ErrOpen:
End Sub

Private Function CorrectLinelistName(ByVal llfileName As String) As String

    'In the file name, replace forbidden characters with an underscore
    Dim fileName As String

    fileName = Replace(llfileName, "<", "_")
    fileName = Replace(fileName, ">", "_")
    fileName = Replace(fileName, ":", "_")
    fileName = Replace(fileName, "|", "_")
    fileName = Replace(fileName, "?", "_")
    fileName = Replace(fileName, "/", "_")
    fileName = Replace(fileName, "\", "_")
    fileName = Replace(fileName, "*", "_")
    fileName = Replace(fileName, ".", "_")
    fileName = Replace(fileName, """", "_")

    CorrectLinelistName = Application.WorksheetFunction.Trim(fileName)
End Function

'Update the status Bar in the main sheet
Private Sub UpdateStatus(ByVal numValue As Integer)

    Const NUMBEROFBARS As Integer = 100

    Dim CurrentStatus As Integer
    Dim currentEvent As Boolean
    Dim sh As Worksheet

    Set sh = Wksh()

    currentEvent = Application.ScreenUpdating
    Application.ScreenUpdating = True
    CurrentStatus = (NUMBEROFBARS) * Round(numValue / 100, 1)
    sh.Range(RNGUPDATE).Value = "[" & String(CurrentStatus, "|") & SPACE(NUMBEROFBARS - CurrentStatus) & "]" & _
                                                                                                          " " & CInt(numValue) & "%"
    Application.ScreenUpdating = currentEvent
End Sub

'Add information on edition to the end user
Private Sub AddInfo(ByVal trads As IDesTranslation, ByVal msgCode As String, Optional ByVal infoName As String = "edition")
    Dim mainsh As Worksheet
    Dim infoRng As Range
    Dim rngName As String

    Select Case infoName
    Case "edition"
        rngName = RNGEDITION
    Case "geopath"
        rngName = RNGPATHGEO
    Case "setuppath"
        rngName = RNGPATHDICO
    Case "lldir"
        rngName = RNGLLDIR
    Case "llname"
        rngName = RNGLLNAME
    Case "setuplang"
        rngName = RNGLANGSETUP
    Case "lllang"
        rngName = RNGLLFORM
    Case "temppath"
        rngName = RNGTEMPPATH
    Case Else
        rngName = RNGEDITION
    End Select

    Set mainsh = Wksh()
    Set infoRng = mainsh.Range(rngName)

    'End user informations
    infoRng.Value = trads.TranslationMsg(msgCode)
    infoRng.Interior.color = vbWhite
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub CheckRequirements(ByVal sh As Worksheet)

    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String

    'Sheet Should Exists
    If sh Is Nothing Then ThrowError "Geo sheet not defined"
    Set rngTable = New BetterArray
    rngTable.Push RNGLLFORM, RNGLANGSETUP, RNGDESIGNLL, _
                  RNGPATHDICO, RNGPATHGEO, RNGLLDIR, RNGLLNAME, _
                  RNGEDITION, RNGMAINREADY, RNGUPDATE, RNGLLPWDOPEN, _
                  RNGTEMPPATH, RNGMAINTEMP, RNGASKOPEN

    For counter = rngTable.LowerBound To rngTable.UpperBound
        rngName = rngTable.Item(counter)
        RangeExists sh, rngName
    Next
End Sub

Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)

    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name
End Sub

Private Property Get IMain_ValueOf(ByVal tagName As String) As String
    IMain_ValueOf = ValueOf(tagName)
End Property

Private Property Get IMain_Ready() As Boolean
    IMain_Ready = Ready()
End Property

Private Sub IMain_CheckReadiness(ByVal trads As IDesTranslation)
    CheckReadiness trads
End Sub

Private Sub IMain_CheckFileExistence(ByVal trads As IDesTranslation)
    CheckFileExistence trads
End Sub

Private Sub IMain_ClearInputRanges(Optional ByVal clearValues As Boolean = False)
    ClearInputRanges clearValues:=clearValues
End Sub

Private Sub IMain_UpdateStatus(ByVal numValue As Integer)
    UpdateStatus numValue
End Sub

Private Sub IMain_AddInfo(ByVal trads As IDesTranslation, ByVal msgCode As String, _
                          Optional ByVal infoName As String = "edition")
    AddInfo trads, msgCode, infoName
End Sub

Private Property Get IMain_LinelistPassword() As String
    IMain_LinelistPassword = LinelistPassWord()
End Property

Private Sub IMain_OpenLL()
    OpenLL
End Sub

Private Property Get IMain_HasTemplate() As Boolean
    IMain_HasTemplate = HasTemplate()
End Property

Private Sub IMain_CheckRibbonExistence(ByVal trads As IDesTranslation)
    CheckRibbonExistence trads
End Sub
