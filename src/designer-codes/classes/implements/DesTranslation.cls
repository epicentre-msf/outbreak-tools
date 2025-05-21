VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Translation of the designer"


'@IgnoreModule UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Translation of the designer")

Option Explicit

'Exposed methods

Implements IDesTranslation

Private Type TDesTranslation
    sh As Worksheet
End Type

Private Const TABSHAPE As String = "T_tradShape"
Private Const TABRANGE As String = "T_tradRange"
Private Const TABSMSG As String = "T_tradMsg"
Private Const TABLANG As String = "T_Lang"
Private Const TABDROP As String = "T_tradDrop"
Private Const TABLANGDICT As String = "T_LanguageDictionary"
Private Const CLASSNAME As String = "DesTranslation"
Private Const RNGMAINLANG As String = "RNG_MainLang"
Private Const RNGMAINLANGCODE As String = "RNG_MainLangCode"
Private Const RNGLANGDICTLIST As String = "LangDictList"

Private this As TDesTranslation

Public Function Create(ByVal sh As Worksheet) As IDesTranslation

    CheckRequirements sh

    With New DesTranslation
        Set .Wksh = sh
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As IDesTranslation
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Private Property Get DesignerLanguage() As String
    Dim sh As Worksheet
    Dim lang As Variant
    Set sh = Wksh()

    lang = sh.Range(RNGMAINLANGCODE).Value

    If (TypeName(lang) = "String") Then
        DesignerLanguage = lang
    Else
        DesignerLanguage = vbNullString
    End If
End Property

Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject

    Dim Lo As listObject
    Dim lang As String
    Dim sh As Worksheet

    Set sh = Wksh()
    lang = DesignerLanguage()

    Select Case trans
    Case TranslationOfMessages
        Set Lo = sh.ListObjects(TABSMSG)
    Case TranslationOfShapes
        Set Lo = sh.ListObjects(TABSHAPE)
    Case TranslationOfRanges
        Set Lo = sh.ListObjects(TABRANGE)
    Case TranslationOfDropdowns
        Set Lo = sh.ListObjects(TABDROP)
    End Select

    Set TransObject = TranslationObject.Create(Lo, lang)
End Property

Private Sub TranslateDesigner(ByVal mainsh As Worksheet)

    Dim tradobj As ITranslationObject
    Dim shp As Shape
    Dim counter As Long
    Dim shapesNames As BetterArray               'Name of all shapes to translate in worksheet main
    Dim rngNames As BetterArray                  'List of all named ranges in the worksheet main
    Dim Lo As listObject
    Dim sh As Worksheet
    Dim langCode As String                       'Language for the designer
    Dim prevFont As String
    Dim rngCode As String                        'Code Name of the range
    Dim mainRng As Range

    Set sh = Wksh()

    langCode = DesignerLanguage()

    If (langCode = vbNullString) Then Exit Sub

     'Translation of shapes
     Set shapesNames = New BetterArray
     Set Lo = sh.ListObjects(TABSHAPE)
     Set tradobj = TransObject(TranslationOfShapes)
     shapesNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For Each shp In mainsh.Shapes
         If shapesNames.Includes(shp.Name) Then
             With shp
                 'keep the previous font
                 prevFont = .TextFrame.Characters.Font.Name
                 .TextFrame.Characters.Text = tradobj.TranslatedValue(shp.Name)
                 .TextFrame.Characters.Font.Name = prevFont
             End With
         End If
    Next

   'Translation of label Ranges
     Set rngNames = New BetterArray
     Set Lo = sh.ListObjects(TABRANGE)
     Set tradobj = TransObject(TranslationOfRanges)
     rngNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For counter = rngNames.LowerBound To rngNames.UpperBound
         rngCode = rngNames.Item(counter)
         'Test if the range exists before changing its value
         On Error Resume Next
         Set mainRng = mainsh.Range(rngCode)
         On Error GoTo 0

       If Not (mainRng Is Nothing) Then
             mainRng.Value = tradobj.TranslatedValue(rngCode)
             Set mainRng = Nothing
         End If
    Next

    'Translation of dropdowns
    Set rngNames = New BetterArray
     Set Lo = sh.ListObjects(TABDROP)
     Set tradobj = TransObject(TranslationOfDropdowns)
     rngNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For counter = rngNames.LowerBound To rngNames.UpperBound
        rngCode = rngNames.Item(counter)
        'Test if the range exists before changing its value
        On Error Resume Next
        Set mainRng = mainsh.Range(rngCode)
        

        If Not (mainRng Is Nothing) Then

            With mainRng.validation
                .Delete
                'for all the others, add an information alert
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="= " & tradobj.TranslatedValue(rngCode)

                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = vbNullString
                .errorTitle = vbNullString
                .InputMessage = vbNullString
                .errorMessage = vbNullString
                .ShowInput = True
                .ShowError = True
            End With
            
            mainRng.ClearContents
            Set mainRng = Nothing
        End If
        On Error GoTo 0
    Next
End Sub

Private Function TranslationMsg(ByVal msgCode As String) As String
    Dim msgObject As ITranslationObject
    Set msgObject = TransObject(TranslationOfMessages)
    TranslationMsg = msgObject.TranslatedValue(msgCode)
End Function

Private Property Get LangListRng() As Range
    Dim sh As Worksheet
    Set sh = Wksh()
    Set LangListRng = sh.Range(RNGLANGDICTLIST).Cells(1, 1)
End Property

'Requirements for instanciation
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim tabl As BetterArray
    If sh Is Nothing Then ThrowError "Unable to instanciate, no linelist translation sheet"
    'Required listObjects fo translations
    Set tabl = New BetterArray
    tabl.Push TABSHAPE, TABRANGE, TABSMSG, TABLANG, TABLANGDICT
    LoExists sh, tabl

    RangeExists sh, RNGMAINLANG
    RangeExists sh, RNGMAINLANGCODE
    RangeExists sh, RNGLANGDICTLIST
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub LoExists(sh As Worksheet, ByVal loNames As BetterArray)
    Dim Lo As listObject
    Dim counter As Long
    Dim lName As String
    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError "Table " & lName & " does not exists in Worksheet " & sh.Name
    Next
End Sub

Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)

    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name

End Sub

Private Property Get IDesTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesTranslation_TransObject = TransObject(trans)
End Property

Private Sub IDesTranslation_TranslateDesigner(ByVal mainsh As Worksheet)
    TranslateDesigner mainsh
End Sub

'Get the range of the languages
Private Property Get IDesTranslation_LangListRng() As Range
    Set IDesTranslation_LangListRng = LangListRng()
End Property

Private Function IDesTranslation_TranslationMsg(ByVal msgCode As String) As String
    IDesTranslation_TranslationMsg = TranslationMsg(msgCode)
End Function
