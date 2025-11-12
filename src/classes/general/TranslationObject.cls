VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TranslationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Translate various objects from one language to another"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Translate various objects from one language to another")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation
'@depends BetterArray, Checking, IChecking, ProjectError


Option Explicit
Implements ITranslationObject

Private Const CLASSNAME As String = "TranslationObject"

Private Type TTranslationObject
    transLo As ListObject
    transLng As String
    hasChecks As Boolean
    languageColumnIndex As Long
End Type

Private this As TTranslationObject

'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

'@label:Create
'@prop-title Create a translation Object
'@details
' A translation object is a listObject with a specific language which is the
' language you want to translate elements to. 
' The first column of the listObject contains the tags for translation. In 
' case tags are not found during the translation process. The tag is returned
' as is without any translation done.
'@return A ITranslationObject
'@export
Public Function Create(ByVal translationLo As ListObject, ByVal translationLang As String) As ITranslationObject

    ValidateListObject translationLo, "Create"

    With New TranslationObject
        Set .Table = translationLo
        .Language = Trim$(translationLang)
        Set Create = .Self
    End With
End Function

'@label:elements
'@section Translation elements

Public Property Get Self() As ITranslationObject
    Set Self = Me
End Property

Public Property Get Table() As ListObject
    Set Table = this.transLo
End Property

Public Property Set Table(ByVal lo As ListObject)
    ValidateListObject lo, "Table"
    Set this.transLo = lo
    ResetCaches
End Property

'@prop-title Translation language to translate elements
'@details
'The languages are headers of the list Object (except the first column)
'@return A String, one of the languages.
Public Property Get Language() As String
    Language = this.transLng
End Property

'@prop-title Translation list of Languages
'@details Returns the language headers excluding the helper tag column.
Public Property Get LanguagesList() As BetterArray
    Dim headerRow As Range
    Dim languages As BetterArray
    Dim idx As Long
    Dim valueText As String

    Set languages = New BetterArray
    languages.LowerBound = 1

    Set headerRow = HeaderRowRange()
    If headerRow Is Nothing Then
        Set LanguagesList = languages
        Exit Property
    End If

    languages.FromExcelRange headerRow, DetectLastRow:=False, DetectLastColumn:=False

    For idx = languages.LowerBound To languages.UpperBound
        valueText = Trim$(CStr(languages.Item(idx)))
        languages.Item(idx) = valueText
    Next idx

    languages.Shift
    languages.LowerBound = 1
    languages.Filter vbNullString, Include:=False

    Set LanguagesList = languages
End Property


Public Property Let Language(ByVal transLng As String)
    this.transLng = Trim$(transLng)
    ResetCaches
End Property

'@label:Name
'@prop-title Name of the translation Object
Private Property Get Name() As String
    If this.transLo Is Nothing Then
        Name = CLASSNAME & " (uninitialised)"
    Else
        Name = "Table: " & this.transLo.Name & " - Language: " & Language
    End If
End Property

'@prop-title where is the language column
'@return A Long
Private Function LanguageIndex() As Long

    Dim headerRow As Range
    Dim found As Range
    Dim currentName As String

    If this.transLo Is Nothing Then Exit Function
    If LenB(this.transLng) = 0 Then Exit Function

    If this.languageColumnIndex > 0 Then
        On Error Resume Next
        currentName = this.transLo.ListColumns(this.languageColumnIndex).Name
        On Error GoTo 0
        If StrComp(currentName, this.transLng, vbBinaryCompare) <> 0 Then
            this.languageColumnIndex = 0
        End If
    End If

    If this.languageColumnIndex <= 0 Then
        this.languageColumnIndex = -1
        Set headerRow = HeaderRowRange()
        If headerRow Is Nothing Then
            LanguageIndex = this.languageColumnIndex
            Exit Function
        End If
        Set found = headerRow.Find(What:=this.transLng, LookAt:=xlWhole, MatchCase:=True)
        If Not found Is Nothing Then
            this.languageColumnIndex = found.Column - headerRow.Column + 1
        End If
    End If

    LanguageIndex = this.languageColumnIndex
End Function

'@label:ValueExists
'@prop-title Check if a tag is present in the first column of the 
'@details
'Check if a tag is present in the first column of the worksheet
'@return A Boolean
'@export
Private Function ValueExists(ByVal Text As String) As Boolean
    Dim translationCell As Range
    Dim sanitized As String

    Set translationCell = ResolveTranslationCell(Text)

    If translationCell Is Nothing Then
        sanitized = SanitizeText(Text)
        If StrComp(sanitized, Text, vbBinaryCompare) <> 0 Then
            Set translationCell = ResolveTranslationCell(sanitized)
        End If
    End If

    ValueExists = Not translationCell Is Nothing
End Function

'@label:TranslatedValue
'@prop-title Translated value of a translation Object
'@details
'In case here is a need to translate like formula, only characters present in quotes "
'are translated. Unquoted characters are not.
'@param textToTranslate A String
'@param containsFormula Optional, Boolean translate the text as as Formula. Default is False
'@return A String
'@export
Private Function TranslatedValue(ByVal textToTranslate As String, _ 
                                 Optional ByVal containsFormula As Boolean = False) As String

    Dim cleanedText As String
    Dim translationCell As Range
    Dim returnedValue As String

    cleanedText = SanitizeText(textToTranslate)

    If containsFormula Then
        TranslatedValue = TranslateFormulaText(cleanedText)
        Exit Function
    End If

    Set translationCell = ResolveTranslationCell(textToTranslate)

    If translationCell Is Nothing Then
        TranslatedValue = cleanedText
    Else
        returnedValue = SanitizeText(CStr(translationCell.Value))
        'If returned value is an empty string, return the text to translate
        TranslatedValue = IIF(LenB(returnedValue) = 0, cleanedText, returnedValue)
    End If
End Function

'@label:operate
'@section Operate with the outside world

'@label:TranslateRange
'@sub-title Translate a Range
'@param colRng A Range Object to translate.
'@param containsFormula Optional, Boolean. Translate each values in the range as formula. Default is False
'@export
Private Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    Dim targetCell As Range

    If colRng Is Nothing Then Exit Sub

    For Each targetCell In colRng.Columns(1).Cells
        targetCell.Value = TranslatedValue(targetCell.Value, containsFormula:=containsFormula)
    Next targetCell
End Sub

'@label:TranslateForm
'@sub-title Translate a Form
'@param UserFrm A Form Object
'@export
Private Sub TranslateForm(ByVal UserFrm As Object)
    Dim formControl As Control
    Dim controlType As String

    If Not SupportsUserFormContract(UserFrm) Then
        LogInfo "Attempt to translate an object of type '" & TypeName(UserFrm) & "' that is not a form", checkingNote
        Exit Sub
    End If

    For Each formControl In UserFrm.Controls
        controlType = TypeName(formControl)
        If StrComp(controlType, "MultiPage", vbTextCompare) = 0 Then
            TranslateMultiPageControl formControl
        ElseIf ShouldTranslateControl(controlType) Then
            If Trim$(CStr(formControl.Caption)) <> vbNullString Then
                formControl.Caption = TranslatedValue(formControl.Name)
            End If
        Else
            LogInfo "Attempt to translate a control that is not supported :'" & controlType & "'", checkingNote
        End If
    Next formControl
End Sub

'@label:checkings
'@section Checkings and Errors

'@sub-title Add elements to the checking object for future notes
'@param Label String, the message to add.
'@param scope Byte. The scope of the checking object.
'@export
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Translation " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new informations to the checking object
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:HasCheckings
'@prop-title Check if the object has checkings
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Get the cheking values
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

Private Sub ShowDebug()
    Debug.Print "Check the informations of Translation object " & Name()
End Sub

'@label:helpers
'@section Helpers

Private Function SanitizeText(ByVal rawValue As String) As String
    Dim cleaned As String

    cleaned = Application.WorksheetFunction.SUBSTITUTE(rawValue, Chr$(60), " ")
    cleaned = Application.WorksheetFunction.CLEAN(cleaned)
    cleaned = Application.WorksheetFunction.TRIM(cleaned)

    SanitizeText = cleaned
End Function

Private Function HeaderRowRange() As Range
    On Error Resume Next
    If Not this.transLo Is Nothing Then
        Set HeaderRowRange = this.transLo.HeaderRowRange
    End If
    On Error GoTo 0
End Function

Private Function KeysDataRange() As Range
    On Error Resume Next
    If Not this.transLo Is Nothing Then
        Set KeysDataRange = this.transLo.ListColumns(1).DataBodyRange
    End If
    On Error GoTo 0
End Function

Private Function ResolveTranslationCell(ByVal token As String) As Range
    Dim keyRange As Range
    Dim found As Range
    Dim bodyRange As Range
    Dim relativeRow As Long
    Dim languageColumn As Long

    If LenB(token) = 0 Then Exit Function
    If this.transLo Is Nothing Then Exit Function

    languageColumn = LanguageIndex()
    If languageColumn <= 0 Then Exit Function

    Set keyRange = KeysDataRange()
    If keyRange Is Nothing Then Exit Function

    Set found = keyRange.Find(What:=token, LookAt:=xlWhole, MatchCase:=True)
    If found Is Nothing Then Exit Function

    Set bodyRange = this.transLo.DataBodyRange
    If bodyRange Is Nothing Then Exit Function

    relativeRow = found.Row - keyRange.Row + 1
    Set ResolveTranslationCell = bodyRange.Cells(relativeRow, languageColumn)
End Function

Private Function TranslateFormulaText(ByVal formulaText As String) As String
    Dim strippedText As String
    Dim workingText As String
    Dim replacements As BetterArray
    Dim quoteStart As Long
    Dim idx As Long
    Dim chunk As String
    Dim translated As String
    Dim DOUBLE_QUOTE As String 'A constant

    DOUBLE_QUOTE = Chr$(34)

    workingText = formulaText
    strippedText = Replace(formulaText, DOUBLE_QUOTE & DOUBLE_QUOTE, vbNullString)

    If InStr(1, strippedText, DOUBLE_QUOTE, vbBinaryCompare) = 0 Then
        TranslateFormulaText = workingText
        Exit Function
    End If

    Set replacements = New BetterArray

    For idx = 1 To Len(strippedText)
        If Mid$(strippedText, idx, 1) = DOUBLE_QUOTE Then
            If quoteStart = 0 Then
                quoteStart = idx + 1
            Else
                chunk = Mid$(strippedText, quoteStart, idx - quoteStart)
                If LenB(chunk) > 0 Then
                    If Not replacements.Includes(chunk) Then
                        translated = TranslatedValue(chunk)
                        workingText = Replace(workingText, DOUBLE_QUOTE & chunk & DOUBLE_QUOTE, DOUBLE_QUOTE & translated & DOUBLE_QUOTE)
                        replacements.Push chunk
                    End If
                End If
                quoteStart = 0
            End If
        End If
    Next idx

    TranslateFormulaText = workingText
End Function

Private Function SupportsUserFormContract(ByVal candidate As Object) As Boolean
    Dim controls As Object

    If candidate Is Nothing Then Exit Function

    On Error Resume Next
    Set controls = candidate.Controls
    SupportsUserFormContract = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ShouldTranslateControl(ByVal controlType As String) As Boolean
    Select Case LCase$(controlType)
        Case "commandbutton", "label", "optionbutton", "page", "frame", "checkbox"
            ShouldTranslateControl = True
    End Select
End Function

Private Sub TranslateMultiPageControl(ByVal multiPageControl As Object)
    Dim pageIndex As Long
    Dim controlName As String

    If multiPageControl Is Nothing Then Exit Sub

    controlName = CStr(multiPageControl.Name)

    If controlName <> "MultiPage1" And controlName <> "MultiPage2" Then Exit Sub

    For pageIndex = 0 To multiPageControl.Pages.Count - 1
        multiPageControl.Pages(pageIndex).Caption = TranslatedValue(multiPageControl.Pages(pageIndex).Name)
    Next pageIndex
End Sub

Private Sub ValidateListObject(ByVal source As ListObject, ByVal sourceMember As String)
    If source Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "ListObject is not defined.", sourceMember
    End If
End Sub

Private Sub ResetCaches(Optional ByVal resetLanguageIndex As Boolean = True)
    If resetLanguageIndex Then
        this.languageColumnIndex = 0
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As ProjectError, ByVal message As String, Optional ByVal sourceMember As String = vbNullString)
    Dim sourceName As String

    sourceName = CLASSNAME
    If LenB(sourceMember) > 0 Then
        sourceName = sourceName & "." & sourceMember
    End If

    Err.Raise CLng(errNumber), sourceName, message
End Sub

'@section Interface

'@jump:elements
'Elements----------------------------------------------------------------------

'@jump:Name
Private Property Get ITranslationObject_Name() As String
    ITranslationObject_Name = Name()
End Property

'@jump:Language
Private Property Get ITranslationObject_Language() As String
    ITranslationObject_Language = Language
End Property

Private Property Get ITranslationObject_LanguagesList() As BetterArray
    Set ITranslationObject_LanguagesList = LanguagesList
End Property

'@jump:ValueExists
Private Property Get ITranslationObject_ValueExists(ByVal text As String) As Boolean
    ITranslationObject_ValueExists = ValueExists(text)
End Property

'@jump:TranslatedValue
Private Function ITranslationObject_TranslatedValue(ByVal Text As String, _ 
                                                   Optional ByVal containsFormula As Boolean = False) As String
    ITranslationObject_TranslatedValue = TranslatedValue(Text, containsFormula)
End Function

'@jump:operate
'Outside world -----------------------------------------------------------------

'@jump:TranslateRange
Private Sub ITranslationObject_TranslateRange(ByVal colRng As Range, _ 
                                             Optional ByVal containsFormula As Boolean = False)
    TranslateRange colRng:=colRng, containsFormula:=containsFormula
End Sub


'@jump:TranslateForm
Private Sub ITranslationObject_TranslateForm(ByVal UserFrm As Object)
    TranslateForm UserFrm
End Sub

'@jump:checkings
'checkings----------------------------------------------------------------------

'@jump:HasCheckings
Private Property Get ITranslationObject_HasCheckings() As Boolean
    ITranslationObject_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
Private Property Get ITranslationObject_CheckingValues() As Object
    Set ITranslationObject_CheckingValues = CheckingValues()
End Property



