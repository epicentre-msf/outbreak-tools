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
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation


Option Explicit
Implements ITranslationObject

Private Const CLASSNAME As String = "TranslationObject"

Private Type TTranslationObject
    transLo As listObject
    transLng As String
    hasChecks As Boolean
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
Public Function Create(ByVal translationLo As listObject, ByVal translationLang As String) As ITranslationObject

    'Test if the listobject is empty
    If translationLo Is Nothing Then ThrowErrorUnableInstanciate "Listobject is not defined"
    With New Translation
        Set .Table = translationLo
        .Language = translationLang
        Set Create = .Self
    End With
End Function

'@label:elements
'@section Translation elements

Public Property Get Self() As ITranslationObject
    Set Self = Me
End Property

Public Property Get Table() As listObject
    Set Table = this.transLo
End Property

Public Property Set Table(ByVal Lo As listObject)
    Set this.transLo = Lo
End Property

'@prop-title Translation language to translate elements
'@details
'The languages are headers of the list Object (except the first column)
'@return A String, one of the languages.
Public Property Get Language() As String
    Language = this.transLng
End Property


Public Property Let Language(ByVal transLng As String)
    this.transLng = transLng
End Property

'@prop-title Name of the translation Object
Private Property Get Name() As String
    Name = "Table: " & Table.Name & " - Language: " & Language
End Property

'@prop-title Check if the provided language exists
Private Property Get LanguageExists() As Boolean
    Dim lng As String
    Dim hRng As Range
    lng = Language
    Set hRng = Table.HeaderRowRange()
    LanguageExists = (Not hRng.Find(What:=lng, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Property



'@prop-title where is the language column
'@return A Long
Private Function LanguageIndex() As Long

    Dim lngIndex As Long
    Dim hRng As Range
    Dim lng As String

    lng = Language
    If LanguageExists() Then
        Set hRng = Table.HeaderRowRange
        lngIndex = hRng.Find(What:=lng, lookAt:=xlWhole, MatchCase:=True).Column -  hRng.Column + 1
    Else
        lngIndex = -1
    End If

    LanguageIndex = lngIndex
End Function

'@label:ValueExists
'@prop-title Check if a tag is present in the first column of the 
'@details
'Check if a tag is present in the first column of the worksheet
'@return A Boolean
'@export
Private Function ValueExists(ByVal Text As String) As Boolean
    Dim textExists As Boolean
    Dim keyRng As Range
    Dim Lo As listObject
    textExists = False
    
    Set Lo = Table()
    Set keyRng = Lo.ListColumns(1).Range

    If LanguageExists() Then
        textExists = (Not keyRng.Find(What:=Text, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If

    ValueExists = textExists
End Function

'@prop-title Line of a value
'@details
'In case a value is present, return the line index relative to the worksheet
'@return Long
Private Function ValueIndex(ByVal Text As String) As Long
    Dim Index As Long
    Dim keyRng As Range

    If ValueExists(Text) Then
        Set keyRng = Table.ListColumns(1).Range
        Index = keyRng.Find(What:=Text, lookAt:=xlWhole, MatchCase:=True).Row - keyRng.Row + 1
    Else
        Index = -1
    End If

    ValueIndex = Index
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
Private Function TranslatedValue(ByVal textTotranslate As String, _ 
                                 Optional ByVal containsFormula As Boolean = False) As String
    
    Dim cleanedText As String
    Dim transText As String
    Dim formTransText As String
    Dim formulaText As String
    Dim lngCol As Long
    Dim valIndex As Long
    Dim LoRng As Range
    Dim counter As Long
    Dim quoteStart As Long
    Dim chunkToTrans As String
    Dim repLists As BetterArray
    
    'Replace all the non breaking spaces with empty character
    cleanedText = Application.WorksheetFunction.SUBSTITUTE(textTotranslate, chr(160), " ")
    cleanedText = Application.WorksheetFunction.CLEAN(cleanedText)
    cleanedText = Application.WorksheetFunction.TRIM(cleanedText)

    If containsFormula Then
        
        formTransText = cleanedText
        formulaText = Replace(cleanedText, chr(34) & chr(34), vbNullString)
        Set repLists = New BetterArray

        'Remove empty characters from the formula.
        If InStr(1, formulaText, chr(34)) > 0 Then
            For counter = 1 To Len(formulaText)
                If Mid(formulaText, counter, 1) = chr(34) Then
                    If quoteStart = 0 Then
                        quoteStart = counter + 1
                    Else
                        chunkToTrans = Mid(formulaText, quoteStart, counter - quoteStart)
                        'Stock the list of elements already translated, and skip if it is translated already
                        If Not (repLists.Includes(chunkToTrans)) Then
                            transText = TranslatedValue(chunkToTrans)
                            formTransText = Replace(formTransText, chr(34) & chunkToTrans & chr(34), chr(34) & transText & chr(34))
                            repLists.Push chunkToTrans
                        End If
                        quoteStart = 0
                    End If
                End If
            Next
        End If

        TranslatedValue = formTransText
    Else
        If ValueExists(cleanedText) Then
            Set LoRng = Table.Range
            valIndex = ValueIndex(cleanedText)
            lngCol = LanguageIndex()
            transText = LoRng.Cells(valIndex, lngCol).Value
            TranslatedValue = transText
        Else
            TranslatedValue = cleanedText
        End If
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
    Dim counter As Long
    Dim EndRow As Long
  
    EndRow = colRng.Rows.Count
    'The range should be a column range, not a row range.
    For counter = 1 To EndRow
        colRng.Cells(counter, 1).Value = TranslatedValue(colRng.Cells(counter, 1).Value, containsFormula:=containsFormula)
    Next
End Sub

'@label:TranslateForm
'@sub-title Translate a Form
'@param UserFrm A Form Object
'@export
Private Sub TranslateForm(ByVal UserFrm As Object)
    Dim counter As Integer
    Dim formControl As Control

    'Be sure the object is a userform
    If Not (TypeOf UserFrm Is UserForm) Then
        LogInfo "Attempt to translate an object that is not a form", checkingNote
        ShowDebug
        Exit Sub
    End If

    For Each formControl In UserFrm.Controls
        If (TypeOf formControl Is MSForms.CommandButton) Or _
                                                         (TypeOf formControl Is MSForms.LABEL) Or _
                                                         (TypeOf formControl Is MSForms.OptionButton) Or _
                                                         (TypeOf formControl Is MSForms.Page) Or _
                                                         (TypeOf formControl Is MSForms.MultiPage) Or _
                                                         (TypeOf formControl Is MSForms.Frame) Or _
                                                         (TypeOf formControl Is MSForms.CheckBox) Then
            If TypeOf formControl Is MSForms.MultiPage Then
                For counter = 0 To formControl.Pages.Count - 1
                    If formControl.Name = "MultiPage1" Then
                        UserFrm.MultiPage1.Pages(counter).Caption = _
                                                                  TranslatedValue(UserFrm.MultiPage1.Pages(counter).Name)
                    End If
                    If formControl.Name = "MultiPage2" Then
                        UserFrm.MultiPage2.Pages(counter).Caption = _
                                                                  TranslatedValue(UserFrm.MultiPage2.Pages(counter).Name)
                    End If
                Next
            Else
                If Trim(formControl.Caption) <> vbNullString Then formControl.Caption = TranslatedValue(formControl.Name)
            End If
        End If
    Next
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
    internalChecks.Add Cstr(checkCounter), label, scope
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

'@sub-title Raise Error on instanciation
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'@section Interface

'@jump:elements
'Elements----------------------------------------------------------------------

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



