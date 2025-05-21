VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("General")
'@ModuleDescription("Translation class for dictionary and the linelist")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ITranslation

Private Const CLASSNAME As String = "Translation"

Private Type TTranslation
    transLo As listObject
    transLng As String
End Type

Private this As TTranslation

'Instanciation

Public Function Create(ByVal translationLo As listObject, ByVal translationLang As String) As ITranslation

    'Test if the listobject is empty
    If translationLo Is Nothing Then ThrowErrorUnableInstanciate "Listobject is not defined"
    With New Translation
        Set .Table = translationLo
        .Language = translationLang
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ITranslation
    Set Self = Me
End Property

Public Property Get Table() As listObject
    Set Table = this.transLo
End Property

Public Property Set Table(ByVal Lo As listObject)
    Set this.transLo = Lo
End Property

Public Property Get Language() As String
    Language = this.transLng
End Property

Public Property Let Language(ByVal transLng As String)
    this.transLng = transLng
End Property

Private Property Get LanguageExists() As Boolean
    Dim lng As String
    Dim hRng As Range
    lng = Language
    Set hRng = Table.HeaderRowRange()
    LanguageExists = (Not hRng.Find(What:=lng, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Property

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

Private Function ValueExists(ByVal Text As String) As Boolean
    Dim textExists As Boolean
    Dim keyRng As Range
    Dim Lo As listObject
    textExists = False
    
    Set Lo = Table()
    Set keyRng = Lo.ListColumns(1).Range

    If LanguageExists() Then textExists = (Not keyRng.Find(What:=Text, lookAt:=xlWhole, MatchCase:=True) Is Nothing)

    ValueExists = textExists
End Function

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

Private Function TranslatedValue(ByVal textTotranslate As String, Optional ByVal containsFormula As Boolean = False) As String
    
    'Test if the language Exists
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

'Translate a Range
Private Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    Dim counter As Long
    Dim EndRow As Long
  
    EndRow = colRng.Rows.Count
    'The range should be a column range, not a row range.
    For counter = 1 To EndRow
        colRng.Cells(counter, 1).Value = TranslatedValue(colRng.Cells(counter, 1).Value, containsFormula:=containsFormula)
    Next
End Sub

'Translate a Form
Private Sub TranslateForm(ByVal UserFrm As Object)
    Dim counter As Integer
    Dim formControl As Control

    'Be sure the object is a userform
    If Not (TypeOf UserFrm Is UserForm) Then Exit Sub

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

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Interface
Public Sub ITranslation_TranslateForm(ByVal UserFrm As Object)
    TranslateForm UserFrm
End Sub

Public Sub ITranslation_TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    TranslateRange colRng:=colRng, containsFormula:=containsFormula
End Sub

Public Function ITranslation_TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
    ITranslation_TranslatedValue = TranslatedValue(Text, containsFormula)
End Function
