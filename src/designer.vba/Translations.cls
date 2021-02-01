VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text

Const TRANSLATION_MAX_COUNT = 10000

Private mItems As Dictionary

Private Sub Class_Initialize()
    Set mItems = New Dictionary
End Sub
 
Private Sub Class_Terminate()
    Set mItems = Nothing
End Sub

Property Get Items() As Dictionary
    Set Items = mItems
End Property

Sub AddTranslationFromRange(oRange As Range)
    Dim oTranslation As Translation
    Dim oCell As Range
    
    For Each oCell In oRange.Cells
        Set oTranslation = New Translation
        oTranslation.EN = oCell.value
        AddTranslation oTranslation
    Next
End Sub

Sub AddTranslation(oTranslation As Translation)
    If oTranslation.EN <> "" And Not TranslationExists(oTranslation.EN) Then
        mItems.add oTranslation.EN, oTranslation
    End If
End Sub

Public Function TranslationExists(ByVal key As String) As Boolean
    TranslationExists = mItems.exists(key)
End Function

Public Sub FillFromRange(oStartCell As Range)

    Dim oTranslation As Translation
    Dim Ni As Long
    
    mItems.RemoveAll
    
    For Ni = 1 To TRANSLATION_MAX_COUNT
        If Not isEmpty(oStartCell.offSet(Ni, 0).value) Then
            Set oTranslation = New Translation
            oTranslation.EN = oStartCell.offSet(Ni, 0).value
            oTranslation.FR = oStartCell.offSet(Ni, 1).value
            oTranslation.SP = oStartCell.offSet(Ni, 2).value
            AddTranslation oTranslation
        End If
    Next
End Sub

Public Sub FillRange(oStartCell As Range)

    Dim oTranslation As Translation
    Dim Ni As Integer
    
    oStartCell.offSet(0, 0).value = "EN"
    oStartCell.offSet(0, 1).value = "FR"
    oStartCell.offSet(0, 2).value = "SP"

    Set mItems = mItems.SortByKey
    
    For Ni = 0 To mItems.Count - 1
        Set oTranslation = mItems.Items(Ni)
        oStartCell.offSet(Ni + 1, 0).value = oTranslation.EN
        oStartCell.offSet(Ni + 1, 1).value = oTranslation.FR
        oStartCell.offSet(Ni + 1, 2).value = oTranslation.SP
    Next
End Sub
