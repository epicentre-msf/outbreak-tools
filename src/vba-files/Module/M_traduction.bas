Attribute VB_Name = "M_traduction"
Option Explicit

Const C_iColFR As Integer = 2
Const C_iColEn As Integer = 3

Sub Translation()

Dim oShape As Object
Dim i As Integer    'cpt ligne Translation
Dim iColLangue As Byte
Dim T_data
Dim D_data As New Scripting.Dictionary
Dim sFont As String
Dim shpVisible As Boolean

Application.ScreenUpdating = False

Select Case [RNG_ChoixLangue1].value
Case "English"
    iColLangue = 1
Case "Français"
    iColLangue = 2
Case Else
    iColLangue = 1
End Select

With Sheets("MAIN")
    'pour les shapes
    T_data = [T_tradShape]
    i = 1
    While i <= UBound(T_data)
        D_data.Add T_data(i, 1), i
        i = i + 1
    Wend
    
    Set oShape = .Shapes
    For Each oShape In .Shapes
        If D_data.Exists(oShape.Name) Then
            If Not oShape.Visible Then  'pour forcer la Translation des shapes invisibles
                oShape.Visible = True
                shpVisible = True
            End If
            .Shapes(oShape.Name).Select
            sFont = Selection.Characters(1, 1).Font.Name
            Selection.Characters.Text = T_data(D_data(oShape.Name), iColLangue + 1)
            Selection.Characters.Font.Name = "calibri"
            Selection.Characters(1, 1).Font.Name = sFont
            If shpVisible Then
                oShape.Visible = False
                shpVisible = False
            End If
        End If
    Next
    Set oShape = Nothing
    Set D_data = Nothing
    ReDim T_data(0)

    'pour les range
    T_data = [T_tradRange]
    i = 1
    On Error Resume Next
    While i <= UBound(T_data)
        .Range(T_data(i, 1)).value = T_data(i, iColLangue + 1)
        i = i + 1
    Wend
    On Error GoTo 0

    .Range("a1").Select
    .Range("RNG_msg").value = TranslateMsg("MSG_Traduit")
End With

Application.ScreenUpdating = True

End Sub

Function TranslateMsg(sIDMsg As String) As String

Dim i As Integer
Dim T_data
Dim D_data As New Scripting.Dictionary
Dim iColLangue As Byte

TranslateMsg = ""

Select Case ThisWorkbook.Sheets("Main").Range("RNG_ChoixLangue1").value
Case "English"
    iColLangue = 1
Case "Français"
    iColLangue = 2
Case Else
    iColLangue = 1
End Select

'Set D_data = CreateObject("scripting.dictionary")
D_data.RemoveAll
T_data = ThisWorkbook.Sheets("translation").[T_tradMsg]
i = 1
While i <= UBound(T_data, 1)
    D_data.Add T_data(i, 1), i
    i = i + 1
Wend

If D_data.Exists(sIDMsg) Then
    TranslateMsg = T_data(D_data(sIDMsg), iColLangue + 1)
End If
ReDim T_data(0)
Set D_data = Nothing

End Function

Sub TranslateForm(sNameForm As String)

Dim i As Integer
Dim T_data
Dim D_data As Scripting.Dictionary

T_data = ThisWorkbook.Sheets("translation").[T_tradForm]

i = 1
While i <= UBound(T_data, 1) And sNameForm <> T_data(0, i)
    i = i + 1
Wend

If sNameForm = T_data(0, i) Then
    While sNameForm = T_data(0, i)

        ThisWorkbook.VBProject.VBComponents(sNameForm).Controls(T_data(1, i)).value = T_data(2, i)
        i = i + 1
    Wend
End If

End Sub
