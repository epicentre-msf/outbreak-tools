Attribute VB_Name = "M_traduction"
Option Explicit

Const C_iColFR As Integer = 2
Const C_iColEn As Integer = 3

Sub Traduction()

Dim oShape As Object
Dim i As Integer    'cpt ligne traduction
Dim iColLangue As Byte
Dim T_data
Dim D_data As New Scripting.Dictionary
Dim sFont As String

Application.ScreenUpdating = False

Select Case [RNG_ChoixLangue1].Value
Case "English"
    iColLangue = 2
Case "Français"
    iColLangue = 1
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
            .Shapes(oShape.Name).Select
            sFont = Selection.Characters(1, 1).Font.Name
            Selection.Characters.Text = T_data(D_data(oShape.Name), iColLangue + 1)
            Selection.Characters.Font.Name = "calibri"
            Selection.Characters(1, 1).Font.Name = sFont
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
        .Range(T_data(i, 1)).Value = T_data(i, iColLangue + 1)
        i = i + 1
    Wend
    On Error GoTo 0

    .Range("a1").Select
End With

Application.ScreenUpdating = True

End Sub

Function TraduireMSG(sIDMsg As String) As String

Dim i As Integer
Dim T_data
Dim D_data As New Scripting.Dictionary
Dim iColLangue As Byte

TraduireMSG = ""

Select Case [RNG_ChoixLangue1].Value
Case "English"
    iColLangue = 2
Case "Français"
    iColLangue = 1
Case Else
    iColLangue = 1
End Select

T_data = [T_tradMsg]
i = 1
While i <= UBound(T_data, 1)
    D_data.Add T_data(i, 1), i
    i = i + 1
Wend

If D_data.Exists(sIDMsg) Then
    TraduireMSG = T_data(D_data(sIDMsg), iColLangue + 1)
End If
ReDim T_data(0)
Set D_data = Nothing


End Function
