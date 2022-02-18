Attribute VB_Name = "Module1"
Sub tessub()
    Dim t As String

    t = Sheets("Main").Cells(1, 0).Address
    MsgBox t
End Sub

