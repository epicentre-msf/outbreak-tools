Attribute VB_Name = "DesignerTests"
Option Explicit
Sub test()
    Dim sText As String
    
   sText = TranslateLLMsg("MSG_PathTooLong")
   
    Debug.Print sText
End Sub

'Test for the
