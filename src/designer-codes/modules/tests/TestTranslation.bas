Attribute VB_Name = "TestTranslation"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private TransObject As ITranslation

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set TransObject = Nothing
End Sub

'This method runs before every test in the module..

'@TestInitialize
Private Sub TestInitialize()

End Sub

'@TestMethod
Private Sub TestTranslation()
    On Error GoTo Fail
    Dim formVal As String

    Dim Lo As ListObject
    Set Lo = ThisWorkbook.Worksheets("LinelistTranslation").ListObjects("T_TradLLMsg")
    Set TransObject = Translation.Create(Lo, "FRA")

    Assert.IsTrue (TransObject.TranslatedValue("MSG_Day") = "Jour"), "Bad translated value"
    Assert.IsTrue (TransObject.TranslatedValue("www&!") = "www&!"), "unfound translated value found"
    formVal = TransObject.TranslatedValue("IF(" & chr(34) & "MSG_Day" & chr(34) & ", " & chr(34) & "MSG_Year" & chr(34) & ")", containsFormula:=True)

    Assert.IsTrue (formVal = "IF(" & chr(34) & "Jour" & chr(34) & ", " & chr(34) & "Ann�e" & chr(34) & ")"), "Bad translated formula : obtained " & formVal
    formVal = TransObject.TranslatedValue("IF(" & chr(34) & "MSG_Day" & chr(34), containsFormula:=True)
    Assert.IsTrue (formVal = "IF(" & chr(34) & "Jour" & chr(34)), "Bad translated formula : obtained " & formVal


    Exit Sub
Fail:
    Assert.Fail "Translation failed: #" & Err.Number & " : " & Err.Description
End Sub
