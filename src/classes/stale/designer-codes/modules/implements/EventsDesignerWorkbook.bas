Attribute VB_Name = "EventsDesignerWorkbook"
Option Explicit

Private Sub Workbook_Open()
    Dim geo As ILLGeo
    Dim destradsh As Worksheet
    Dim geosh As Worksheet

    On Error GoTo ErrManage

    With ThisWorkbook
        Set destradsh = .Worksheets("DesignerTranslation")
        Set geosh = .Worksheets("Geo")
    End With

    destradsh.Range("LangDictList").ClearContents

    'Clear the geobase when opening the workbook
    Set geo = LLGeo.Create(geosh)
    geo.Clear
    geo.Translate rawNames:=True

ErrManage:
End Sub
