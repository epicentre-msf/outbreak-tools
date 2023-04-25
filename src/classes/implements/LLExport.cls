VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Export Object of the linelist")
'@IgnoreModule

Option Explicit

Implements ILLExport

Private Type TLLExport
    expData As IDataSheet
    sh As Worksheet
End Type

Private this As TLLExport
Private SPECSHEADER As BetterArray

Public Function Create(ByVal sh As Worksheet) As ILLExport

    'Set creation of datasheet
    Dim customData As IDataSheet
    Set customData = DataSheet.Create(sh, 1, 1)

    With New LLExport
        Set .Wksh = sh
        Set .Data = customData
        Set Create = .Self()
    End With

End Function

'Initialize the Arrays
Private Sub ArraysInitialize()
    Set SPECSHEADER = New BetterArray
    SPECSHEADER.Push "export number", "label button", "password", _
                     "status", "export metadata", "export translation", _
                     "file format", "file name", "export header"
End Sub

'Data for the exports
Public Property Get Data() As IDataSheet
    Set Data = this.expData
End Property

Public Property Set Data(ByVal customData As IDataSheet)
    Set this.expData = customData
End Property

Public Property Get Self() As ILLExport
    Set Self = Me
End Property

Public Property Set Wksh( ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'Export specifications to another workbook
Private Sub ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Dim expData As IDataSheet
    Set expData = Data()
    expData.export toWkb, Hide:=Hide
End Sub

Private Sub Translate(ByVal trans As ITranslation)
    Dim expData As IDataSheet
    Dim rng As Range
    Set expData = Data()
    Set rng = expData.DataRange("label button")
    trans.TranslateRange rng
End Sub

'Import from another worksheet
Private Sub ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Dim expData As IDataSheet
    Dim impExpData As IDataSheet
    Dim sh As Worksheet

    Set expData = Data()
    Set sh = Wksh()
    If SPECSHEADER Is Nothing Then ArraysInitialize
    'Here start row and start columns are just 1, 1
    SPECSHEADER.ToExcelRange sh.Cells(1, 1), TransposeValues:=True
    Set impExpData = DataSheet.Create(fromWksh, fromStartRow, fromStartCol)
    expData.Import impExpData
End Sub

Public Sub ILLExport_Translate(ByVal trans As ITranslation)
    Translate trans
End Sub

Public Sub ILLExport_ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    ExportSpecs toWkb, Hide:=Hide
End Sub

Public Sub ILLExport_ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    ImportSpecs fromWksh, fromStartRow, fromStartcol
End Sub

