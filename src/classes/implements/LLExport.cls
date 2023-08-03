VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Export Object of the linelist"



'@PredeclaredId
'@Interface
'@Folder("Linelist")
'@ModuleDescription("Export Object of the linelist")


Option Explicit

Implements ILLExport

Private Type TLLExport
    expData As IDataSheet
    sh As Worksheet
End Type

Private this As TLLExport
Private SPECSHEADER As BetterArray
Private Const ACTIVESTATUS As String = "active"

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
                     "export pcodes", "export analysis", _
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

'Get a value from export as well as status of export
Private Property Get Value(ByVal colName As String, ByVal exportNumber As Integer) As String
    Dim expData As IDataSheet
    Dim returnedRng As Range
    Dim returnedVal As String

    returnedVal = vbNullString

    Set expData = Data()

    'It is possible that the column does not exists, bypass and continue, you
    'should return vbNullString in that case
    On Error Resume Next
    Set returnedRng = expData.DataRange(colName)
    On Error GoTo 0
    
    If Not (returnedRng Is Nothing) Then _
    returnedVal = returnedRng.Cells(exportNumber, 1).Value

    Value = returnedVal
End Property


'Test if an export is active
Private Property Get IsActive(ByVal exportNumber As Integer) As Boolean
    IsActive = (Value("status", exportNumber) = ACTIVESTATUS)
End Property

Private Function ClearChunk(ByVal chunk As String) As String

    'In the file name, replace forbidden characters with an underscore
    Dim newChunk As String

    
    newChunk = Replace(chunk, "<", "_")
    newChunk = Replace(newChunk, Chr(34), vbNullString)
    newChunk = Replace(newChunk, ">", "_")
    newChunk = Replace(newChunk, ":", "_")
    newChunk = Replace(newChunk, "|", "_")
    newChunk = Replace(newChunk, "?", "_")
    newChunk = Replace(newChunk, "/", "_")
    newChunk = Replace(newChunk, "\", "_")
    newChunk = Replace(newChunk, "*", "_")
    newChunk = Replace(newChunk, ".", "_")
    newChunk = Replace(newChunk, """", "_")
    newChunk = Replace(newChunk, "&", "_")
    

    ClearChunk = Application.WorksheetFunction.Trim(newChunk)
End Function


'Find a path of an export (using the dictionary)
Private Property Get ExportFileName(ByVal scope As Byte, _
                            ByVal dict As ILLdictionary, _
                            Byval pass As ILLPasswords) As String

    Dim vars As ILLVariables
    Dim PathData As BetterArray
    Dim fileName As String
    Dim outFile As String 'Final file name
    Dim pathCounter As Long
    Dim fun As WorksheetFunction
    Dim varSheetName As String
    Dim varSheetType As String
    Dim chunkValue As String
    Dim rangeValue As String
    Dim sh As Worksheet
    Dim vlistSh As Worksheet
    Dim wb As Workbook

    Set sh = Wksh()
    Set wb = sh.Parent

    If scope = ExportAll Then
        outFile = Replace(wb.Name, ".xlsb", vbNullString) & _
                  "_export_data_" & Format(Now, "yyyymmdd-HhNn")
    ElseIf scope = ExportAna Then
        outFile = Replace(wb.Name, ".xlsb", vbNullString) & _
                  "_analysis_outputs_" & Format(Now, "yyyymmdd-HhNn")
    Else
        Set PathData = New BetterArray
        Set fun = Application.WorksheetFunction
        Set vars = LLVariables.Create(dict)

        fileName = Value("file name", CInt(scope))

        PathData.Items = Split(fileName, "+")

        For pathCounter = 1 To PathData.UpperBound

            'Remove the other + signs in the path name
            chunkValue = fun.TRIM(Replace(PathData.Item(pathCounter), "+", vbNullString))
            rangeValue = vbNullString

            varSheetName = vars.Value(varName:=chunkValue, colName:="sheet name")
            varSheetType = vars.Value(varName:=chunkValue, colName:="sheet type")

            If (varSheetName <> vbNullString) And (varSheetType = "vlist1D") Then
                On Error Resume Next
                    Set vlistSh = wb.Worksheets(varSheetName)
                    rangeValue = vlistSh.Range(chunkValue).Value
                    rangeValue = ClearChunk(rangeValue)
                On Error GoTo 0
            End If

            If rangeValue <> vbNullString Then
                PathData.Item(pathCounter) = rangeValue
            Else 'unable to find the rangeValue
                PathData.Item(pathCounter) = ClearChunk(chunkValue)
            End If
        Next

        outFile = PathData.ToString(Separator:="__", OpeningDelimiter:=vbNullString,  _
                                    ClosingDelimiter:=vbNullString, QuoteStrings:=False) & _
                                    "__" & "v" & pass.Value("version") & "-" & _ 
                                    pass.Value("publickey") & "__" &  _
                                    Format(Now, "yyyymmdd-HhNn")
    End If

    ExportFileName = ClearChunk(outFile)
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
    Set rng = expData.DataRange("file name")
    trans.TranslateRange rng, containsFormula:=True
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


Private Sub ILLExport_Translate(ByVal trans As ITranslation)
    Translate trans
End Sub

Private Sub ILLExport_ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    ExportSpecs toWkb, Hide:=Hide
End Sub

Private Sub ILLExport_ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    ImportSpecs fromWksh, fromStartRow, fromStartcol
End Sub

Private Property Get ILLExport_Value(Byval colName As String,  _
                                     ByVal exportNumber As Integer) As String
    ILLExport_Value = Value(colName:=colName, exportNumber:=exportNumber)
End Property

Private Property Get ILLExport_IsActive(ByVal exportNumber As Integer) As Boolean
    ILLExport_IsActive = IsActive(exportNumber)
End Property

Private Property Get ILLExport_ExportFileName(ByVal scope As Byte, _
                                              ByVal dict As ILLdictionary, _
                                              Byval pass As ILLPasswords) As String
    ILLExport_ExportFileName = ExportFileName(scope, dict, pass)
End Property
