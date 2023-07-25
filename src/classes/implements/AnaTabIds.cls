VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "AnaTabIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Table Ids for export of analysis"

'@Folder("Linelist")
'@ModuleDescription("Table Ids For export of analysis")

Option Explicit
Implements IAnaTabIds

Private Type TAnaIds
  sh As Worksheet
End Type

Private this As TAnaIds
Private Const CLASSNAME As String = "AnaTabIds"
'ListObjects of the table Ids sheet
Private Const TABLISTGS As String = "GS_analysis_ids"
Private Const TABLISTUA As String = "UA_analysis_ids"
Private Const TABLISTBA As String = "BA_analysis_ids"
Private Const TABLISTTS As String = "TS_analysis_ids"
Private Const TABLISTSP As String = "SP_analysis_ids"
Private Const TABLISTSPTEMP As String = "SPTEMP_analysis_ids"
'Ranges That should be present
Private Const RNGUASHEET As String = "RNG_SheetUAName"
Private Const RNGTSSHEET As String = "RNG_SheetTSName"
Private Const RNGSPSHEET As String = "RNG_SheetSPName"
Private Const RNGSPTEMPSHEET As String = "RNG_SheetSPTempName"

'Exposed methods

Public Function Create(ByVal idsh As Worksheet) As IAnaTabIds

  CheckRequirements idsh

  With New AnaTabIds
    Set .Wksh = idsh
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As IAnaTabIds
  Set Self = Me
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

'Get univariate, time series or spatial sheets or spatio temporal sheets
'analysisName can take up to 4 values: ua for univariate, ts for time series
'sp for spatial and sptemp for spatio-temporal.
Private Property Get AnalysisSheet(ByVal analysisName As String) As Worksheet

  Dim rngName As String
  Dim currwb As Workbook
  Dim sh As Worksheet

  Select Case analysisName
  Case "UA"
    rngName = RNGUASHEET
  Case "TS"
    rngName = RNGTSSHEET
  Case "SP"
    rngName = RNGSPSHEET
  Case "SPT"
    rngName = RNGSPTEMPSHEET
  End Select
  
  Set sh = Wksh()
  Set currwb = sh.Parent
  Set AnalysisSheet = currwb.Worksheets(rngName)
End Property

'Add a table Id to a Worksheet
Private Sub AddTable(ByVal specs As ITableSpecs)

  Dim tabId As String
  Dim tabType As Byte
  Dim loName As String
  Dim Lo As ListObject
  Dim sh As Worksheet
  Dim rng As Range

  Set sh = Wksh()
  tabType = specs.TableType()

  Select Case tabType
  Case TypeGlobalSummary
    loName = TABLISTGS
  Case TypeUnivariate
    loName = TABLISTUA
  Case TypeBivariate
    loName = TABLISTBA
  Case TypeTimeSeries
    loName = TABLISTTS
  Case TypeSpatial
    loName = TABLISTSP
  Case TypeSpatioTemporal
    loName = TABLISTSPTEMP
  End Select

  Set Lo = sh.ListObjects(loName)
  tabId = tabObj.TableId()
  'Add the value of the id at the end of ListObject and resize
  If Lo.Range.Cells(2, 1).Value = vbNullString Then
     Lo.Range.Cells(2, 1).Value = tabId
  Else
    Lo.Range.Cells(Lo.Range.Rows.Count + 1, 1).Value = tabId
    Set rng = sh.Range(Lo.Range.Cells(1, 1), _ 
                       Lo.Range.Cells(Lo.Range.Rows.Count + 1, 1))
    Lo.Resize rng
  End If
End Sub

Private Sub TransferRange(ByVal rngName As String, ByVal tag As String, _ 
                          ByVal outsh As Worksheet)
  

End Sub

'Export a table to a worksheet
Private Sub ExportTable(ByVal tabId As String, ByVal outsh As String)

  'I know the type of the table using the id, I can also get the corresponding
  'analysis worksheet. I just need to export the correct Ranges to the output
  'worksheet

  Dim anash As Worksheet
  Dim tag As String 'Know which sheet
  'Input and Output ranges
  Dim inRng As Range
  Dim outRng As Range
  Dim namesArray As BetterArray

  tag = Split(tabId, "_")(0)
  Set anash = AnalysisSheet(tag)
  
  If Not RangeExists(outsh, "InfoFiltered") Then
    'Get the information on filtered/or not filtered data
    outsh.Cells(1, 2).Name = "InfoFiltered"
    outsh.Cells(1, 2).Value = anash.Cells(2, 5).Value 
  End If

  Select Case tag
  
  Case "GS"
  'Global summary Ranges

    'Title for Global summary
    If Not RangeExists(outsh, "ROWSGS_SET") Then
      On Error Resume Next
      Set inRng = anash.Range("ROWSGS_SET")
      Set outRng = outsh.Range(inRng.Address)
      outRng.Value = inRng.Value
      outRng.Name = "ROWSGS_SET"
      outRng.Font.Color = 
      outRng.Font.Size = 13
      outRng.Font.Bold = True
      On Error GoTo 0
    End If

    'Add the other Ranges
    'Start column
    If RangeExists(anash, "STARTCOL_" & tabId) Then
       Set inRng = anash.Range("STARTCOL_" & tabId)
       Set outRng = outsh.Range(inRng.Address)
       outRng.Value = inRng.Value
       outRng.Font.Bold = inRng.Font.Bold
    End If

    'StartRow
    If RangeExists(anash, "STARTROW_"& tabId) Then
      Set inRng = anash.Range("STARTROW_" & tabId)
      Set outRng = outsh.Range(inRng.Address)
      outRng.Value = 

    End If

  Case "UA"


  Case "BA"

  Case "TS"

  Case "SP"

  Case "SPT"

  End Select
  
  'For univariate, time series or bivariate


End Sub




'Check Requirements and Helpers functions

'Test if a listobject exists
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
  Dim Lo As ListObject

  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0
  LoExists = Not (Lo Is Nothing)
End Function

'Test if a range exits in a worksheet
Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
  Dim rng As Range
  On Error Resume Next
    Set rng = sh.Range(rngName)
  On Error GoTo 0
  RangeExists = Not (rng Is Nothing)
End Function

'Check Requirements of the class
Private Sub CheckRequirements(ByVal sh As Worksheet)
  Dim loArray As BetterArray
  Dim counter As Long
  Dim wb As Workbook
  Dim anash As Worksheet
  Dim actName As String

  
  If sh Is Nothing Then ThrowError "Unable to instanciate, Please provide a worksheet"
  Set loArray = New BetterArray
  loArray.Push TABLISTGS, TABLISTUA, TABLISTBA, TABLISTTS, TABLISTSP, TABLISTSPTEMP

  'Test if LisObjects are presents
  For counter = loArray.LowerBound To loArray.UpperBound
    actName = loArray.Item(counter)
    If Not LoExists(sh, actName) Then _ 
      ThrowError "Unable to instanciate AnaTabIds class. Table " & actName & " is missing."
  Next

  'Test if Ranges for sheetnames (ua, ts, sp, and sptemp) are presents
  loArray.Clear
  loArray.Push RNGUASHEET, RNGTSSHEET, RNGSPSHEET, RNGSPTEMPSHEET
  Set wb = sh.Parent

  For counter = loArray.LowerBound To loArray.UpperBound
    'actual Range name
    actName = loArray.Item(counter)
    Set anash = Nothing
    If Not RangeExists(sh, actName) Then
      ThrowError "Unable to instanciate AnaTabIds class. Range " & actName & " is missing."
    End If
  Next
End Sub

'Throw Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub