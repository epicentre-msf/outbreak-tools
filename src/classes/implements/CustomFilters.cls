VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "CustomFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Create and Manage custom filters in a linelist"

'@Folder("Linelist")
'@ModuleDescription("Create and Manage custom filters in a linelist")

Option Explicit
Implements ICustomFilters

Private Const CSTABLE As String = "csfilter__" 
Private TABINFOS As BetterArray
Private this As TCSFilt

'This will contain the output/input worksheet, the translations
'and the F_Filters Form (prefer to adopt Object as name.)

Private Type TCSFilt
  sh As Worksheet
  trads As ITranslation
  fobj As Object
End Type

'Exposed methods
Private Function Create(ByVal sh As Worksheet, _ 
                        ByVal trads As ITranslation, _ 
                        ByVal fObj As Object) As ICustomFilters

  CheckRequirements sh, trads, formObj

  On Error Resume Next
  sh.ListObjects(1).DataBodyRange.Delete
  On Error GoTo 0

  With New CustomFilters
    Set .Wksh = sh
    Set .TransObject = trads
    Set .FormObject = fObj
    Set Create = .Self()
  End With

End Function

Public Property Get Self() As ICustomFilters
  Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Proprety Set TransObject(ByVal trads As ITranslation)
  Set this.trads = trads
End Property

Public Property Get TransObject() As ITranslation
  Set TransObject = this.trads
End Property

Private Function InitializeTable
  If TABINFOS Is Nothing Then
    Set TABINFOS = New BetterArray
    TABINFOS.LowerBound = 1
  End If
End Function


'Load the current form.
Private Sub Load(ByVal formObj As Object)
  
  Dim rng As Range
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim listObj As Object


  Set sh = Wksh()
  Set Lo = sh.ListObjects(1)
  
  If formObj Is Nothing Then ThrowError "Unable to load: Please provide The Filter form"
  If TypeName(formObj= <> "F_Filters" Then ThrowError "Unable to load: Please provide the form F_Filters"
  If Lo.DataBodyRange Is Nothing Then Exit Sub

  Set rng = Lo.DataBodyRange
  
  'avoid the index
  Set rng = sh.Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 2))
  Set listObj = formObj.LST_FiltersList
  
  listObj.BoundColumn = 2
  listObj.ColumnCount = 3
  listObj.List = rng.Value

  formObj.Show
End Sub


'Rename a filter
Private Sub RenameFilter(index As )
  
  Dim Lo As ListObject
  Dim trads As ITranslation
  Dim sh As Worksheet
  Dim lstFiltersName As BetterArray
  

  Set trads = TransObject()
  Set sh = Wksh()
  Set Lo = sh.ListObjects(1)
  Set lstFiltersName = New BetterArray

  Set rng = Lo.ListColumns(2).Range


End Sub


Private Sub SaveFilter(ByVal filterName As String)




End Filter


Private Sub CheckRequirements(ByVal sh As Worksheet)
  
  If sh Is Nothing Then _ 
    ThrowError "Unable to instanciate, Please provide a worksheet"
  
  If sh.Name <> "custom_filters__" Then _ 
    ThrowError "The provided worksheet is not custom_filters__"
  
  If Not LoExists(sh, CSTABLE) Then _ 
    ThrowError "Unable to find csfilter__ listobject on custom filter worksheet"

End Sub

'Test if a listobject exists
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean

  Dim Lo As ListObject

  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0

  LoExists = Not (Lo Is Nothing)
End Function

'Throw Error
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
