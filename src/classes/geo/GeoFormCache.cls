VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoFormCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "In-memory lists of the geo form"

'@PredeclaredId
'@Folder("Geo")
'@ModuleDescription("In-memory lists of the geo form")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class GeoFormCache
'@description
'Holds the four lists the geo form searches: the concatenated admin places,
'the concatenated health facilities, and the two historic lists. Each list is
'read from its workbook name on the first ask and held, so a keystroke in a
'search box of F_Geo scans memory. The concat lists are sorted once at that
'first read, so the hits of a search arrive in order with no sort per
'keystroke.
'
'THE DEFAULT INSTANCE CARRIES THE SHARED CACHE
'-------------------------------------------------------------------------------
'GeoModule.LoadGeo calls LoadFrom on the default instance on every form open,
'and FormLogicGeo reads the same instance on every search. LoadFrom drops the
'held lists, so each open reads each list at most once, and an open after a
'geobase import shows the new places. Create answers a separate instance for a
'caller that wants its own, which is what the tests build.
'
'AFTER A STATE RESET THE CACHE ANSWERS EMPTY LISTS
'-------------------------------------------------------------------------------
'Instance state dies on any unhandled error and any edit to the project, and
'the form outlives both. A read in that state answers an empty list, so a
'search shows nothing and raises nothing; the next LoadGeo fills the cache
'again. A read whose workbook name is missing raises, because that is a broken
'workbook the user has to hear about, and the callers that load the form hold
'a handler for it.
'
'THE ANSWERED LISTS ARE HELD, SO TREAT THEM AS READ ONLY
'-------------------------------------------------------------------------------
'ConcatList and HistoricList answer the held BetterArray itself. A copy per
'ask would pay the length of the geobase on every keystroke. A caller that
'wants to reorder or grow the answer builds its own copy.
'@depends BetterArray

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "GeoFormCache"

' The workbook names the lists are read from. They are written at build time,
' workbook scoped, pointing at table columns of the Geo worksheet.
Private Const CONCAT_GEO_NAME As String = "adm4_concat"
Private Const CONCAT_HF_NAME As String = "hf_concat"
Private Const HISTORIC_GEO_NAME As String = "histo_geo"
Private Const HISTORIC_HF_NAME As String = "histo_hf"

'@section State
'===============================================================================

Private Type TGeoFormCache
    sourceBook As Workbook
    concatGeo As BetterArray
    concatHF As BetterArray
    historicGeo As BetterArray
    historicHF As BetterArray
End Type

Private this As TGeoFormCache

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new GeoFormCache over one workbook
'@details
'Factory method that answers a separate instance loaded from the given
'workbook. Called on the predeclared instance. The shared cache the form
'reads is the predeclared instance itself, filled through LoadFrom.
'@param sourceBook Workbook. The workbook whose names are read.
'@return GeoFormCache. A loaded instance.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Function Create(ByVal sourceBook As Workbook) As GeoFormCache
    Dim instance As GeoFormCache

    Set instance = New GeoFormCache
    instance.LoadFrom sourceBook
    Set Create = instance
End Function

'@section Loading
'===============================================================================

'@label:load-from
'@sub-title Point the cache at a workbook and drop the held lists
'@details
'Records the workbook and empties the cache, so the next ask of each list
'reads it fresh. LoadGeo calls this once per form open, which is what keeps
'the form current after a geobase import.
'@param sourceBook Workbook. The workbook whose names are read.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
Public Sub LoadFrom(ByVal sourceBook As Workbook)
    If sourceBook Is Nothing Then _
        ThrowError ProjectError.ObjectNotInitialized, "Workbook not provided"

    Set this.sourceBook = sourceBook
    Refresh
End Sub

'@label:refresh
'@sub-title Drop the held lists and keep the workbook
'@details
'The next ask of each list reads it fresh from its workbook name.
Public Sub Refresh()
    Set this.concatGeo = Nothing
    Set this.concatHF = Nothing
    Set this.historicGeo = Nothing
    Set this.historicHF = Nothing
End Sub

'@label:clear
'@sub-title Drop the held lists and the workbook
'@details
'Puts the cache back in its unloaded state, where every ask answers an empty
'list.
Public Sub Clear()
    Refresh
    Set this.sourceBook = Nothing
End Sub

'@section Reading
'===============================================================================

'@label:concat-list
'@prop-title The concatenated places of one scope, sorted
'@details
'The admin scope reads "adm4_concat" and the facility scope reads "hf_concat".
'The list is read and sorted on the first ask and held until LoadFrom,
'Refresh or Clear drops it. An unloaded cache answers an empty list.
'@param scope Byte. GeoScopeAdmin (0) or GeoScopeHF (1).
'@return BetterArray. The held list. Treat it as read only.
'@throws ProjectError.InvalidArgument When the scope is unknown.
Public Function ConcatList(ByVal scope As Byte) As BetterArray
    GuardScope scope

    If scope = GeoScopeAdmin Then
        If this.concatGeo Is Nothing Then _
            Set this.concatGeo = ReadList(CONCAT_GEO_NAME, sorted:=True)
        Set ConcatList = this.concatGeo
    Else
        If this.concatHF Is Nothing Then _
            Set this.concatHF = ReadList(CONCAT_HF_NAME, sorted:=True)
        Set ConcatList = this.concatHF
    End If
End Function

'@label:historic-list
'@prop-title The historic places of one scope, in table order
'@details
'The admin scope reads "histo_geo" and the facility scope reads "histo_hf".
'The historic table shows in the order its rows hold, which is the order the
'user validated places in, so this list is held unsorted.
'@param scope Byte. GeoScopeAdmin (0) or GeoScopeHF (1).
'@return BetterArray. The held list. Treat it as read only.
'@throws ProjectError.InvalidArgument When the scope is unknown.
Public Function HistoricList(ByVal scope As Byte) As BetterArray
    GuardScope scope

    If scope = GeoScopeAdmin Then
        If this.historicGeo Is Nothing Then _
            Set this.historicGeo = ReadList(HISTORIC_GEO_NAME, sorted:=False)
        Set HistoricList = this.historicGeo
    Else
        If this.historicHF Is Nothing Then _
            Set this.historicHF = ReadList(HISTORIC_HF_NAME, sorted:=False)
        Set HistoricList = this.historicHF
    End If
End Function

'@section Implementation
'===============================================================================

'@sub-title Read one list from its workbook name
'@details
'One crossing per list per load. An unloaded cache answers an empty list; a
'loaded cache whose name is missing lets the resolve raise.
'@param rangeName String. The workbook scoped name to read.
'@param sorted Boolean. True to sort the list after the read.
'@return BetterArray. The list read, or an empty one with no workbook held.
Private Function ReadList(ByVal rangeName As String, _
                          ByVal sorted As Boolean) As BetterArray
    Dim valuesTable As BetterArray

    Set valuesTable = New BetterArray

    If Not this.sourceBook Is Nothing Then
        valuesTable.FromExcelRange this.sourceBook.Names(rangeName).RefersToRange
        If sorted Then valuesTable.Sort
    End If

    Set ReadList = valuesTable
End Function

'@sub-title Refuse a scope outside the two the form knows
'@details
'GeoScopeBoth exists on the enum and has no list in the form, so an unknown
'scope raises here. An unknown scope used to mean facility in silence
'elsewhere in the geo code, and this class refuses it the way the module now
'does.
'@param scope Byte. The scope asked for.
'@throws ProjectError.InvalidArgument When the scope is unknown.
Private Sub GuardScope(ByVal scope As Byte)
    If scope <> GeoScopeAdmin And scope <> GeoScopeHF Then _
        ThrowError ProjectError.InvalidArgument, "The cache knows no scope " & scope
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errNumber Long. The ProjectError error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
