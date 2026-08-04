VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "What the Metadata sheet of a file being imported says about it"

'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("What the Metadata sheet of a file being imported says about it")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ProcedureNotUsed

Option Explicit

'@class ImportMetadata
'@description
'Reads the variable and value block of the Metadata sheet of a workbook being
'imported, once, and answers every question the import has about that file.
'LoadFrom reads the sheet, Exists says whether the block is there, Tag answers
'one tag by name, and Language, ExportType, IsMigrationExport, EpiWeekStart and
'UpdatesOnImport answer the named facts. ExportOption reads one option of the
'options line, CarriesVariableNames says whether the file holds variable names,
'and HeaderRow says which row the header sits on.
'
'ONE READ, AND NO Range.Find
'-------------------------------------------------------------------------------
'The block is read into two arrays and every tag is looked up in memory. Four
'places used to build a DataSheet over this sheet and run a Range.Find with
'LookIn, SearchOrder and SearchDirection left unset. Excel fills those in from
'the last Find of the session, including one a user ran by hand.
'
'THE OPTIONS LINE
'-------------------------------------------------------------------------------
'`export_options` holds the options the export was written with, as
'`name: value` pairs joined with a pipe. `ExportOption` splits it back apart.
'The two pairs that decide whether the file can be read at all are `variable
'names` and `header row`: a file whose header row carries labels alone holds no
'variable names anywhere, and a file carrying both has its variable names on
'row 2.
'
'A file written before the options line existed carries none of it. Every
'answer then falls back to the shape every export used to have, variable names
'on row 1, which is the shape of a migration export and of the default custom
'export.
'@depends DataSheet, BetterArray, Checking
'@version 1.0 (2026-08-01)


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "ImportMetadata"

Private Const METADATASHEET As String = "Metadata"
Private Const VARIABLECOLUMN As String = "variable"

' Tag names read off the Metadata sheet
Private Const TAG_LANGUAGE As String = "language"
Private Const TAG_EXPORT_TYPE As String = "export_type"
Private Const TAG_EPIWEEK As String = "epiweekstart"
Private Const TAG_OPTIONS As String = "export_options"

' The value a tag carries when the import is asked to act on it
Private Const UPDATE_ON_IMPORT As String = "update on import"

' The export type a migration file carries
Private Const TYPE_MIGRATION As String = "export for migrations"

' Separators of the export options line
Private Const OPTION_SEPARATOR As String = "|"
Private Const OPTION_PAIR_SEPARATOR As String = ":"


'@section State
'===============================================================================

Private Type TImportMetadata
    names As BetterArray
    values As BetterArray
    optionNames As BetterArray
    optionValues As BetterArray
    exists As Boolean
    optionsRead As Boolean
    sealed As Boolean
End Type

Private this As TImportMetadata


'@section Factory
'===============================================================================

'@label:create
'@sub-title Read the Metadata sheet of a workbook being imported
'@details
'A workbook with no Metadata sheet gives an instance whose Exists answers False
'and whose every tag answers an empty string. That is a real case - a file
'exported with the metadata sheets switched off carries none - so it is
'answered rather than raised.
'@param impwb Workbook. The opened import workbook.
'@return ImportMetadata. A ready instance.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Function Create(ByVal impwb As Workbook) As ImportMetadata

    Dim newMeta As ImportMetadata

    If impwb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ImportMetadata requires the workbook being imported"
    End If

    Set newMeta = New ImportMetadata
    newMeta.LoadFrom impwb
    newMeta.Seal
    Set Create = newMeta
End Function

'@label:load-from
'@sub-title Read the variable/value block into memory
'@details
'Public because the factory calls it on the instance it just made. A Private
'procedure cannot be reached by dotted access in VBA, so the factory could not
'fill the object at all.
'@param impwb Workbook. The opened import workbook.
Public Sub LoadFrom(ByVal impwb As Workbook)

    Dim metaSh As Worksheet
    Dim metaTab As DataSheet
    Dim varRng As Range

    GuardNotSealed "LoadFrom"

    Set this.names = New BetterArray
    Set this.values = New BetterArray
    this.names.LowerBound = 1
    this.values.LowerBound = 1

    Set metaSh = Nothing
    On Error Resume Next
    Set metaSh = impwb.Worksheets(METADATASHEET)
    On Error GoTo 0
    If metaSh Is Nothing Then Exit Sub

    Set metaTab = DataSheet.Create(metaSh, 1, 1)

    ' A worksheet named Metadata with no `variable` header is not one of ours.
    ' DataRange raises on the missing column, and that is answered as a file
    ' carrying no metadata rather than as a failed import.
    Set varRng = Nothing
    On Error Resume Next
    Set varRng = metaTab.DataRange(VARIABLECOLUMN)
    On Error GoTo 0
    If varRng Is Nothing Then Exit Sub

    this.names.FromExcelRange varRng, DetectLastRow:=False, DetectLastColumn:=False
    this.values.FromExcelRange varRng.Offset(, 1), _
                               DetectLastRow:=False, DetectLastColumn:=False

    this.exists = True
End Sub

'@label:seal
'@prop-title Prevent further changes to setup-only members
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only member is used after sealing
'@param memberName String. The name of the member being guarded.
Private Sub GuardNotSealed(ByVal memberName As String)
    If this.sealed Then _
        ThrowError ProjectError.ElementShouldNotExists, _
                   memberName & " is used at creation only"
End Sub


'@section What the file says
'===============================================================================

'@label:exists
'@prop-title Whether the file carries a Metadata sheet at all
'@return Boolean. True when the sheet was found and read.
'@export
Public Property Get Exists() As Boolean
    Exists = this.exists
End Property

'@label:tag
'@prop-title The value of one tag of the Metadata sheet
'@details
'The name is matched without regard to case, the way the Range.Find this
'replaced did. A tag the file does not carry answers an empty string.
'@param tagName String. The tag to read.
'@return String. What the file holds for it, empty when it holds nothing.
'@export
Public Property Get Tag(ByVal tagName As String) As String

    Dim counter As Long

    If Not this.exists Then Exit Property
    If LenB(tagName) = 0 Then Exit Property

    For counter = this.names.LowerBound To this.names.UpperBound
        If StrComp(CStr(this.names.Item(counter)), tagName, vbTextCompare) = 0 Then
            Tag = CStr(this.values.Item(counter))
            Exit Property
        End If
    Next counter
End Property

'@label:language
'@prop-title The language the file was written in
'@return String. The language code, empty when the file names none.
'@export
Public Property Get Language() As String
    Language = Tag(TAG_LANGUAGE)
End Property

'@label:export-type
'@prop-title What kind of export the file is
'@return String. "export for migrations" or "custom export: export N", empty
'  when the file names neither.
'@export
Public Property Get ExportType() As String
    ExportType = Tag(TAG_EXPORT_TYPE)
End Property

'@label:is-migration-export
'@prop-title Whether the file is an export for migration
'@return Boolean. True when the export type names a migration.
'@export
Public Property Get IsMigrationExport() As Boolean
    IsMigrationExport = (StrComp(ExportType, TYPE_MIGRATION, vbTextCompare) = 0)
End Property

'@label:epi-week-start
'@prop-title The epiweek start day the file was written with
'@return String. The day number, empty when the file carries none.
'@export
Public Property Get EpiWeekStart() As String
    EpiWeekStart = Tag(TAG_EPIWEEK)
End Property

'@label:updates-on-import
'@prop-title Whether one tag asks the import to act on it
'@details
'`show_hide` and `editable_labels` are the two tags written this way. Both hold
'"update on import" or "no update on import".
'@param tagName String. The tag to read.
'@return Boolean. True when the tag asks for an update.
'@export
Public Property Get UpdatesOnImport(ByVal tagName As String) As Boolean
    UpdatesOnImport = (StrComp(Tag(tagName), UPDATE_ON_IMPORT, vbTextCompare) = 0)
End Property


'@section The export options line
'===============================================================================

'@label:export-option
'@prop-title One option of the line the export wrote
'@details
'The pairs are split on the pipe and each pair on the first colon, so a value
'holding a colon of its own survives. An option the file does not carry answers
'an empty string, and the properties below turn that into the shape every export
'had before the line existed.
'@param optionName String. The option to read.
'@return String. Its value, empty when the file carries no such option.
'@export
Public Property Get ExportOption(ByVal optionName As String) As String

    Dim counter As Long

    ReadOptions

    If this.optionNames.Length = 0 Then Exit Property
    If LenB(optionName) = 0 Then Exit Property

    For counter = this.optionNames.LowerBound To this.optionNames.UpperBound
        If StrComp(CStr(this.optionNames.Item(counter)), optionName, vbTextCompare) = 0 Then
            ExportOption = CStr(this.optionValues.Item(counter))
            Exit Property
        End If
    Next counter
End Property

'@label:carries-variable-names
'@prop-title Whether the data sheets of the file name their columns
'@details
'A file whose header row holds the labels a user reads carries no variable name
'anywhere, so no column of it can be matched to this linelist. A file written
'before the options line existed is taken as carrying them, which is what every
'export used to do.
'@return Boolean. True when the columns can be matched.
'@export
Public Property Get CarriesVariableNames() As Boolean

    Dim answer As String

    answer = ExportOption("variable names")
    If LenB(answer) = 0 Then
        CarriesVariableNames = True
        Exit Property
    End If

    CarriesVariableNames = (StrComp(answer, "yes", vbTextCompare) = 0)
End Property

'@label:header-row
'@prop-title The row of a data sheet that holds the variable names
'@details
'Row 2 for a file written with labels above the names, row 1 for every other
'shape and for any file written before the options line existed.
'@return Long. The 1-based row index.
'@export
Public Property Get HeaderRow() As Long

    Dim answer As String

    HeaderRow = 1

    answer = ExportOption("header row")
    If LenB(answer) = 0 Then Exit Property

    On Error Resume Next
    HeaderRow = CLng(answer)
    On Error GoTo 0

    If HeaderRow < 1 Then HeaderRow = 1
End Property

'@label:read-options
'@sub-title Split the options line into names and values, once
Private Sub ReadOptions()

    Dim pairs As Variant
    Dim onePair As String
    Dim cutAt As Long
    Dim counter As Long
    Dim optionsLine As String

    If this.optionsRead Then Exit Sub
    this.optionsRead = True

    Set this.optionNames = New BetterArray
    Set this.optionValues = New BetterArray
    this.optionNames.LowerBound = 1
    this.optionValues.LowerBound = 1

    optionsLine = Tag(TAG_OPTIONS)
    If LenB(optionsLine) = 0 Then Exit Sub

    pairs = Split(optionsLine, OPTION_SEPARATOR)

    For counter = LBound(pairs) To UBound(pairs)
        onePair = Trim$(CStr(pairs(counter)))
        cutAt = InStr(1, onePair, OPTION_PAIR_SEPARATOR)
        If cutAt > 1 Then
            this.optionNames.Push Trim$(Left$(onePair, cutAt - 1))
            this.optionValues.Push Trim$(Mid$(onePair, cutAt + 1))
        End If
    Next counter
End Sub


'@section Private - Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a class level error
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong, for a person reading it.
'@throws ProjectError Always raises.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
