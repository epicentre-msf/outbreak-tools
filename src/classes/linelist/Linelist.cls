VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist workbook facade: creates, populates and saves the output workbook"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Linelist workbook facade: creates, populates and saves the output workbook")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class Linelist
'@description
'Facade for the output linelist workbook. Prepare adds every worksheet the
'linelist needs, prepares the temporary sheets with their hidden names and
'ListObjects, formats the admin sheet and transfers the VBA code, and SaveLL
'writes the .xlsb the user receives. AddOutputSheet adds one sheet,
'SheetExists and SheetNames answer what is there, Wksh, PrintWksh and CRFWksh
'reach a data entry sheet and its two companions, and Dropdown and Pivots
'answer the two managers the class builds on first use.
'LinelistSpecs.Prepare has to have run first: it is what creates the output
'workbook and points the domain managers at it.
'
'The class uses VB_Creatable = True, so callers build one with
'`Set ll = Linelist.Create(specs)` over a `New Linelist`.
'
'SHEET NAMES AND THEIR SCOPE
'-------------------------------------------------------------------------------
'A data entry sheet can carry two companions, and each companion is the same
'name under a prefix: `print_` for the printable copy, `crf_` for the case
'report form. ScopedName is the one place that puts a prefix on a name and cuts
'the result to what Excel accepts, and AddOutputSheet, ResolveWksh and the
'SheetExists calls all go through it. A sheet created under a shortened name is
'found again under the same shortened name.
'
'A name the cut actually shortens is filed as a report entry, once per name, so
'the user learns which sheet came out under a different name.
'
'WHAT THIS CLASS REPORTS
'-------------------------------------------------------------------------------
'Prepare files its own entries and folds in the ones CodeTransfer filed, so
'HasCheckings and CheckingValues answer over both. clickGenerate flushes them
'into the generation report right after Prepare returns.
'@depends LinelistSpecs, LLdictionary, LLSheets, TranslationObject,
'  LLFormat, DropdownLists, CustomPivotTable, TemporaryRepos,
'  Buttons, CodeTransfer, Passwords, HiddenNames, AnaTabIds,
'  ShowHideStore, AnalysisRanges, Checking, BetterArray
'@version 1.4 (2026-08-01)



'@section State
'===============================================================================
'@description Backing UDT, module-level constants, and cached instance variable.

Private Const CLASS_NAME As String = "Linelist"
Private Const PRINT_PREFIX As String = "print_"
Private Const CRF_PREFIX As String = "crf_"

'Excel accepts 31 characters in a worksheet name. The constant read 30, and the
'cut was applied when a sheet was created and skipped when one was looked up, so
'a dictionary sheet name of 25 characters created `print_` + 30 characters and
'PrintWksh then asked for the 31-character name and raised.
Private Const SHEET_NAME_MAX_LENGTH As Long = 31

Private Const WORKBOOK_MODULE_NAME As String = "EventLinelistWorkbook"
Private Const SHEET_MODULE_NAME As String = "EventLinelistSheet"
Private Const SHEET_GEO As String = "Geo"

Private Type TLinelist
    outWkb As Workbook
    specs As LinelistSpecs
    tempRepos As TemporaryRepos
    codeMover As CodeTransfer
    standardDropdowns As DropdownLists
    customDropdowns As DropdownLists
    pivotTables As CustomPivotTable
    sheetNames As BetterArray
    sheetInfo As LLSheets
    checks As Checking
    shortened As Collection
End Type

Private this As TLinelist

'@section Factory
'===============================================================================
'@description Instance creation via VB_Creatable = True factory pattern.

'@label:create
'@sub-title Instantiate a Linelist facade bound to specifications
'@details
'Validates that the supplied specifications object is not Nothing, creates
'a new Linelist instance via `New Linelist`, binds the specifications
'through the Friend property, and returns an Linelist interface reference.
'@param specs LinelistSpecs. The linelist specifications providing
'dictionary, translations, design format, passwords, and all other setup
'data required to build the output workbook.
'@return Linelist. A ready-to-use linelist facade.
'@throws ProjectError.ObjectNotInitialized When specs is Nothing.
'@export
Public Function Create(ByVal specs As LinelistSpecs) As Linelist
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist specifications are required"
    End If

    Dim instance As Linelist
    Set instance = New Linelist

    Set instance.BinSpecs = specs
    Set Create = instance
End Function

'@section Bindings
'===============================================================================
'@description Friend setters used by the factory to inject dependencies.

'@label:bin-specs
'@prop-title Bind the linelist specifications to this instance
'@details
'Called by Create to inject the LinelistSpecs dependency into the
'backing UDT. Exposed as Friend to prevent external callers from
'resetting specifications after construction.
'@param value LinelistSpecs. The specifications to bind.
Friend Property Set BinSpecs(ByVal value As LinelistSpecs)
    Set this.specs = value
End Property

'@section Public properties
'===============================================================================
'@description Worksheet resolution and collaborator accessors exposed through
'the Linelist contract.

'@label:linelist-data
'@prop-title Specifications backing this linelist
'@details
'Returns the LinelistSpecs instance that was supplied during factory
'construction, providing access to dictionary, choices, translations,
'and all other setup data.
'@return LinelistSpecs. The linelist specifications.
'@export
Public Property Get LinelistData() As LinelistSpecs
    Set LinelistData = this.specs
End Property

'@label:dictionary
'@prop-title Dictionary from the specifications
'@details
'Convenience accessor that delegates to specs.Dictionary, avoiding
'the extra indirection at call sites.
'@return LLdictionary. The linelist dictionary.
'@export
Public Property Get Dictionary() As LLdictionary
    Set Dictionary = this.specs.Dictionary
End Property

'@label:sheet-names
'@prop-title The dictionary's unique sheet names
'@details
'One walk of the dictionary's sheet name column, kept for the whole build.
'Prepare, TransferAllCode and clickGenerate each derived this list on their own,
'and every walk reads the whole column.
'@return BetterArray. The unique sheet names of the dictionary.
'@export
Public Property Get SheetNames() As BetterArray
    Set SheetNames = DictionarySheetNames()
End Property

'@label:sheet-info-manager
'@prop-title The LLSheets over the dictionary, shared by the whole build
'@details
'One LLSheets instance over the specifications dictionary, built on first read
'and kept for the whole build. TransferAllCode and the sheet-building loop of
'clickGenerate each created their own over the same dictionary, so every row
'resolution and every cache slot was computed twice.
'@return LLSheets. The sheet information manager of the dictionary.
'@export
Public Property Get SheetInfoManager() As LLSheets
    Set SheetInfoManager = DictionarySheetInfo()
End Property

'@label:wksh
'@prop-title Worksheet from the output workbook by name
'@details
'Returns the worksheet matching the given name from the output workbook.
'Delegates to ResolveWksh with scope 1 (no prefix). Raises an error
'if the sheet does not exist.
'@param sheetName String. The worksheet name to look up.
'@return Worksheet. The resolved worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get Wksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 1, Wksh
End Property

'@label:print-wksh
'@prop-title Printed companion worksheet by base name
'@details
'Returns the printed companion worksheet (prefixed with "print_") from
'the output workbook. Used for HList sheets that have a printable
'duplicate. Delegates to ResolveWksh with scope 2.
'@param sheetName String. The base sheet name (without the print_ prefix).
'@return Worksheet. The printed companion worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 2, PrintWksh
End Property

'@label:crf-wksh
'@prop-title CRF companion worksheet by base name
'@details
'Returns the CRF companion worksheet (prefixed with "crf_") from the
'output workbook. Used for case report form layouts. Delegates to
'ResolveWksh with scope 3.
'@param sheetName String. The base sheet name (without the crf_ prefix).
'@return Worksheet. The CRF companion worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get CRFWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 3, CRFWksh
End Property

'@label:dropdown
'@prop-title Dropdown lists manager for the output workbook
'@details
'Returns the DropdownLists instance backed by the appropriate list sheet.
'Scope 1 uses the temporary __dropdown_lists sheet with the "dropdown_" header
'prefix; scope 2 uses the custom choice sheet with no prefix.
'
'Each scope is built once and kept. DropdownLists accumulates its report entries
'in instance state, and this property used to hand back a fresh object every
'time, so every warning the data-entry build recorded went to a throwaway
'instance and clickGenerate's phase 2b flush pushed nothing. Keeping the
'instance also takes a worksheet lookup and an object construction out of the
'per-sheet build loop, which asks for both scopes on every sheet.
'@param scope Optional Byte. 1 = standard dropdowns, 2 = custom choice
'dropdowns. Defaults to 1.
'@return DropdownLists. The dropdown lists manager.
'@throws ProjectError.InvalidArgument When the scope is neither 1 nor 2.
'@export
Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As DropdownLists
    Select Case scope
        Case 1
            If this.standardDropdowns Is Nothing Then
                Set this.standardDropdowns = BuildDropdown(1)
            End If
            Set Dropdown = this.standardDropdowns

        Case 2
            If this.customDropdowns Is Nothing Then
                Set this.customDropdowns = BuildDropdown(2)
            End If
            Set Dropdown = this.customDropdowns

        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "There is no dropdown scope '" & CStr(scope) & "' in a linelist"
    End Select
End Property

'@label:pivots
'@prop-title Pivot tables manager for the output workbook
'@details
'Returns the CustomPivotTable instance backed by the translated custom pivot
'table sheet. It is built once and kept, for the reason the dropdown managers
'are: the class holds the metadata of the blocks it has stacked, and a fresh
'instance per call starts empty.
'@return CustomPivotTable. The pivot tables manager.
'@export
Public Property Get Pivots() As CustomPivotTable
    Dim trans As TranslationObject
    Dim sh As Worksheet

    If this.pivotTables Is Nothing Then
        Set trans = this.specs.TransObject
        Set sh = Wksh(trans.TranslatedValue("LLSHEET_CustomPivotTable"))
        Set this.pivotTables = CustomPivotTable.Create(sh)
    End If

    Set Pivots = this.pivotTables
End Property

'@label:has-checkings
'@prop-title Whether this linelist has anything to report
'@details
'Answers True once an entry has been filed. Prepare files a warning for every
'sheet name it had to shorten, and folds in the entries CodeTransfer filed for
'the components it replaced in the output workbook.
'@return Boolean. True when there is something to read.
'@export
Public Property Get HasCheckings() As Boolean
    If this.checks Is Nothing Then Exit Property
    HasCheckings = (this.checks.Length() > 0)
End Property

'@label:checking-values
'@prop-title The report entries of this linelist
'@details
'Answers the Checking object holding the entries, and Nothing while none have
'been filed. clickGenerate flushes it into the generation report right after
'Prepare returns.
'@return Checking. The entries, or Nothing.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.checks
End Property

'@label:sheet-exists
'@sub-title Check whether a sheet exists in the output workbook
'@details
'Tests whether a worksheet with the given name exists in the output
'workbook without raising an error. The name is taken as it stands: callers
'that work with a scope pass the answer of ScopedName.
'@param sheetName String. The worksheet name to test.
'@return Boolean. True when the sheet exists, False otherwise.
'@export
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = OutputWkb.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = (Not sh Is Nothing)
End Function

'@label:add-output-sheet
'@sub-title Add a worksheet to the output workbook
'@details
'Creates a worksheet in the output workbook under the name ScopedName builds
'from the base name and the scope. When beforeSheet is supplied the new sheet is
'inserted immediately before it; otherwise the sheet is appended at the default
'position.
'
'A sheet that is already there is left alone. Template workbooks ship with some
'of these sheets already made, and rebuilding one would throw away what the
'template author put in it.
'@param sheetName String. The base name for the new worksheet.
'@param Hide Optional Long. Excel visibility constant. Defaults to
'xlSheetVeryHidden.
'@param sheetScope Optional Byte. 1 = no prefix, 2 = print_ prefix,
'3 = crf_ prefix. Defaults to 1.
'@param beforeSheet Optional Worksheet. An existing worksheet of the output
'workbook to insert before. Defaults to Nothing (append at default position).
'@export
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1, _
                          Optional ByVal beforeSheet As Worksheet)
    Dim actSheetName As String
    Dim outWkb As Workbook

    actSheetName = ScopedName(sheetName, sheetScope)

    If SheetExists(actSheetName) Then Exit Sub

    Set outWkb = OutputWkb

    With outWkb
        If Not beforeSheet Is Nothing Then
            .Worksheets.Add(Before:=beforeSheet).Name = actSheetName
        Else
            .Worksheets.Add.Name = actSheetName
        End If
        .Worksheets(actSheetName).Visible = Hide
    End With
End Sub

'@section Lifecycle
'===============================================================================
'@description Build, save, and error-manage the output workbook.

'@label:prepare
'@sub-title Build the full linelist output workbook from specifications
'@details
'Runs the whole build of the output workbook:
'1. Takes the temporary repository the specifications created.
'2. Resolves the Geo worksheet once. Every sheet this routine adds goes in
'   before it, and the anchor is passed down as a worksheet.
'3. Writes the two language hidden names and every translated message and sheet
'   name onto the output workbook.
'4. Adds one data-entry worksheet per dictionary sheet name.
'5. Adds the six analysis and custom worksheets.
'6. Adds the six temporary worksheets and prepares them with their hidden names
'   and ListObjects.
'7. Moves the instruction sheet to the front (template workbooks only).
'8. Adds and formats the admin sheet (non-template only).
'9. Transfers the VBA components into the workbook.
'10. Folds the code transfer's report entries into this instance's.
'
'The default worksheet Workbooks.Add made is dropped by LinelistSpecs.Prepare,
'before any exporter writes to the workbook.
'@throws ProjectError.ErrorUnexpectedState When the specifications have not been
'prepared.
'@export
Public Sub Prepare()
    Dim trans As TranslationObject
    Dim dictSheets As BetterArray
    Dim extraSheets As BetterArray
    Dim counter As Long
    Dim geoAnchor As Worksheet
    Dim codeTrans As CodeTransfer
    Dim codeChecks As Checking
    Dim ownChecks As Checking
    Dim instrName As String
    Dim instrSheet As Worksheet

    If Not (this.specs.Prepared) Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "Linelist specifications need to be prepared before Linelist creation"
    End If

    Set this.tempRepos = this.specs.OutputTemporaryRepos
    Set geoAnchor = GeoAnchor()

    InitialiseWorkbookNames
    InitialiseTranslationNames

    Set trans = this.specs.TransObject

    'Data entry sheets from the dictionary, before Geo
    Set dictSheets = DictionarySheetNames()

    For counter = dictSheets.LowerBound To dictSheets.UpperBound
        AddOutputSheet CStr(dictSheets.Item(counter)), xlSheetVisible, _
                       beforeSheet:=geoAnchor
    Next

    'Analysis and custom sheets, before Geo. Each one is inserted immediately
    'before the anchor, so the sheets land in the order they are listed here.
    Set extraSheets = New BetterArray
    extraSheets.Push trans.TranslatedValue("LLSHEET_Analysis"), _
                     trans.TranslatedValue("LLSHEET_TemporalAnalysis"), _
                     trans.TranslatedValue("LLSHEET_SpatialAnalysis"), _
                     trans.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"), _
                     trans.TranslatedValue("LLSHEET_CustomChoice"), _
                     trans.TranslatedValue("LLSHEET_CustomPivotTable")

    For counter = extraSheets.LowerBound To extraSheets.UpperBound
        AddOutputSheet CStr(extraSheets.Item(counter)), xlSheetVisible, _
                       beforeSheet:=geoAnchor
    Next

    'The six temporary sheets, before Geo. The linelist runtime reads all six,
    'so the list stays six.
    extraSheets.Clear
    extraSheets.Push this.specs.TemporarySheetName(TempSheetSingle), _
                     this.specs.TemporarySheetName(TempSheetList), _
                     this.specs.TemporarySheetName(TempSheetAnalysis), _
                     this.specs.TemporarySheetName(TempSheetImport), _
                     this.specs.TemporarySheetName(TempSheetSpatial), _
                     this.specs.TemporarySheetName(TempSheetShowHide)

    For counter = extraSheets.LowerBound To extraSheets.UpperBound
        AddOutputSheet CStr(extraSheets.Item(counter)), xlSheetVeryHidden, _
                       beforeSheet:=geoAnchor
    Next

    PrepareTemporarySheets
    PrepareAnalysisLists

    'Instruction sheet (template only)
    If this.specs.HasTemplate() Then
        instrName = trans.TranslatedValue("INSTSHEETNAME")

        On Error Resume Next
        Set instrSheet = OutputWkb.Worksheets(instrName)
        On Error GoTo 0

        If Not instrSheet Is Nothing Then
            'Unhide before Move, because a hidden sheet cannot be moved
            instrSheet.Visible = xlSheetVisible
            instrSheet.Move Before:=OutputWkb.Worksheets(1)

            If this.specs.Value("instruct") = "No" Then
                instrSheet.Visible = xlSheetVeryHidden
            End If
        End If
    End If

    'Admin sheet (non-template only)
    If Not this.specs.HasTemplate() Then
        AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlSheetVisible, _
                       beforeSheet:=geoAnchor
        PrepareAdmin
    End If

    'Transfer code. The instance is kept, because it files a report entry for
    'every component it had to replace in the target.
    Set codeTrans = CodeTransfer.Create(this.specs.DesWorkbook, OutputWkb, this.tempRepos)
    Set this.codeMover = codeTrans
    TransferAllCode codeTrans, dictSheets

    If codeTrans.HasCheckings Then
        Set ownChecks = Checks
        Set codeChecks = codeTrans.CheckingValues
        ownChecks.Append codeChecks
    End If
End Sub

'@label:save-ll
'@sub-title Save the output workbook as .xlsb with password protection
'@details
'Saves the output workbook in .xlsb format with the configured password,
'protects the workbook structure via Passwords.Protect, activates an
'appropriate visible sheet (admin or instruction), closes the workbook,
'and cleans up temporary files via TemporaryRepos.Reset.
'
'Both cached references to the workbook are dropped once it is closed. This
'instance held one and LinelistSpecs held the other, and a read of either after
'the close raised an automation error. The read that mattered was ErrorManage,
'reached from clickGenerate's cleanup, so the failure landed inside the error
'handler and hid the error that got it there.
'@export
Public Sub SaveLL()
    Dim outWkb As Workbook
    Dim passwd As String
    Dim outPath As String
    Dim passObj As Passwords
    Dim msgTrads As TranslationObject

    Set outWkb = OutputWkb
    Set passObj = this.specs.Password
    Set msgTrads = this.specs.TransObject

    passwd = this.specs.Value("llpassword")
    outPath = this.specs.Value("lldir") & Application.PathSeparator & _
              this.specs.Value("llname") & ".xlsb"

    'Select an appropriate visible sheet before saving
    On Error Resume Next
    outWkb.Worksheets(msgTrads.TranslatedValue("LLSHEET_Admin")).Activate
    outWkb.Worksheets(msgTrads.TranslatedValue("INSTSHEETNAME")).Activate
    On Error GoTo 0

    'Protect the workbook structure, add debugging mechanism
    passObj.Protect outWkb
    passObj.EnsureDebugExitHandler outWkb

    outWkb.SaveAs fileName:=outPath, fileFormat:=xlExcel12, Password:=passwd, _
                   ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outWkb.Close

    Set this.outWkb = Nothing
    this.specs.ReleaseLLWorkbook

    If Not this.tempRepos Is Nothing Then this.tempRepos.Reset
End Sub

'@label:error-manage
'@sub-title Display an error dialog for incomplete linelist operations
'@details
'Shows a message box with the error details and asks the user whether
'they want to view the incomplete linelist workbook or close it without
'saving. If the user chooses Yes the workbook's Application is made
'visible; otherwise the workbook is closed without saving.
'
'A workbook SaveLL has already closed leaves nothing to show, and the message
'says so. This routine is reached from a caller's error handler, so it has to
'finish whatever state the build left behind.
'@param textMessage Optional String. Error description to display.
'Defaults to vbNullString.
'@export
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim outWkb As Workbook
    Dim answer As Long

    Set outWkb = this.outWkb
    If outWkb Is Nothing Then Set outWkb = this.specs.LLWorkbook

    If outWkb Is Nothing Then
        MsgBox "An internal error occurred with specified information: " & Chr(10) & _
               textMessage & Chr(10) & _
               "The linelist workbook is already closed.", _
               vbExclamation + vbOKOnly, "ERROR"
        Exit Sub
    End If

    answer = MsgBox("An internal error occurred with specified information: " & Chr(10) & _
                    textMessage & Chr(10) & _
                    "Do you want to see the uncompleted linelist?", _
                    vbExclamation + vbYesNo, "ERROR")

    If answer = vbYes Then
        outWkb.Application.Visible = True
    Else
        DiscardBuild
    End If
End Sub

'@label:discard-build
'@sub-title Close the incomplete output workbook and drop the references to it
'@details
'Closes the output workbook of a build that stopped before SaveLL, with its
'changes thrown away, and drops both cached references to it: the one this
'instance holds and the one LinelistSpecs holds. The multi generation driver
'calls it after a failed row, so the loop moves to the next row with no
'workbook left open behind it; ErrorManage routes its close branch here. When
'there is no output workbook the call exits quietly.
'@export
Public Sub DiscardBuild()
    Dim outWkb As Workbook

    Set outWkb = this.outWkb
    If outWkb Is Nothing Then Set outWkb = this.specs.LLWorkbook
    If outWkb Is Nothing Then Exit Sub

    outWkb.Close savechanges:=False
    Set this.outWkb = Nothing
    this.specs.ReleaseLLWorkbook
End Sub

'@section Private - Workbook management
'===============================================================================
'@description Internal helpers for reaching the output workbook and for turning
'a base sheet name plus a scope into the name Excel actually carries.

'@label:output-wkb
'@prop-title Return the output workbook from LinelistSpecs
'@details
'On first access, retrieves the output workbook that LinelistSpecs.Prepare
'created and caches it locally.
'
'The class used to record Worksheets(1).Name here so Prepare could delete it
'later. First access is well after every exporter has written to the workbook,
'and LLGeo inserts its two sheets near the front, so the name recorded was a
'real, populated sheet. LinelistSpecs.Prepare drops the default worksheet now,
'holding the reference it took before the first exporter ran.
'@return Workbook. The output workbook.
'@throws ProjectError.ErrorUnexpectedState When there is no output workbook.
Private Property Get OutputWkb() As Workbook
    If this.outWkb Is Nothing Then
        Set this.outWkb = this.specs.LLWorkbook
    End If

    If this.outWkb Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "There is no output workbook to work on: either Prepare has " & _
                   "not created it, or SaveLL has already closed it"
    End If

    Set OutputWkb = this.outWkb
End Property

'@label:scoped-name
'@sub-title The worksheet name a base name carries under a scope
'@details
'Puts the scope's prefix on the base name and cuts the result to the 31
'characters Excel accepts. Every caller that creates a sheet and every caller
'that looks one up goes through here, which is what keeps the two sides
'agreeing.
'
'A cut that shortens the name is filed as a report entry, once per resulting
'name, so a build finishes and the user reads which sheets came out shorter.
'@param sheetName String. The base sheet name.
'@param sheetScope Byte. 1 = no prefix, 2 = print_ prefix, 3 = crf_ prefix.
'@return String. The worksheet name to create or to look up.
Private Function ScopedName(ByVal sheetName As String, ByVal sheetScope As Byte) As String
    Dim fullName As String
    Dim shortName As String

    Select Case sheetScope
        Case 2
            fullName = PRINT_PREFIX & sheetName
        Case 3
            fullName = CRF_PREFIX & sheetName
        Case Else
            fullName = sheetName
    End Select

    If Len(fullName) <= SHEET_NAME_MAX_LENGTH Then
        ScopedName = fullName
        Exit Function
    End If

    shortName = Left$(fullName, SHEET_NAME_MAX_LENGTH)
    ReportShortName fullName, shortName
    ScopedName = shortName
End Function

'@label:report-short-name
'@sub-title File one report entry for a name that had to be shortened
'@details
'The same name is scoped many times during a build, so the entry is filed once
'and the names already filed are kept in a keyed Collection.
'@param fullName String. The name the prefix and the base name make.
'@param shortName String. The name the worksheet actually carries.
Private Sub ReportShortName(ByVal fullName As String, ByVal shortName As String)
    If this.shortened Is Nothing Then Set this.shortened = New Collection
    If AlreadyReported(shortName) Then Exit Sub

    this.shortened.Add shortName, shortName

    LogCheck shortName, _
             "The worksheet name " & fullName & " is longer than the " & _
             CStr(SHEET_NAME_MAX_LENGTH) & " characters Excel allows, so the " & _
             "sheet was created as " & shortName & ". Shorten the sheet name in " & _
             "the dictionary to choose the name yourself.", _
             checkingWarning
End Sub

'@label:already-reported
'@sub-title Whether a shortened name has been filed already
'@details
'A Collection raises on a key it does not hold, which is how a miss is read.
'@param shortName String. The shortened worksheet name.
'@return Boolean. True when the name has been filed.
Private Function AlreadyReported(ByVal shortName As String) As Boolean
    Dim marker As Variant

    On Error Resume Next
        marker = this.shortened.Item(shortName)
    On Error GoTo 0

    AlreadyReported = Not IsEmpty(marker)
End Function

'@label:resolve-wksh
'@sub-title Resolve a worksheet from the output workbook with scope prefix
'@details
'Builds the worksheet name through ScopedName, verifies the sheet exists, and
'returns a reference through the ByRef result parameter.
'@param sheetName String. The base sheet name.
'@param sheetScope Byte. 1 = no prefix, 2 = print_ prefix, 3 = crf_ prefix.
'@param result Worksheet. ByRef output parameter receiving the resolved sheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
Private Sub ResolveWksh(ByVal sheetName As String, ByVal sheetScope As Byte, _
                        ByRef result As Worksheet)
    Dim actSheetName As String
    Dim outWkb As Workbook

    actSheetName = ScopedName(sheetName, sheetScope)
    Set outWkb = OutputWkb

    If Not SheetExists(actSheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet " & actSheetName & " does not exist in the output workbook"
    End If

    Set result = outWkb.Worksheets(actSheetName)
End Sub

'@label:geo-anchor
'@sub-title The Geo worksheet every added sheet goes in before
'@details
'Prepare used to pass the name "Geo" down to more than twenty AddOutputSheet
'calls, and each one resolved it again. A missing anchor made AddOutputSheet
'fall back to a bare Worksheets.Add, which inserts before whatever sheet is
'active and gives a different final sheet order on every run. The anchor is
'resolved once here and a missing one stops the build.
'@return Worksheet. The Geo worksheet of the output workbook.
'@throws ProjectError.ElementNotFound When the Geo worksheet is missing.
Private Function GeoAnchor() As Worksheet
    Dim anchor As Worksheet
    Dim outWkb As Workbook

    Set outWkb = OutputWkb

    On Error Resume Next
        Set anchor = outWkb.Worksheets(SHEET_GEO)
    On Error GoTo 0

    If anchor Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The " & SHEET_GEO & " worksheet is missing from the output " & _
                   "workbook, so there is no place to put the new worksheets"
    End If

    Set GeoAnchor = anchor
End Function

'@label:dictionary-sheet-names
'@sub-title The dictionary's sheet names, walked once per build
'@details
'Prepare and TransferAllCode both need this list and each used to derive it,
'which walks the whole dictionary.
'@return BetterArray. The unique sheet names of the dictionary.
Private Function DictionarySheetNames() As BetterArray
    If this.sheetNames Is Nothing Then
        Set this.sheetNames = this.specs.Dictionary.UniqueValues("sheet name")
    End If
    Set DictionarySheetNames = this.sheetNames
End Function

'@label:dictionary-sheet-info
'@sub-title The dictionary's sheet manager, built on first read
'@details
'The dictionary is meaningful once specs.Prepare has run, and a build on first
'read keeps an early ask safe.
'@return LLSheets. The sheet information manager of the dictionary.
Private Function DictionarySheetInfo() As LLSheets
    If this.sheetInfo Is Nothing Then
        Set this.sheetInfo = LLSheets.Create(this.specs.Dictionary)
    End If
    Set DictionarySheetInfo = this.sheetInfo
End Function

'@label:build-dropdown
'@sub-title Build the dropdown manager of one scope
'@param scope Byte. 1 = standard dropdowns, 2 = custom choice dropdowns.
'@return DropdownLists. The manager over the sheet that scope reads.
Private Function BuildDropdown(ByVal scope As Byte) As DropdownLists
    Dim listSheetName As String
    Dim headerPrefix As String
    Dim trans As TranslationObject

    If scope = 2 Then
        Set trans = this.specs.TransObject
        listSheetName = trans.TranslatedValue("LLSHEET_CustomChoice")
    Else
        listSheetName = this.specs.TemporarySheetName(TempSheetList)
        headerPrefix = "dropdown_"
    End If

    Set BuildDropdown = DropdownLists.Create(Wksh(listSheetName), headerPrefix)
End Function

'@section Private - Prepare helpers
'===============================================================================
'@description Internal routines called by Prepare to set up temporary sheets,
'the admin sheet, and code transfer.

'@label:prepare-temporary-sheets
'@sub-title Set up hidden names and ListObjects on temporary sheets
'@details
'Performs three setup tasks on the output workbook:
'1. Creates workbook-level hidden names for the global flags: update list auto,
'   import done, geo update manual, show all optionals, export init, number of
'   exports and epi week start.
'2. Creates four ListObjects on the import report sheet for tracking sheets
'   and variables not imported/found.
'3. Asks AnaTabIds to build its registry table on the analysis sheet. That
'   class reads the columns, so it writes the headers.
Private Sub PrepareTemporarySheets()
    Dim outWkb As Workbook
    Dim sh As Worksheet
    Dim rng As Range
    Dim sheetName As String
    Dim wkbNames As HiddenNames

    Set outWkb = OutputWkb

    'Workbook-level hidden names for the global flags
    Set wkbNames = HiddenNames.Create(outWkb)
    wkbNames.EnsureName "RNG_UpdateListAuto", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_UpdateListAuto", vbNullString
    wkbNames.EnsureName "RNG_ImportDone", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ImportDone", vbNullString
    wkbNames.EnsureName "RNG_GeoUpdateManual", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_GeoUpdateManual", vbNullString
    wkbNames.EnsureName "RNG_ShowAllOptionals", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ShowAllOptionals", vbNullString
    wkbNames.EnsureName "RNG_ExportInit", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ExportInit", vbNullString
    wkbNames.EnsureName "RNG_NumberExports", this.specs.Value("numberofexports"), HiddenNameTypeString
    wkbNames.SetValue "RNG_NumberExports", this.specs.Value("numberofexports")
    wkbNames.EnsureName "RNG_EpiWeekStart", this.specs.Value("epiweekstart"), HiddenNameTypeString
    wkbNames.SetValue "RNG_EpiWeekStart", this.specs.Value("epiweekstart")

    'Import report sheet - create ListObjects
    sheetName = this.specs.TemporarySheetName(TempSheetImport)
    Set sh = outWkb.Worksheets(sheetName)

    With sh
        .Cells(1, 1).Value = "sheetNotImported"
        .Cells(1, 3).Value = "sheetNotTouched"
        .Cells(1, 5).Value = "variableNotImported"
        .Cells(1, 6).Value = "correspondingSheet"
        .Cells(1, 8).Value = "variableNotFound"
        .Cells(1, 9).Value = "correspondingSheet"

        Set rng = .Range(.Cells(1, 1), .Cells(2, 1))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotImp"

        Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotTouch"

        Set rng = .Range(.Cells(1, 5), .Cells(2, 6))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotImp"

        Set rng = .Range(.Cells(1, 8), .Cells(2, 9))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotTouch"
    End With

    'Analysis sheet - one registry table, built by the class that reads it.
    'The analysis sheet names used to live here as worksheet-level hidden names
    'as well. The export passes the sheet it copies from, and the workbook-level
    'set LLTranslation writes is the one every other reader uses.
    sheetName = this.specs.TemporarySheetName(TempSheetAnalysis)
    Set sh = outWkb.Worksheets(sheetName)
    AnaTabIds.PrepareSheet sh

    'Show/hide sheet - the table the user's visibility choices are kept in.
    'The sheet used to be provisioned empty and the only code that ever built a
    'table on it wrote into the export workbook, so every choice a user made was
    'thrown away when the workbook closed.
    sheetName = this.specs.TemporarySheetName(TempSheetShowHide)
    Set sh = outWkb.Worksheets(sheetName)
    ShowHideStore.CreateOnSheet sh
End Sub

'@label:prepare-analysis-lists
'@sub-title Add the choice lists every analysis sheet shares
'@details
'ONE LIST FOR THE WHOLE WORKBOOK
'-------------------------------------------------------------------------------
'The time units a temporal table can group by are one workbook-level dropdown on
'the list sheet. Every temporal table binds its own time unit cell to it, and
'GetAgg in CustomLinelistFunctions reads the same list to turn the chosen label
'back into a time unit.
'
'Each analysis sheet used to build a copy of its own, which is why there were
'two spellings of the same five words. A reader that named one of them answered
'wrongly on the sheet carrying the other, and answered nothing at all in a
'workbook that carried neither.
'
'The words come from the message translations, so the list is written in the
'language of the linelist. The dropdown is added once; DropdownLists refuses a
'second one under the same name and files the reason.
Private Sub PrepareAnalysisLists()
    Dim listValues As BetterArray
    Dim trans As TranslationObject
    Dim drop As DropdownLists

    Set trans = this.specs.TransObject
    Set drop = Dropdown(1)

    Set listValues = New BetterArray
    listValues.LowerBound = 1
    listValues.Push trans.TranslatedValue("MSG_Day"), _
                    trans.TranslatedValue("MSG_Week"), _
                    trans.TranslatedValue("MSG_Month"), _
                    trans.TranslatedValue("MSG_Quarter"), _
                    trans.TranslatedValue("MSG_Year")

    drop.Add listValues, AnalysisRanges.TimeUnitListName
End Sub

'@label:initialise-workbook-names
'@sub-title Check the two language selections before the build starts
'@details
'The two language codes reach the output workbook through
'LLBuild.SyncDesignerLanguageNames and ExportDesignerHiddenNames, during
'specs.Prepare. This routine used to write them a second time, deriving the
'interface code with a bare Split on "-" where the build service guards the
'split with InStr. Both answer the same string for every input, including an
'empty one, so the second write added nothing.
'
'What it does add is the check: an empty language selection is a mistake the
'user can fix, and the build service says nothing about it. That check is what
'is left here.
'@throws ProjectError.ObjectNotInitialized When either language is empty.
Private Sub InitialiseWorkbookNames()
    Dim dictLang As String
    Dim llLang As String

    dictLang = this.specs.Value("setuplang")
    llLang = Split(this.specs.Value("lllang"), "-")(0)

    If LenB(dictLang) = 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary language (RNG_LangSetup) is empty on the Main sheet. " & _
                   "Please select a setup language before generating."
    End If

    If LenB(llLang) = 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Interface language (RNG_LLForm) is empty on the Main sheet. " & _
                   "Please select an interface language before generating."
    End If
End Sub

'@label:initialise-translation-names
'@sub-title Create the translated message and sheet names on the output workbook
'@details
'Calls InitialiseHiddenNames to create all RNG_* translated message and sheet
'names on the output workbook.
'
'A second copy of the two language codes used to be written onto the translation
'worksheet here. LLTranslation reads its languages from the workbook store, so
'that copy was written and never read.
Private Sub InitialiseTranslationNames()
    Dim transObj As LLTranslation

    Set transObj = this.specs.Translations
    transObj.InitialiseHiddenNames OutputWkb
End Sub


'@label:prepare-admin
'@sub-title Format the admin sheet and add Advanced/Export buttons
'@details
'Applies the linelist design format to the admin sheet, then creates two
'large-scope buttons (Advanced and Export) with translated labels. After
'placing the buttons, re-suppresses UI updates and activates the admin
'sheet. Finally protects the active sheet via Passwords.Protect with
'the "_active" scope marker.
'@remarks
'The trads object is switched to TranslationOfShapes scope to fetch
'shape-specific translated labels for the buttons.
Private Sub PrepareAdmin()
    Const ADVANCED_COMMAND As String = "ClickAdvanced"
    Const EXPORT_COMMAND As String = "ClickExport"
    Const SHP_ADVANCED As String = "SHP_Advanced"
    Const SHP_EXPORT As String = "SHP_Export"

    Dim sh As Worksheet
    Dim trads As TranslationObject
    Dim design As LLFormat
    Dim pass As Passwords
    Dim but As Buttons
    Dim rng As Range

    Set design = this.specs.DesignFormat
    Set trads = this.specs.TransObject
    Set pass = this.specs.Password

    Set sh = Wksh(trads.TranslatedValue("LLSHEET_Admin"))
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Switch to shapes translation scope
    Set trads = this.specs.TransObject(TranslationOfShapes)

    'Advanced button
    Set rng = sh.Cells(2, 3)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_ADVANCED, scope:=ButtonScopeLarge)
    but.Add actionCommand:=ADVANCED_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_ADVANCED)
    but.Format design

    'Export button
    Set rng = sh.Cells(2, 5)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_EXPORT, scope:=ButtonScopeLarge)
    but.Add actionCommand:=EXPORT_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_EXPORT)
    but.Format design

    sh.Activate

    pass.Protect "_active"
End Sub

'@label:transfer-all-code
'@sub-title Transfer all VBA code components to the output workbook
'@details
'Moves the VBA the generated linelist runs on, in five steps:
'1. 30 class modules. ChoiceKeys, ImportMetadata, ImportReport and
'   CheckingOutput joined the list with the import work: LLExporter types
'   ChoiceKeys, LLImporter types the first three, ImportChecking types
'   CheckingOutput and F_ImportRep reads ImportReport, so a linelist without
'   them fails project compilation. LLLog joined with the user log:
'   EventLinelist holds one and the button and form modules write their
'   outcome lines through it.
'2. 9 standard modules. Five carry the linelist toolbar and its helpers, and
'   four hold the code behind the user forms: FormLogicAdvanced,
'   FormLogicExportMig, FormLogicExport and FormLogicEpiWeek. Those four are the
'   only ones of the nine linelistform modules that export a Public procedure,
'   and EventsLinelistButtons calls three of them, so a linelist without them
'   fails project compilation on the first line of VBA that runs in it.
'   ImportChecking is the fifth of the first group and FormLogicAdvanced calls
'   it.
'3. The ribbon module. A template workbook already carries EventsLinelistRibbon,
'   so its code text is overwritten; otherwise the module is imported.
'4. 9 user forms.
'5. The workbook code module, then EventLinelistSheet into the code module of
'   every hlist2D data-entry worksheet.
'
'scripts/devtools/transferred-code-scan.R checks that every procedure the
'transferred modules call is defined inside the transferred set. Run it after
'any change to the three lists below.
'@param codeTransfer CodeTransfer. The code transfer engine.
'@param sheetNames BetterArray. The dictionary's unique sheet names, walked once
'by Prepare and passed down.
Private Sub TransferAllCode(ByVal codeTransfer As CodeTransfer, _
                            ByVal sheetNames As BetterArray)
    Dim codesList As BetterArray
    Dim counter As Long
    Dim codeName As String
    Dim llshs As LLSheets

    Set codesList = New BetterArray

    'Transfer classes (33 classes)
    codesList.Push "BetterArray", "DataSheet", "LLExport", _
                   "LLChoices", _
                   "LLdictionary", "LLGeo", _
                   "Passwords", _
                   "Checking", _
                   "LLSheets", "LLTranslation", _
                   "OSFiles", _
                   "TranslationObject", "LLVariables", _
                   "LLSpatial", _
                   "DropdownLists", _
                   "ShowHide", "ShowHideLayout", "ShowHideStore", _
                   "CustomTable", _
                   "LLExporter", _
                    "LLImporter", _
                   "ChoiceKeys", "ImportMetadata", "ImportReport", _
                   "CheckingOutput", _
                   "Graphs", "AnaTabIds", _
                   "ExportButton", "FilteredData", _
                   "HiddenNames", _
                   "EventLinelist", _
                   "LLLog", _
                   "ApplicationState"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferClass codeName
    Next

    'Transfer modules (9 modules, plus the ribbon just below)
    codesList.Clear
    codesList.Push "CustomLinelistFunctions", _
                   "GeoModule", "LinelistEventsManager", _
                   "EventsLinelistButtons", "ImportChecking", _
                   "FormLogicAdvanced", "FormLogicExportMig", _
                   "FormLogicExport", "FormLogicEpiWeek"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferModule codeName
    Next

    'Transfer or copy ribbon module depending on template status
    If this.specs.HasTemplate() Then
        'Template already contains EventsLinelistRibbon; overwrite its code
        codeTransfer.CopyModuleText "EventsLinelistRibbon"
    Else
        'Non-template: import EventsLinelistRibbon as a new module
        codeTransfer.TransferModule "EventsLinelistRibbon"
    End If

    'Transfer forms (10 forms)
    codesList.Clear
    codesList.Push "F_Advanced", "F_Export", "F_ExportMig", _
                   "F_Geo", "F_ImportRep", _
                   "F_ShowHideLL", "F_ShowHidePrint", "F_ShowHideSave", _
                   "F_ShowVarLabels", "F_EpiWeek"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferForm codeName
    Next

    'Transfer workbook code
    codeTransfer.TransferWorkbookCode WORKBOOK_MODULE_NAME

    'Inject the worksheet event handlers into each HList worksheet's code
    'module. The dictionary identifies HList sheets by type "hlist2D". The
    'source template module (EventLinelistSheet) is read from the designer
    'workbook and its code is written into each HList worksheet's CodeModule in
    'the output workbook.
    Set llshs = SheetInfoManager

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        codeName = CStr(sheetNames.Item(counter))
        If llshs.SheetInfo(codeName) = "hlist2D" Then
            codeTransfer.TransferWorksheetCode SHEET_MODULE_NAME, codeName
        End If
    Next
End Sub

'@section Checkings
'===============================================================================
'@description What this class reports into the generation report. clickGenerate
'reads HasCheckings right after Prepare and flushes CheckingValues.

'@label:checks
'@prop-title The report entries of this instance
'@return Checking. The entries, created on first use.
Private Property Get Checks() As Checking
    If this.checks Is Nothing Then Set this.checks = Checking.Create(CLASS_NAME)
    Set Checks = this.checks
End Property

'@label:log-check
'@sub-title File one report entry
'@details
'Checking.Add raises on a duplicate key, and one report carries the entries of
'several classes, so the key names this class and the thing the entry is about.
'@param checkKey String. What the entry is about, used to build the key.
'@param message String. The message text.
'@param scope Optional Byte. The checking scope. Defaults to checkingNote.
Private Sub LogCheck(ByVal checkKey As String, ByVal message As String, _
                     Optional ByVal scope As Byte = checkingNote)
    Dim entries As Checking

    Set entries = Checks
    entries.Add CLASS_NAME & "-" & checkKey, message, scope
End Sub

'@section Private - Error handling
'===============================================================================
'@description Centralised error-raising helper.

'@label:throw-error
'@sub-title Raise a project error with class context
'@details
'Wraps Err.Raise with CLASS_NAME as the source, providing consistent
'error attribution across all methods in this class.
'@param errNumber Long. The ProjectError constant to raise.
'@param errorMessage String. Descriptive message for the error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, errorMessage
End Sub
