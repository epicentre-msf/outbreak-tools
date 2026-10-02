VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenerationLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Owns the record of one generation run and writes it to the __check worksheet"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Owns the record of one generation run and writes it to the __check worksheet")

Option Explicit

'@class GenerationLog
'@description
'The single owner of a generation run's record. Start opens the run on
'the designer __check worksheet; Collect and Harvest take Checking
'bundles, stamp every entry with the flush time, keep the entries in an
'in-memory record and flush the bundle to the worksheet through
'CheckingOutput, which appends across renders. Finish writes the closing
'bundle (outcome and elapsed seconds, plus the sheet and variable counts
'when the driver knows them), installs the filter handler and shows the
'worksheet. ExportText writes the whole record to a text file, one line
'per entry, so the run stays readable outside the workbook.
'
'THE RUN WINDOW
'-------------------------------------------------------------------------------
'Entries land between Start and Finish. A bundle collected outside that
'window is ignored, the behaviour every driver leans on: the build core
'runs the same with or without an open log. The in-memory record stays
'alive after Finish, which is what a later text re-export reads. Start
'resets the worksheet, the render marker and the record, so a re-run
'begins clean.
'
'SECTIONS AND SUBSECTIONS
'-------------------------------------------------------------------------------
'A run that builds several linelists opens one section per linelist,
'through OpenSection and CloseSection. While a section is open every
'bundle is re-headed -- the section title becomes the bundle title and
'the bundle's own title becomes its subtitle -- and held back. Closing
'the section writes the held bundles in one pass, so the section title
'lands once as a level one heading and every part of that build reads as
'a subsection under it. Without an open section a bundle is written on
'its own, which is what the single build does.
'
'BUNDLES CARRIED AS TEXT
'-------------------------------------------------------------------------------
'The build runs as steps, and a step may run in another Excel instance
'where the log cannot reach. BundleText turns one bundle into a string
'that ends on a form feed, so the texts of several bundles are joined by
'plain concatenation; CollectText takes such a string back and collects
'every bundle in it, record-only flag included. Both sit on this class
'so the two sides of the carriage share one shape. A label carrying a
'tab or a vertical tab is cut at that character.
'
'SHOWING THE REPORT
'-------------------------------------------------------------------------------
'Finish closes the run and leaves the worksheet where it is. The driver
'shows it once, at the end, through ShowReport, with the screen back on.
'@depends CheckingOutput, Checking, HiddenNames, BetterArray, LinelistSpecs


'@section State
'===============================================================================

Private Type TGenerationLog
    designerBook As Workbook
    writer As CheckingOutput
    record As BetterArray
    startedAt As Date
    sealed As Boolean
    sectionTitle As String
    sectionBatch As BetterArray
End Type

Private this As TGenerationLog

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "GenerationLog"
Private Const SHEET_CHECKING As String = "__check"
Private Const WRITER_NAME As String = "Generation Report"
Private Const HEADER_TITLE As String = "Generation run"
Private Const CLOSING_TITLE As String = "Generation outcome"
Private Const EXPORT_SUFFIX As String = "-generation.txt"

'Render marker CheckingOutput keeps on its worksheet (ROW_MARKER_NAME
'there). Start drops it so a cleared sheet is formatted as a first
'render on every run.
Private Const ROW_MARKER_NAME As String = "CheckingOutputStartingRow"

'The separators of a bundle carried as text. BundleText writes them and
'CollectText reads them: a form feed closes a bundle, a vertical tab
'closes a line inside it, and a tab splits the fields of a line. The
'first line of a bundle is its title, subtitle and record-only flag;
'every other line is one entry: key, label, scope word.
Private Const BUNDLE_END As String = vbFormFeed
Private Const LINE_END As String = vbVerticalTab
Private Const FIELD_SEP As String = vbTab
Private Const FLAG_RECORD_ONLY As String = "1"
Private Const FLAG_WORKSHEET As String = "0"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a log bound to the designer workbook
'@details
'The workbook is where the __check worksheet lives. The run itself opens
'later, through Start.
'@param designerBook Workbook. The designer workbook.
'@return GenerationLog. A ready instance.
'@throws ProjectError.ElementNotFound When designerBook is Nothing.
'@export
Public Function Create(ByVal designerBook As Workbook) As GenerationLog
    If designerBook Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
            "Unable to create the log, please provide a valid workbook"
    End If

    Dim newLog As GenerationLog
    Set newLog = New GenerationLog

    Set newLog.DesignerBook = designerBook
    newLog.Seal
    Set Create = newLog
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-seal mutation
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub


'@section PublicAccessors
'===============================================================================

'@label:designer-book
'@prop-title The designer workbook the run writes on
'@return Workbook. The workbook given at creation.
Public Property Get DesignerBook() As Workbook
    Set DesignerBook = this.designerBook
End Property

'@label:designer-book-set
'@prop-title Assign the designer workbook
'@param designerBook Workbook. The workbook to store.
Public Property Set DesignerBook(ByVal designerBook As Workbook)
    GuardNotSealed "DesignerBook"
    Set this.designerBook = designerBook
End Property

'@label:record-length
'@prop-title The number of entries in the in-memory record
'@details
'Counts every stamped entry the run has collected, the header and
'closing bundles included. Answers 0 before the first Start.
'@return Long. The entry count.
Public Property Get RecordLength() As Long
    If this.record Is Nothing Then Exit Property
    RecordLength = this.record.Length
End Property

'@label:record-line
'@prop-title One record entry as its text line
'@details
'The line shape is the one the text export writes:
'hh:mm:ss  [scope]  bundle-title  entry-label.
'@param index Long. The 1-based position of the entry.
'@return String. The formatted line.
'@throws ProjectError.ObjectNotInitialized When the run has never started.
'@throws ProjectError.InvalidArgument When index is outside the record.
Public Property Get RecordLine(ByVal index As Long) As String
    Dim entry As Variant

    If this.record Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "The record is empty, start the log first"
    End If
    If index < this.record.LowerBound Or index > this.record.UpperBound Then
        ThrowError ProjectError.InvalidArgument, _
            "The record holds no entry at position " & CStr(index)
    End If

    'BetterArray.Push rebases a stored array to the outer lower bound,
    'so the fields are read from the entry's own base.
    Dim fieldBase As Long
    entry = this.record.Item(index)
    fieldBase = LBound(entry)
    RecordLine = CStr(entry(fieldBase)) & "  [" & CStr(entry(fieldBase + 4)) & "]  " & _
                 CStr(entry(fieldBase + 1)) & "  " & CStr(entry(fieldBase + 3))
End Property

'@label:section-title
'@prop-title The title of the section that is open
'@details
'Answers an empty string when no section is open, which is the state of
'a single build from the first bundle to the last.
'@return String. The open section title.
Public Property Get SectionTitle() As String
    SectionTitle = this.sectionTitle
End Property


'@section Run lifecycle
'===============================================================================

'@label:start
'@sub-title Open the run on the __check worksheet
'@details
'Resolves the __check worksheet, creating it very hidden when absent,
'clears it and drops the render marker of the previous run so the writer
'formats a first render. Opens a fresh writer, resets the in-memory
'record, starts the wall clock and stamps the run header as the first
'bundle: the start time and the platform of the run, plus the setup path
'and the linelist name when the caller knows them.
'@param setupPath Optional String. The setup file path of the run.
'@param linelistName Optional String. The output linelist name of the run.
Public Sub Start(Optional ByVal setupPath As String = vbNullString, _
                 Optional ByVal linelistName As String = vbNullString)
    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim header As Checking

    Set sh = ResolveCheckingSheet()
    sh.Cells.Clear

    'The marker survives a Cells.Clear. Keeping it makes the writer resume
    'on the cleared sheet and skip the first-render formatting.
    Set store = HiddenNames.Create(sh)
    If store.HasName(ROW_MARKER_NAME) Then store.RemoveName ROW_MARKER_NAME

    Set this.writer = CheckingOutput.Create(sh, WRITER_NAME)
    Set this.record = New BetterArray
    this.record.LowerBound = 1
    this.startedAt = Now

    'A run opens with no section. The header and the closing bundle of the
    'run are written on their own, and a driver that builds several
    'linelists opens one section per linelist between them.
    this.sectionTitle = vbNullString
    Set this.sectionBatch = Nothing

    Set header = Checking.Create(HEADER_TITLE)
    'The platform rides on the start entry rather than taking an entry of its
    'own. Readers of the record address its lines by position, so an extra
    'header entry would move every one of them.
    header.Add "started", Format$(this.startedAt, "yyyy-mm-dd hh:nn:ss") & _
                          " on " & PlatformTag(), checkingInfo
    If LenB(setupPath) > 0 Then header.Add "setup file", setupPath, checkingInfo
    If LenB(linelistName) > 0 Then header.Add "linelist name", linelistName, checkingInfo
    Collect header
End Sub

'@label:platform-tag
'@sub-title The platform of the run, on one tag
'@details
'A generation report is read on another machine and long after the run, and
'the two platforms do not behave the same: a geobase export Windows accepts
'has refused on a Mac. The name and the bitness come from the compile
'constants, so they are the build that is running and cannot be misread.
'The Excel version comes from the application.
'@return String. Something of the shape "mac-64 excel-16.90".
Private Function PlatformTag() As String
    Dim osName As String
    Dim bits As String

    #If Mac Then
        osName = "mac"
    #Else
        osName = "win"
    #End If

    #If Win64 Then
        bits = "64"
    #Else
        bits = "32"
    #End If

    PlatformTag = osName & "-" & bits & " excel-" & Application.Version
End Function

'@label:open-section
'@sub-title Open one section of the report
'@details
'A section is one whole piece of work, and for a run that builds several
'linelists that is one linelist. While the section is open every bundle
'Collect takes is held back and re-headed: the section title becomes the
'bundle title and the bundle's own title becomes its subtitle. Closing
'the section writes the held bundles in one pass, which puts the section
'title on the sheet once, at level one, with every part of the build
'under it at level two.
'
'Opening a section closes the one that was open before it, so a driver
'that walks a table calls this once per row and the rows never mix. An
'empty title closes the open section and opens none. A call outside the
'run window is ignored, like every other call here.
'@param sectionTitle String. The title of the section.
Public Sub OpenSection(ByVal sectionTitle As String)
    CloseSection

    If this.writer Is Nothing Then Exit Sub
    If LenB(sectionTitle) = 0 Then Exit Sub

    this.sectionTitle = sectionTitle
    Set this.sectionBatch = New BetterArray
    this.sectionBatch.LowerBound = 1
End Sub

'@label:close-section
'@sub-title Write the open section to the worksheet and close it
'@details
'The held bundles go to the writer as one batch, which is what makes the
'section title land once. A section that took no bundle writes nothing.
'Calling this with no section open is a no-op, so a driver may close
'after every row and after the loop.
Public Sub CloseSection()
    If LenB(this.sectionTitle) = 0 Then
        Set this.sectionBatch = Nothing
        Exit Sub
    End If

    If Not this.writer Is Nothing Then
        If Not this.sectionBatch Is Nothing Then
            If this.sectionBatch.Length > 0 Then
                this.writer.PrintOutput this.sectionBatch
            End If
        End If
    End If

    this.sectionTitle = vbNullString
    Set this.sectionBatch = Nothing
End Sub

'@label:collect
'@sub-title Take one Checking bundle into the run
'@details
'Stamps every entry of the bundle with the flush time, appends the
'entries to the in-memory record and flushes the bundle to the worksheet.
'A Nothing or empty bundle is skipped, and so is any bundle arriving
'outside the run window (before Start or after Finish).
'
'THE RECORD-ONLY PATH
'-------------------------------------------------------------------------------
'recordOnly keeps a bundle out of the worksheet and puts it in the
'in-memory record alone, so it reaches the text file and nothing else.
'It carries the detail bundles: one entry per variable written is a few
'hundred entries on a normal linelist, and CheckingOutput ends every
'row it writes with an EntireColumn.AutoFit. The worksheet keeps the
'sheet, section and analysis milestones plus every fault, which is what
'a reader opens it for.
'
'THE BUNDLE OF AN OPEN SECTION
'-------------------------------------------------------------------------------
'The bundle is held in the section, and CloseSection writes the whole
'section in one pass. The record line carries the section title in front
'of the bundle title, so the text file reads the same way the worksheet
'does.
'@param checks Checking. The bundle to take.
'@param recordOnly Optional Boolean. True keeps the bundle off the worksheet.
Public Sub Collect(ByVal checks As Checking, _
                   Optional ByVal recordOnly As Boolean = False)
    Dim stampText As String
    Dim keys As BetterArray
    Dim keyIndex As Long
    Dim keyName As String
    Dim batch As BetterArray
    Dim headingText As String

    If this.writer Is Nothing Then Exit Sub
    If checks Is Nothing Then Exit Sub
    If checks.Length = 0 Then Exit Sub

    'The flush time is the honest stamp: Checking entries carry no clock,
    'and the drivers flush at their milestones.
    stampText = Format$(Now, "hh:nn:ss")

    headingText = checks.Heading
    If LenB(this.sectionTitle) > 0 Then
        headingText = this.sectionTitle & " - " & headingText
    End If

    Set keys = checks.ListOfKeys
    For keyIndex = keys.LowerBound To keys.UpperBound
        keyName = CStr(keys.Item(keyIndex))
        this.record.Push Array(stampText, headingText, keyName, _
                               checks.ValueOf(keyName, checkingLabel), _
                               ScopeText(checks, keyName))
    Next keyIndex

    If recordOnly Then Exit Sub

    If LenB(this.sectionTitle) > 0 Then
        this.sectionBatch.Push SubsectionOf(checks)
        Exit Sub
    End If

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push checks
    this.writer.PrintOutput batch
End Sub

'@label:subsection-of
'@sub-title One bundle re-headed as a subsection of the open section
'@details
'A copy is taken because the caller owns the bundle it handed over, and
'because a Checking is sealed at creation: its title is written by the
'factory alone. Append carries the entries across with their scopes.
'@param checks Checking. The bundle to re-head.
'@return Checking. The copy, headed by the section and subheaded by the bundle.
Private Function SubsectionOf(ByVal checks As Checking) As Checking
    Dim headed As Checking

    Set headed = Checking.Create(this.sectionTitle, checks.Heading)
    headed.Append checks
    Set SubsectionOf = headed
End Function

'@label:harvest
'@sub-title Collect the bundles of the specification collaborators
'@details
'Pulls the checkings of the dictionary, the choices, the exports, the
'analysis, the passwords, the format and the specifications themselves,
'in that order. The pull runs after LinelistSpecs.Prepare, when the
'collaborators have filed their entries.
'@param specs LinelistSpecs. The specifications after Prepare.
'@throws ProjectError.InvalidArgument When specs is Nothing.
Public Sub Harvest(ByVal specs As LinelistSpecs)
    Dim bundles As BetterArray
    Dim index As Long

    If specs Is Nothing Then
        ThrowError ProjectError.InvalidArgument, _
            "Harvest needs the specifications of the run"
    End If
    If this.writer Is Nothing Then Exit Sub

    Set bundles = SpecificationBundles(specs)
    For index = bundles.LowerBound To bundles.UpperBound
        Collect bundles.Item(index)
    Next index
End Sub

'@label:specification-bundles
'@sub-title The bundles the specification collaborators filed, in harvest order
'@details
'The list Harvest walks, answered on its own so a build step running
'away from the log can carry the same bundles as text. The order is the
'dictionary, the choices, the exports, the analysis, the passwords, the
'format and the specifications themselves. A collaborator with nothing
'filed is skipped. Works on the predeclared instance: it reads no state.
'@param specs LinelistSpecs. The specifications after Prepare.
'@return BetterArray. The Checking bundles, 1-based. Empty when nothing was filed.
'@throws ProjectError.InvalidArgument When specs is Nothing.
'@export
Public Function SpecificationBundles(ByVal specs As LinelistSpecs) As BetterArray
    Dim bundles As BetterArray

    If specs Is Nothing Then
        ThrowError ProjectError.InvalidArgument, _
            "The specification bundles need the specifications of the run"
    End If

    Set bundles = New BetterArray
    bundles.LowerBound = 1

    If specs.Dictionary.HasCheckings Then bundles.Push specs.Dictionary.CheckingValues
    If specs.Choices.HasCheckings Then bundles.Push specs.Choices.CheckingValues
    If specs.ExportObject.HasCheckings Then bundles.Push specs.ExportObject.CheckingValues
    If specs.AnalysisObject.HasCheckings Then bundles.Push specs.AnalysisObject.CheckingValues
    If specs.Password.HasCheckings Then bundles.Push specs.Password.CheckingValues
    If specs.DesignFormat.HasCheckings Then bundles.Push specs.DesignFormat.CheckingValues

    'The specifications themselves. A geobase the user asked for that would
    'not open is filed here, and it used to leave the linelist with no
    'geographic data and nothing said.
    If specs.HasCheckings Then bundles.Push specs.CheckingValues

    Set SpecificationBundles = bundles
End Function

'@label:bundle-text
'@sub-title One bundle as text, ready to be joined with others
'@details
'The title line, then one line per entry, then the bundle end. An empty
'bundle answers an empty string, so a caller may append the answer
'without looking at it. Works on the predeclared instance.
'@param checks Checking. The bundle to carry.
'@param recordOnly Optional Boolean. True marks the bundle for the record alone.
'@return String. The bundle text, or empty when the bundle holds nothing.
'@export
Public Function BundleText(ByVal checks As Checking, _
                           Optional ByVal recordOnly As Boolean = False) As String
    Dim keys As BetterArray
    Dim keyIndex As Long
    Dim keyName As String
    Dim flagText As String
    Dim textOut As String

    If checks Is Nothing Then Exit Function
    If checks.Length = 0 Then Exit Function

    flagText = FLAG_WORKSHEET
    If recordOnly Then flagText = FLAG_RECORD_ONLY

    textOut = checks.CheckingTitle & FIELD_SEP & checks.CheckingSubTitle & _
              FIELD_SEP & flagText & LINE_END

    Set keys = checks.ListOfKeys
    For keyIndex = keys.LowerBound To keys.UpperBound
        keyName = CStr(keys.Item(keyIndex))
        textOut = textOut & keyName & FIELD_SEP & _
                  checks.ValueOf(keyName, checkingLabel) & FIELD_SEP & _
                  ScopeText(checks, keyName) & LINE_END
    Next keyIndex

    BundleText = textOut & BUNDLE_END
End Function

'@label:collect-text
'@sub-title Take every bundle carried in one text into the run
'@details
'The text is what one or more BundleText answers make when joined. Each
'bundle is rebuilt as a Checking with its title, subtitle and scopes and
'handed to Collect with its record-only flag. A line short of its three
'fields is skipped, and a key the bundle already carries is skipped too.
'An empty text, and a run that is closed, take nothing.
'@param bundlesText String. The joined bundle texts.
'@export
Public Sub CollectText(ByVal bundlesText As String)
    Dim bundles() As String
    Dim bundleIndex As Long

    If LenB(bundlesText) = 0 Then Exit Sub
    If this.writer Is Nothing Then Exit Sub

    bundles = Split(bundlesText, BUNDLE_END)
    For bundleIndex = LBound(bundles) To UBound(bundles)
        If LenB(bundles(bundleIndex)) > 0 Then CollectOneText bundles(bundleIndex)
    Next bundleIndex
End Sub

'@label:collect-one-text
'@sub-title Rebuild one bundle from its text and collect it
'@param bundleText String. One bundle without its end mark.
Private Sub CollectOneText(ByVal bundleText As String)
    Dim lines() As String
    Dim fields() As String
    Dim lineIndex As Long
    Dim checks As Checking
    Dim recordOnly As Boolean

    lines = Split(bundleText, LINE_END)

    fields = Split(lines(LBound(lines)), FIELD_SEP)
    If UBound(fields) < 2 Then Exit Sub

    Set checks = Checking.Create(fields(0), fields(1))
    recordOnly = (fields(2) = FLAG_RECORD_ONLY)

    For lineIndex = LBound(lines) + 1 To UBound(lines)
        If LenB(lines(lineIndex)) > 0 Then
            fields = Split(lines(lineIndex), FIELD_SEP)
            If UBound(fields) >= 2 Then
                If Not checks.KeyExists(fields(0)) Then
                    checks.Add fields(0), fields(1), ScopeFromWord(fields(2))
                End If
            End If
        End If
    Next lineIndex

    Collect checks, recordOnly
End Sub

'@label:scope-from-word
'@sub-title The scope a plain scope word stands for
'@details
'The reverse of ScopeText. A word this class does not know answers the
'note scope, the way Checking reads an unknown type label.
'@param scopeWord String. Error, Warning, Info, Success or Note.
'@return Byte. The CheckingScope value.
Private Function ScopeFromWord(ByVal scopeWord As String) As Byte
    Select Case LCase$(Trim$(scopeWord))
        Case "error"
            ScopeFromWord = checkingError
        Case "warning"
            ScopeFromWord = checkingWarning
        Case "info"
            ScopeFromWord = checkingInfo
        Case "success"
            ScopeFromWord = checkingSuccess
        Case Else
            ScopeFromWord = checkingNote
    End Select
End Function

'@label:finish
'@sub-title Close the run
'@details
'Writes the closing bundle (the outcome, the elapsed seconds, and the
'sheet and variable counts when the driver passes them), installs the
'worksheet filter handler, then releases the writer. A workbook that
'refuses the code injection the filter needs gets a warning entry in the
'closing bundle and the run still closes whole. The record stays alive
'for the text export. A log that never started, or already finished,
'ignores the call. The worksheet is left as it is: the driver shows it
'through ShowReport once its screen is back on.
'@param outcome String. The outcome text of the run.
'@param sheetCount Optional Long. Data entry sheets built. A value below 0 skips the entry.
'@param variableCount Optional Long. Variables written. A value below 0 skips the entry.
Public Sub Finish(ByVal outcome As String, _
                  Optional ByVal sheetCount As Long = -1, _
                  Optional ByVal variableCount As Long = -1)
    Dim closing As Checking
    Dim filterFault As Boolean

    If this.writer Is Nothing Then Exit Sub

    'A run that ended while a section was open still writes that section.
    'The closing bundle belongs to the run, so it lands after it, on its own.
    CloseSection

    'The filter handler makes the sheet interactive. The install runs
    'before the closing bundle so a refusal can speak inside it: the
    'report of a finished run has to close whole even on a workbook
    'that refuses the code injection the handler needs.
    On Error Resume Next
    this.writer.EnsureWorksheetChangeHandler
    filterFault = Err.Number <> 0
    On Error GoTo 0

    Set closing = Checking.Create(CLOSING_TITLE)
    closing.Add "outcome", outcome, checkingInfo
    If sheetCount >= 0 Then _
        closing.Add "sheets", CStr(sheetCount) & " data entry sheet(s) built", checkingInfo
    If variableCount >= 0 Then _
        closing.Add "variables", CStr(variableCount) & " variable(s) written", checkingInfo
    If filterFault Then _
        closing.Add "filter", "The severity filter could not be installed on the report sheet", checkingWarning
    closing.Add "elapsed", CStr(DateDiff("s", this.startedAt, Now)) & " seconds", checkingInfo
    Collect closing

    Set this.writer = Nothing
End Sub

'@label:show-report
'@sub-title Bring the report worksheet to the front
'@details
'The one owner of showing the report. The drivers call it once, at the end
'of a run, and the ribbon calls it whenever the user asks for the last
'report, so a report opened either way opens the same way.
'
'WHY THE WORKBOOK IS ACTIVATED FIRST
'-------------------------------------------------------------------------------
'Worksheet.Activate raises 1004 when the sheet belongs to a workbook that is not
'the active one. A generation ends with the built linelist open and in front, so
'the activation Finish used to do raised every time and its On Error Resume Next
'swallowed it. That is why the report never appeared at the end of a run.
'
'The sheet is created very hidden and Activate raises on a hidden sheet, so the
'visibility write comes before it.
'
'An empty sheet is not a report. A designer that has never generated anything
'carries __check as a blank very hidden sheet, and showing that would leave the
'user looking at an empty grid with no way back.
'@return Boolean. True when a report was found and brought to the front.
'@export
Public Function ShowReport() As Boolean
    Dim sh As Worksheet

    If this.designerBook Is Nothing Then Exit Function

    On Error Resume Next
    Set sh = this.designerBook.Worksheets(SHEET_CHECKING)
    On Error GoTo 0

    If sh Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(sh.UsedRange) = 0 Then Exit Function

    On Error Resume Next
    this.designerBook.Activate
    sh.Visible = xlSheetVisible
    sh.Activate
    On Error GoTo 0

    ShowReport = (sh.Visible = xlSheetVisible)
End Function

'@label:export-text
'@sub-title Write the whole record to a text file
'@details
'One line per entry, in the order the run collected them:
'hh:mm:ss  [scope]  bundle-title  entry-label. The file lands in the
'given folder as <baseName>-generation.txt, which puts it beside the
'generated linelist when the caller passes the output folder and the
'linelist name.
'@param folderPath String. The folder taking the file.
'@param baseName String. The file name without its suffix.
'@return String. The full path of the written file.
'@throws ProjectError.ObjectNotInitialized When the run has never started.
Public Function ExportText(ByVal folderPath As String, ByVal baseName As String) As String
    Dim filePath As String
    Dim fileNumber As Long
    Dim index As Long

    If this.record Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "The record is empty, start the log first"
    End If

    filePath = folderPath
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    filePath = filePath & baseName & EXPORT_SUFFIX

    fileNumber = FreeFile

    'The handler closes the file on a mid-write fault; a locked or
    'missing folder raises to the caller with the file released.
    On Error GoTo CloseFile
    Open filePath For Output As #fileNumber
    For index = this.record.LowerBound To this.record.UpperBound
        Print #fileNumber, RecordLine(index)
    Next index
    Close #fileNumber
    On Error GoTo 0

    ExportText = filePath
    Exit Function

CloseFile:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0

    ThrowError errNumber, errDesc
End Function


'@section Internal helpers
'===============================================================================

'@label:resolve-checking-sheet
'@sub-title Resolve the __check worksheet, creating it when absent
'@details
'A created worksheet lands very hidden at the end of the workbook.
'Finish makes it visible once the run closes.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveCheckingSheet() As Worksheet
    Dim sh As Worksheet
    Dim wb As Workbook

    Set wb = this.designerBook

    On Error Resume Next
    Set sh = wb.Worksheets(SHEET_CHECKING)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = SHEET_CHECKING
        sh.Visible = xlSheetVeryHidden
    End If

    Set ResolveCheckingSheet = sh
End Function

'@label:scope-text
'@sub-title The plain scope word of one entry
'@details
'Reads the type descriptor of the entry and strips the icon characters,
'leaving Error, Warning, Info, Note or Success for the record and the
'text lines.
'@param checks Checking. The bundle holding the entry.
'@param keyName String. The entry key.
'@return String. The plain scope word.
Private Function ScopeText(ByVal checks As Checking, ByVal keyName As String) As String
    Dim cleaned As String

    cleaned = checks.ValueOf(keyName, checkingType)
    cleaned = Replace(cleaned, ChrW(10060), vbNullString)
    cleaned = Replace(cleaned, ChrW(9888), vbNullString)
    cleaned = Replace(cleaned, ChrW(8505), vbNullString)
    cleaned = Replace(cleaned, ChrW(9998), vbNullString)
    cleaned = Replace(cleaned, ChrW(10004), vbNullString)

    ScopeText = Trim$(cleaned)
End Function


'@section ErrorHandling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-coded runtime error
'@param errNumb Long. ProjectError enumeration value.
'@param errorMessage String. Human-readable description.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASS_NAME, errorMessage
End Sub
