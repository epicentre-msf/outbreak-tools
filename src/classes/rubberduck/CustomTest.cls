VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Harness providing assertion helpers and worksheet output for custom tests"


'@PredeclaredId
'@Folder("Rubberduck")
'@ModuleDescription("Harness providing assertion helpers and worksheet output for custom tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit
Implements ICustomTest

Private Type THarnessState
    targetWorkbook As workbook
    moduleName As String
    outputSheetName As String
    pendingTestName As String
    pendingTestSubtitle As String
    currentChecking As IChecking
    Results As BetterArray
    assertCounter As Long
End Type

Private this As THarnessState

Private Const CLASSNAME As String = "CustomTest"
Private Const DEFAULT_OUTPUT_SHEET As String = "TestOutput"
Private Const DEFAULT_TEST_LABEL As String = "test"
Private Const KEY_PREFIX As String = "ASSERT_"

'@section Factory
'===============================================================================

'@label:Create
'@prop-title Build and configure a harness instance
'@param targetWorkbook Workbook under test that receives logging output
'@param outputSheetName Optional name of the worksheet used for results
'@return ICustomTest harness ready for use
'@export
Public Function Create(ByVal targetWorkbook As workbook, _
                       Optional ByVal outputSheetName As String = vbNullString) As ICustomTest
    With New CustomTest
        .Configure targetWorkbook, outputSheetName
        Set Create = .Self
    End With
End Function

'@label:Self
'@prop-title Access the current harness reference
'@return ICustomTest reference for fluent usage
'@export
Public Property Get Self() As ICustomTest
    Set Self = Me
End Property

'@section ICustomTest implementation
'===============================================================================

'@label:ICustomTest_Configure
'@prop-title Forward interface calls to public Configure
'@param targetWorkbook Workbook that stores test output
'@param outputSheetName Optional target worksheet name
Private Sub ICustomTest_Configure(ByVal targetWorkbook As workbook, _
                                  Optional ByVal outputSheetName As String = vbNullString)
    Configure targetWorkbook, outputSheetName
End Sub

'@label:ICustomTest_ReleaseResources
'@prop-title Bridge interface ReleaseResources to public member
Private Sub ICustomTest_ReleaseResources()
    ReleaseResources
End Sub

'@label:ICustomTest_SetModuleName
'@prop-title Bridge interface SetModuleName to public member
Private Sub ICustomTest_SetModuleName(ByVal moduleName As String)
    SetModuleName moduleName
End Sub

'@label:ICustomTest_SetTestName
'@prop-title Bridge interface SetTestName to public member
Private Sub ICustomTest_SetTestName(ByVal testName As String)
    SetTestName testName
End Sub

'@label:ICustomTest_SetTestSubtitle
'@prop-title Bridge interface SetTestSubtitle to public member
Private Sub ICustomTest_SetTestSubtitle(ByVal testSubtitle As String)
    SetTestSubtitle testSubtitle
End Sub

'@label:ICustomTest_BeginTest
'@prop-title Bridge interface BeginTest to public member
Private Sub ICustomTest_BeginTest(Optional ByVal resetNames As Boolean = False)
    BeginTest resetNames
End Sub

'@label:ICustomTest_Reset
'@prop-title Bridge interface Reset to public member
Private Sub ICustomTest_ResetInstance()
    ResetInstance
End Sub

'@label:ICustomTest_FlushCurrentTest
'@prop-title Return results snapshot for interface callers
Private Function ICustomTest_FlushCurrentTest(Optional ByVal resetNames As Boolean = True) As BetterArray
    Set ICustomTest_FlushCurrentTest = FlushCurrentTest(resetNames)
End Function

'@label:ICustomTest_Flush
'@prop-title Bridge interface Flush to public member
Private Sub ICustomTest_Flush(Optional ByVal resetNames As Boolean = True)
    Flush resetNames
End Sub

'@label:ICustomTest_Results
'@prop-title Provide a clone of accumulated results through the interface
Private Function ICustomTest_Results() As BetterArray
    Set ICustomTest_Results = Results()
End Function

'@label:ICustomTest_PrintResults
'@prop-title Bridge interface PrintResults to public member
Private Sub ICustomTest_PrintResults(Optional ByVal outputSheetName As String = vbNullString)
    PrintResults outputSheetName
End Sub

'@label:ICustomTest_AreEqual
'@prop-title Bridge interface AreEqual to public member
Private Sub ICustomTest_AreEqual(expected As Variant, _
                                 actual As Variant, _
                                    ByVal message As String)
    AreEqual expected, actual, message
End Sub

'@label:ICustomTest_AreNotEqual
'@prop-title Bridge interface AreNotEqual to public member
Private Sub ICustomTest_AreNotEqual(expected As Variant, _
                                    actual As Variant, _
                                       ByVal message As String)
    AreNotEqual expected, actual, message
End Sub

'@label:ICustomTest_IsTrue
'@prop-title Bridge interface IsTrue to public member
Private Sub ICustomTest_IsTrue(ByVal condition As Boolean, _
                                   ByVal message As String)
    IsTrue condition, message
End Sub

'@label:ICustomTest_IsFalse
'@prop-title Bridge interface IsFalse to public member
Private Sub ICustomTest_IsFalse(ByVal condition As Boolean, _
                                    ByVal message As String)
    IsFalse condition, message
End Sub

'@label:ICustomTest_ObjectExists
'@prop-title Bridge interface ObjectExists to public member
Private Sub ICustomTest_ObjectExists(ByVal instance As Object, _
                                     ByVal expectedTypeName As String, _
                                     ByVal message As String)
    ObjectExists instance, expectedTypeName, message
End Sub

'@label:ICustomTest_IsNothing
'@prop-title Bridge interface IsNothing to helper member
Private Sub ICustomTest_IsNothing(ByVal instance As Object, ByVal message As String)
    IsNothing instance, message
End Sub

'@label:ICustomTest_IsNotNothing
'@prop-title Bridge interface IsNotNothing to public member
Private Sub ICustomTest_IsNotNothing(ByVal instance As Object, ByVal message As String)
    IsNotNothing instance, message
End Sub

'@label:ICustomTest_LogFailure
'@prop-title Bridge interface LogFailure to public member
Private Sub ICustomTest_LogFailure(ByVal message As String)
    LogFailure message
End Sub

'@label:ICustomTest_Fail
'@prop-title Bridge interface for Fail, same as Logfailure, for compatibility purposes
Private Sub ICustomTest_Fail(ByVal message As String)
    LogFailure message
End Sub

'@label:ICustomTest_LogSuccesses
'@prop-title Bridge interface LogSuccesses to public member
Private Sub ICustomTest_LogSuccesses(ByVal message As String)
    LogSuccesses message
End Sub

'@section Public members
'===============================================================================

'@label:Configure
'@prop-title Configure the harness with the workbook and target sheet
'@param targetWorkbook Workbook that supplies and receives output
'@param outputSheetName Optional sheet name where results will be written
'@throws ProjectError.InvalidArgument When workbook reference is invalid
'@export
Public Sub Configure(ByVal targetWorkbook As workbook, _
                     Optional ByVal outputSheetName As String = vbNullString)
    If targetWorkbook Is Nothing Then
        RaiseError ProjectError.InvalidArgument, "Configure requires a valid workbook reference"
    End If

    Set this.targetWorkbook = targetWorkbook
    this.outputSheetName = ResolveSheetName(outputSheetName)
    this.moduleName = DEFAULT_TEST_LABEL
    this.pendingTestName = DEFAULT_TEST_LABEL
    this.pendingTestSubtitle = DEFAULT_TEST_LABEL
    this.assertCounter = 0
    Set this.currentChecking = Nothing
    Set this.Results = Nothing
    EnsureResults
End Sub

'@label:ReleaseResources
'@prop-title Clear references and reset harness state
'@export
Private Sub ReleaseResources()
    Set this.currentChecking = Nothing
    Set this.Results = Nothing
    Set this.targetWorkbook = Nothing
    this.moduleName = DEFAULT_TEST_LABEL
    this.pendingTestName = DEFAULT_TEST_LABEL
    this.pendingTestSubtitle = DEFAULT_TEST_LABEL
    this.outputSheetName = DEFAULT_OUTPUT_SHEET
    this.assertCounter = 0
End Sub

'@label:SetModuleName
'@prop-title Set the VBA module name used in rendered output
'@param moduleName Name displayed in result tables
'@export
Public Sub SetModuleName(ByVal moduleName As String)
    this.moduleName = ResolveTitle(moduleName)
End Sub

'@label:SetTestName
'@prop-title Set the pending test name for subsequent assertions
'@param testName Descriptive identifier of the current test
'@export
Public Sub SetTestName(ByVal testName As String)
    this.pendingTestName = ResolveTitle(testName)
End Sub

'@label:SetTestSubtitle
'@prop-title Set an optional subtitle for the pending test
'@param testSubtitle Additional context for the test case
'@export
Public Sub SetTestSubtitle(ByVal testSubtitle As String)
    this.pendingTestSubtitle = ResolveTitle(testSubtitle)
End Sub

'@label:BeginTest
'@prop-title Persist previous checking and start logging a new test
'@export
Public Sub BeginTest(Optional ByVal resetNames As Boolean = False)
    PersistActiveChecking resetNames
    If resetNames Then
        ResetPendingTitles
    End If
    InitialiseCurrentChecking
End Sub

'@label:Reset
'@prop-title Clear current checking and accumulated results
'@export
Public Sub ResetInstance()
    ClearCurrentChecking True
    Set this.Results = Nothing
    EnsureResults
End Sub

'@label:FlushCurrentTest
'@prop-title Persist active checking and return accumulated results
'@param resetNames Optional Boolean specifying whether to reset names after clearing
'@return BetterArray containing all persisted checkings
'@export
Public Function FlushCurrentTest(Optional ByVal resetNames As Boolean = True) As BetterArray
    PersistActiveChecking resetNames
    Set FlushCurrentTest = Results()
End Function

'@label:Flush
'@prop-title Persist active checking without returning the accumulated results
'@param resetNames Optional Boolean specifying whether to reset names after clearing
'@export
Public Sub Flush(Optional ByVal resetNames As Boolean = True)
    PersistActiveChecking resetNames
End Sub

'@label:Results
'@prop-title Clone accumulated checking results
'@return BetterArray clone of collected checkings
'@export
Public Function Results() As BetterArray
    If this.Results Is Nothing Then
        EnsureResults
    End If
    Set Results = this.Results.Clone
End Function

'@label:PrintResults
'@prop-title Write the results to a worksheet using CheckingOutput
'@param outputSheetName Optional override for worksheet name
'@depends CheckingOutput
'@export
Public Sub PrintResults(Optional ByVal outputSheetName As String = vbNullString)
    Dim sheetName As String
    Dim targetSheet As Worksheet
    Dim writer As ICheckingOutput

    If this.targetWorkbook Is Nothing Then Exit Sub
    If this.Results Is Nothing Then Exit Sub
    If this.Results.Length = 0 Then Exit Sub

    sheetName = this.outputSheetName
    If LenB(outputSheetName) > 0 Then
        sheetName = ResolveSheetName(outputSheetName)
        this.outputSheetName = sheetName
    End If

    Set targetSheet = PrintOutputWorksheet(this.targetWorkbook, sheetName)
    Set writer = CheckingOutput.Create(targetSheet, this.moduleName)
    writer.PrintOutput this.Results
    writer.EnsureWorksheetChangeHandler
End Sub

'@label:AreEqual
'@prop-title Log an assertion comparing expected and actual values
'@param expected Expected value
'@param actual Actual value
'@param message Message describing the assertion context
'@export
Public Sub AreEqual(expected As Variant, _
                    actual As Variant, _
                       ByVal message As String)
    Dim label As String
    Dim scope As Byte

    EnsureCurrentChecking
    If VariantsAreEqual(expected, actual) Then
        scope = checkingSuccess
        label = FormatSuccessMessage(message)
    Else
        scope = checkingError
        label = BuildFailureMessage(message, expected, actual)
    End If

    LogEntry label, scope
End Sub

'@label:AreNotEqual
'@prop-title Log an assertion confirming expected and actual values differ
'@param expected Expected value
'@param actual Actual value
'@param message Message describing the assertion context
'@export
Public Sub AreNotEqual(expected As Variant, _
                       actual As Variant, _
                          ByVal message As String)
    Dim label As String
    Dim scope As Byte

    EnsureCurrentChecking
    If VariantsAreEqual(expected, actual) Then
        scope = checkingError
        label = BuildFailureMessage(message, "Values to differ", _
                                    "Values matched: " & VariantToString(actual))
    Else
        scope = checkingSuccess
        label = FormatSuccessMessage(message)
    End If

    LogEntry label, scope
End Sub

'@label:IsTrue
'@prop-title Log that a boolean condition holds true
'@param condition Condition to evaluate
'@param message Message describing the assertion context
'@export
Public Sub IsTrue(ByVal condition As Boolean, _
                      ByVal message As String)
    EnsureCurrentChecking
    If condition Then
        LogEntry FormatSuccessMessage(message), checkingSuccess
    Else
        LogEntry ResolveFailurePrefix(message), checkingError
    End If
End Sub

'@label:IsFalse
'@prop-title Log that a boolean condition evaluates to False
'@param condition Condition to evaluate
'@param message Message describing the assertion context
'@export
Public Sub IsFalse(ByVal condition As Boolean, _
                       ByVal message As String)
    EnsureCurrentChecking
    If condition Then
        LogEntry ResolveFailurePrefix(message), checkingError
    Else
        LogEntry FormatSuccessMessage(message), checkingSuccess
    End If
End Sub

'@label:ObjectExists
'@prop-title Confirm an object reference exists and matches the expected type name
'@param instance Object reference to validate
'@param expectedTypeName Expected VBA type name
'@param message Message describing the assertion context
'@export
Public Sub ObjectExists(ByVal instance As Object, _
                        ByVal expectedTypeName As String, _
                        ByVal message As String)
    Dim resolvedType As String
    Dim actualType As String
    Dim expectation As String

    EnsureCurrentChecking

    resolvedType = Trim$(expectedTypeName)
    If LenB(resolvedType) = 0 Then
        resolvedType = "Object"
    End If

    expectation = "Instance of type '" & resolvedType & "'"

    If instance Is Nothing Then
        LogEntry BuildFailureMessage(message, expectation, "Nothing"), checkingError
        Exit Sub
    End If

    actualType = TypeName(instance)
    If StrComp(actualType, resolvedType, vbTextCompare) = 0 Then
        LogEntry FormatSuccessMessage(message), checkingSuccess
    Else
        LogEntry BuildFailureMessage(message, expectation, _
                                     "Instance of type '" & actualType & "'"), checkingError
    End If
End Sub

'@label:IsNothing
'@prop-title Validate that an object reference is Nothing
'@param instance Object reference to validate
'@param message Message describing the assertion context
Private Sub IsNothing(ByVal instance As Object, ByVal message As String)
    If instance Is Nothing Then
        LogEntry FormatSuccessMessage(message), checkingSuccess
    Else
        LogEntry BuildFailureMessage(message, "Nothing", _
                                     "Instance of type '" & TypeName(instance) & "'"), checkingError
    End If
End Sub

'@label:IsNotNothing
'@prop-title Validate that an object is not Nothing
'@param instance Object reference to validate
'@param message Message describing the assertion context
Private Sub IsNotNothing(ByVal instance As Object, ByVal message As String)
    If instance Is Nothing Then
        LogEntry BuildFailureMessage(message, "Expected: Not Nothing", "Actual: Nothing"), checkingError
    Else
        LogEntry FormatSuccessMessage(message), checkingSuccess
    End If
End Sub

'@label:LogFailure
'@prop-title Log a failure message directly to the current checking
'@param message Message describing the failure
'@export
Public Sub LogFailure(ByVal message As String)
    EnsureCurrentChecking
    LogEntry ResolveFailurePrefix(message), checkingError
End Sub

'@label:LogSuccesses
'@prop-title Log a success message directly to the current checking
'@param message Message describing the success
'@export
Public Sub LogSuccesses(ByVal message As String)
    EnsureCurrentChecking
    LogEntry FormatSuccessMessage(message), checkingSuccess
End Sub

'@section Internal helpers
'===============================================================================

'@label:InitialiseCurrentChecking
'@prop-title Instantiate a fresh checking object for the active test
Private Sub InitialiseCurrentChecking()
    Dim testName As String
    Dim testSubtitle As String

    testName = ResolveTitle(this.pendingTestName)
    testSubtitle = ResolveTitle(this.pendingTestSubtitle)
    Set this.currentChecking = checking.Create(testName, testSubtitle)
    this.assertCounter = 0
End Sub

'@label:EnsureCurrentChecking
'@prop-title Lazily initialise the current checking container
Private Sub EnsureCurrentChecking()
    If this.currentChecking Is Nothing Then
        InitialiseCurrentChecking
    End If
End Sub

'@label:ClearCurrentChecking
'@prop-title Drop the current checking reference and optionally reset names
'@param resetNames When True also reset pending titles to defaults
Private Sub ClearCurrentChecking(Optional ByVal resetNames As Boolean = True)
    Set this.currentChecking = Nothing
    this.assertCounter = 0
    If resetNames Then
        ResetPendingTitles
    End If
End Sub

Private Sub ResetPendingTitles()
    this.pendingTestName = DEFAULT_TEST_LABEL
    this.pendingTestSubtitle = DEFAULT_TEST_LABEL
End Sub

'@label:PersistActiveChecking
'@prop-title Store the active checking into the results buffer when it has entries
Private Sub PersistActiveChecking(Optional ByVal resetNames As Boolean = False)
    Dim persisted As IChecking

    If this.currentChecking Is Nothing Then Exit Sub
    If this.currentChecking.Length = 0 Then
        ClearCurrentChecking resetNames
        Exit Sub
    End If

    Set persisted = this.currentChecking.Clone
    EnsureResults.Push persisted
    ClearCurrentChecking resetNames
End Sub

'@label:EnsureResults
'@prop-title Create the BetterArray container when missing and return it
'@return BetterArray used to hold persisted checkings
Private Function EnsureResults() As BetterArray
    If this.Results Is Nothing Then
        Set this.Results = New BetterArray
        this.Results.lowerBound = 1
    End If
    Set EnsureResults = this.Results
End Function

'@label:LogEntry
'@prop-title Append a formatted assertion outcome to the current checking
'@param message Text describing the result
'@param scope Checking scope (success, error, ...)
Private Sub LogEntry(ByVal message As String, ByVal scope As Byte)
    Dim key As String

    EnsureCurrentChecking
    this.assertCounter = this.assertCounter + 1
    key = KEY_PREFIX & Format$(this.assertCounter, "000000")
    this.currentChecking.Add key, message, scope
End Sub

'@label:FormatSuccessMessage
'@prop-title Build the message stored when an assertion succeeds
'@param mess Optional message provided by the caller
'@return Formatted message prefixed with Success
Private Function FormatSuccessMessage(ByVal mess As String) As String
    Dim message As String

    message = Trim$(mess)
    If LenB(message) = 0 Then
        FormatSuccessMessage = "Success"
    Else
        FormatSuccessMessage = "Success: " & message
    End If
End Function

'@label:BuildFailureMessage
'@prop-title Build the failure message including expected and actual values
'@param message Assertion context supplied by caller
'@param expected Expected value
'@param actual Actual value
'@return Failure message containing expected/actual details
Private Function BuildFailureMessage(ByVal message As String, _
                                     expected As Variant, _
                                     actual As Variant) As String
    Dim prefix As String
    prefix = ResolveFailurePrefix(message)
    BuildFailureMessage = prefix & " (expected: " & VariantToString(expected) & _
                           ", actual: " & VariantToString(actual) & ")"
End Function

'@label:ResolveTitle
'@prop-title Normalise names used for modules and tests
'@param val Candidate name to store
'@return Sanitised name or default label
Private Function ResolveTitle(ByVal val As String) As String
    Dim value As String

    value = Trim$(val)
    If LenB(value) = 0 Then
        ResolveTitle = DEFAULT_TEST_LABEL
    Else
        ResolveTitle = value
    End If
End Function

'@label:ResolveFailurePrefix
'@prop-title Normalize assertion failure prefix text
'@param mess Optional message provided by the caller
'@return Prefix used for failure messages
Private Function ResolveFailurePrefix(ByVal mess As String) As String
    Dim message As String
    message = Trim$(mess)
    If LenB(message) = 0 Then
        ResolveFailurePrefix = "Assertion failed"
    Else
        ResolveFailurePrefix = message
    End If
End Function

'@label:ResolveSheetName
'@prop-title Normalise worksheet names used for output
'@param shName Optional sheet name override
'@return Validated sheet name or default output sheet
Private Function ResolveSheetName(ByVal shName As String) As String
    Dim sheetName As String
    sheetName = Trim$(shName)
    If LenB(sheetName) = 0 Then
        ResolveSheetName = DEFAULT_OUTPUT_SHEET
    Else
        ResolveSheetName = sheetName
    End If
End Function

'@label:PrintOutputWorksheet
'@prop-title Retrieve or create the worksheet that receives test output
'@param targetWorkbook Workbook where results will be printed
'@param sheetName Worksheet name to resolve
'@return Worksheet ready for output
Private Function PrintOutputWorksheet(ByVal targetWorkbook As workbook, _
                                 ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    On Error Resume Next
        Set sh = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = targetWorkbook.Worksheets.Add
        sh.Name = sheetName
    End If

    Set PrintOutputWorksheet = sh
End Function

'@label:VariantsAreEqual
'@prop-title Compare two Variant values using tolerant rules
'@param leftValue First value to compare
'@param rightValue Second value to compare
'@return True when the values are equal
Private Function VariantsAreEqual(leftValue As Variant, _
                                  rightValue As Variant) As Boolean
    On Error GoTo NumericCompare
    If VarType(leftValue) = vbString Or VarType(rightValue) = vbString Then
        VariantsAreEqual = (StrComp(CStr(leftValue), CStr(rightValue), vbTextCompare) = 0)
        Exit Function
    End If
    VariantsAreEqual = (leftValue = rightValue)
    Exit Function
NumericCompare:
    VariantsAreEqual = (CStr(leftValue) = CStr(rightValue))
End Function

'@label:VariantToString
'@prop-title Convert a Variant into a readable String representation
'@param value Variant to convert
'@return String representation of the value
Private Function VariantToString(value As Variant) As String
    On Error GoTo Clean
    Select Case VarType(value)
        Case vbString: VariantToString = CStr(value)
        Case vbBoolean: VariantToString = IIf(CBool(value), "True", "False")
        Case vbDate: VariantToString = Format$(value, "yyyy-mm-dd hh:nn:ss")
        Case vbNull: VariantToString = "Null"
        Case vbError: VariantToString = "#Error"
        Case Else: VariantToString = CStr(value)
    End Select
    Exit Function
Clean:
    VariantToString = "<error>"
End Function

'@label:RaiseError
'@prop-title Wrap Err.Raise to keep consistent error source
'@param errorCode Project error identifier
'@param description Message supplied to Err.Raise
Private Sub RaiseError(ByVal errorCode As ProjectError, ByVal description As String)
    Err.Raise errorCode, CLASSNAME, description
End Sub
