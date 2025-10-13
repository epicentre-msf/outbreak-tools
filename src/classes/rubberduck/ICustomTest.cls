VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract implemented by CustomTest"

'@Interface
'@Folder("Rubberduck")
'@ModuleDescription("Contract implemented by CustomTest")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle configuration
'===============================================================================

'@label:Configure
'@prop-title Configure the harness to target a workbook and output sheet
'@param targetWorkbook Workbook under test that collects printed output
'@param outputSheetName Optional worksheet name override
Public Sub Configure(ByVal targetWorkbook As workbook, _
                     Optional ByVal outputSheetName As String = vbNullString)
End Sub

'@label:Release
'@prop-title Release workbook references and reset state
Public Sub ReleaseResources()
End Sub

'@label:SetModuleName
'@prop-title Store the VBA module name shown in output
'@param moduleName Friendly module name
Public Sub SetModuleName(ByVal moduleName As String)
End Sub

'@label:SetTestName
'@prop-title Set the label used for the next test entry
'@param testName Descriptive test label
Public Sub SetTestName(ByVal testName As String)
End Sub

'@label:SetTestSubtitle
'@prop-title Add a subtitle for the next test entry
'@param testSubtitle Additional context for the test
Public Sub SetTestSubtitle(ByVal testSubtitle As String)
End Sub

'@label:BeginTest
'@prop-title Persist previous results and prepare a new checking instance
Public Sub BeginTest(Optional ByVal resetNames As Boolean = False)
End Sub

'@label:Reset
'@prop-title Reset pending checking and accumulated results
Public Sub ResetInstance()
End Sub

'@section Results
'===============================================================================

'@label:FlushCurrentTest
'@prop-title Persist the current checking and retrieve all results
'@param resetNames Optional Boolean specifying whether to reset names after clearing
'@return BetterArray containing each Checking snapshot
Public Function FlushCurrentTest(Optional ByVal resetNames As Boolean = True) As BetterArray: End Function

'@label:Flush
'@prop-title Persist the current checking without returning the result buffer
'@param resetNames Optional Boolean specifying whether to reset names after clearing
Public Sub Flush(Optional ByVal resetNames As Boolean = True)
End Sub

'@label:Results
'@prop-title Clone accumulated results without modifying state
'@return BetterArray clone of the stored results
Public Function Results() As BetterArray: End Function

'@label:PrintResults
'@prop-title Write accumulated results to the workbook
'@param outputSheetName Optional worksheet name override
Public Sub PrintResults(Optional ByVal outputSheetName As String = vbNullString)
End Sub

'@section Assertions
'===============================================================================

'@label:AreEqual
'@prop-title Compare expected and actual values and log the outcome
'@param expected Expected value
'@param actual Actual value
'@param message Message describing the assertion context
Public Sub AreEqual(expected As Variant, _
                    actual As Variant, _
                    ByVal message As String)
End Sub

'@label:AreNotEqual
'@prop-title Confirm expected and actual values differ and log the outcome
'@param expected Expected value
'@param actual Actual value
'@param message Message describing the assertion context
Public Sub AreNotEqual(expected As Variant, _
                       actual As Variant, _
                       ByVal message As String)
End Sub

'@label:IsTrue
'@prop-title Validate a boolean condition and log the outcome
'@param condition Condition under test
'@param message Message describing the assertion context
Public Sub IsTrue(ByVal condition As Boolean, ByVal message As String)
End Sub

'@label:IsFalse
'@prop-title Validate that a boolean condition evaluates to False
'@param condition Condition under test
'@param message Message describing the assertion context
Public Sub IsFalse(ByVal condition As Boolean, ByVal message As String)
End Sub

'@label:ObjectExists
'@prop-title Validate that an object is set and matches the expected type name
'@param instance Object reference to validate
'@param expectedTypeName Expected VBA type name
'@param message Message describing the assertion context
Public Sub ObjectExists(ByVal instance As Object, _
                        ByVal expectedTypeName As String, _
                        ByVal message As String)
End Sub

'@label:IsNothing
'@prop-title Validate that an object reference is Nothing
'@param instance Object reference to validate
'@param message Message describing the assertion context
Public Sub IsNothing(ByVal instance As Object, ByVal message As String)
End Sub

'@jump:IsNotNothing
'@prop-title Validate that an object is not Nothing
'@param instance Object reference to validate
'@param message Message describing the assertion context
Public Sub IsNotNothing(ByVal instance As Object, ByVal message As String)
End Sub

'@label:LogFailure
'@prop-title Log a failure message to the current checking
'@param message Message describing the failure
Public Sub LogFailure(ByVal message As String)
End Sub

'@label:Fail
'@prop-title Log a failure message to the current checking; for compatibility purposes
Public Sub Fail(ByVal message As String)
End Sub

'@label:LogSuccesses
'@prop-title Log a success message to the current checking
'@param message Message describing the success
Public Sub LogSuccesses(ByVal message As String)
End Sub
