# OBT Documentation Template

This file shows the canonical layout for a fully documented implementation
class and its matching interface. Use it as a blueprint when documenting any
OutbreakTools VBA class.

---

## Table of Contents

1. [Implementation Class Template](#1-implementation-class-template)
2. [Interface Class Template](#2-interface-class-template)
3. [Section Ordering Conventions](#3-section-ordering-conventions)
4. [Minimal Doc Block (Private Helpers)](#4-minimal-doc-block-private-helpers)
5. [Full Doc Block (Public Members)](#5-full-doc-block-public-members)
6. [Common Patterns](#6-common-patterns)

---

## 1. Implementation Class Template

Below is the full layout of a documented implementation class. Every tag shown
with `<angle brackets>` is a placeholder you fill in. Comments starting with
`' >>` are instructions -- do not include them in the real file.

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "<ClassName>"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "<One-line summary matching @ModuleDescription>"
Option Explicit


' >> Rubberduck annotations (preserve as-is if already present)
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("<TopicFolder>")
'@ModuleDescription("<One-line summary>")


' >> Interface declaration (if this class implements an interface)
Implements I<ClassName>

' >> Class-level documentation block
'@class <ClassName>
'@description
'<2-5 sentence overview. Explain:
' - What this class does
' - Who consumes it (which modules or workflows call it)
' - How it fits into the project architecture
' - Any important design decisions (e.g. wraps a DataSheet for
'   consistent range access)
'Write so a beginner VBA developer can understand the class's role.>
'@depends <ClassA>, <ClassB>, <ClassC>
'@version <version or date>
'@author <author if known>


' >> Private state (UDT backing store and module-level variables)
Private Type T<ClassName>
    <field1> As <Type>
    <field2> As <Type>
End Type

Private this As T<ClassName>

' >> Constants
Private Const CLASS_NAME As String = "<ClassName>"
Private Const <CONSTANT_NAME> As <Type> = <value>

' >> Checking support (if the class uses IChecking)
Private checkCounter As Long
Private internalChecks As IChecking


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a <ClassName> instance
'@details
'<Explain what Create does in 2-4 sentences. Describe the factory pattern:
'this is a predeclared method that returns a new instance through the
'interface. Mention what gets initialised and any validation.>
'@param <param1Name> <Type>. <Description.>
'@param <param2Name> <Type>. <Description.>
'@param <param3Name> Optional <Type>. <Description. Defaults to <value>.>
'@return I<ClassName>. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When <condition>.
'@depends <ClassA>, <ClassB>
'@export
Public Function Create(<params>) As I<ClassName>
    ' ... implementation ...
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the interface
'reference from the predeclared Create method.
'@return I<ClassName>. The current instance cast to the interface.
Public Property Get Self() As I<ClassName>
    Set Self = Me
End Property


'@section Data Elements
'===============================================================================
'@description Properties that expose the internal state of the class.

'@label:data
'@prop-title <PropertyName> backing store
'@details
'<Explain what this property exposes and why callers need it.>
'@return <Type>. <Description.>
Public Property Get <PropertyName>() As <Type>
    ' ... implementation ...
End Property

'@label:data-set
'@prop-title Assign the <PropertyName> backing store
'@param <paramName> <Type>. <Description.>
Public Property Set <PropertyName>(ByVal <paramName> As <Type>)
    ' ... implementation ...
End Property


'@section Core Operations
'===============================================================================
'@description Methods that perform the primary business logic of this class.

'@label:<operation-id>
'@sub-title <Imperative summary of what the method does>
'@details
'<Detailed explanation covering:
' - Step-by-step behaviour
' - What happens with edge-case inputs (Nothing, empty string, zero)
' - Side effects (worksheet mutations, cache resets, persisted values)
' - Why this method exists (context for a newcomer)>
'@param <paramName> <Type>. <Description.>
'@return <Type>. <Description.>
'@throws ProjectError.<ErrorName> When <condition>.
'@remarks <Implementation note for maintainers, if any.>
'@note <Important caveat for callers, if any.>
Public Function <MethodName>(<params>) As <ReturnType>
    ' ... implementation ...
End Function


'@section Helpers
'===============================================================================
'@description Private helper methods that support the public API.

'@label:<helper-id>
'@sub-title <Brief summary>
'@details
'<Explain what the helper does and why it is factored out.>
'@param <paramName> <Type>. <Description.>
Private Sub <HelperName>(<params>)
    ' ... implementation ...
End Sub


'@section Checking
'===============================================================================
'@description Logging and validation support via the IChecking interface.

'@label:log-info
'@sub-title Log an informational or warning message
'@param message String. The message to record.
'@param logType Byte. Severity level from CheckingLogType.
Private Sub LogInfo(ByVal message As String, ByVal logType As Byte)
    ' ... implementation ...
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber ProjectError. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members that satisfy the I<ClassName> contract.
'These methods forward calls to the corresponding Public members above.

Private Property Get I<ClassName>_<MemberName>() As <Type>
    ' delegates to public member
End Property

Private Sub I<ClassName>_<MethodName>(<params>)
    ' delegates to public member
End Sub
```

---

## 2. Interface Class Template

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "I<ClassName>"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of <ClassName> class"
Option Explicit


'@Interface
'@Folder("<TopicFolder>")
'@ModuleDescription("Interface of <ClassName> class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@class I<ClassName>
'@description
'<1-2 sentence summary of the interface contract. Example:
'"Defines the public contract for the <ClassName> class, exposing
'methods for <main capability 1>, <main capability 2>, and <main
'capability 3>.">


'@section Data Access
'===============================================================================

'@jump:data
'@prop-title <PropertyName> backing store
'@details
'<Same description as the implementation, written for the caller's perspective.>
'@return <Type>. <Description.>
Public Property Get <PropertyName>() As <Type>: End Property

'@jump:data-set
'@prop-title Assign the <PropertyName> backing store
'@param <paramName> <Type>. <Description.>
Public Property Set <PropertyName>(ByVal <paramName> As <Type>): End Property


'@section Core Operations
'===============================================================================

'@jump:<operation-id>
'@sub-title <Imperative summary>
'@details
'<Explain what this does from the caller's perspective. The interface doc
'should focus on WHAT, not HOW, since the implementation details are in
'the implementation class.>
'@param <paramName> <Type>. <Description.>
'@return <Type>. <Description.>
Public Function <MethodName>(<params>) As <ReturnType>: End Function


'@section Data Exchange
'===============================================================================

'@jump:import
'@sub-title Import data from another worksheet
'@param fromWksh Worksheet. Source worksheet to import from.
'@param fromStartRow Long. Starting row of the source data (1-based).
'@param fromStartcol Long. Starting column of the source data (1-based).
'@param clearSheet Optional Boolean. When True, clears the target before importing. Defaults to False.
Public Sub Import(ByVal fromWksh As Worksheet, _
                  ByVal fromStartRow As Long, _
                  ByVal fromStartcol As Long, _
                  Optional ByVal clearSheet As Boolean = False)
End Sub
```

---

## 3. Section Ordering Conventions

Sections should appear in a logical progression from creation to use to
cleanup. The recommended order is:

1. **Instantiation** -- `Create`, `Self`, constructors
2. **Data Elements** -- Properties that expose internal state
3. **Core Operations** -- The primary business-logic methods
4. **Specialised Views** -- Filtered or computed read-only accessors
5. **Preparation** -- Setup or initialisation workflows
6. **Data Exchange** -- Import, export, translate
7. **Validation** -- Checks, assertions, guards
8. **Helpers** -- Private utility methods
9. **Checking** -- Logging and error-raising helpers
10. **Interface Implementation** -- `IClassName_*` delegation stubs

Not every class needs all sections. Use only what applies, but keep the
ordering consistent across files.

---

## 4. Minimal Doc Block (Private Helpers)

For simple private methods (one-liners, pure delegation), a minimal block is
acceptable:

```vba
'@label:start-row
'@prop-title Header row index
'@return Long. The 1-based row index of the header.
Private Property Get StartRow() As Long
    StartRow = Data().DataStartRow
End Property
```

The minimum is: `@label` + `@prop-title` or `@sub-title` + `@return` (if it
returns a value).

---

## 5. Full Doc Block (Public Members)

For public members, every applicable tag must be present. Here is a fully
documented example:

```vba
'@label:create
'@sub-title Create a dictionary object
'@details
'Creates a linelist dictionary wrapper around the supplied worksheet. The
'routine initialises a backing DataSheet so all subsequent access goes through
'a consistent abstraction and callers interact with a predeclared instance.
'When no export count is supplied, the method falls back to the class
'constant DEFAULTNUMBEROFEXPORTS (20). The returned object is cast to the
'ILLdictionary interface to enforce programming against the contract rather
'than the concrete class.
'@param dictWksh Worksheet. The worksheet hosting the dictionary data.
'@param dictStartRow Long. Header row of the dictionary (1-based).
'@param dictStartColumn Long. First column of the dictionary (1-based).
'@param numberOfExports Optional Long. Number of export columns to configure. Defaults to 20.
'@return ILLdictionary. A fully initialised dictionary instance ready for use.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@depends DataSheet, Checking
'@export
Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartRow As Long, _
                       ByVal dictStartColumn As Long, _
                       Optional ByVal numberOfExports As Long = DEFAULTNUMBEROFEXPORTS) As ILLdictionary

    Dim customDataSheet As IDataSheet

    Set customDataSheet = DataSheet.Create(dictWksh, dictStartRow, dictStartColumn)

    With New LLdictionary
        Set .Data = customDataSheet
        .InitialiseTotalExports numberOfExports
        Set Create = .Self
    End With
End Function
```

---

## 6. Common Patterns

### Factory Pattern (Create + Self)

Most OutbreakTools classes use predeclared instances with a `Create` factory:

```vba
'@label:create
'@sub-title Create an <X> instance
'@details ...
'@export
Public Function Create(...) As IClassName
    With New ClassName
        ' set properties
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@return IClassName. The current instance.
Public Property Get Self() As IClassName
    Set Self = Me
End Property
```

### Property Get/Set Pair

Document both halves, linking them through related labels:

```vba
'@label:exports
'@prop-title Number of export columns
'@details
'Number of export slots currently configured. Returns the default when
'the stored value is invalid.
'@return Long. The current export column count.
Public Property Get TotalNumberOfExports() As Long
    ...
End Property

'@label:exports-set
'@prop-title Update the number of export columns
'@details
'Persist the new count while falling back to the default for invalid values.
'@param numberOfExports Long. Requested number of export slots.
Public Property Let TotalNumberOfExports(ByVal numberOfExports As Long)
    ...
End Property
```

### Checking/Logging Pattern

```vba
'@label:log-info
'@sub-title Log an informational or warning message
'@details
'Routes the message to the internal IChecking instance when available.
'Silently exits when no checking object is configured, ensuring the
'method is always safe to call.
'@param message String. The message to record.
'@param logType Byte. Severity level from CheckingLogType enum values.
Private Sub LogInfo(ByVal message As String, ByVal logType As Byte)
    ...
End Sub
```

### Error Handling Pattern

```vba
'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber ProjectError. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
```

### Interface Delegation

The implementation stubs at the bottom of the file typically need only a
minimal doc block because the public member they delegate to is already
documented:

```vba
'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the I<ClassName> contract.
'See the corresponding Public members above for full documentation.

Private Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = Data
End Property

Private Sub ILLdictionary_AddColumn(ByVal colName As String)
    AddColumn colName
End Sub
```

These delegation stubs do not need individual `@label`, `@details`, or
`@param` tags -- the full documentation lives on the public member.
