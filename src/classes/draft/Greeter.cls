VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Greeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds greeting strings; a probe class for the AppleScript test loop"

'@PredeclaredId
'@Folder("Draft")
'@ModuleDescription("Builds greeting strings; a probe class for the AppleScript test loop")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

Private Type TGreeter
    name As String
End Type

Private this As TGreeter

'@section Factory
'===============================================================================

'@sub-title Build a Greeter with an optional starting name.
'@param startName Optional String. Person to greet. Defaults to "world".
'@return Greeter ready for use.
Public Function Create(Optional ByVal startName As String = "world") As Greeter
    Dim instance As Greeter
    Set instance = New Greeter
    instance.Configure startName
    Set Create = instance
End Function

'@sub-title Set the name on a fresh instance.
'@param startName String. Person to greet.
Public Sub Configure(ByVal startName As String)
    this.name = startName
End Sub

'@section Members
'===============================================================================

'@prop-title The name currently being greeted.
'@return String.
Public Property Get name() As String
    name = this.name
End Property

'@sub-title Change the name being greeted.
'@param newName String. The replacement name.
Public Sub Rename(ByVal newName As String)
    this.name = newName
End Sub

'@sub-title Build the greeting for the current name.
'@return String. "Hello, <name>!"
Public Function Greeting() As String
    Greeting = "Hello, " & this.name & "!"
End Function

'@sub-title Build the greeting in upper case.
'@return String. The greeting shouted.
Public Function LoudGreeting() As String
    LoudGreeting = UCase$(Greeting())
End Function

'@sub-title Report whether no usable name is set.
'@return Boolean. True when the name is blank or whitespace only.
Public Function IsAnonymous() As Boolean
    IsAnonymous = (Len(Trim$(this.name)) = 0)
End Function
