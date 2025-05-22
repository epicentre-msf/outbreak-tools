VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Passwords"

'@Interface
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of Passwords")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit


'@jump:Create

'@jump:elements
'@section Elements

'available values are : "debuggingpassword", "publickey", "privatekey", "version"
Public Property Get Value(ByVal tagName As String) As String: End Property
Public Property Get TableRange(ByVal tabName As String) As Range: End Property

Public Sub UnProtect(Optional ByVal sheetName As String = "_active"): End Sub

Public Sub Protect(Optional ByVal sheetObject As Variant, _
                   Optional ByVal allowShapes As Boolean = False, _ 
                   Optional ByVal allowDeletingRows As Boolean = True)
End Sub

Public Sub EnterDebugMode(): End Sub

Public Sub LeaveDebugMode(): End Sub

'Export password sheet
Public Sub Export(ByVal toWkb As Workbook): End Sub

'Protect and unprotect a given workbook to avoid sheet name changes.
Public Sub ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
End Sub

Public Sub UnProtectWkb(ByVal wb As Workbook): End Sub

'Generate new key and display current keys
Public Sub DisplayPrivateKey(ByVal trads As ITranslationObject): End Sub

Public Sub GenerateKey(ByVal trads As ITranslationObject): End Sub
Public Sub Import(ByVal pass As IPasswords): End Sub


'@jump:HasCheckings
'Test if the object has checkings enable
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Get the checking values
Public Property Get CheckingValues() As Object: End Property