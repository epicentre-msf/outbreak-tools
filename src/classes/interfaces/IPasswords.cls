VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface of Password")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Public Property Get Password() As String: End Property

'available values are : "debuggingpassword", "publickey", "privatekey", "version"
Public Property Get Value(ByVal tagName As String) As String: End Property
Public Property Get TableRange() As Range: End Property

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
