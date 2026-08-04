VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoiceKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The key a custom dropdown is stored under on a Choices sheet"

'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("The key a custom dropdown is stored under on a Choices sheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ProcedureNotUsed

Option Explicit

'@class ChoiceKeys
'@description
'The one place the export side and the import side build the name a custom
'dropdown is written under on a Choices sheet. CustomChoiceName builds the key
'from a dropdown name, IsCustomChoice says whether a key belongs to a custom
'dropdown, and ListNameFromChoice reads the dropdown name back out of a key.
'Both ends call these, so the two ends of a round trip agree.
'
'THE NAME THE KEY IS BUILT FROM
'-------------------------------------------------------------------------------
'The dropdown name held in the DropdownLists registry, exactly as the user
'typed it. Nothing is trimmed, nothing is cleaned and no counter is added.
'`DropdownLists.AllDropdowns` answers that list and `DropdownLists.Update`
'takes the same name, so a key built this way survives the trip in both
'directions.
'@version 1.0 (2026-08-01)


'@section Constants
'===============================================================================

'The prefix every custom dropdown carries on a Choices sheet
Private Const CUSTOM_PREFIX As String = "__choice_custom_"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a key builder
'@details
'The class holds no state. It is a class because both LLExporter and LLImporter
'need the same answer and a class is how this project shares behaviour.
'@return ChoiceKeys. A ready instance.
'@export
Public Function Create() As ChoiceKeys
    Dim newKeys As ChoiceKeys
    Set newKeys = New ChoiceKeys
    Set Create = newKeys
End Function


'@section Keys
'===============================================================================

'@label:custom-choice-name
'@sub-title The Choices sheet key of a custom dropdown
'@details
'An empty dropdown name answers an empty key. `LLChoices.Categories` reads an
'empty criteria as "blank cells" rather than as "nothing", so handing it a bare
'prefix would answer rows that belong to no dropdown.
'@param listName String. The dropdown name held in the DropdownLists registry.
'@return String. The key the choice is stored under, empty when the name is empty.
'@export
Public Function CustomChoiceName(ByVal listName As String) As String
    If LenB(listName) = 0 Then Exit Function
    CustomChoiceName = CUSTOM_PREFIX & listName
End Function

'@label:is-custom-choice
'@sub-title Whether a Choices sheet key names a custom dropdown
'@param choiceName String. A key read off a Choices sheet.
'@return Boolean. True when the key carries the custom dropdown prefix.
'@export
Public Function IsCustomChoice(ByVal choiceName As String) As Boolean
    IsCustomChoice = (InStr(1, choiceName, CUSTOM_PREFIX, vbTextCompare) = 1)
End Function

'@label:list-name-from-choice
'@sub-title The dropdown name a Choices sheet key was built from
'@details
'The inverse of CustomChoiceName. A key that carries no custom prefix answers an
'empty string, so a reader walking a whole Choices sheet can tell the custom
'dropdowns apart from the choices the setup file authored.
'@param choiceName String. A key read off a Choices sheet.
'@return String. The dropdown name, empty when the key is not a custom one.
'@export
Public Function ListNameFromChoice(ByVal choiceName As String) As String
    If Not IsCustomChoice(choiceName) Then Exit Function
    ListNameFromChoice = Mid$(choiceName, Len(CUSTOM_PREFIX) + 1)
End Function
