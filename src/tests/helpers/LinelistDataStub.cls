VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistDataStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist data stub exposing TransObject and Categories for the analysis tests"

'@Folder("Tests.Helpers")
'@ModuleDescription("Linelist data stub exposing TransObject and Categories for the analysis tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class LinelistDataStub
'@description
'Stands in for the linelist specifications object that CrossTable is handed at
'creation. CrossTable asks that object for two things only: TransObject, for
'the translated labels it writes, and Categories, which TableSpecs asks on its
'behalf when it resolves the row and column categories of a table. This stub
'answers both, and a test seeds each one so the layout under test is driven by
'known data.
'
'WHY A STUB AND NOT THE REAL CLASS
'-------------------------------------------------------------------------------
'LinelistSpecs is 1,536 lines and its own dependency closure names six more
'classes across four folders. CrossTable.Create types the parameter Object for
'that reason, so the test workbook needs none of them.
'@depends BetterArray, TranslationObject

Private Type TLinelistDataStub
    translations As TranslationObject
    categoryLists As Collection
End Type

Private this As TLinelistDataStub

'@section Seeding
'===============================================================================
'@description Members a test calls to load the stub before handing it over.

'@label:set-trans-object
'@sub-title Store the translation object the stub hands back.
'@param trans TranslationObject. The translation object to store.
Public Sub SetTransObject(ByVal trans As TranslationObject)
    Set this.translations = trans
End Sub

'@label:set-categories
'@sub-title Store the categories of one variable.
'@details
'Later calls for the same variable name replace the earlier list. Collection
'keys are matched without regard to case, which is how the dictionary matches
'a variable name too.
'
'The list is copied into a one-based array whatever bounds it arrived with,
'because LLChoices.Categories sets LowerBound to 1 and everything that reads a
'category list reads it from there. A stub answering a zero-based list would
'test the readers against a shape they never meet.
'@param varName String. The variable name the categories belong to.
'@param categories BetterArray. The category list to answer for that variable.
Public Sub SetCategories(ByVal varName As String, ByVal categories As BetterArray)
    Dim stored As BetterArray
    Dim idx As Long

    If this.categoryLists Is Nothing Then Set this.categoryLists = New Collection

    Set stored = New BetterArray
    stored.LowerBound = 1
    If Not categories Is Nothing Then
        For idx = categories.LowerBound To categories.UpperBound
            stored.Push categories.Item(idx)
        Next idx
    End If

    On Error Resume Next
    this.categoryLists.Remove varName
    On Error GoTo 0

    this.categoryLists.Add stored, varName
End Sub

'@label:set-categories-verbatim
'@sub-title Store the categories of one variable with the bounds they arrived with.
'@details
'SetCategories normalises to a one-based list because that is what the real
'choices class answers. This member skips the normalisation, so a test can hand
'a reader a list starting at zero and check that the reader copes. One reader
'used to walk index one to Length regardless of where the list started, which
'read one past the end of a zero-based list.
'@param varName String. The variable name the categories belong to.
'@param categories BetterArray. The category list to answer, bounds untouched.
Public Sub SetCategoriesVerbatim(ByVal varName As String, _
                                 ByVal categories As BetterArray)
    If this.categoryLists Is Nothing Then Set this.categoryLists = New Collection

    On Error Resume Next
    this.categoryLists.Remove varName
    On Error GoTo 0

    this.categoryLists.Add categories, varName
End Sub

'@label:clear-categories
'@sub-title Forget every seeded category list.
Public Sub ClearCategories()
    Set this.categoryLists = New Collection
End Sub

'@section Members CrossTable reads
'===============================================================================

'@label:trans-object
'@prop-title The translation object for locale-aware labels.
'@return TranslationObject. The stored translation object.
Public Property Get TransObject() As TranslationObject
    Set TransObject = this.translations
End Property

'@label:categories
'@sub-title Categories of one variable.
'@details
'Answers the list a test seeded for varName. A variable with nothing seeded
'answers an empty list, which is what the real class answers for a variable
'that carries no choices.
'@param varName String. The variable whose categories are wanted.
'@return BetterArray. The seeded categories, or an empty list.
Public Function Categories(ByVal varName As String) As BetterArray
    Dim found As BetterArray

    If Not this.categoryLists Is Nothing Then
        On Error Resume Next
        Set found = this.categoryLists.Item(varName)
        On Error GoTo 0
    End If

    If found Is Nothing Then
        Set found = New BetterArray
        found.LowerBound = 1
    End If

    Set Categories = found
End Function
