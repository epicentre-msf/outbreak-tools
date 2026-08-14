Attribute VB_Name = "FormLogicShowVarLabels"
Attribute VB_Description = "Form code-behind for F_ShowVarLabels"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowVarLabels")
'@depends EventLinelist, BetterArray, TranslationObject, LLTranslation

'@description
'The code behind F_ShowVarLabels. The form shows one row per hlist2D variable
'of the dictionary: the title of the pivot block its table was given, the
'variable name and the main label.
'
'Nothing is passed in. The rows come from EventLinelist.VarLabelTable, which
'reads the held dictionary and variable reader, so the form fills itself and
'the button that opens it only has to show it.

Option Explicit

'Pivot title, variable name and main label -- one per header label the form
'carries above its list.
Private Const VARLAB_COLUMNS As Long = 3

Private tradform As TranslationObject


'The one event service of the running linelist.
Private Function LinelistService() As EventLinelist
    Set LinelistService = LinelistEventsManager.EventLinelistService()
End Function

'The translation helper is the one EventLinelist holds. This module used to
'build its own on every open, and LLTranslation.Create validates all five
'translation tables per build.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    Set linelistEvents = LinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicShowVarLabels", _
                  "This linelist carries no usable translation sheet"

    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

'Put the pivot title, variable name and label of every hlist2D variable in the
'three-column list.
'
'VarLabelTable answers a BetterArray of one-dimensional rows, and BetterArray
'reports that shape as JAGGED, so `Items` gives back an array OF ARRAYS. A
'jagged array is not something a ListBox can take: the button that opens this
'form used to assign `List = table.Items` and the form came up showing nothing
'at all. The rows go in one cell at a time, the way
'FormLogicImportRep.FillTwoColumnList already does it for the same reason.
Private Sub LoadVarLabelList()
    Dim linelistEvents As EventLinelist
    Dim labelRows As BetterArray
    Dim listControl As Object
    Dim entry As Variant
    Dim counter As Long
    Dim rowIndex As Long
    Dim firstColumn As Long

    Set listControl = Me.LST_CustomTabList
    listControl.Clear
    listControl.ColumnCount = VARLAB_COLUMNS

    Set linelistEvents = LinelistService()
    If linelistEvents Is Nothing Then Exit Sub

    Set labelRows = linelistEvents.VarLabelTable()
    If labelRows Is Nothing Then Exit Sub

    For counter = labelRows.LowerBound To labelRows.UpperBound
        entry = labelRows.Item(counter)
        If IsArray(entry) Then
            firstColumn = LBound(entry)
            listControl.AddItem CStr(entry(firstColumn))
            listControl.List(rowIndex, 1) = CStr(entry(firstColumn + 1))
            listControl.List(rowIndex, 2) = CStr(entry(firstColumn + 2))
            rowIndex = rowIndex + 1
        End If
    Next counter
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 700
    Me.Height = 380
End Sub

'The list is filled here and not in Initialize. Initialize runs once per form
'instance, and this form is shown on the predeclared instance and hidden rather
'than unloaded, so a second click would show whatever the first click read. An
'import or a rebuild between the two opens changes those rows.
Private Sub UserForm_Activate()
    LoadVarLabelList
End Sub
