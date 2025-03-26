VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for dictionary class"

'@Folder("Dictionary")
'@ModuleDescription("Interface for dictionary class")
'@PredeclaredId

Option Explicit
Implements ILLdictionary

'Types of the class
Private Type TDictionary
    dictDataSheet As IDataSheet                  'Worksheet data of the dictionary
End Type

Private this As TDictionary
Private DICTIONARYCOLUMNS As BetterArray

Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartRow As Long, _
                       ByVal dictStartColumn As Long) As ILLdictionary

    Dim customDataSheet As IDataSheet

    'Will throw an error if something went bad with the creation
    Set customDataSheet = DataSheet.Create(dictWksh, dictStartRow, dictStartColumn)

    With New LLdictionary
        Set .Data = customDataSheet
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLdictionary
    Set Self = Me
End Property

'Data sheet
Public Property Get Data() As IDataSheet
    Set Data = this.dictDataSheet
End Property

Friend Property Set Data(ByVal customDataSheet As IDataSheet)
    Set this.dictDataSheet = customDataSheet
End Property

'Worksheet
Public Property Get Wksh() As Worksheet
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    Set Wksh = customDataSheet.Wksh()
End Property

'End row and columns of the dictionary
Private Property Get DictEndRow() As Long
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    DictEndRow = customDataSheet.DataEndRow()
End Property

'Initialize the Headers  -----------------------------------------------------
Private Sub ClassArraysInitialize()

    Set DICTIONARYCOLUMNS = New BetterArray

    'colnames of the dictionary
    DICTIONARYCOLUMNS.Push "variable name", "main label", "sub label", "note", _
                           "sheet name", _
                           "sheet type", "main section", "sub section", "status", "personal identifier", "type", "format", _
                           "control", "control details", "unique", _
                           "min", "max", "alert", "message", "table name", "column index", _
                           "list auto", "visibility"
End Sub

'Dictionary data Range
Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional includeHeaders As Boolean = False) As Range
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    Set DataRange = customDataSheet.DataRange(colName, includeHeaders)
End Property

'Functions ---------------------------------------------------------------------
'Check if colname exists
Private Function ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean = False) As Boolean
    Dim customDataSheet As IDataSheet
    Dim exist As Boolean

    Set customDataSheet = Data()
    exist = customDataSheet.ColumnExists(colName)

    If exist And checkValidity Then
        If DICTIONARYCOLUMNS Is Nothing Then ClassArraysInitialize
        exist = (DICTIONARYCOLUMNS.Includes(colName) Or (Split(colName, " ")(0) = "export"))
    End If

    ColumnExists = exist
End Function

'Insert a column
Private Sub InsertColumn(ByVal colName As String, ByVal After As String)

    Dim prevRng As Range
    Dim hRng As Range
    Dim sh As Worksheet

    'Test if the column to add one after exists
    If Not ColumnExists(after) Then Exit Sub
    If ColumnExists(colName) Then Exit Sub

    Set prevRng = DataRange(after, includeHeaders:=True)
    Set sh = Wksh()

    'insert a column
    prevRng.Offset(, 1).EntireColumn.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
    prevRng.Offset(, 1).Cells(1, 1).Value = colName
    Set hRng = sh.Range(prevRng.Cells(-1, 1), prevRng.Cells(0,2))
    On Error Resume Next
        hRng.Merge
    On Error GoTo 0
End Sub

'Remove a column
Private Sub RemoveColumn(ByVal colName As String)

    Dim actRng As Range

    'Test if the column exists before removing it
    If Not ColumnExists(colName) Then Exit Sub
    
    Set actRng = DataRange(colName)
    On Error Resume Next
    actRng.EntireColumn.Delete
    On Error GoTo 0
End Sub

'Check Existence of a variable
Private Function VariableExists(ByVal varName As String) As Boolean
    Dim varRng As Range
    Set varRng = DataRange("variable name")

    VariableExists = Not (varRng.Find(What:=varName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

'Add or remove Rows to the dictionary
Private Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim tabl As ICustomTable

    Set sh = Wksh()

    On Error Resume Next
    Set Lo = sh.ListObjects(1)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(Lo)
    If del Then
        tabl.RemoveRows
    Else
        tabl.AddRows
    End If
End Sub

'Get the choices Var, GeoVar or time vars
Private Property Get SpecialVars(ByVal first_condition As String, _
                                Optional ByVal second_condition As String = vbNullString, _
                                Optional ByVal conditionName As String = "Control") As BetterArray
    Dim lsVars As BetterArray
    Dim ctrlRng As Range
    Dim varRng As Range
    Dim endrw As Long
    Dim counter As Long
    Dim ctrlValue As String
    Dim sheetRng As Range
    Dim sheetValue As String

    Set lsVars = New BetterArray
    If (ColumnExists(conditionName) And ColumnExists("Variable Name") And ColumnExists("Sheet Type")) Then
        Set ctrlRng = DataRange(conditionName)
        Set varRng = DataRange("Variable Name")
        Set sheetRng = DataRange("Sheet Type")
        endrw = DictEndRow()

        For counter = 1 To endrw
            ctrlValue = ctrlRng.Cells(counter, 1).Value
            sheetValue = sheetRng.Cells(counter, 1).Value
            'Two conditions on control
            If (second_condition <> vbNullString) Then
                If (sheetValue = "hlist2D") And ((ctrlValue = first_condition) Or _
                    (ctrlValue = second_condition)) Then _
                    lsVars.Push varRng.Cells(counter, 1).Value
            Else
            'only one condition on control
                If (sheetValue = "hlist2D") And (ctrlValue = first_condition) _
                 Then lsVars.Push varRng.Cells(counter, 1).Value
            End If
        Next
    End If

    Set SpecialVars = lsVars.Clone()
End Property

' INTERFACE __________________________________________________________________________________________________________________________

Private Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                            Optional ByVal includeHeaders As Boolean = False) As Range
    Set ILLdictionary_DataRange = DataRange(colName, includeHeaders)
End Property

Private Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = Data()
End Property

'Functions
Private Function ILLdictionary_ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean) As Boolean
    ILLdictionary_ColumnExists = ColumnExists(colName, checkValidity)
End Function

Private Function ILLdictionary_VariableExists(ByVal varName As String) As Boolean
    ILLdictionary_VariableExists = VariableExists(varName)
End Function

Private Sub ILLdictionary_AddRows()
    ManageRows
End Sub

Private Sub ILLdictionary_RemoveRows()
    ManageRows del:=True
End Sub

Private Property Get ILLdictionary_ChoicesVars() As BetterArray
    Set ILLdictionary_ChoicesVars = SpecialVars(first_condition:="choice_manual", _
                                                second_condition:="choice_formula")
End Property

Private Property Get ILLdictionary_GeoVars() As BetterArray
    Set ILLdictionary_GeoVars = SpecialVars(first_condition:="geo", _
                                            second_condition:="hf")
End Property

Private Property Get ILLdictionary_TimeVars() As BetterArray
    Set ILLdictionary_TimeVars = SpecialVars(first_condition:="date", _
                                             conditionName:="Variable Type")
End Property

Private Property Get ILLdictionary_SpecialVars( _ 
                                ByVal first_condition As String, _
                                Optional ByVal second_condition As String = vbNullString, _
                                Optional ByVal conditionName As String = "Control") As BetterArray

    Set ILLdictionary_SpecialVars = SpecialVars(first_condition, _ 
                                        second_condition, _ 
                                        conditionName)

End Property

Private Sub ILLdictionary_InsertColumn(ByVal colName As String, ByVal After As String)
    InsertColumn colName:=colName, After:=After    
End Sub

Private Sub ILLdictionary_RemoveColumn(ByVal colName As String)
    RemoveColumn colName:=colName
End Sub
