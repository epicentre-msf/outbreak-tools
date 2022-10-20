VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Dictionary Class"

'@Folder("Dictionary")
'@ModuleDescription("Interface for dicitonary in the linelist")
'@IgnoreModule
'@PredeclareId
Option Explicit
Implements ILLdictionary

'Types of the class
Private Type TDictionary
    dictDataSheet As IDataSheet 'Worksheet data of the dictionary
End Type

Private this As TDictionary
Private DICTIONARYCOLUMNS As BetterArray
Private PRESERVEDSHEETSNAMES As BetterArray
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"


Private Const CLASS_NAME As String = "ILLdictionary"

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

'StartRow
Public Property Get StartRow() As Long
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    StartRow = customDataSheet.StartRow()
End Property

'StartColumn
Public Property Get StartColumn() As Long
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    StartColumn = customDataSheet.StartColumn()
End Property

'End row and columns of the dictionary
Private Property Get DictEndRow() As Long
   Dim customDataSheet As IDataSheet
   Set customDataSheet = Data()
   DictEndRow = customDataSheet.DataEndRow()
End Property

Private Property Get DictEndColumn() As Long
    Dim customDataSheet As IDataSheet
   Set customDataSheet = Data()
   DictEndColumn = customDataSheet.DataEndColumn()
End Property

'Is the dictionary prepared for linelist?
Public Property Get Prepared() As Boolean
    Dim dictWksh As Worksheet
    Set dictWksh = Wksh()
    Prepared = (dictWksh.Cells(DictEndRow + 1, StartColumn).Font.Color = vbBlue) And _
                ColumnExists("table name") And ColumnExists("column index") And _
                (Not ColumnExists("main section sort")) And (Not ColumnExists("main section index"))
End Property

'Initialize the Headers  -----------------------------------------------------
Private Sub ClassArraysInitialize()

    Set DICTIONARYCOLUMNS = New BetterArray

    'colnames of the dictionary
    DICTIONARYCOLUMNS.Push "variable name", "main label", "sub label", "note", "sheet name", _
     "sheet type", "main section", "sub section", "status", "personal identifier", "type", "control", "control details", "unique", _
     "export 1", "export 2", "export 3", "export 4", "export 5", "min", "max", "alert", "message", "table name", "column index"

    'Preserved names
    Set PRESERVEDSHEETSNAMES = New BetterArray
    PRESERVEDSHEETSNAMES.Push "GEO", "ControleFormule", "Password", "Temp_", _
    "LinelistTranslation", "List_auto_", "Dictionary", "Exports", "Choices", "Translations", _
    "Metadata", "Ana_temp_", "Import_temp_", "Admin", "Debug___"

End Sub

'Dictionary data Range
Public Property Get DataRange(Optional colName As String = "__all__", Optional includeHeaders As Boolean = False) As Range
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    Set DataRange = customDataSheet.DataRange(colName, includeHeaders)
End Property

Private Property Get variables() As ILLVariables
    Dim var As ILLVariables
    Set var = LLVariables.Create(Self)
End Property

'Functions ---------------------------------------------------------------------
'Check if colname exists
Private Function ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean = False) As Boolean
    Dim customDataSheet As IDataSheet
    Dim Exist As Boolean

    Set customDataSheet = Data()
    Exist = customDataSheet.ColumnExists(colName)

    If Exist And checkValidity Then
        If DICTIONARYCOLUMNS Is Nothing Then ClassArraysInitialize
        Exist = DICTIONARYCOLUMNS.Includes(colName)
    End If

    ColumnExists = Exist
End Function

'Column Index within the data range of the dictionary
Private Function ColumnIndex(ByVal colName As String, Optional ByVal inDataRange As Boolean = False) As Long
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    ColumnIndex = customDataSheet.ColumnIndex(colName, inDataRange, shouldExist:=True)
End Function

'Get Unique values of one column
Private Function UniqueValues(colName As String) As BetterArray

    Dim varColIndex As Long
    Dim varValue As String
    Dim counter As Long
    Dim endRow As Long
    Dim varData As BetterArray
    Dim dictWksh As Worksheet

    endRow = DictEndRow
    counter = StartRow + 1
    varColIndex = ColumnIndex(colName)
    Set varData = New BetterArray
    Set dictWksh = Wksh()

    With dictWksh
        Do While counter < endRow
            varValue = .Cells(counter, varColIndex).Value
            If Not varData.Includes(varValue) Then varData.Push varValue
            counter = counter + 1
        Loop
    End With

    Set UniqueValues = varData.Clone()
End Function

'Check Existence of a variable
Private Function VariableExists(varName As String) As Boolean
    Dim varRng As Range
    Set varRng = DataRange("variable name")

    VariableExists = Not (varRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

'------ Prepare and clean the dictionary
Private Sub Clean()
    Dim endCol As Long
    Dim counter As Long
    Dim dictWksh As Worksheet

    Set dictWksh = Wksh()
    endCol = DictEndColumn
    counter = endCol

    With dictWksh
        Do While counter >= 1
            If Not ColumnExists(.Cells(StartRow, counter).Value, checkValidity:=True) Then
                .Columns(counter).EntireColumn.Delete
            End If

            counter = counter - 1
         Loop
    End With
End Sub


'Prepare the dictionary for designer
Private Sub Prepare(Optional ByVal tablePrefix As String = "table")
    Dim endCol As Long
    Dim dictWksh As Worksheet
    Dim counter As Long
    Dim controlRng As Range

    If Not Prepared() Then
            'Test if the dictionary is prepared before doing the whole loop
            If Not ColumnExists("table name") Then
                 AppendNumberColumn onColName:="sheet name", tablePrefix:=tablePrefix, _
                                    newColumn:="table name"    'Add table name
            End If

            If Not ColumnExists("column index") Then AppendColumnIndex 'Add Column Index
            'Testing control to update the Geo
            Set controlRng = DataRange("control")

            If (Not controlRng Is Nothing) Then
                If (Not controlRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True) Is Nothing) And _
                    (controlRng.Find(What:="geo2", LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
                    'Update the Geo
                    AppendGeoLines
                End If
            End If
            'Sort the dictionary accordingly and make unique varNames
            UniqueVarNames
            GoodSheetNames
            Clean
            'Add color to the last line for prepare indication
            Set dictWksh = Wksh()
            dictWksh.Cells(DictEndRow + 1, StartColumn).Font.Color = vbBlue
    End If
End Sub


Private Sub GoodSheetNames()
    Dim dictWksh As Worksheet
    Dim sheetNameIndex As Long

    Dim endRow As Long
    Dim counter As Long

    endRow = DictEndRow
    Set dictWksh = Wksh()
    sheetNameIndex = ColumnIndex("sheet name")

    If (PRESERVEDSHEETSNAMES Is Nothing) Then ClassArraysInitialize

    With dictWksh
        For counter = StartRow + 1 To endRow
            If PRESERVEDSHEETSNAMES.Includes(.Cells(counter, sheetNameIndex).Value) Then
                .Cells(counter, sheetNameIndex).Value = .Cells(counter, sheetNameIndex).Value & "_"
                'Designer Loggin
            End If
        Next
    End With

End Sub

'Make unique variable names
Private Sub UniqueVarNames()

    Dim dictWksh As Worksheet
    Dim endRow As Long
    Dim varNameIndex As Long
    Dim counter As Long
    Dim varName As String

    Dim varTrackData As BetterArray

    Set varTrackData = New BetterArray
    Set dictWksh = Wksh()
    endRow = DictEndRow
    varNameIndex = ColumnIndex("variable name")

    With dictWksh
        For counter = StartRow + 1 To endRow
            'remove no breakable spaces and clean the name of the variable. replace spaces with _

            varName = Application.WorksheetFunction.SUBSTITUTE(varName, Chr(160), " ")
            varName = Application.WorksheetFunction.Clean(varName)
            varName = Application.WorksheetFunction.Trim(.Cells(counter, varNameIndex).Value)
            varName = Replace(varName, " ", "_")
            If Len(varName) < 3 Then varName = "var_" & varName 'Should be at least 4 letters.

            If varTrackData.Includes(varName) Then varName = varName & "_"

            varTrackData.Push varName
            .Cells(counter, varNameIndex).Value = varName
        Next
    End With
End Sub


'Append Geo Lines
Private Sub AppendGeoLines()

    Dim dictRng As Range
    Dim geoRng As Range
    Dim pastedGeoRng As Range
    Dim dictWksh As Worksheet
    Dim geoRowIndex As Long
    Dim controlIndex As Long
    Dim varNameIndex As Long
    Dim subSectionIndex As Long
    Dim mainLabelIndex As Long
    Dim varName As String
    Dim endRow As Long
    Dim endCol As Long

    Dim counter As Long

    Set dictWksh = Wksh()
    Set dictRng = DataRange("control")
    controlIndex = ColumnIndex("control")
    varNameIndex = ColumnIndex("variable name")
    mainLabelIndex = ColumnIndex("main label")
    subSectionIndex = ColumnIndex("sub section")

    endRow = DictEndRow
    endCol = DictEndColumn
    geoRowIndex = StartRow + 1

    With dictWksh
        Do While (Not dictRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True) Is Nothing)

            geoRowIndex = dictRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True).Row

            Set geoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
            varName = .Cells(geoRowIndex, varNameIndex).Value

            'Replace sub-section by main label (if required)
            If IsEmpty(.Cells(geoRowIndex, subSectionIndex)) And (Not IsEmpty(.Cells(geoRowIndex, mainLabelIndex))) Then
                .Cells(geoRowIndex, subSectionIndex).Value = .Cells(geoRowIndex, mainLabelIndex).Value
                .Cells(geoRowIndex, mainLabelIndex).Value = ""
            End If

            .Cells(geoRowIndex, varNameIndex).Value = "adm1" & "_" & varName

            'Insert the others geo lines
            geoRowIndex = geoRowIndex + 1
            For counter = 4 To 2 Step -1
                .Rows(geoRowIndex).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Set pastedGeoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
                pastedGeoRng.Value = geoRng.Value 'Paste dictionary values
                .Cells(geoRowIndex, varNameIndex).Value = "adm" & counter & "_" & varName
                .Cells(geoRowIndex, controlIndex).Value = "geo" & counter
            Next

            geoRowIndex = geoRowIndex + 3
            Set dictRng = .Range(.Cells(geoRowIndex, controlIndex), .Cells(endRow, controlIndex))
        Loop
    End With
End Sub

Private Sub AppendNumberColumn(Optional ByVal onColName As String = "sheet name", _
                               Optional ByVal tablePrefix As String = "table", _
                               Optional ByVal newColumn As String = "table name")

    Dim endRow As Long
    Dim endCol As Integer
    Dim tableOrder As Long
    Dim sheetNameIndex As Long
    Dim sheetName As String
    Dim counter As Long

    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim sheetTrackData As BetterArray
    Dim tableTrackData As BetterArray

    Set sheetTrackData = New BetterArray
    Set tableTrackData = New BetterArray

    endRow = DictEndRow
    endCol = DictEndColumn
    tableOrder = 1
    sheetNameIndex = ColumnIndex(onColName)

    Set dictWksh = Wksh()

    With dictWksh
        .Cells(StartRow, endCol + 1).Value = "number_" & newColumn
        sheetTrackData.Push .Cells(StartRow + 1, sheetNameIndex).Value
        tableTrackData.Push tableOrder

        For counter = StartRow + 1 To endRow
            sheetName = .Cells(counter, sheetNameIndex).Value
            If sheetTrackData.Includes(sheetName) Then
                .Cells(counter, endCol + 1).Value = tableTrackData.Items(sheetTrackData.IndexOf(sheetName))
            Else
                tableOrder = tableOrder + 1
                sheetTrackData.Push sheetName
                tableTrackData.Push tableOrder
                .Cells(counter, endCol + 1).Value = tableOrder
            End If
        Next

        'Sort on table name
        Set dictRng = DataRange
        dictRng.Sort key1:=DataRange("number_" & newColumn), order1:=xlAscending

        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).Value = newColumn
        .Cells(StartRow + 1, endCol + 1).Formula = "= " & Chr(34) & tablePrefix & Chr(34) & " " & Chr(38) & " " & .Cells(StartRow + 1, endCol).Address(Rowabsolute:=False)
        Set dictRng = DataRange(newColumn)

        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues
        dictRng.Value = dictRng.Value 'remove the link to the formulas

    End With
End Sub


Private Sub AppendColumnIndex()

    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim endCol As Long
    Dim endRow As Long
    Dim counter As Long
    Dim varOrder As Long
    Dim sheetTypeIndex As Long
    Dim sheetNameIndex As Long
    Dim controlIndex As Long
    Dim mainSectionIndex As Long

    Dim firstsheetNameCell As String
    Dim firstmainSectionCell As String
    Dim sheetName As String

    Set dictWksh = Wksh()
    endCol = DictEndColumn
    endRow = DictEndRow
    sheetTypeIndex = ColumnIndex("sheet type")
    sheetNameIndex = ColumnIndex("sheet name")
    mainSectionIndex = ColumnIndex("main section")
    controlIndex = ColumnIndex("control")

    sheetName = "&&&&&&;-------"
    varOrder = 1

    With dictWksh
        'Sorting
        .Cells(StartRow, endCol + 1).Value = "main section sort"
        firstsheetNameCell = .Cells(StartRow + 1, sheetNameIndex).Address(Rowabsolute:=False) & " "
        firstmainSectionCell = .Cells(StartRow + 1, mainSectionIndex).Address(Rowabsolute:=False)
        .Cells(StartRow + 1, endCol + 1).Formula = "= IF(" & firstmainSectionCell & "=" & Chr(34) & Chr(34) & _
                                                    ", ROW(" & firstmainSectionCell & ")" & ", " & _
                                                   firstsheetNameCell & Chr(38) & Chr(34) & " - " & Chr(34) & Chr(38) & firstmainSectionCell & _
                                                   ")"
        Set dictRng = DataRange("main section sort")
        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues

        'Now append and sort
        AppendNumberColumn onColName:="main section sort", tablePrefix:="main section", newColumn:="main section index"

        'Filling column index
        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).Value = "column index"
        For counter = StartRow + 1 To endRow
            If .Cells(counter, sheetNameIndex).Value <> sheetName Then
                'New sheet name
                varOrder = IIf(.Cells(counter, sheetTypeIndex).Value = SHEETTYPEVERTICAL, 4, 1)
                sheetName = .Cells(counter, sheetNameIndex).Value
            Else
                'Old sheet name
                varOrder = IIf(.Cells(counter - 1, controlIndex).Value = "geo", varOrder + 4, varOrder + 1)
            End If
            .Cells(counter, endCol + 1).Value = varOrder 'variable order
        Next
    End With
End Sub


' INTERFACE __________________________________________________________________________________________________________________________

'Properties
Public Property Get ILLdictionary_Prepared() As Boolean
    ILLdictionary_Prepared = Prepared
End Property

Public Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                            Optional ByVal includeHeaders As Boolean = False) As Range
    Set ILLdictionary_DataRange = DataRange(colName, includeHeaders)
End Property

Public Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = Data()
End Property

'Public Property Get ILLdictionary_Variables() As ILLVariables
 '   Set ILLdictionary_Variables = variables
'End Property

'Functions
Public Function ILLdictionary_ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean) As Boolean
    ILLdictionary_ColumnExists = ColumnExists(colName, checkValidity)
End Function

Public Function ILLdictionary_UniqueValues(ByVal colName As String) As BetterArray
    Set ILLdictionary_UniqueValues = UniqueValues(colName)
End Function

Public Function ILLdictionary_VariableExists(ByVal varName As String) As Boolean
    ILLdictionary_VariableExists = VariableExists(varName)
End Function

'void methods
Public Sub ILLdictionary_Prepare(Optional ByVal tablePrefix As String = "tab")
   Prepare tablePrefix
End Sub

Public Sub ILLdictionary_Clean()
    Clean
End Sub

