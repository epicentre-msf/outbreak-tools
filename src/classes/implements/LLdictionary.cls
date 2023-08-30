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

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Interface for dicitonary in the linelist")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLdictionary

'Types of the class
Private Type TDictionary
    dictDataSheet As IDataSheet                  'Worksheet data of the dictionary
End Type

Private this As TDictionary
Private DICTIONARYCOLUMNS As BetterArray
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"
Private Const CLASSNAME As String = "ILLdictionary"

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

Public Property Set Data(ByVal customDataSheet As IDataSheet)
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
    Prepared = (dictWksh.Cells(DictEndRow + 1, StartColumn).Font.color = vbBlue) And _
                                                                                 ColumnExists("table name") And ColumnExists("column index") And _
                                                                                 ColumnExists("visibility") And _
                                                                                 (Not ColumnExists("main section sort")) And (Not ColumnExists("main section index"))
End Property

'Initialize the Headers  -----------------------------------------------------
Private Sub ClassArraysInitialize()

    Set DICTIONARYCOLUMNS = New BetterArray

    'colnames of the dictionary
    DICTIONARYCOLUMNS.Push "variable name", "main label", "editable label", "sub label", "note", "sheet name", _
                           "sheet type", "main section", "sub section", "status", _
                            "personal identifier", "variable type", "variable format", _
                            "control", "control details", "unique", _
                           "export 1", "export 2", "export 3", "export 4", _
                           "export 5", "min", "max", "alert", "message", _
                           "formatting condition", "register book", _
                           "formatting values", "lock cells", "table name", "column index", _
                           "list auto", "visibility"
End Sub

'Dictionary data Range
Public Property Get DataRange(Optional colName As String = "__all__", Optional includeHeaders As Boolean = False) As Range
    Dim customDataSheet As IDataSheet
    Set customDataSheet = Data()
    Set DataRange = customDataSheet.DataRange(colName, includeHeaders)
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
    Dim EndRow As Long
    Dim varData As BetterArray
    Dim dictWksh As Worksheet

    EndRow = DictEndRow
    counter = StartRow + 1
    varColIndex = ColumnIndex(colName)
    Set varData = New BetterArray
    Set dictWksh = Wksh()

    With dictWksh
        Do While counter < EndRow
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

    VariableExists = Not (varRng.Find(What:=varName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

'------ Prepare and clean the dictionary
Private Sub Clean(Optional ByVal removeAddedColumns As Boolean = False)
    Dim endCol As Long
    Dim counter As Long
    Dim dictWksh As Worksheet
    Dim rng As Range
    Dim colName As String
    Dim removeAdded As Boolean

    Set dictWksh = Wksh()
    Set rng = dictWksh.UsedRange
    counter = rng.Column + rng.Columns.Count - 1

    With dictWksh
        Do While counter >= 1
            colName = .Cells(StartRow, counter).Value
            removeAdded = removeAddedColumns And _
                ((colName = "table name") Or _
                 (colName = "column index") Or _
                 (colName = "list auto") Or _
                 (colName = "visibility"))
            If (Not ColumnExists(colName, checkValidity:=True)) Or _
                 removeAdded Then _
                .Columns(counter).EntireColumn.Delete
            counter = counter - 1
        Loop
    End With
End Sub

'Prepare the dictionary for designer
Private Sub Prepare(ByVal PreservedSheetsNames As BetterArray, ByVal geoObject As ILLGeo, _
                    ByVal msgTrads As ITranslation, _
                    Optional ByVal tablePrefix As String = "table")
    Dim endCol As Long
    Dim dictWksh As Worksheet
    Dim counter As Long
    Dim controlRng As Range

    If Not Prepared() Then
        'Test if the dictionary is prepared before doing the whole loop
        If Not ColumnExists("table name") Then
            AppendNumberColumn onColName:="sheet name", tablePrefix:=tablePrefix, _
                               newColumn:="table name" 'Add table name
        End If

        If Not ColumnExists("column index") Then AppendColumnIndex 'Add Column Index

        'Testing control to update the Geo
        Set controlRng = DataRange("control")
        If (Not controlRng Is Nothing) Then
            If (Not controlRng.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True) Is Nothing) And _
                                                                                               (controlRng.Find(What:="geo2", lookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
                'Update the Geo
                AppendGeoLines geoObject
            End If
        End If
        'Sort the dictionary accordingly and make unique varNames
        UniqueVarNames
        GoodSheetNames PreservedSheetsNames

        'Add visibility (after adding geoLines)
        If Not ColumnExists("visibility") Then AddVisibility msgTrads 'Add visibility for show/Hide

        Clean
        'Add color to the last line for prepare indication
        Set dictWksh = Wksh()
        dictWksh.Cells(DictEndRow + 1, StartColumn).Font.color = vbBlue
    End If
End Sub

Private Sub GoodSheetNames(ByVal PreservedSheetsNames As BetterArray)
    Dim dictWksh As Worksheet
    Dim sheetNameIndex As Long

    Dim EndRow As Long
    Dim counter As Long

    EndRow = DictEndRow
    Set dictWksh = Wksh()
    sheetNameIndex = ColumnIndex("sheet name")

    With dictWksh
        For counter = StartRow + 1 To EndRow
            If PreservedSheetsNames.Includes(.Cells(counter, sheetNameIndex).Value) Then
                .Cells(counter, sheetNameIndex).Value = .Cells(counter, sheetNameIndex).Value & "_"
                'Trace here for logging
            End If
        Next
    End With

End Sub

'Make unique variable names
Private Sub UniqueVarNames()

    Dim dictWksh As Worksheet
    Dim EndRow As Long
    Dim varNameIndex As Long
    Dim counter As Long
    Dim varName As String

    Dim varTrackData As BetterArray

    Set varTrackData = New BetterArray
    Set dictWksh = Wksh()
    EndRow = DictEndRow
    varNameIndex = ColumnIndex("variable name")

    With dictWksh
        For counter = StartRow + 1 To EndRow
            'remove no breakable spaces and clean the name of the variable. replace spaces with _

            varName = Application.WorksheetFunction.SUBSTITUTE(varName, chr(160), " ")
            varName = Application.WorksheetFunction.Clean(varName)
            varName = Application.WorksheetFunction.Trim(.Cells(counter, varNameIndex).Value)
            varName = Replace(varName, " ", "_")
            'varName Should be at least 4 letters.
            If Len(varName) < 4 Then ThrowError _
             "Aborting, variable names should be at least of length 4," & _
             " please change the name of variable " & _
              varName

            If varTrackData.Includes(varName) Then varName = varName & "_"

            varTrackData.Push varName
            .Cells(counter, varNameIndex).Value = varName
        Next
    End With
End Sub

'Add the visibility for show Hide
Private Sub AddVisibility(ByVal msgTrads As ITranslation)

    Dim dictWksh As Worksheet
    Dim EndRow As Long
    Dim endCol As Long
    Dim counter As Long
    Dim statusIndex As Long
    Dim colIndex As Long
    Dim statusValue As String


    Set dictWksh = Wksh()
    EndRow = DictEndRow()
    endCol = DictEndColumn()

    statusIndex = ColumnIndex("status")
    colIndex = ColumnIndex("column index")

    With dictWksh
        .Cells(StartRow, endCol + 1).Value = "visibility"
        For counter = StartRow + 1 To EndRow
            statusValue = .Cells(counter, statusIndex).Value

            Select Case statusValue

            Case "hidden"

                .Cells(counter, endCol + 1).Value = msgTrads.TranslatedValue("MSG_DesHidden")

            Case "mandatory"

                .Cells(counter, endCol + 1) = msgTrads.TranslatedValue("MSG_Mandatory")

            Case "optional"

                .Cells(counter, endCol + 1) = msgTrads.TranslatedValue("MSG_Shown")
            End Select
        Next
    End With

End Sub

'Append Geo Lines
Private Sub AppendGeoLines(ByVal geoObject As ILLGeo)

    Dim dictRng As Range
    Dim geoRng As Range
    Dim pastedGeoRng As Range
    Dim dictWksh As Worksheet
    Dim geoRowIndex As Long
    Dim controlIndex As Long
    Dim varNameIndex As Long
    Dim statusIndex As Long
    Dim subSectionIndex As Long
    Dim mainLabelIndex As Long
    Dim registerIndex As Long
    Dim varName As String
    Dim EndRow As Long
    Dim endCol As Long
    Dim varPositionIndex As Long
    Dim varPosition As Long
    Dim controlDetailsIndex As Long
    Dim formulaValue As String

    Dim counter As Long

    Set dictWksh = Wksh()
    Set dictRng = DataRange("control")
    controlIndex = ColumnIndex("control")
    varNameIndex = ColumnIndex("variable name")
    mainLabelIndex = ColumnIndex("main label")
    subSectionIndex = ColumnIndex("sub section")
    registerIndex = ColumnIndex("register book")
    statusIndex = ColumnIndex("status")
    varPositionIndex = ColumnIndex("column index")
    controlDetailsIndex = ColumnIndex("control details")

    EndRow = DictEndRow
    endCol = DictEndColumn
    geoRowIndex = StartRow + 1

    With dictWksh

        'Geo lines
        Do While (Not dictRng.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True) Is Nothing)

            geoRowIndex = dictRng.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True).Row

            Set geoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
            varName = .Cells(geoRowIndex, varNameIndex).Value
            varPosition = .Cells(geoRowIndex, varPositionIndex).Value

            'Replace sub-section by main label (if required)
            If IsEmpty(.Cells(geoRowIndex, subSectionIndex)) And (Not IsEmpty(.Cells(geoRowIndex, mainLabelIndex))) Then
                .Cells(geoRowIndex, subSectionIndex).Value = .Cells(geoRowIndex, mainLabelIndex).Value
                .Cells(geoRowIndex, mainLabelIndex).Value = ""
            End If

            .Cells(geoRowIndex, varNameIndex).Value = "adm1" & "_" & varName
            .Cells(geoRowIndex, mainLabelIndex).Value = geoObject.GeoNames("adm1_name")
            .Cells(geoRowIndex, controlIndex).Value = "geo1"

            'Insert the others geo lines
            geoRowIndex = geoRowIndex + 1
            For counter = 12 To 2 Step -1
                .Rows(geoRowIndex).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Set pastedGeoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
                pastedGeoRng.Value = geoRng.Value 'Paste dictionary values
                .Cells(geoRowIndex, varNameIndex).Value = "adm" & counter & "_" & varName
                .Cells(geoRowIndex, controlIndex).Value = "geo" & counter
                .Cells(geoRowIndex, mainLabelIndex).Value = geoObject.GeoNames("adm" & counter & "_" & "name")
                .Cells(geoRowIndex, varPositionIndex).Value = varPosition + counter - 1
            Next

            geoRowIndex = geoRowIndex + 2

            'Format the pcodes
            For counter = 1 To 4

                'Hide the pcode when everything is done (after the admin names)
                .Cells(geoRowIndex + counter, statusIndex).Value = "hidden"
                .Cells(geoRowIndex + counter, registerIndex).Value = "hidden"
                .Cells(geoRowIndex + counter, controlIndex).Value = "formula"
                .Cells(geoRowIndex + counter, mainLabelIndex).Value = "geo_pcode_adm" & counter
                .Cells(geoRowIndex + counter, varNameIndex).Value = "pcode" & "_" & "adm" & counter & "_" & varName

                'Formula for pcode
                formulaValue = "INDEX(T_ADM" & counter & _
                ", MATCH(concat_adm" & counter & "_" & varName & _
                ", adm" & counter & "_concat, 0), " & (counter + 2) & ")"

                .Cells(geoRowIndex + counter, controlDetailsIndex) = "IF(ISNA(" & formulaValue & "), " & chr(34) & chr(34) & ", " & formulaValue & ")"
            Next

            geoRowIndex = geoRowIndex + 4

            'Formulas for concatenate values
            For counter = 1 To 4

                'Hide the pcode when everything is done (after the admin names)
                .Cells(geoRowIndex + counter, statusIndex).Value = "hidden"
                .Cells(geoRowIndex + counter, registerIndex).Value = "hidden"
                .Cells(geoRowIndex + counter, controlIndex).Value = "formula"
                .Cells(geoRowIndex + counter, mainLabelIndex).Value = "geo_concat_adm" & counter
                .Cells(geoRowIndex + counter, varNameIndex).Value = "concat" & "_" & "adm" & counter & "_" & varName

                'Formula for pcode
                .Cells(geoRowIndex + counter, controlDetailsIndex) = "GEOCONCAT(adm1_" & varName & ", " & counter & ")"
            Next

            geoRowIndex = geoRowIndex + 5

            Set dictRng = .Range(.Cells(geoRowIndex, controlIndex), .Cells(EndRow, controlIndex))
        Loop

        'Hf Lines
        'On Hf lines I only insert the pcode value
        Set dictRng = DataRange("control")
        Do While (Not dictRng.Find(What:="hf", lookAt:=xlWhole, MatchCase:=True) Is Nothing)
            geoRowIndex = dictRng.Find(What:="hf", lookAt:=xlWhole, MatchCase:=True).Row
            Set geoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
            varName = .Cells(geoRowIndex, varNameIndex).Value
            varPosition = .Cells(geoRowIndex, varPositionIndex).Value

            'Just insert one line for the pcode of health facility
            .Cells(geoRowIndex, varNameIndex).Value = "hf_" & varName
             geoRowIndex = geoRowIndex + 1
             .Rows(geoRowIndex).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
             Set pastedGeoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
             pastedGeoRng.Value = geoRng.Value 'Paste dictionary values

             'Change the pcode and others value
            .Cells(geoRowIndex, statusIndex).Value = "hidden"
            .Cells(geoRowIndex, controlIndex).Value = "formula"
            .Cells(geoRowIndex, mainLabelIndex).Value = "hf_pcode"
            .Cells(geoRowIndex, varNameIndex).Value = "hf_pcode" & "_" & varName
            'Column index
            .Cells(geoRowIndex, varPositionIndex).Value = varPosition + 1

            formulaValue = "INDEX(T_HF" & _
                ", MATCH(hf_" & varName & _
                ", hf" & "_concat, 0)," & "5 )"

            .Cells(geoRowIndex, controlDetailsIndex) = "IF(ISNA(" & formulaValue & "), " & chr(34) & chr(34) & ", " & formulaValue & ")"

            geoRowIndex = geoRowIndex + 1
            Set dictRng = .Range(.Cells(geoRowIndex, controlIndex), .Cells(EndRow, controlIndex))
        Loop

    End With
End Sub

Private Sub AppendNumberColumn(Optional ByVal onColName As String = "sheet name", _
                               Optional ByVal tablePrefix As String = "table", _
                               Optional ByVal newColumn As String = "table name")

    Dim EndRow As Long
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

    EndRow = DictEndRow
    endCol = DictEndColumn
    tableOrder = 1
    sheetNameIndex = ColumnIndex(onColName)

    Set dictWksh = Wksh()

    With dictWksh
        .Cells(StartRow, endCol + 1).Value = "number_" & newColumn
        sheetTrackData.Push .Cells(StartRow + 1, sheetNameIndex).Value
        tableTrackData.Push tableOrder

        For counter = StartRow + 1 To EndRow
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
        dictRng.Sort key1:=DataRange("number_" & newColumn), Order1:=xlAscending

        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).Value = newColumn
        .Cells(StartRow + 1, endCol + 1).formula = "= " & chr(34) & tablePrefix & _
                                                     chr(34) & " " & chr(38) & _
                                                      " " & .Cells(StartRow + 1, _
                                                      endCol).Address(Rowabsolute:=False)
        Set dictRng = DataRange(newColumn)

        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues
        .calculate                               'Calculate the worksheet
        dictRng.Value = dictRng.Value            'remove the link to the formulas

    End With
End Sub

Private Sub AppendColumnIndex()

    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim endCol As Long
    Dim EndRow As Long
    Dim counter As Long
    Dim varOrder As Long
    Dim sheetTypeIndex As Long
    Dim sheetNameIndex As Long
    Dim controlIndex As Long
    Dim mainSectionIndex As Long

    Dim firstsheetNameCell As String
    Dim firstmainSectionCell As String
    Dim sheetName As String
    Dim prevControl As String

    Set dictWksh = Wksh()
    endCol = DictEndColumn
    EndRow = DictEndRow
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
        .Cells(StartRow + 1, endCol + 1).formula = "= IF(" & firstmainSectionCell & "=" & chr(34) & chr(34) & _
                                                   ", ROW(" & firstmainSectionCell & ")" & ", " & _
                                                   firstsheetNameCell & chr(38) & chr(34) & _
                                                   " - " & chr(34) & chr(38) & firstmainSectionCell & _
                                                   ")"
        Set dictRng = DataRange("main section sort")
        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues

        'Now append and sort
        AppendNumberColumn onColName:="main section sort", _
        tablePrefix:="main section", newColumn:="main section index"

        'Filling column index
        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).Value = "column index"
        For counter = StartRow + 1 To EndRow
            If .Cells(counter, sheetNameIndex).Value <> sheetName Then
                'New sheet name
                varOrder = IIf(.Cells(counter, sheetTypeIndex).Value = SHEETTYPEVERTICAL, 4, 1)
                sheetName = .Cells(counter, sheetNameIndex).Value
            Else
                prevControl = .Cells(counter - 1, controlIndex).Value
                'Old sheet name
                varOrder = Switch(prevControl = "geo", varOrder + 12, _
                                  prevControl = "hf", varOrder + 2, _
                                  True, varOrder + 1)
            End If
            .Cells(counter, endCol + 1).Value = varOrder 'variable order
        Next
    End With
End Sub

'Import and export subs
'Export
Private Sub Export(ByVal Wkb As Workbook, _
                   Optional ByVal exportType As String = "__all__", _
                   Optional ByVal addListObject As Boolean = True, _
                   Optional ByVal Hide As Long = xlSheetHidden)

    Dim dictData As IDataSheet
    Dim rng As Range
    Dim sh As Worksheet
    Dim sheetName As String

    Set dictData = Data()
    If ColumnExists(exportType) Then
        dictData.Export Wkb, filteredVarName:=exportType, Hide:=Hide
    ElseIf exportType = "__all__" Then
        dictData.Export Wkb, Hide:=Hide
    End If

    'Add a list object to the dictionnary.
    If addListObject Then
        sheetName = Wksh.Name
        Set sh = Wkb.Worksheets(sheetName)
        Set rng = DataRange(includeHeaders:=True)
        Set rng = sh.Range(rng.Address)

        'Add a listObject to the dictionary (for compatibility purposes with previous codes)
        With sh
            .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "o" & LCase(sheetName)
            'prepare the dictionary
            .Cells(DictEndRow + 1, StartColumn).Font.color = vbBlue
        End With
    End If
End Sub

'Import
Private Sub Import(ByVal fromWksh As Worksheet, _
                   ByVal fromStartRow As Long, _
                   ByVal fromStartcol As Long)

    Dim dictData As IDataSheet
    Dim impdictData As IDataSheet
    Dim impRng As Range
    Dim sh As Worksheet 'Worksheet of the dictionary

    Set sh = Wksh()
    Set dictData = Data()
    Set impdictData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    'Clear the current worksheet before importing
    sh.Cells.Clear
    If DICTIONARYCOLUMNS Is Nothing Then ClassArraysInitialize
    DICTIONARYCOLUMNS.ToExcelRange sh.Cells(StartRow, StartColumn), TransposeValues:=True

    'Clean the current data sheet (remove unwanted columns)
    Clean removeAddedColumns:=True
    dictData.Import impdictData

    'Import the formats of the Range Formatting Values
    Set impRng = impdictData.DataRange("Formatting Values", _
                                      includeHeaders:=True)

    'Import the format for formatting values
    dictData.ImportFormat impRng
End Sub

'Translation
Private Sub Translate(ByVal TransObject As ITranslation)
    Dim transCols As BetterArray
    Dim rng As Range
    Dim counter As Long

    Set transCols = New BetterArray

    transCols.Push "main label", "sub label", "sheet name", "note", _
                   "main section", "sub section", "message"

    For counter = transCols.LowerBound To transCols.UpperBound
        Set rng = DataRange(transCols.Item(counter))
        TransObject.TranslateRange rng
    Next

    Set rng = DataRange("control details")
    TransObject.TranslateRange rng, containsFormula:=True
End Sub

'Add a column to the dictionary

Private Sub AddColumn(ByVal colName As String)
    Dim dictWksh As Worksheet
    Dim endCol As Long
    Set dictWksh = Wksh()
    endCol = DictEndColumn()
    If colName <> vbNullString Then
        dictWksh.Cells(StartRow, endCol + 1).Value = colName
    End If
End Sub

'Throw Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
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
Public Sub ILLdictionary_Prepare(ByVal PreservedSheetsNames As BetterArray, _
                                 ByVal geoObject As ILLGeo, ByVal msgTrads As ITranslation, _
                                 Optional ByVal tablePrefix As String = "table")
    Prepare PreservedSheetsNames, geoObject, msgTrads, tablePrefix
End Sub

Public Sub ILLdictionary_Clean()
    Clean
End Sub

Public Sub ILLdictionary_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub

Public Sub ILLdictionary_Export(ByVal toWkb As Workbook, _
                                Optional ByVal exportType As String = "__all__", _
                                Optional ByVal addListObject As Boolean = True, _
                                Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, exportType, addListObject, Hide
End Sub

Public Sub ILLdictionary_Translate(ByVal TransObject As ITranslation)
    Translate TransObject
End Sub

Public Sub ILLdictionary_AddColumn(ByVal colName As String)
    AddColumn colName
End Sub
