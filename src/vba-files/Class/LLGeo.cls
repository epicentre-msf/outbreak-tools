VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Geobase management class")
'@IgnoreModule

Option Explicit
Implements ILLGeo

Private Type TGeo
    geosh As Worksheet
End Type

'Exposed methods
Private this As TGeo
Private Const CLASSNAME As String = "Geo"
Private Const ADMIN1    As String = "ADM1"
Private Const ADMIN2    As String = "ADM2"
Private Const ADMIN3    As String = "ADM3"
Private Const ADMIN4    As String = "ADM4"
Private Const HF        As String = "HF"
Private Const HISTOGEO  As String = "HISTOGEO"
Private Const HISTOHF   As String = "HISTOHF"
Private Const NAMES     As String = "NAMES"
Private Const METADATA  As String = "METADATA"

Public Function Create(ByVal sh As Worksheet) As ILLGeo
    CheckRequirements sh
    With New LLGeo
        Set .Wksh = sh
        Set Create = .Self()
    End With
End Function

Public Property Get Self() As ILLGeo
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.geosh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.geosh = sh
End Property

Public Property Get TableNames() As BetterArray
    Dim tabl As BetterArray
    Set tabl = New BetterArray
    tabl.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
              "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF
    Set TableNames = tabl.Clone()
End Property

Private Property Get GeoNames(ByVal admname As String) As String
    Dim sh As Worksheet
    Dim actualName As Variant
    Set sh = Wksh()
    Select Case admname
    Case "adm1_name"
        actualName = sh.Range("RNG_ADM1NAME").Value
    Case "adm2_name"
        actualName = sh.Range("RNG_ADM2NAME").Value
    Case "adm3_name"
        actualName = sh.Range("RNG_ADM3NAME").Value
    Case "adm4_name"
        actualName = sh.Range("RNG_ADM4NAME").Value
    Case "hf_name"
        actualName = sh.Range("RNG_HFNAME").Value
    Case Else
        actualName = admname
    End Select

    If TypeName(actualName) = "String" Then
        GeoNames = actualName
    Else
        GeoNames = admname
    End If
End Property

Private Function GeoColumnIndex(ByVal colName As String, ByVal adminLevel As GeoLevel) As Long
    Dim hRng As Range
    Dim loName As String
    Dim tabl As BetterArray
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim actualColName As String

    Set sh = Wksh()
    Set tabl = TableNames()
    loName = tabl.Item(adminLevel)
    actualColName = colName
    If sh.Range("RNG_GeoUpdated").Value = "updated, translated" Then actualColName = GeoNames(colName)
    Set hRng = sh.ListObjects(loName).HeaderRowRange
    colIndex = ColumnIndex(actualColName, hRng)
    If colIndex < 0 Then ThrowError "Column " & actualColName & " not found in table " & loName
    GeoColumnIndex = colIndex
End Function


Private Function HFColumnIndex(ByVal colName As String) As Long
    Dim hRng As Range
    Dim sh As Worksheet
    Dim actualColName As String
    Dim colIndex As Long
    Set sh = Wksh()
    Set hRng = sh.ListObjects("T_" & HF).HeaderRowRange
    actualColName = colName
    If sh.Range("RNG_GeoUpdated").Value = "updated, translated" Then actualColName = GeoNames(colName)
    colIndex = ColumnIndex(actualColName, hRng)
    If colIndex < 0 Then ThrowError "Column " & actualColName & " not found in table T_HF"
    HFColumnIndex = colIndex
End Function

Private Function ColumnIndex(ByVal colName As String, ByVal headerRng As Range, _
                             Optional ByVal inRange As Boolean = True) As Long
    Dim index As Long
    If Not (headerRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
        index = headerRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True).Column
        If inRange Then index = index - headerRng.Column + 1
    Else
        index = -1
    End If
    ColumnIndex = index
End Function

Private Function UniqueValues() As BetterArray
    Dim EndRow As Long
    Dim endCol As Long
    Dim sh As Worksheet
    Dim pasteRng As Range
    Dim returnedValues As BetterArray

    Set returnedValues = New BetterArray
    Set sh = Wksh()

    With sh
        endCol = .Range("RNG_PastingGeoCol").Column
        EndRow = .Cells(.Rows.Count, endCol).End(xlUp).Row
        Set pasteRng = .Range(.Cells(1, endCol), .Cells(EndRow, endCol))
        pasteRng.RemoveDuplicates Columns:=1, Header:=xlNo
        EndRow = .Cells(.Rows.Count, endCol).End(xlUp).Row 'New endrow after removing duplicates
        Set pasteRng = .Range(.Cells(1, endCol), .Cells(EndRow, endCol))
        returnedValues.FromExcelRange pasteRng, DetectLastRow:=False, _
                                      DetectLastColumn:=False
    End With

    Set UniqueValues = returnedValues.Clone()
End Function

Private Function HFRange(ByVal colName As String) As Range
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    colIndex = HFColumnIndex(colName)
    If colIndex < 0 Then ThrowError colName & " is not in HF Table"
    Set HFRange = sh.ListObjects("T_" & HF).ListColumns(colIndex).DataBodyRange()
End Function

Private Property Get GeoLevel(ByVal adminLevel As GeoLevel, _
                              Optional ByVal geoType As CustomGeoType = CustomTypeGeo, _
                              Optional ByVal adminNames As Variant = vbNullString _
                              ) As BetterArray

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outLevels As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim rng As Range
    Dim loName As String
    Dim colIndex As Long
    Dim counter As Long
    Dim returnedRng As Range

    Set sh = Wksh()
    Set tabl = TableNames()

    If geoType = CustomTypeGeo Then
        loName = tabl.Item(adminLevel)
    Else
        loName = tabl.Item(4)
    End If

    'Work only on Geobase with values withing
    If (sh.Range("RNG_GeoUpdated").Value = "empty") Then Exit Property

    'Test the type of the adminNames
    If (adminLevel > 1) And (Not TypeName(adminNames) = "BetterArray") Then
        ThrowError "Admin names should be a list of all previous admin levels for admin " & adminLevel
    ElseIf (adminLevel = LevelAdmin1 Or adminLevel = LevelAdmin2) And (Not TypeName(adminNames) = "String") Then
        ThrowError "No need to add admin names If you want the list of all admin"
    ElseIf (TypeName(adminNames) = "BetterArray") Then
        If (adminNames.Length <> CLng(adminLevel)) Then ThrowError "You should provide a table of length " & CLng(adminLevel) & " for admin " & (CLng(adminLevel) + 1)
    End If

    'Output Levels
    Set outLevels = New BetterArray
    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range

    'If the Listobject is Empty, throw error
    If Lo.DataBodyRange Is Nothing Then ThrowError "The listobject " & loName & "has no data"

    Select Case geoType

    Case CustomTypeGeo

        Select Case adminLevel

        Case LevelAdmin1                         'Admin 1 of Geo

            colIndex = GeoColumnIndex("adm1_name", LevelAdmin1)
            Set rng = Lo.ListColumns(colIndex).DataBodyRange
            outLevels.FromExcelRange rng, DetectLastRow:=False, _
                                     DetectLastColumn:=False
        Case LevelAdmin2
            colIndex = GeoColumnIndex("adm1_name", LevelAdmin2)
            With LoRng
                .AutoFilter Field:=colIndex, Criteria1:=adminNames
            End With
            'returned Range
            colIndex = GeoColumnIndex("adm2_name", LevelAdmin2)

        Case LevelAdmin3, LevelAdmin4
            With LoRng
                For counter = 1 To adminNames.Length
                    colIndex = GeoColumnIndex("adm" & counter & "_name", adminLevel)
                    .AutoFilter Field:=colIndex, Criteria1:=adminNames.Item(counter)
                Next
            End With
            colIndex = GeoColumnIndex("adm" & (CLng(adminLevel) + 1) & "_name", adminLevel)
        End Select

    Case CustomTypeHF

        Select Case adminLevel

        Case LevelAdmin1

            Set rng = HFRange("adm1_name")
            Set outLevels = UniqueValues(rng)

        Case LevelAdmin2

            colIndex = HFColumnIndex("adm1_name")
            With LoRng
                .AutoFilter Field:=colIndex, Criteria1:=adminNames
            End With
            colIndex = HFColumnIndex("adm2_name")

        Case LevelAdmin3, LevelAdmin4
            With LoRng
                For counter = 1 To adminNames.Length
                    colIndex = HFColumnIndex("adm" & counter & "_name")
                    .AutoFilter Field:=colIndex, Criteria1:=adminNames.Item(counter)
                Next
            End With
            If adminLevel = LevelAdmin3 Then
                colIndex = HFColumnIndex("adm3_name")
            Else                                 'adm4, equivalent to hf_name
                colIndex = HFColumnIndex("hf_name")
            End If
        End Select
    End Select

    If adminLevel <> LevelAdmin1 Then            'Take filtered values If not on admin 1

        On Error Resume Next
        Set returnedRng = Lo.ListColumns(colIndex).DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not returnedRng Is Nothing Then
            returnedRng.Copy sh.Range("RNG_PastingGeoCol")
            Lo.AutoFilter.ShowAllData
            Set outLevels = UniqueValues()
        End If
    End If

    Lo.AutoFilter.ShowAllData
    'Clear the pasting column
    sh.Range("RNG_PastingGeoCol").EntireColumn.Cells.Clear
    Set GeoLevel = outLevels.Clone()
End Property

Private Property Get Population(ByVal adminLevel As GeoLevel, _
                                ByVal adminNames As Variant) As String

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim counter As Long
    Dim returnedRng As Range
    Dim loName As String
    Dim colIndex As Long

    Set sh = Wksh()
    Set tabl = TableNames()
    loName = tabl.Item(adminLevel)
    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range

    If (TypeName(adminNames) = "BetterArray") Then
        If (adminNames.Length <> CLng(adminLevel) + 1) Then _
                                                       ThrowError "You should provide a table of length " & CLng(adminLevel) + 1 & _
                                                       " for admin " & (CLng(adminLevel) + 1)
    ElseIf (adminLevel = 0) And (Not TypeName(adminNames) = "String") Then
        ThrowError "You should provide only one admin1 name for population of admin1"
    End If

    Select Case adminLevel
    Case LevelAdmin1
        With LoRng
            colIndex = GeoColumnIndex("adm1_name", LevelAdmin1)
            .AutoFilter Field:=colIndex, Criteria1:=adminNames
            colIndex = GeoColumnIndex("adm1_pop", LevelAdmin1)
        End With
    Case LevelAdmin2, LevelAdmin3, LevelAdmin4
        With LoRng
            adminNames.LowerBound = 1
            For counter = 1 To adminNames.Length
                colIndex = GeoColumnIndex("adm" & counter & "_name", adminLevel)
                .AutoFilter Field:=colIndex, Criteria1:=adminNames.Item(counter)
            Next
            colIndex = GeoColumnIndex("adm" & (counter - 1) & "_pop", adminLevel)
        End With
    End Select

    'Return filtered value
    On Error Resume Next
    Set returnedRng = Lo.ListColumns(colIndex).DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    Lo.AutoFilter.ShowAllData

    If Not returnedRng Is Nothing Then Population = returnedRng.Value
End Property

Private Sub Translate(Optional ByVal rawNames As Boolean = False)
    'Headers for the geo sheet
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim colIndex As Long
    Dim counter As Long
    Dim counter2 As Long
    Dim rng As Range
    Dim Lo As ListObject
    Dim loName As String
    Dim colName As String
    Dim admname As String
    Dim transGeoName As String                   'Translated geo Name
    Dim hRng As Range                            'Header Range

    Set sh = Wksh()
    Set rng = sh.Range("RNG_GeoUpdated")
    If Not rawNames And ((rng.Value = "updated, translated") Or (rng.Value = "empty")) Then Exit Sub
    If rawNames And (rng.Value = "updated, not translated") Then Exit Sub

    Set tabl = TableNames()
    tabl.LowerBound = 1
    For counter = 1 To 5
        loName = tabl.Item(counter)
        Set Lo = sh.ListObjects(loName)
        Set hRng = Lo.HeaderRowRange
        'Admin 1 to 4
        For counter2 = 1 To counter
            admname = "adm" & counter2 & "_name"
            transGeoName = GeoNames(admname)
            colName = IIf(rawNames, transGeoName, admname)
            colIndex = ColumnIndex(colName, hRng, inRange:=False)
            If colIndex > 0 Then sh.Cells(hRng.Row, colIndex).Value = IIf(rawNames, admname, transGeoName)
        Next

        If counter = 5 Then
            'HF
            admname = "hf_name"
            transGeoName = GeoNames(admname)
            colName = IIf(rawNames, transGeoName, admname)
            colIndex = ColumnIndex(colName, hRng, inRange:=False)
            If colIndex > 0 Then sh.Cells(hRng.Row, colIndex).Value = IIf(rawNames, admname, transGeoName)
        End If

    Next
    rng.Value = IIf(rawNames, "updated, not translated", "updated, translated")
End Sub

Private Sub Clear()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = Wksh()
    For Each Lo In sh.ListObjects
        On Error Resume Next
        If Lo.Name <> "T_LevelNames" Then Lo.DataBodyRange.Cells.Clear
        On Error GoTo 0
    Next

    sh.Range("RNG_GeoUpdated").Value = "empty"
End Sub

Private Sub ClearHistoric()
    Dim sh As Worksheet
    Set sh = Wksh()
    On Error Resume Next
    sh.ListObjects("T_" & HISTOGEO).DataBodyRange.Cells.Clear
    sh.ListObjects("T_" & HISTOHF).DataBodyRange.Cells.Clear
    On Error GoTo 0
End Sub

Private Sub ExportToFile(outputWkb As Workbook, Optional ByVal onlyHistory As Boolean = False)

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outputRng As Range
    Dim inputRng As Range
    Dim sheetName As String
    Dim loName As String
    Dim counter As Long

    Set tabl = New BetterArray
    Set sh = Wksh()
    Translate rawNames:=True                     'Translate with raw names

    If onlyHistory Then
        tabl.Push "T_" & HISTOGEO, "T_" & HISTOHF
    Else
        Set tabl = TableNames()
    End If

    With outputWkb
        'geo metadata
        .Worksheets(1).Name = METADATA
        Set inputRng = sh.ListObjects("T_" & METADATA).Range
        With .Worksheets(METADATA)
            Set outputRng = .Range(.Cells(1, 1), .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
        End With
        outputRng.Value = inputRng.Value
        sheetName = METADATA

        'Add the others tables to the worksheet
        For counter = tabl.LowerBound To tabl.UpperBound
            loName = tabl.Item(counter)
            Set inputRng = sh.ListObjects(loName).Range
            .Worksheets.Add(after:=.Worksheets(sheetName)).Name = Replace(loName, "T_", vbNullString)
            sheetName = Replace(loName, "T_", vbNullString)
            With .Worksheets(sheetName)
                Set outputRng = .Range(.Cells(1, 1), .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
            End With
            outputRng.Value = inputRng.Value
        Next
    End With

    Translate rawNames:=False
End Sub

'Export by keeping the actual format of the Worksheet
Private Sub ExportToWkb(ByVal outputWkb As Workbook)

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As ListObject
    Dim expAddress As String
    Dim expRange As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String

    'Test if a sheet named as the geo sheet already exists
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(sh.Name)
    On Error GoTo 0


    If Not outsh Is Nothing Then ThrowError "Unable to export, already a sheet named " & sh.Name & " in the workbook"
    Set sh = Wksh()
    Translate rawNames:=False

    'Now move every listObjects
    outputWkb.Worksheets.Add.Name = sh.Name
    Set outsh = outputWkb.Worksheets(sh.Name)

    'Export the listObjects
    For Each Lo In sh.ListObjects
        expAddress = Lo.Range.Address
        Set expRange = outsh.Range(expAddress)
        expRange.Value = Lo.Range.Value
        'Add the listObject
        With outsh
            .ListObjects.Add(xlSrcRange, expRange, , xlYes).Name = Lo.Name
            .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
        End With
    Next

    'Export all the ranges
    Set rngTable = New BetterArray
    rngTable.Push "RNG_GeoName", "RNG_GeoUpdated", "RNG_PastingGeoCol", "RNG_GeoLangCode", _
                  "RNG_HFNAME", "RNG_ADM1NAME", "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME"

    For counter = rngTable.LowerBound To rngTable.UpperBound
        rngName = rngTable.Item(counter)
        expAddress = sh.Range(rngName).Address
        Set expRange = outsh.Range(expAddress)
        expRange.Name = rngName
        expRange.Value = sh.Range(rngName).Value
        If sh.Range(rngName).formula <> vbNullString Then
            On Error Resume Next
            expRange.formula = sh.Range(rngName).formula
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub Import(ByVal fromWkb As Workbook)
    Dim sh As Worksheet
    Dim impsh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim admNames As BetterArray
    Dim counter As Long
    Dim colName As String
    Dim colIndex As Long
    Dim lastRow As Long
    Dim impRng As Range
    Dim actRng As Range


    Set admNames = New BetterArray
    admNames.Push ADMIN1, ADMIN2, ADMIN3, ADMIN4, HF, HISTOGEO, HISTOHF, NAMES, METADATA

    Set sh = Wksh()
    'Clear the Current Geo
    Clear
    Translate rawNames:=True

    For Each impsh In fromWkb.Worksheets
        'Only sheet with valid names are used for import
        If admNames.Includes(impsh.Name) Then
            Set Lo = sh.ListObjects("T_" & impsh.Name) 'Lo Name
            Set hRng = Lo.HeaderRowRange
            counter = 1
            Do While impsh.Cells(1, counter) <> vbNullString
                With impsh
                    lastRow = .UsedRange.Rows.Count
                    colName = .Cells(1, counter).Value
                    colIndex = ColumnIndex(colName, hRng, inRange:=False)

                    If colIndex > 0 Then
                        Set impRng = .Range(.Cells(2, counter), .Cells(lastRow, counter))
                        Set actRng = sh.Range(sh.Cells(2, colIndex), sh.Cells(lastRow, colIndex))
                        actRng.Value = impRng.Value
                    End If
                End With
                counter = counter + 1
            Loop
        End If
    Next
    Translate rawNames:=False
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub CheckRequirements(ByVal sh As Worksheet)

    Dim admTableNames As BetterArray

    'Sheet Should Exists
    If sh Is Nothing Then ThrowError "Geo sheet not defined"
    'Check for existence of all the required tables
    Set admTableNames = New BetterArray
    admTableNames.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
                       "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF, "T_" & METADATA
    LoExists admTableNames, sh
    'Check if ranges for geo all exists
    RangeExists sh, "RNG_GeoName"
    RangeExists sh, "RNG_GeoUpdated"
    RangeExists sh, "RNG_PastingGeoCol"
    RangeExists sh, "RNG_GeoLangCode"
    RangeExists sh, "RNG_HFNAME"
    RangeExists sh, "RNG_ADM1NAME"
    RangeExists sh, "RNG_ADM2NAME"
    RangeExists sh, "RNG_ADM3NAME"
    RangeExists sh, "RNG_ADM4NAME"
End Sub

Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)

    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name

End Sub

Private Sub LoExists(ByVal loNames As BetterArray, sh As Worksheet)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String
    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError "Table " & lName & " does not exists in Worksheet " & sh.Name
    Next
End Sub

'Interfaces
Public Property Get ILLGeo_GeoLevel(ByVal adminLevel As GeoLevel, _
                                  Optional ByVal geoType As CustomGeoType = CustomTypeGeo, _
                                  Optional ByVal adminNames As Variant = vbNullString) As BetterArray

    Set ILLGeo_GeoLevel = GeoLevel(adminLevel, geoType, adminNames)
End Property

'Convert the code name to the actual name of the current language
Public Property Get ILLGeo_GeoNames(ByVal admname As String) As String
    ILLGeo_GeoNames = GeoNames(admname)
End Property

'Get the Population of one administrative level (only for geoType)
Public Property Get ILLGeo_Population(ByVal adminLevel As GeoLevel, _
                                    ByVal adminNames As Variant) As String
    ILLGeo_Population = Population(adminLevel, adminNames)
End Property

'Translation
Public Sub ILLGeo_Translate(Optional ByVal rawNames As Boolean = False)
    Translate rawNames
End Sub

'Import a new Geobase from a workbook
Public Sub ILLGeo_Import(ByVal fromWkb As Workbook)
    Import fromWkb
End Sub

'Export to a workbook
Public Sub ILLGeo_ExportToFile(ByVal outputWkb As Workbook, Optional ByVal onlyHistory As Boolean = False)
    ExportToFile outputWkb, onlyHistory
End Sub

Public Sub ILLGeo_ExportToWkb(ByVal outputWkb As Workbook)
    ExportToWkb outputWkb
End Sub

Public Sub ILLGeo_Clear()
    Clear
End Sub

Public Sub ILLGeo_ClearHistoric()
    ClearHistoric
End Sub
