VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Geobase management class"

'@PredeclaredId
'@Folder("Geo")
'@ModuleDescription("Geobase management class")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing

Option Explicit
Implements ILLGeo

'@section State
'===============================================================================

Private Type TGeo
    geosh As Worksheet
    namesCacheKeys As BetterArray
    namesCacheValues As BetterArray
    namesCacheLoaded As Boolean
End Type

Private this As TGeo
Private Const CLASS_NAME As String = "LLGeo"

'Table name constants
Private Const ADMIN1 As String = "ADM1"
Private Const ADMIN2 As String = "ADM2"
Private Const ADMIN3 As String = "ADM3"
Private Const ADMIN4 As String = "ADM4"
Private Const HF As String = "HF"
Private Const HISTOGEO As String = "HISTOGEO"
Private Const HISTOHF As String = "HISTOHF"
Private Const NAMES As String = "NAMES"
Private Const METADATA As String = "METADATA"

'@section Factory
'===============================================================================

' @description Create a new LLGeo instance from a geobase worksheet.
'              Validates that all required tables and named ranges exist.
' @param sh The geobase worksheet (must contain T_ADM1..T_ADM4, T_HF, etc.)
' @return ILLGeo interface pointer
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

'@section Table Names
'===============================================================================

' @description Get the list of all geobase table names in order:
'              T_ADM1, T_ADM2, T_ADM3, T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF
Public Property Get TableNames() As BetterArray
    Dim tabl As BetterArray
    Set tabl = New BetterArray
    tabl.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
              "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF
    Set TableNames = tabl.Clone()
End Property

'@section Translation Cache
'===============================================================================

' @description Ensure the translation name cache is loaded. The cache maps raw
'              admin names (adm1_name, adm2_name, etc.) to their translated values
'              from the geobase named ranges. Loaded lazily on first access.
Private Sub EnsureNamesCache()
    Dim sh As Worksheet

    If this.namesCacheLoaded Then Exit Sub

    Set sh = Wksh()
    sh.Calculate

    Set this.namesCacheKeys = New BetterArray
    Set this.namesCacheValues = New BetterArray

    CacheOneName sh, "adm1_name", "RNG_ADM1NAME"
    CacheOneName sh, "adm2_name", "RNG_ADM2NAME"
    CacheOneName sh, "adm3_name", "RNG_ADM3NAME"
    CacheOneName sh, "adm4_name", "RNG_ADM4NAME"
    CacheOneName sh, "hf_name", "RNG_HFNAME"
    CacheOneName sh, "geobase", "RNG_GeoName"

    this.namesCacheLoaded = True
End Sub

' @description Cache one translation: raw name -> resolved value from named range
Private Sub CacheOneName(ByVal sh As Worksheet, ByVal rawName As String, _
                         ByVal rngName As String)
    Dim actualName As Variant
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push rawName
        Exit Sub
    End If

    actualName = rng.Value
    If TypeName(actualName) = "String" Then
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push CStr(actualName)
    Else
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push rawName
    End If
End Sub

' @description Reset the translation cache. Call after Import or Translate.
Private Sub ResetNamesCache()
    this.namesCacheLoaded = False
    Set this.namesCacheKeys = Nothing
    Set this.namesCacheValues = Nothing
End Sub

' @description Convert a raw admin column name to its translated display name.
'              Uses the cached translation table. Falls back to the raw name
'              if no translation mapping exists.
Private Property Get GeoNames(ByVal admname As String) As String
    Dim counter As Long

    EnsureNamesCache

    For counter = this.namesCacheKeys.LowerBound To this.namesCacheKeys.UpperBound
        If this.namesCacheKeys.Item(counter) = admname Then
            GeoNames = CStr(this.namesCacheValues.Item(counter))
            Exit Property
        End If
    Next

    GeoNames = admname
End Property

'@section Column Helpers
'===============================================================================

' @description Find the 1-based column index of a column name within a header range.
' @param colName The column name to search for
' @param headerRng The header row range to search in
' @param inRange True to return the position relative to the range start,
'                False to return the absolute column number
' @return Column index (1-based), or -1 if not found
Private Function ColumnIndex(ByVal colName As String, ByVal headerRng As Range, _
                             Optional ByVal inRange As Boolean = True) As Long
    Dim matchCell As Range

    Set matchCell = headerRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=True)

    If matchCell Is Nothing Then
        ColumnIndex = -1
    ElseIf inRange Then
        ColumnIndex = matchCell.Column - headerRng.Column + 1
    Else
        ColumnIndex = matchCell.Column
    End If
End Function

' @description Resolve the column index for a geo admin table, applying
'              translation if the geobase has been translated.
Private Function GeoColumnIndex(ByVal colName As String, ByVal adminLevel As Byte) As Long
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim loName As String
    Dim actualColName As String
    Dim colIdx As Long

    Set sh = Wksh()
    Set tabl = TableNames()
    loName = tabl.Item(adminLevel)
    actualColName = colName

    If sh.Range("RNG_GeoUpdated").Value = "updated, translated" Then
        actualColName = GeoNames(colName)
    End If

    colIdx = ColumnIndex(actualColName, sh.ListObjects(loName).HeaderRowRange)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & actualColName & " not found in table " & loName

    GeoColumnIndex = colIdx
End Function

' @description Resolve the column index for the HF table, applying translation.
Private Function HFColumnIndex(ByVal colName As String) As Long
    Dim sh As Worksheet
    Dim actualColName As String
    Dim colIdx As Long

    Set sh = Wksh()
    actualColName = colName

    If sh.Range("RNG_GeoUpdated").Value = "updated, translated" Then
        actualColName = GeoNames(colName)
    End If

    colIdx = ColumnIndex(actualColName, sh.ListObjects("T_" & HF).HeaderRowRange)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & actualColName & " not found in table T_HF"

    HFColumnIndex = colIdx
End Function

' @description Extract unique values from a range into a BetterArray.
Private Function UniqueRangeValue(ByVal rng As Range) As BetterArray
    Dim cellRng As Range
    Dim retVal As BetterArray

    Set retVal = New BetterArray

    For Each cellRng In rng
        If Not retVal.Includes(cellRng.Value) Then retVal.Push cellRng.Value
    Next

    Set UniqueRangeValue = retVal.Clone()
End Function

' @description Get the data body range for a named column in the HF table.
Private Function HFRange(ByVal colName As String) As Range
    Dim colIdx As Long
    colIdx = HFColumnIndex(colName)
    Set HFRange = Wksh().ListObjects("T_" & HF).ListColumns(colIdx).DataBodyRange()
End Function

'@section GeoLevel
'===============================================================================

' @description Get the list of values at a given admin level, optionally filtered
'              by parent admin selections. Uses AutoFilter on the appropriate
'              ListObject for hierarchical cascade (adm1 -> adm2 -> adm3 -> adm4).
Private Property Get GeoLevel(ByVal adminLevel As Byte, _
                              Optional ByVal geoType As Byte = GeoScopeAdmin, _
                              Optional adminNames As Variant = vbNullString) As BetterArray

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outLevels As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim rng As Range
    Dim loName As String
    Dim colIdx As Long
    Dim counter As Long
    Dim returnedRng As Range
    Dim prevCalc As Boolean

    Set sh = Wksh()
    Set tabl = TableNames()
    Set outLevels = New BetterArray

    'Disable calculation for performance during filtering
    prevCalc = sh.EnableCalculation
    sh.EnableCalculation = False

    On Error GoTo Cleanup

    If geoType = GeoScopeAdmin Then
        loName = tabl.Item(adminLevel)
    Else
        loName = tabl.Item(4) 'T_HF
    End If

    'Empty geobase: return empty array
    If sh.Range("RNG_GeoUpdated").Value = "empty" Then GoTo Cleanup

    'Validate adminNames
    If (adminLevel > 1) And (Not TypeName(adminNames) = "BetterArray") Then
        ThrowError InvalidArgument, _
            "Admin names should be a list of all previous admin levels for admin " & adminLevel
    ElseIf (adminLevel = LevelAdmin1 Or adminLevel = LevelAdmin2) _
            And (Not TypeName(adminNames) = "String") Then
        ThrowError InvalidArgument, _
            "No need to add admin names if you want the list of all admin"
    ElseIf TypeName(adminNames) = "BetterArray" Then
        If adminNames.Length <> CLng(adminLevel) Then
            ThrowError InvalidArgument, _
                "You should provide a table of length " & CLng(adminLevel) & _
                " for admin " & (CLng(adminLevel) + 1)
        End If
        If adminNames.LowerBound <> 1 Then
            ThrowError InvalidArgument, _
                "Please ensure your BetterArray has 1 as LowerBound, it is mandatory"
        End If
    End If

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range

    If Lo.DataBodyRange Is Nothing Then
        ThrowError ErrorUnexpectedState, "The listobject " & loName & " has no data"
    End If

    Select Case geoType

    Case GeoScopeAdmin

        Select Case adminLevel

        Case LevelAdmin1
            colIdx = GeoColumnIndex("adm1_name", LevelAdmin1)
            Set rng = Lo.ListColumns(colIdx).DataBodyRange
            outLevels.FromExcelRange rng, DetectLastRow:=False, DetectLastColumn:=False

        Case LevelAdmin2
            colIdx = GeoColumnIndex("adm1_name", LevelAdmin2)
            LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames
            colIdx = GeoColumnIndex("adm2_name", LevelAdmin2)

        Case LevelAdmin3, LevelAdmin4
            For counter = 1 To adminNames.Length
                colIdx = GeoColumnIndex("adm" & counter & "_name", adminLevel)
                LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames.Item(counter)
            Next
            colIdx = GeoColumnIndex("adm" & (CLng(adminLevel) + 1) & "_name", adminLevel)
        End Select

    Case GeoScopeHF

        Select Case adminLevel

        Case LevelAdmin1
            Set rng = HFRange("adm1_name")
            Set outLevels = UniqueRangeValue(rng)

        Case LevelAdmin2
            colIdx = HFColumnIndex("adm1_name")
            LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames
            colIdx = HFColumnIndex("adm2_name")

        Case LevelAdmin3, LevelAdmin4
            For counter = 1 To adminNames.Length
                colIdx = HFColumnIndex("adm" & counter & "_name")
                LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames.Item(counter)
            Next
            If adminLevel = LevelAdmin3 Then
                colIdx = HFColumnIndex("adm3_name")
            Else
                colIdx = HFColumnIndex("hf_name")
            End If
        End Select
    End Select

    'For levels beyond admin1, extract filtered visible cells
    If adminLevel <> LevelAdmin1 Then
        On Error Resume Next
        Set returnedRng = Lo.ListColumns(colIdx).DataBodyRange _
                            .SpecialCells(xlCellTypeVisible)
        On Error GoTo Cleanup

        If Not returnedRng Is Nothing Then
            Lo.AutoFilter.ShowAllData
            Set outLevels = UniqueRangeValue(returnedRng)
        End If

        On Error Resume Next
        Lo.AutoFilter.ShowAllData
        On Error GoTo Cleanup
    End If

Cleanup:
    Set GeoLevel = outLevels.Clone()
    sh.EnableCalculation = prevCalc
End Property

'@section Translate
'===============================================================================

' @description Translate or un-translate all geobase table headers. When rawNames
'              is False, headers are translated to display names. When True, headers
'              are reverted to raw column names.
Private Sub Translate(Optional ByVal rawNames As Boolean = False)
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim counter As Long
    Dim counter2 As Long
    Dim Lo As ListObject
    Dim loName As String
    Dim admname As String
    Dim transGeoName As String
    Dim hRng As Range
    Dim rng As Range
    Dim cellRng As Range

    Set sh = Wksh()
    sh.Calculate
    Set rng = sh.Range("RNG_GeoUpdated")

    'Skip if already in the requested state
    If Not rawNames And (InStr(1, rng.Value, ", translated") > 0) Then Exit Sub
    If rawNames And (InStr(1, rng.Value, ", not translated") > 0) Then Exit Sub

    'Reset cache so GeoNames reads fresh values
    ResetNamesCache

    Set tabl = TableNames()
    tabl.LowerBound = 1

    'Translate admin tables (T_ADM1 through T_ADM4)
    For counter = 1 To 4
        loName = tabl.Item(counter)
        Set Lo = sh.ListObjects(loName)
        Set hRng = Lo.HeaderRowRange

        For counter2 = 1 To counter
            admname = "adm" & counter2 & "_name"
            transGeoName = GeoNames(admname)
            hRng.Cells(1, counter2).Value = IIf(rawNames, admname, transGeoName)
        Next
    Next

    'Translate HF table headers
    Set Lo = sh.ListObjects("T_HF")
    Set hRng = Lo.HeaderRowRange
    transGeoName = GeoNames("hf_name")
    hRng.Cells(1, 1).Value = IIf(rawNames, "hf_name", transGeoName)
    Set cellRng = hRng.Cells(1, 5)

    For counter = 1 To 3
        admname = "adm" & counter & "_name"
        transGeoName = GeoNames(admname)
        cellRng.Offset(, -counter).Value = IIf(rawNames, admname, transGeoName)
    Next

    'Update translation state flag
    rng.Value = Replace(rng.Value, ", not translated", vbNullString)
    rng.Value = Replace(rng.Value, ", translated", vbNullString)
    rng.Value = IIf(rawNames, rng.Value & ", not translated", rng.Value & ", translated")

    'Reset cache after translation change
    ResetNamesCache
End Sub

'@section Clear
'===============================================================================

' @description Clear all geobase data. Clears all ListObject data bodies except
'              T_LevelNames and T_CONCAT, and sets the update flag to "empty".
Private Sub Clear()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = Wksh()

    For Each Lo In sh.ListObjects
        On Error Resume Next
        If Lo.Name <> "T_LevelNames" And Lo.Name <> "T_CONCAT" Then
            Lo.DataBodyRange.Cells.Clear
        End If
        On Error GoTo 0
    Next

    sh.Range("RNG_GeoUpdated").Value = "empty"
    ResetNamesCache
End Sub

' @description Clear historic entries. scope: 0 = geo, 1 = hf, 3 = both.
Private Sub ClearHistoric(Optional ByVal scope As Byte = 3)
    Dim sh As Worksheet
    Set sh = Wksh()

    On Error Resume Next
    If (scope = 0) Or (scope = 3) Then
        sh.ListObjects("T_" & HISTOGEO).DataBodyRange.Cells.Clear
    End If
    If (scope = 1) Or (scope = 3) Then
        sh.ListObjects("T_" & HISTOHF).DataBodyRange.Cells.Clear
    End If
    On Error GoTo 0
End Sub

'@section Import
'===============================================================================

' @description Import geobase data from an external workbook. Each worksheet name
'              matching a known table name (ADM1..ADM4, HF, NAMES, HISTOGEO, etc.)
'              is imported column-by-column into the corresponding ListObject.
'              Concat formulas are regenerated for filtered lookup support.
Private Sub Import(ByVal fromWkb As Workbook, _
                   Optional ByVal histoOnly As Boolean = False)

    Dim sh As Worksheet
    Dim impsh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim admNames As BetterArray
    Dim counter As Long
    Dim colName As String
    Dim colIdx As Long
    Dim lastRow As Long
    Dim impRng As Range
    Dim actRng As Range
    Dim concatFormula As String
    Dim sep As String
    Dim concatRng As Range

    Set admNames = New BetterArray
    admNames.Push HISTOGEO, HISTOHF, METADATA

    If histoOnly Then
        ClearHistoric
    Else
        Clear
        admNames.Push ADMIN1, ADMIN2, ADMIN3, ADMIN4, HF, NAMES
    End If

    Set sh = Wksh()
    sep = " & " & Chr(34) & " | " & Chr(34) & " & "

    Translate rawNames:=True

    For Each impsh In fromWkb.Worksheets
        With impsh
            If admNames.Includes(.Name) Then
                lastRow = Application.WorksheetFunction.Max(.UsedRange.Rows.Count, 2)
                Set Lo = sh.ListObjects("T_" & impsh.Name)
                Set hRng = Lo.HeaderRowRange

                'Import data column by column
                counter = 1
                Do While impsh.Cells(1, counter) <> vbNullString
                    colName = .Cells(1, counter).Value
                    colIdx = ColumnIndex(colName, hRng, inRange:=False)
                    If colIdx > 0 Then
                        Set impRng = .Range(.Cells(2, counter), .Cells(lastRow, counter))
                        Set actRng = sh.Range(sh.Cells(2, colIdx), sh.Cells(lastRow, colIdx))
                        actRng.Value = impRng.Value
                    End If
                    counter = counter + 1
                Loop

                'Build concat formula for the last column of the table
                Set impRng = Lo.HeaderRowRange.Offset(1)

                Select Case Lo.Name
                Case "T_ADM1"
                    concatFormula = "=" & impRng.Cells(1, 1).Address(RowAbsolute:=False)

                Case "T_ADM2"
                    concatFormula = "=" & impRng.Cells(1, 1).Address(RowAbsolute:=False) & _
                                   sep & impRng.Cells(1, 2).Address(RowAbsolute:=False)

                Case "T_ADM3"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False)

                Case "T_ADM4"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 4).Address(RowAbsolute:=False)

                Case "T_HF"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 4).Address(RowAbsolute:=False)

                Case Else
                    concatFormula = vbNullString
                End Select

                'Clear concat column before resizing
                Set concatRng = sh.Range( _
                    sh.Cells(2, hRng.Column + hRng.Cells.Count - 1), _
                    sh.Cells(Lo.Range.Rows.Count, hRng.Column + hRng.Cells.Count - 1))

                If concatFormula <> vbNullString Then concatRng.ClearContents

                'Resize the ListObject to fit imported data
                Set actRng = sh.Range( _
                    sh.Cells(1, hRng.Column), _
                    sh.Cells(lastRow, hRng.Column + hRng.Cells.Count - 1))
                Lo.Resize actRng

                'Apply the concat formula
                If concatFormula <> vbNullString Then
                    sh.Cells(impRng.Row, hRng.Column + hRng.Cells.Count - 1).Formula = _
                        concatFormula
                End If
            End If
        End With
    Next

    sh.Range("RNG_GeoName").Value = Replace(fromWkb.Name, ".xlsx", vbNullString)
    sh.Range("RNG_GeoUpdated").Value = "updated"
    Translate rawNames:=False

    UpdateMetadata
    ResetNamesCache
End Sub

'@section Export
'===============================================================================

' @description Export geobase tables to a new workbook with one sheet per table.
Private Sub ExportToFile(ByVal outputWkb As Workbook, _
                         Optional ByVal onlyHistory As Boolean = False)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outputRng As Range
    Dim inputRng As Range
    Dim sheetName As String
    Dim loName As String
    Dim counter As Long

    Set tabl = New BetterArray
    Set sh = Wksh()
    Translate rawNames:=True

    If onlyHistory Then
        tabl.Push "T_" & HISTOGEO, "T_" & HISTOHF
    Else
        Set tabl = TableNames()
    End If

    With outputWkb
        'Export metadata sheet first
        .Worksheets(1).Name = METADATASHEET
        Set inputRng = sh.ListObjects("T_" & METADATA).Range

        With .Worksheets(METADATASHEET)
            Set outputRng = .Range(.Cells(1, 1), _
                .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
        End With
        outputRng.Value = inputRng.Value
        sheetName = METADATASHEET

        'Export each table to its own sheet
        For counter = tabl.LowerBound To tabl.UpperBound
            loName = tabl.Item(counter)
            Set inputRng = sh.ListObjects(loName).Range
            .Worksheets.Add(After:=.Worksheets(sheetName)).Name = _
                Replace(loName, "T_", vbNullString)
            sheetName = Replace(loName, "T_", vbNullString)

            With .Worksheets(sheetName)
                Set outputRng = .Range(.Cells(1, 1), _
                    .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
            End With
            outputRng.Value = inputRng.Value
        Next
    End With

    Translate rawNames:=False
End Sub

' @description Export the geobase worksheet to a workbook preserving the layout,
'              including ListObjects, named ranges, formulas, pcode references,
'              concat references, and metadata.
Private Sub ExportToWkb(ByVal outputWkb As Workbook, _
                        Optional ByVal llName As String = vbNullString, _
                        Optional ByVal setupPath As String = vbNullString)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As ListObject
    Dim expAddress As String
    Dim expRange As Range
    Dim rng As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String
    Dim metaKeys As BetterArray
    Dim metaValues As BetterArray
    Dim setupName As String

    Set sh = Wksh()

    'Check that target sheets do not already exist
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(sh.Name)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & sh.Name & " in the workbook"
    End If

    On Error Resume Next
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & METADATASHEET & " in the workbook"
    End If

    Translate rawNames:=False

    'Create the geo sheet and copy ListObjects
    outputWkb.Worksheets.Add.Name = sh.Name
    Set outsh = outputWkb.Worksheets(sh.Name)

    For Each Lo In sh.ListObjects
        expAddress = Lo.Range.Address
        Set expRange = outsh.Range(expAddress)
        expRange.Value = Lo.Range.Value

        With outsh
            .ListObjects.Add( _
                SourceType:=xlSrcRange, _
                Source:=expRange, _
                XlListObjectHasHeaders:=xlYes _
            ).Name = Lo.Name
            .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
        End With
    Next

    outsh.Range("A:AZ").Locked = False
    outsh.Range("A:AZ").EntireColumn.AutoFit

    'Export named ranges
    Set rngTable = New BetterArray
    rngTable.Push "RNG_GeoName", "RNG_GeoUpdated", "RNG_PastingGeoCol", _
                  "RNG_GeoLangCode", "RNG_HFNAME", "RNG_ADM1NAME", _
                  "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME", _
                  "RNG_FormLoaded", "RNG_MetaLang"

    For counter = rngTable.LowerBound To rngTable.UpperBound
        rngName = rngTable.Item(counter)
        expAddress = sh.Range(rngName).Address
        Set expRange = outsh.Range(expAddress)

        expRange.Name = rngName
        expRange.Value = sh.Range(rngName).Value
        expRange.Locked = True

        If sh.Range(rngName).Formula <> vbNullString Then
            On Error Resume Next
            expRange.Formula = sh.Range(rngName).Formula
            On Error GoTo 0
        End If
    Next

    'Create metadata sheet
    outputWkb.Worksheets.Add.Name = METADATASHEET
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)

    setupName = Split(setupPath, Application.PathSeparator)( _
        UBound(Split(setupPath, Application.PathSeparator)))

    Set metaKeys = New BetterArray
    Set metaValues = New BetterArray

    metaKeys.Push "variable", "language", "linelist_creation_date", _
                  "linelist_creation_name", "used_setup_name", "used_designer_name"

    metaValues.Push "value", sh.Range("RNG_MetaLang").Value, _
                    Format(Now(), "yyyy-mm-dd Hh:Nn"), _
                    llName, setupName, ThisWorkbook.Name

    With outsh
        metaKeys.ToExcelRange .Cells(1, 1)
        metaValues.ToExcelRange .Cells(1, 2)

        Set expRange = .Cells(metaKeys.Length + 1, 1)
        expRange.Name = "RNG_Meta_Anchor"
        Set expRange = .Range(expRange, expRange.Offset(rng.Rows.Count - 1, 1))
        expRange.Value = rng.Value
        .Visible = xlSheetVeryHidden
    End With

    'Add workbook-level named references for pcodes, concat, population, and historic
    With outputWkb.Names
        .Add Name:="adm1_pcode", RefersToR1C1:="=T_ADM1[adm1_pcode]"
        .Add Name:="adm2_pcode", RefersToR1C1:="=T_ADM2[adm2_pcode]"
        .Add Name:="adm3_pcode", RefersToR1C1:="=T_ADM3[adm3_pcode]"
        .Add Name:="adm4_pcode", RefersToR1C1:="=T_ADM4[adm4_pcode]"
        .Add Name:="hf_pcode", RefersToR1C1:="=T_HF[hf_pcode]"

        .Add Name:="adm1_concat", RefersToR1C1:="=T_ADM1[adm1_concat]"
        .Add Name:="adm2_concat", RefersToR1C1:="=T_ADM2[adm2_concat]"
        .Add Name:="adm3_concat", RefersToR1C1:="=T_ADM3[adm3_concat]"
        .Add Name:="adm4_concat", RefersToR1C1:="=T_ADM4[adm4_concat]"
        .Add Name:="hf_concat", RefersToR1C1:="=T_HF[hf_concat]"

        .Add Name:="adm1_pop", RefersToR1C1:="=T_ADM1[adm1_pop]"
        .Add Name:="adm2_pop", RefersToR1C1:="=T_ADM2[adm2_pop]"
        .Add Name:="adm3_pop", RefersToR1C1:="=T_ADM3[adm3_pop]"
        .Add Name:="adm4_pop", RefersToR1C1:="=T_ADM4[adm4_pop]"

        .Add Name:="histo_geo", RefersToR1C1:="=T_HISTOGEO[HistoGeo]"
        .Add Name:="histo_hf", RefersToR1C1:="=T_HISTOHF[HistoFacility]"
    End With
End Sub

'@section Update
'===============================================================================

' @description Update linelist geo column headers and admin1 dropdown after
'              geobase change. Iterates all HList worksheets, finds geo1-4/hf
'              control columns, and updates their labels with translated names.
Private Sub Update(ByVal pass As IPasswords)

    Const DROPDOWNSHEET As String = "dropdown_lists__"

    Dim sh As Worksheet
    Dim wb As Workbook
    Dim cellRng As Range
    Dim controlValue As String
    Dim geoData As BetterArray
    Dim drop As IDropdownLists
    Dim cellPrintRng As Range
    Dim printsh As Worksheet

    Set wb = Wksh.Parent

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            pass.UnProtect sh.Name

            With sh
                Set cellRng = .Range(.Cells(1, 4).Value & "_START")
                Set printsh = wb.Worksheets("print_" & sh.Name)
                Set cellPrintRng = printsh.Range(cellRng.Address)

                Do While Not IsEmpty(cellRng)
                    controlValue = cellRng.Offset(-4).Value

                    Select Case controlValue
                    Case "geo1"
                        cellRng.Offset(-1).Value = GeoNames("adm1_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm1_name")
                    Case "geo2"
                        cellRng.Offset(-1).Value = GeoNames("adm2_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm2_name")
                    Case "geo3"
                        cellRng.Offset(-1).Value = GeoNames("adm3_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm3_name")
                    Case "geo4"
                        cellRng.Offset(-1).Value = GeoNames("adm4_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm4_name")
                    Case "hf"
                        cellRng.Offset(-1).Value = GeoNames("hf_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("hf_name")
                    End Select

                    Set cellRng = cellRng.Offset(, 1)
                    Set cellPrintRng = cellPrintRng.Offset(, 1)
                Loop
            End With

            pass.Protect sh.Name
        End If
    Next

    'Update the admin1 dropdown
    On Error Resume Next
    Set sh = wb.Worksheets(DROPDOWNSHEET)
    On Error GoTo 0

    If sh Is Nothing Then Exit Sub

    Set drop = DropdownLists.Create(sh)
    Set geoData = GeoLevel(LevelAdmin1)
    drop.Update geoData, "admin1"
End Sub

' @description Update dictionary main labels for geo-related variables.
Private Sub UpdateDict(ByVal dict As ILLdictionary)
    Dim controlRng As Range
    Dim mainlabRng As Range
    Dim counter As Long

    Set controlRng = dict.DataRange("control")
    Set mainlabRng = dict.DataRange("main label")

    For counter = 1 To controlRng.Rows.Count
        Select Case controlRng.Cells(counter, 1).Value
        Case "geo1"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm1_name")
        Case "geo2"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm2_name")
        Case "geo3"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm3_name")
        Case "geo4"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm4_name")
        Case "hf"
            mainlabRng.Cells(counter, 1).Value = GeoNames("hf_name")
        End Select
    Next
End Sub

' @description Update show/hide tables with translated geo labels.
Private Sub UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
    Dim csTab As ICustomTable
    Dim vars As ILLVariables
    Dim Lo As ListObject
    Dim counter As Long
    Dim mainlabRng As Range
    Dim varRng As Range
    Dim controlValue As String
    Dim varValue As String
    Dim subsecValue As String

    Set vars = LLVariables.Create(dict)

    For Each Lo In sh.ListObjects
        Set csTab = CustomTable.Create(Lo)
        Set mainlabRng = csTab.DataRange("main label")
        Set varRng = csTab.DataRange("variable name")

        For counter = 1 To varRng.Rows.Count
            varValue = varRng.Cells(counter, 1).Value
            controlValue = vars.Value(colName:="control", varName:=varValue)
            subsecValue = vars.Value(colName:="sub section", varName:=varValue)

            Select Case controlValue
            Case "geo1"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm1_name")
            Case "geo2"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm2_name")
            Case "geo3"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm3_name")
            Case "geo4"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm4_name")
            Case "hf"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("hf_name")
            End Select
        Next
    Next
End Sub

'@section Historic
'===============================================================================

' @description Add a new entry to the historic data, resizing the ListObject
'              and removing duplicates.
Private Sub UpdateHistoric(ByVal newValue As String, _
                           Optional ByVal scope As Byte = 0)
    Dim sh As Worksheet
    Dim loName As String
    Dim LoRng As Range
    Dim cellRng As Range
    Dim Lo As ListObject

    Set sh = Wksh()
    loName = IIf(scope = 0, HISTOGEO, HISTOHF)
    loName = "T_" & loName

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range
    Set cellRng = LoRng.Cells(LoRng.Rows.Count + 1, 1)

    If Lo.DataBodyRange Is Nothing Then
        Set cellRng = LoRng.Cells(2, 1)
    ElseIf IsEmpty(Lo.DataBodyRange) Then
        Set cellRng = LoRng.Cells(2, 1)
    End If

    cellRng.Value = newValue
    Lo.Resize sh.Range(LoRng.Cells(1, 1), cellRng)

    On Error Resume Next
    LoRng.RemoveDuplicates Columns:=1, Header:=xlYes
    LoRng.Sort key1:=LoRng, Header:=xlYes
    On Error GoTo 0
End Sub

'@section AdminCode
'===============================================================================

' @description Resolve the admin code from a translated admin display name.
Private Function AdminCode(ByVal selectedAdmin As String) As String
    Select Case selectedAdmin
    Case GeoNames("adm1_name")
        AdminCode = "adm1"
    Case GeoNames("adm2_name")
        AdminCode = "adm2"
    Case GeoNames("adm3_name")
        AdminCode = "adm3"
    Case GeoNames("adm4_name")
        AdminCode = "adm4"
    Case Else
        AdminCode = "adm1"
    End Select
End Function

' @description Whether the geobase has no imported data.
Private Property Get HasNoData() As Boolean
    HasNoData = (Wksh().Range("RNG_GeoUpdated").Value = "empty")
End Property

'@section Internal Helpers
'===============================================================================

' @description Update the metadata sheet from the geo metadata table.
Private Sub UpdateMetadata()
    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim metash As Worksheet
    Dim rng As Range
    Dim expRng As Range
    Dim lastRow As Long

    On Error GoTo EndUpdate

    Set sh = Wksh()
    Set metash = sh.Parent.Worksheets(METADATASHEET)

    With metash
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set expRng = .Range("RNG_Meta_Anchor")

        If lastRow >= expRng.Row Then
            Set rng = .Range(expRng, .Cells(lastRow, 2))
            rng.ClearContents
        End If

        Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)
        Set expRng = .Range(expRng, expRng.Offset(rng.Rows.Count - 1, 1))
        expRng.Value = rng.Value
    End With

EndUpdate:
End Sub

'@section Validation
'===============================================================================

' @description Validate that the worksheet has all required tables and named ranges.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim admTableNames As BetterArray

    If sh Is Nothing Then
        ThrowError ObjectNotInitialized, "Geo sheet not defined"
    End If

    Set admTableNames = New BetterArray
    admTableNames.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
                       "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF, _
                       "T_" & METADATA
    LoExists admTableNames, sh

    RangeExists sh, "RNG_GeoName"
    RangeExists sh, "RNG_GeoUpdated"
    RangeExists sh, "RNG_PastingGeoCol"
    RangeExists sh, "RNG_GeoLangCode"
    RangeExists sh, "RNG_HFNAME"
    RangeExists sh, "RNG_ADM1NAME"
    RangeExists sh, "RNG_ADM2NAME"
    RangeExists sh, "RNG_ADM3NAME"
    RangeExists sh, "RNG_ADM4NAME"
    RangeExists sh, "RNG_FormLoaded"
    RangeExists sh, "RNG_MetaLang"
End Sub

Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then
        ThrowError ElementNotFound, _
            "Range " & rngName & " does not exist in worksheet " & sh.Name
    End If
End Sub

Private Sub LoExists(ByVal loNames As BetterArray, ByVal sh As Worksheet)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String

    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        Set Lo = Nothing

        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0

        If Lo Is Nothing Then
            ThrowError ElementNotFound, _
                "Table " & lName & " does not exist in worksheet " & sh.Name
        End If
    Next
End Sub

'@section Error Handling
'===============================================================================

' @description Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================

Private Property Get ILLGeo_GeoLevel(ByVal adminLevel As Byte, _
                                     Optional ByVal geoType As Byte = GeoScopeAdmin, _
                                     Optional adminNames As Variant = vbNullString _
                                     ) As BetterArray
    Set ILLGeo_GeoLevel = GeoLevel(adminLevel, geoType, adminNames)
End Property

Private Property Get ILLGeo_GeoNames(ByVal admname As String) As String
    ILLGeo_GeoNames = GeoNames(admname)
End Property

Private Property Get ILLGeo_HasNoData() As Boolean
    ILLGeo_HasNoData = HasNoData()
End Property

Private Sub ILLGeo_Translate(Optional ByVal rawNames As Boolean = False)
    Translate rawNames
End Sub

Private Sub ILLGeo_Import(ByVal fromWkb As Workbook, _
                          Optional ByVal histoOnly As Boolean = False)
    Import fromWkb, histoOnly
End Sub

Private Sub ILLGeo_ExportToFile(ByVal outputWkb As Workbook, _
                                Optional ByVal onlyHistory As Boolean = False)
    ExportToFile outputWkb, onlyHistory
End Sub

Private Sub ILLGeo_ExportToWkb(ByVal Wkb As Workbook, _
                               Optional ByVal llName As String = vbNullString, _
                               Optional ByVal setupPath As String = vbNullString)
    ExportToWkb Wkb, llName, setupPath
End Sub

Private Sub ILLGeo_Clear()
    Clear
End Sub

Private Sub ILLGeo_ClearHistoric(Optional ByVal scope As Byte = 3)
    ClearHistoric scope
End Sub

Private Sub ILLGeo_UpdateHistoric(ByVal newValue As String, _
                                  Optional ByVal scope As Byte = 0)
    UpdateHistoric newValue, scope
End Sub

Private Sub ILLGeo_Update(ByVal pass As IPasswords)
    Update pass
End Sub

Private Sub ILLGeo_UpdateDict(ByVal dict As ILLdictionary)
    UpdateDict dict
End Sub

Private Sub ILLGeo_UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
    UpdateShowHide sh, dict
End Sub

Private Function ILLGeo_AdminCode(ByVal selectedAdmin As String) As String
    ILLGeo_AdminCode = AdminCode(selectedAdmin)
End Function
