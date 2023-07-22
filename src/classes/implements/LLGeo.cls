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
'@Folder("Geobase")
'@ModuleDescription("Geobase management class")

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
    sh.calculate
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
    Case "geobase"
        actualName = sh.Range("RNG_GeoName").Value
    Case Else
        actualName = admname
    End Select

    'Important to test the type because those ranges are formulas.
    If TypeName(actualName) = "String" Then
        GeoNames = actualName
    Else
        GeoNames = admname
    End If
End Property

Private Function GeoColumnIndex(ByVal colName As String, ByVal adminLevel As Byte) As Long
    Dim hRng As Range
    Dim LoName As String
    Dim tabl As BetterArray
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim actualColName As String

    Set sh = Wksh()
    Set tabl = TableNames()
    LoName = tabl.Item(adminLevel)
    actualColName = colName
    If sh.Range("RNG_GeoUpdated").Value = "updated, translated" Then actualColName = GeoNames(colName)
    Set hRng = sh.ListObjects(LoName).HeaderRowRange
    colIndex = ColumnIndex(actualColName, hRng)
    If colIndex < 0 Then ThrowError "Column " & actualColName & " not found in table " & LoName
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
    Dim Index As Long
    If Not (headerRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
        Index = headerRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True).Column
        If inRange Then Index = Index - headerRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
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

Private Property Get GeoLevel(ByVal adminLevel As Byte, _
                              Optional ByVal geoType As Byte = GeoScopeAdmin, _
                              Optional ByVal adminNames As Variant = vbNullString _
                              ) As BetterArray

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outLevels As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim rng As Range
    Dim LoName As String
    Dim colIndex As Long
    Dim counter As Long
    Dim returnedRng As Range


    Set sh = Wksh()
    Set tabl = TableNames()

    On Error GoTo Err

    sh.EnableCalculation = False

    If geoType = GeoScopeAdmin Then
        LoName = tabl.Item(adminLevel)
    Else
        LoName = tabl.Item(4)
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
        If adminNames.LowerBound <> 1 Then ThrowError "Please ensure your BetterArray has 1 as Lowerbound, it is mandatory"
    End If

    'Output Levels
    Set outLevels = New BetterArray
    Set Lo = sh.ListObjects(LoName)
    Set LoRng = Lo.Range

    'If the Listobject is Empty, throw error
    If Lo.DataBodyRange Is Nothing Then ThrowError "The listobject " & LoName & "has no data"

    Select Case geoType

    Case GeoScopeAdmin

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

    Case GeoScopeHF

        Select Case adminLevel

        Case LevelAdmin1

            Set rng = HFRange("adm1_name")
            rng.Copy sh.Range("RNG_PastingGeoCol")
            Set outLevels = UniqueValues()

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
        On Error GoTo Err
        If Not returnedRng Is Nothing Then
            returnedRng.Copy sh.Range("RNG_PastingGeoCol")
            Lo.AutoFilter.ShowAllData
            Set outLevels = UniqueValues()
        End If

        Lo.AutoFilter.ShowAllData
    End If

    sh.Range("RNG_PastingGeoCol").EntireColumn.Cells.Clear

Err:
    Set GeoLevel = outLevels.Clone()
    sh.EnableCalculation = True
End Property


Private Property Get ValueOf(ByVal adminLevel As Byte, _
        Optional ByVal concatValue As String = vbNullString, _
        Optional ByVal returnedSuffix As String = "pop") As String

    'Return the population or the pcode of a admin level, by default, return population
    'Returned Suffix is the suffix of the column to return.

    Dim sh As Worksheet
    Dim Lo As listObject
    Dim LoName As String
    Dim colIndex As Long
    Dim concatRng As Range
    Dim tabl As BetterArray
    Dim returnedValue As String
    Dim cellRng As Range
    Dim retRng As Range

    Set sh = Wksh()
    Set tabl = TableNames()

    LoName = IIf(adminLevel = LevelHF, "T_HF", tabl.Item(adminLevel))
    Set Lo = sh.ListObjects(LoName)

    If Lo.DataBodyRange Is Nothing Then Exit Property

    'returned column index
    If adminLevel <> LevelHF Then
        colIndex = GeoColumnIndex("adm" & (adminLevel + 1) & "_" & returnedSuffix, adminLevel)
    Else
        colIndex = HFColumnIndex("hf_" & returnedSuffix)
    End If

    Select Case adminLevel

        Case LevelGeoAdmin1

        Set concatRng = Lo.ListColumns("adm1_concat").Range

        Case LevelGeoAdmin2

        Set concatRng = Lo.ListColumns("adm2_concat").Range

        Case LevelGeoAdmin3

        Set concatRng = Lo.ListColumns("adm3_concat").Range

        Case LevelGeoAdmin4

        Set concatRng = Lo.ListColumns("adm4_concat").Range

        Case LevelHF

        Set concatRng = Lo.ListColumns("hf_concat").Range
    End Select

    'Use index match to get the result
    Set cellRng = concatRng.Cells(1, 1)
    Set retRng = Lo.Range.Cells(1, colIndex)

    Do While (cellRng.Value <> concatValue) And _ 
             (cellRng.Row <= concatRng.Cells(concatRng.Rows.Count, 1).Row) 
        Set cellRng = cellRng.Offset(1)
        Set retRng = retRng.Offset(1)   
    Loop

    returnedValue = vbNullString

    If (cellRng.Row <= concatRng.Cells(concatRng.Rows.Count, 1).Row)  Then
        On Error Resume Next
        returnedValue = CStr(retRng.Value)
        On Error GoTo 0
    End If

    ValueOf = returnedValue
End Property


Private Sub Translate(Optional ByVal rawNames As Boolean = False)
    'Headers for the geo sheet
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim counter As Long
    Dim counter2 As Long
    Dim rng As Range
    Dim cellRng As Range
    Dim Lo As ListObject
    Dim LoName As String
    Dim admname As String
    Dim transGeoName As String                   'Translated geo Name
    Dim hRng As Range                            'Header Range

    Set sh = Wksh()
    sh.calculate
    Set rng = sh.Range("RNG_GeoUpdated")
    If Not rawNames And (InStr(1, rng.Value, ", translated") > 0) Then Exit Sub
    If rawNames And (InStr(1, rng.Value, ", not translated") > 0) Then Exit Sub

    Set tabl = TableNames()
    tabl.LowerBound = 1
    For counter = 1 To 4 'admin 1, 2, 3, 4
        LoName = tabl.Item(counter)
        Set Lo = sh.ListObjects(LoName)
        Set hRng = Lo.HeaderRowRange

        'Admin 1 to 5, and hf
        For counter2 = 1 To counter
            admname = "adm" & counter2 & "_name"
            transGeoName = GeoNames(admname)
            'I have to take in account hf, since hf table is reversed
            hRng.Cells(1, counter2).Value = IIf(rawNames, admname, transGeoName)
        Next
    Next

    'Update informations for health facility
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

    'Informations for next translations
    rng.Value = Replace(rng.Value, ", not translated", "")
    rng.Value = Replace(rng.Value, ", translated", "")
    rng.Value = IIf(rawNames, rng.Value & ", not translated", rng.Value & ", translated")
End Sub

Private Sub Clear()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = Wksh()
    For Each Lo In sh.ListObjects
        On Error Resume Next
        If Lo.Name <> "T_LevelNames" And Lo.Name <> "T_CONCAT" Then Lo.DataBodyRange.Cells.Clear
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
    Dim LoName As String
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
            LoName = tabl.Item(counter)
            Set inputRng = sh.ListObjects(LoName).Range
            .Worksheets.Add(After:=.Worksheets(sheetName)).Name = Replace(LoName, "T_", vbNullString)
            sheetName = Replace(LoName, "T_", vbNullString)
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


    'Test if a sheet named as the geo sheet already exists
    Set sh = Wksh()
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(sh.Name)
    On Error GoTo 0

    If Not outsh Is Nothing Then ThrowError "Unable to export, already a sheet named " & sh.Name & " in the workbook"

    'Test if there is a metadata sheet
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    On Error GoTo 0

    If Not outsh Is Nothing Then ThrowError "Unable to export, already a sheet named " & METADATASHEET & " in the workbook"


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

    'Unlock some columns
    outsh.Range("A:AZ").Locked = False
    outsh.Range("A:AZ").EntireColumn.AutoFit

    'Export all the ranges
    Set rngTable = New BetterArray
    rngTable.Push "RNG_GeoName", "RNG_GeoUpdated", "RNG_PastingGeoCol", "RNG_GeoLangCode", _
                  "RNG_HFNAME", "RNG_ADM1NAME", "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME", _
                  "RNG_FormLoaded", "RNG_MetaLang"

    For counter = rngTable.LowerBound To rngTable.UpperBound

        rngName = rngTable.Item(counter)
        expAddress = sh.Range(rngName).Address
        Set expRange = outsh.Range(expAddress)

        expRange.Name = rngName
        expRange.Value = sh.Range(rngName).Value
        expRange.Locked = True

        If sh.Range(rngName).formula <> vbNullString Then
            On Error Resume Next
            expRange.formula = sh.Range(rngName).formula
            On Error GoTo 0
        End If

    Next

    'Add the metadata sheet
    outputWkb.Worksheets.Add.Name = METADATASHEET
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    Set rng = sh.ListObjects("T_" & METADATA).Range

    With outsh
        Set expRange = .Range(.Cells(1, 1), .Cells(rng.Rows.Count, rng.Columns.Count))
        expRange.Value = rng.Value

        'Linelist language
        .Cells(rng.Rows.Count + 1, 1).Value = "language"
        .Cells(rng.Rows.Count + 1, 2).Value = sh.Range("RNG_MetaLang").Value

        'Linelist creation date
        .Cells(rng.Rows.Count + 2, 1).Value = "linelist_creation_date"
        .Cells(rng.Rows.Count + 2, 2).Value = Format(Now, "yyyy-mm-dd Hh:Nn")

        .Visible = xlSheetVeryHidden

    End With

    'Add the names to the output workbook

    'pcodes
    outputWkb.NAMES.Add Name:="adm1_pcode", RefersToR1C1:="=" & "T_ADM1" & "[" & "adm1_pcode" & "]"
    outputWkb.NAMES.Add Name:="adm2_pcode", RefersToR1C1:="=" & "T_ADM2" & "[" & "adm2_pcode" & "]"
    outputWkb.NAMES.Add Name:="adm3_pcode", RefersToR1C1:="=" & "T_ADM3" & "[" & "adm3_pcode" & "]"
    outputWkb.NAMES.Add Name:="adm4_pcode", RefersToR1C1:="=" & "T_ADM4" & "[" & "adm4_pcode" & "]"
    outputWkb.NAMES.Add Name:="hf_pcode", RefersToR1C1:="=" & "T_HF" & "[" & "hf_pcode" & "]"

    'concat values
    outputWkb.NAMES.Add Name:="adm1_concat", RefersToR1C1:="=" & "T_ADM1" & "[" & "adm1_concat" & "]"
    outputWkb.NAMES.Add Name:="adm2_concat", RefersToR1C1:="=" & "T_ADM2" & "[" & "adm2_concat" & "]"
    outputWkb.NAMES.Add Name:="adm3_concat", RefersToR1C1:="=" & "T_ADM3" & "[" & "adm3_concat" & "]"
    outputWkb.NAMES.Add Name:="adm4_concat", RefersToR1C1:="=" & "T_ADM4" & "[" & "adm4_concat" & "]"
    outputWkb.NAMES.Add Name:="hf_concat", RefersToR1C1:="=" & "T_HF" & "[" & "hf_concat" & "]"

    'population values
    outputWkb.NAMES.Add Name:="adm1_pop", RefersToR1C1:="=" & "T_ADM1" & "[" & "adm1_pop" & "]"
    outputWkb.NAMES.Add Name:="adm2_pop", RefersToR1C1:="=" & "T_ADM2" & "[" & "adm2_pop" & "]"
    outputWkb.NAMES.Add Name:="adm3_pop", RefersToR1C1:="=" & "T_ADM3" & "[" & "adm3_pop" & "]"
    outputWkb.NAMES.Add Name:="adm4_pop", RefersToR1C1:="=" & "T_ADM4" & "[" & "adm4_pop" & "]"


End Sub

Private Sub Import(ByVal fromWkb As Workbook, Optional ByVal histoOnly As Boolean = False)
    Dim sh As Worksheet
    Dim impsh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim admNames As BetterArray
    Dim counter As Long
    Dim colName As String
    Dim colIndex As Long
    Dim LastRow As Long
    Dim impRng As Range
    Dim actRng As Range
    Dim concatFormula As String
    Dim sep As String
    Dim concatRng As Range


    Set admNames = New BetterArray
    admNames.Push HISTOGEO, HISTOHF, METADATA
    'Clear historic or all the geobase
    If histoOnly Then
        ClearHistoric
    Else
        Clear
        admNames.Push ADMIN1, ADMIN2, ADMIN3, ADMIN4, HF, NAMES
    End If

    Set sh = Wksh()
    'Separator for concatenate
    sep = " & " & chr(34) & " | " & chr(34) & " & "

    'Clear the Current Geo

    Translate rawNames:=True

    For Each impsh In fromWkb.Worksheets
        'Only sheet with valid names are used for import
        With impsh
            If admNames.Includes(.Name) Then
                'If the lastRow is < 2, then the lastRow is 2. To skip/avoid errors on
                'Empty worksheets
                LastRow = Application.WorksheetFunction.Max(.UsedRange.Rows.Count, 2)
                Set Lo = sh.ListObjects("T_" & impsh.Name) 'Lo Name
                Set hRng = Lo.HeaderRowRange
                counter = 1
                Do While impsh.Cells(1, counter) <> vbNullString
                    colName = .Cells(1, counter).Value
                    colIndex = ColumnIndex(colName, hRng, inRange:=False)
                    If colIndex > 0 Then
                        Set impRng = .Range(.Cells(2, counter), .Cells(LastRow, counter))
                        Set actRng = sh.Range(sh.Cells(2, colIndex), sh.Cells(LastRow, colIndex))
                        actRng.Value = impRng.Value
                    End If
                    counter = counter + 1
                Loop
                'Add a concat table (for admin selections)
                Set impRng = Lo.HeaderRowRange.Offset(1)

                'Clear the adm concat range before writing the formulas.
                
                Select Case Lo.Name

                Case "T_ADM1"

                    'Creating the formula
                    concatFormula = "=" & impRng.Cells(1, 1).Address(Rowabsolute:=False)

                Case "T_ADM2"

                    concatFormula = "=" & impRng.Cells(1, 2).Address(Rowabsolute:=False) & _
                                                                                         sep & _
                                                                                         impRng.Cells(1, 1).Address(Rowabsolute:=False)
                Case "T_ADM3"

                    concatFormula = "=" & _
                                    impRng.Cells(1, 3).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 2).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 1).Address(Rowabsolute:=False)

                Case "T_ADM4"

                    concatFormula = "=" & _
                                    impRng.Cells(1, 4).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 3).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 2).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 1).Address(Rowabsolute:=False)

                Case "T_HF"

                    concatFormula = "=" & _
                                    impRng.Cells(1, 1).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 2).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 3).Address(Rowabsolute:=False) & _
                                    sep & _
                                    impRng.Cells(1, 4).Address(Rowabsolute:=False)

                Case Else

                    concatFormula = vbNullString

                End Select

                'clear the concat range before resizing and setting the formula
                Set concatRng = sh.Range( _ 
                                    sh.Cells(2,  _ 
                                             hRng.Column + hRng.Cells.Count - 1), _ 
                                    sh.Cells(Lo.Range.Rows.Count, _ 
                                             hRng.Column + hRng.Cells.Count - 1) _ 
                                )
                'Clear the concat Range only if there is a concat formula
                'You need to clear the concat Range before adding the formula
                'Or resizing the listObject.
                If concatFormula <> vbNullString Then concatRng.ClearContents

                'I know that each geo listobject starts on the first line
                Set actRng = sh.Range(sh.Cells(1, hRng.Column), sh.Cells(LastRow, hRng.Column + hRng.Cells.Count - 1))

                'Resize the listObject and add a concatenate formula
                Lo.Resize actRng

                If concatFormula <> vbNullString Then _
                sh.Cells(impRng.Row, _ 
                        hRng.Column + hRng.Cells.Count - 1).formula = _ 
                        concatFormula
                
            End If
        End With
    Next

    'Add the workbook name (Geobase Name)
    sh.Range("RNG_GeoName").Value = Replace(fromWkb.Name, ".xlsx", vbNullString)
    sh.Range("RNG_GeoUpdated").Value = "updated"
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
    RangeExists sh, "RNG_FormLoaded"
    RangeExists sh, "RNG_MetaLang"
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

'Update the geobase admin names in the linelist
Private Sub Update(ByVal pass As ILLPasswords)

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

        'Test if it is a Hlist
        If sh.Cells(1, 3).Value = "HList" Then


            'Unprotect and update
            pass.UnProtect sh.Name

            With sh

                'Start of the table in the HList
                '.Cells(1, 4) contains the name of table in the linelist HList sheet

                Set cellRng = .Range(.Cells(1, 4).Value & "_" & "START")
                Set printsh = wb.Worksheets("print_" & sh.Name)
                Set cellPrintRng = printsh.Range(cellRng.Address)

                Do While Not IsEmpty(cellRng.Offset(-1))

                    controlValue = cellRng.Offset(-5).Value

                    Select Case controlValue

                    Case "geo1"

                        cellRng.Offset(-2).Value = GeoNames("adm1_name")
                        cellPrintRng.Offset(-2).Value = GeoNames("adm1_name")

                    Case "geo2"

                        cellRng.Offset(-2).Value = GeoNames("adm2_name")
                        cellPrintRng.Offset(-2).Value = GeoNames("adm2_name")

                    Case "geo3"

                        cellRng.Offset(-2).Value = GeoNames("adm3_name")
                        cellPrintRng.Offset(-2).Value = GeoNames("adm3_name")

                    Case "geo4"

                        cellRng.Offset(-2).Value = GeoNames("adm4_name")
                        cellPrintRng.Offset(-2).Value = GeoNames("adm4_name")

                    Case "hf"

                        cellRng.Offset(-2).Value = GeoNames("hf_name")
                        cellPrintRng.Offset(-2).Value = GeoNames("hf_name")

                    End Select

                    Set cellRng = cellRng.Offset(, 1)
                    Set cellPrintRng = cellPrintRng.Offset(, 1)

                Loop

            End With

            pass.Protect sh.Name
        End If
    Next


    On Error Resume Next
    'Update the dropdown for geo1
    Set sh = wb.Worksheets(DROPDOWNSHEET)
    On Error GoTo 0

    If (sh Is Nothing) Then Exit Sub
    Set drop = DropdownLists.Create(sh)
    Set geoData = GeoLevel(LevelAdmin1)

    'Now update the dropdown
    drop.Update geoData, "admin1"

End Sub

Private Sub UpdateDict(ByVal dict As ILLdictionary)
    
    Dim controlRng As Range
    Dim mainlabRng As Range
    Dim counter As Long

    Set controlRng = dict.DataRange("control")
    Set mainlabRng = dict.DataRange("main label")

    For counter = 1 to controlRng.Rows.Count
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

    For Each Lo in sh.ListObjects

        Set csTab = CustomTable.Create(Lo)
        Set mainlabRng = csTab.DataRange("main label")
        Set varRng = csTab.DataRange("variable name")

        For counter = 1 To varRng.Rows.Count

            varValue = varRng.Cells(counter, 1).Value
            controlValue = vars.Value(colName:="control", varName:=varValue)
            subsecValue = vars.Value(colName:="sub section", varName:=varValue)

            Select Case controlValue
            case "geo1"
                mainlabRng.Cells(counter, 1).Value = subsecValue & ": " & GeoNames("adm1_name")
            Case "geo2"
                mainlabRng.Cells(counter, 1).Value = subsecValue & ": " & GeoNames("adm2_name")    
            Case "geo3"
                mainlabRng.Cells(counter, 1).Value = subsecValue & ": " & GeoNames("adm3_name")
            Case "geo4"
                mainlabRng.Cells(counter, 1).Value = subsecValue & ": " & GeoNames("adm4_name")
            Case "hf"
                mainlabRng.Cells(counter, 1).Value = subsecValue & ": " & GeoNames("hf_name")
            End Select
        Next
    Next

End Sub

'Find the admin level given one admin name
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

'Interfaces
Private Property Get ILLGeo_GeoLevel(ByVal adminLevel As Byte, _
                                    Optional ByVal geoType As Byte = GeoScopeAdmin, _
                                    Optional ByVal adminNames As Variant = vbNullString) As BetterArray

    Set ILLGeo_GeoLevel = GeoLevel(adminLevel, geoType, adminNames)
End Property

'Convert the code name to the actual name of the current language
Private Property Get ILLGeo_GeoNames(ByVal admname As String) As String
    ILLGeo_GeoNames = GeoNames(admname)
End Property

'Get the Population of one administrative level (only for geoType)
Private Property Get ILLGeo_Population(ByVal adminLevel As Byte, Optional ByVal concatValue As String = vbNullString) As String
    ILLGeo_Population = ValueOf(adminLevel:=adminLevel, concatValue:=concatValue, returnedSuffix:="pop")
End Property

Private Property Get ILLGeo_Pcode(ByVal adminLevel As Byte, Optional ByVal concatValue As String = vbNullString) As String
    ILLGeo_Pcode = ValueOf(adminLevel:=adminLevel, concatValue:=concatValue, returnedSuffix:="pcode")
End Property

'Translation
Private Sub ILLGeo_Translate(Optional ByVal rawNames As Boolean = False)
    Translate rawNames
End Sub

'Import a new Geobase from a workbook
Private Sub ILLGeo_Import(ByVal fromWkb As Workbook, Optional ByVal histoOnly As Boolean = False)
    Import fromWkb, histoOnly
End Sub

'Export to a workbook
Private Sub ILLGeo_ExportToFile(ByVal outputWkb As Workbook, Optional ByVal onlyHistory As Boolean = False)
    ExportToFile outputWkb, onlyHistory
End Sub

Private Sub ILLGeo_ExportToWkb(ByVal outputWkb As Workbook)
    ExportToWkb outputWkb
End Sub

Private Sub ILLGeo_Clear()
    Clear
End Sub

Private Sub ILLGeo_ClearHistoric()
    ClearHistoric
End Sub

Private Sub ILLGeo_Update(ByVal pass As ILLPasswords)
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

