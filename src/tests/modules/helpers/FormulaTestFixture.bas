Attribute VB_Name = "FormulaTestFixture"
Attribute VB_Description = "Fixture helpers for formula tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Fixture helpers for formula tests")

'@section Fixture Data
'===============================================================================

'@description Header row for the allowed Excel functions table.
Public Function FormulaFunctionsHeaderRow() As Variant
    FormulaFunctionsHeaderRow = Array("ENG")
End Function

'@description Complete list of allowed Excel functions (from draft1.csv).
Public Function FormulaFunctionsValues() As Variant
    Dim parts As Collection
    Dim segments As Variant

    Set parts = New Collection

    parts.Add "ABS|ACCRINT|ACCRINTM|ACOS|ACOSH|ACOT|ACOTH|ADRESSE|AGGREGATE|AMORDEGRC|AMORLINC|AND|ARABIC|AREAS|ASC|ASIN|ASINH|ATAN|ATAN2|ATANH|AVEDEV"
    parts.Add "AVERAGE|AVERAGEA|AVERAGEIF|AVERAGEIFS|BAHTTEXT|BASE|BESSELI|BESSELJ|BESSELK|BESSELY|BETA.DIST|BETA.INV|BETADIST|BETAINV|BIN2DEC|BIN2HEX|BIN2OCT|BINOM.DIST|BINOM.DIST.RANGE|BINOM.INV|BINOMDIST"
    parts.Add "BITAND|BITLSHIFT|BITOR|BITRSHIFT|BITXOR|CEILING|CEILING.MATH|CEILING.PRECISE|CELL|CHAR|CHIDIST|CHIINV|CHISQ.DIST|CHISQ.DIST.RT|CHISQ.INV|CHISQ.INV.RT|CHISQ.TEST|CHITEST|CHOOSE|CLEAN|CODE"
    parts.Add "COLUMN|COLUMNS|COMBIN|COMBINA|COMPLEX|CONCATENATE|CONFIDENCE|CONFIDENCE.NORM|CONFIDENCE.T|CONVERT|CORREL|COS|COSH|COT|COTH|COUNT|COUNTA|COUNTBLANK|COUNTIF|COUNTIFS|COUPDAYBS"
    parts.Add "COUPDAYS|COUPDAYSNC|COUPNCD|COUPNUM|COUPPCD|COVAR|COVARIANCE.P|COVARIANCE.S|CRITBINOM|CSC|CSCH|CUBEKPIMEMBER|CUBEMEMBER|CUBEMEMBERPROPERTY|CUBERANKEDMEMBER|CUBESET|CUBESETCOUNT|CUBEVALUE|CUMIPMT|CUMPRINC|DATE"
    parts.Add "DATEDIF|DATEVALUE|DAVERAGE|DAY|DAYS|DAYS360|DB|DCOUNT|DCOUNTA|DDB|DEC2BIN|DEC2HEX|DEC2OCT|DECIMAL|DEGREES|DELTA|DEVSQ|DGET|DISC|DMAX|DMIN"
    parts.Add "DOLLAR|DOLLARDE|DOLLARFR|DPRODUCT|DSTDEV|DSTDEVP|DSUM|DURATION|DVAR|DVARP|EDATE|EFFECT|ENCODEURL|EOMONTH|EPIWEEK|ERF|ERF.PRECISE|ERFC|ERFC.PRECISE|ERROR.TYPE|EVEN"
    parts.Add "EXACT|EXP|EXPON.DIST|EXPONDIST|F.DIST|F.DIST.RT|F.INV|F.INV.RT|F.TEST|FACT|FACTDOUBLE|FALSE|FDIST|FILTERXML|FIND|FINDB|FINV|FISHER|FISHERINV|FIXED|FLOOR"
    parts.Add "FLOOR.MATH|FLOOR.PRECISE|FORECAST|FORMULATEXT|FREQUENCY|FTEST|FV|FVSCHEDULE|GAMMA|GAMMA.DIST|GAMMA.INV|GAMMADIST|GAMMAINV|GAMMALN|GAMMALN.PRECISE|GAUSS|GCD|GEOMEAN|GESTEP|GETPIVOTDATA|GROWTH"
    parts.Add "HARMEAN|HEX2BIN|HEX2DEC|HEX2OCT|HLOOKUP|HOUR|HYPERLINK|HYPGEOM.DIST|HYPGEOMDIST|IF|IFERROR|IFNA|IMABS|IMAGINARY|IMARGUMENT|IMCONJUGATE|IMCOS|IMCOSH|IMCOT|IMCSC|IMCSCH"
    parts.Add "IMDIV|IMEXP|IMLN|IMLOG10|IMLOG2|IMPOWER|IMPRODUCT|IMREAL|IMSEC|IMSECH|IMSIN|IMSINH|IMSQRT|IMSUB|IMSUM|IMTAN|INDEX|INDIRECT|INFO|INT|INTERCEPT"
    parts.Add "INTRATE|IPMT|IRR|ISBLANK|ISERR|ISERROR|ISEVEN|ISFORMULA|ISLOGICAL|ISNA|ISNONTEXT|ISNUMBER|ISO.CEILING|ISODD|ISOWEEKNUM|ISPMT|ISREF|ISTEXT|JIS|KURT|LARGE"
    parts.Add "LCM|LEFT|LEFTB|LEN|LENB|LINEST|LN|LOG|LOG10|LOGEST|LOGINV|LOGNORM.DIST|LOGNORM.INV|LOGNORMDIST|LOOKUP|LOWER|MATCH|MAX|MAXA|MDETERM|MDURATION"
    parts.Add "MEDIAN|MID|MIDB|MIN|MINA|MINUTE|MINVERSE|MIRR|MMULT|MOD|MODE|MODE.MULT|MODE.SNGL|MONTH|MROUND|MULTINOMIAL|MUNIT|NA|NEGBINOM.DIST|NEGBINOMDIST"
    parts.Add "NETWORKDAYS|NETWORKDAYS.INTL|NOMINAL|NORM.DIST|NORM.INV|NORM.S.DIST|NORM.S.INV|NORMDIST|NORMINV|NORMSDIST|NORMSINV|NOT|NOW|NPER|NPV|NUMBERVALUE|OCT2BIN|OCT2DEC|OCT2HEX|ODD|ODDFPRICE"
    parts.Add "ODDFYIELD|ODDLPRICE|ODDLYIELD|OFFSET|OR|PDURATION|PEARSON|PERCENTILE|PERCENTILE.EXC|PERCENTILE.INC|PERCENTRANK|PERCENTRANK.EXC|PERCENTRANK.INC|PERMUT|PERMUTATIONA|PHI|PHONETIC|PI|PMT|POISSON|POISSON.DIST"
    parts.Add "POWER|PPMT|PRICE|PRICEDISC|PRICEMAT|PROB|PRODUCT|PROPER|PV|QUARTILE|QUARTILE.EXC|QUARTILE.INC|QUOTIENT|RADIANS|RAND|RANDBETWEEN|RANK|RANK.AVG|RANK.EQ|RATE|RECEIVED"
    parts.Add "REPLACE|REPLACEB|REPT|RIGHT|RIGHTB|ROMAN|ROUND|ROUNDDOWN|ROUNDUP|ROW|ROWS|RRI|RSQ|RTD|SEARCH|SEARCHB|SEC|SECH|SECOND|SERIESSUM|SHEET"
    parts.Add "SHEETS|SIGN|SIN|SINH|SKEW|SKEW.P|SLN|SLOPE|SMALL|SQRT|SQRTPI|STANDARDIZE|STDEV|STDEV.P|STDEV.S|STDEVA|STDEVP|STDEVPA|STEYX|SUBSTITUTE|SUBTOTAL"
    parts.Add "SUM|SUMIF|SUMIFS|SUMPRODUCT|SUMSQ|SUMX2MY2|SUMX2PY2|SUMXMY2|SYD|T|T.DIST|T.DIST.2T|T.DIST.RT|T.INV|T.INV.2T|T.TEST|TAN|TANH|TBILLEQ|TBILLPRICE|TBILLYIELD"
    parts.Add "TDIST|TEXT|TIME|TIMEVALUE|TINV|TODAY|TRANSPOSE|TREND|TRIM|TRIMMEAN|TRUE|TRUNC|TTEST|TYPE|UNICHAR|UNICODE|UPPER|USDOLLAR|VALUE|VAR|VAR.P"
    parts.Add "VAR.S|VARA|VARP|VARPA|VDB|VLOOKUP|WEBSERVICE|WEEKDAY|WEEKNUM|WEIBULL|WEIBULL.DIST|WORKDAY|WORKDAY.INTL|XIRR|XNPV|XOR|YEAR|YEARFRAC|YIELD|YIELDDISC|YIELDMAT"
    parts.Add "Z.TEST|ZTEST|VALUE_OF|DATE_RANGE|MEAN|GEOPCODE|GEOCONCAT|HFPCODE|T_ADM1|T_ADM2|T_ADM3|T_ADM4|ADM1_PCODE|ADM2_PCODE|ADM3_PCODE|ADM4_PCODE|HF_PCODE|ADM1_CONCAT|ADM2_CONCAT|ADM3_CONCAT|ADM4_CONCAT"
    parts.Add "HF_CONCAT|T_HF|HF_CONCAT|EPIWEEK"

    segments = TestHelpers.CollectionToArray(parts)
    FormulaFunctionsValues = Split(Join(segments, "|"), "|")
End Function

'@description Convert the list of function names into single-column rows.
Public Function FormulaFunctionsRows() As Variant
    FormulaFunctionsRows = TestHelpers.SingleColumnRows(FormulaFunctionsValues())
End Function

'@description Header row for the allowed special characters table.
Public Function FormulaCharactersHeaderRow() As Variant
    FormulaCharactersHeaderRow = Array("ASCII", "TEXT")
End Function

'@description Allowed special characters (from draft2.csv).
Public Function FormulaCharactersRows() As Variant
    Dim rows As Collection

    Set rows = New Collection

    rows.Add Array(33, "!")
    rows.Add Array(35, "#")
    rows.Add Array(36, "$")
    rows.Add Array(37, "%")
    rows.Add Array(38, "&")
    rows.Add Array(39, "'")
    rows.Add Array(40, "(")
    rows.Add Array(41, ")")
    rows.Add Array(42, "*")
    rows.Add Array(43, "+")
    rows.Add Array(44, ",")
    rows.Add Array(45, "-")
    rows.Add Array(47, "/")
    rows.Add Array(58, ":")
    rows.Add Array(59, ";")
    rows.Add Array(60, "<")
    rows.Add Array(61, "=")
    rows.Add Array(62, ">")
    rows.Add Array(63, "?")
    rows.Add Array(64, "@")
    rows.Add Array(91, "[")
    rows.Add Array(92, "\")
    rows.Add Array(93, "]")
    rows.Add Array(94, "^")
    rows.Add Array(96, "`")
    rows.Add Array(123, "{")
    rows.Add Array(124, "|")
    rows.Add Array(125, "}")
    rows.Add Array(126, "~")

    FormulaCharactersRows = TestHelpers.CollectionToArray(rows)
End Function

'@description Seed a worksheet with formula fixtures and return it.
'@param sheetName Worksheet name to create/reset.
'@param formulasTableName Optional ListObject name for the functions table.
'@param charactersTableName Optional ListObject name for the character table.
'@return Worksheet populated with both tables.
Public Function PrepareFormulaFixtureSheet(ByVal sheetName As String, _
                                           Optional ByVal formulasTableName As String = "T_XlsFonctions", _
                                           Optional ByVal charactersTableName As String = "T_ascii", _ 
                                           Optional ByVal outwb As Workbook) As Worksheet

    Dim sh As Worksheet
    Dim formulaHeader As Variant
    Dim formulaRows As Variant
    Dim characterHeader As Variant
    Dim characterRows As Variant
    Dim formulaHeaderMatrix As Variant
    Dim formulaRowsMatrix As Variant
    Dim characterHeaderMatrix As Variant
    Dim characterRowsMatrix As Variant
    Dim formulaRange As Range
    Dim characterRange As Range
    Dim totalFormulaRows As Long
    Dim totalCharacterRows As Long
    Dim formulaCols As Long
    Dim characterCols As Long
    Dim formulaTable As ListObject
    Dim characterTable As ListObject
    Dim formulaHeaderRows As Long
    Dim formulaDataRows As Long
    Dim characterHeaderRows As Long
    Dim characterDataRows As Long

    If outwb Is Nothing Then
        Set sh = TestHelpers.EnsureWorksheet(sheetName, visibility:=xlSheetHidden)
    Else
        Set sh = TestHelpers.EnsureWorksheet(sheetName, targetBook:=outwb)
    End If

    formulaHeader = FormulaFunctionsHeaderRow()
    formulaRows = FormulaFunctionsRows()
    formulaHeaderMatrix = TestHelpers.RowsToMatrix(Array(formulaHeader))
    formulaRowsMatrix = TestHelpers.RowsToMatrix(formulaRows)
    TestHelpers.WriteMatrix sh.Range("A1"), formulaHeaderMatrix
    TestHelpers.WriteMatrix sh.Range("A2"), formulaRowsMatrix

    formulaCols = UBound(formulaHeaderMatrix, 2) - LBound(formulaHeaderMatrix, 2) + 1
    formulaHeaderRows = UBound(formulaHeaderMatrix, 1) - LBound(formulaHeaderMatrix, 1) + 1
    If IsEmpty(formulaRowsMatrix) Then
        formulaDataRows = 0
    Else
        formulaDataRows = UBound(formulaRowsMatrix, 1) - LBound(formulaRowsMatrix, 1) + 1
    End If
    totalFormulaRows = formulaHeaderRows + formulaDataRows
    Set formulaRange = sh.Range("A1").Resize(totalFormulaRows, formulaCols)

    characterHeader = FormulaCharactersHeaderRow()
    characterRows = FormulaCharactersRows()
    characterHeaderMatrix = TestHelpers.RowsToMatrix(Array(characterHeader))
    characterRowsMatrix = TestHelpers.RowsToMatrix(characterRows)
    TestHelpers.WriteMatrix sh.Range("C1"), characterHeaderMatrix
    TestHelpers.WriteMatrix sh.Range("C2"), characterRowsMatrix

    characterCols = UBound(characterHeaderMatrix, 2) - LBound(characterHeaderMatrix, 2) + 1
    characterHeaderRows = UBound(characterHeaderMatrix, 1) - LBound(characterHeaderMatrix, 1) + 1
    If IsEmpty(characterRowsMatrix) Then
        characterDataRows = 0
    Else
        characterDataRows = UBound(characterRowsMatrix, 1) - LBound(characterRowsMatrix, 1) + 1
    End If
    totalCharacterRows = characterHeaderRows + characterDataRows
    Set characterRange = sh.Range("C1").Resize(totalCharacterRows, characterCols)

    Set formulaTable = sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=formulaRange, XlListObjectHasHeaders:=xlYes)
    formulaTable.Name = formulasTableName

    Set characterTable = sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=characterRange, XlListObjectHasHeaders:=xlYes)
    characterTable.Name = charactersTableName

    Set PrepareFormulaFixtureSheet = sh
End Function

