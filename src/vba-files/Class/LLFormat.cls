VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Format of the linelist")
'@IgnoreModule

Option Explicit
Implements ILLFormat
Private Const CLASSNAME As String = "LLFormat"

Private Type TLLFormat
    design As String
    sh As Worksheet
End Type

Private this As TLLFormat

'instanciation:

Public Function Create(formatSh As Worksheet, designType As String) As ILLFormat
    Dim Lo As ListObject
    Dim rng As Range
    Dim design As String

    If formatSh Is Nothing Then ThrowError "Format sheet does not exist, please set it."
    Set Lo = formatSh.ListObjects(1)
    If Lo Is Nothing Then ThrowError "No format table found in " & formatSh.Name & " worksheet. Do you have a table in this sheet?"
    design = designType

    'Test if design 1 or default is present
    Set rng = Lo.HeaderRowRange
    If (Not ValueExist("design 1", rng)) Then ThrowError "Not default (design 1) found in format sheet, can you seet a design 1 on table " & Lo.Name & "?"
    If Lo.Range.Rows.Count = 1 Then ThrowError "The table " & Lo.Name & "is empty"
    If (Not ValueExist("label", rng)) Then ThrowError "No label column found in the design table " & Lo.Name
    If (Not ValueExist(designType, rng)) Then
        'Trace or log here on the design (not choosen)
        design = "design 1"
    End If

    With New LLFormat
        Set .Wksh = formatSh
        .designName = design
        Set Create = .Self
    End With
End Function

Private Function ValueExist(ByVal Value As String, ByVal rng As Range) As Boolean
    ValueExist = (Not rng.Find(What:=Value, LookAt:=xlPart, MatchCase:=False) Is Nothing)
End Function

Public Property Get Self() As ILLFormat
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get designName() As String
    designName = this.design
End Property

Public Property Let designName(ByVal design As String)
    this.design = design
End Property

'Get the colors:

Private Function DesignValue(ByVal labelName As String, _
                             Optional ByVal returnColor As Boolean = True) As Variant

    Dim labelColumnRng As Range
    Dim headerRng As Range
    Dim desIndex As Long
    Dim labIndex As Long
    Dim design As String
    Dim sh As Worksheet


    Set sh = Wksh()
    design = designName
    Dim retdesValue As Variant

    With sh
        Set labelColumnRng = .ListObjects(1).ListColumns("label").Range
        Set headerRng = .ListObjects(1).HeaderRowRange
        desIndex = headerRng.Find(What:=design, LookAt:=xlPart, MatchCase:=False).Column
        If ValueExist(labelName, labelColumnRng) Then

            labIndex = labelColumnRng.Find(What:=labelName, LookAt:=xlPart, MatchCase:=False).Row

            If returnColor Then
                retdesValue = .Cells(labIndex, desIndex).Interior.color
            Else
                retdesValue = .Cells(labIndex, desIndex).Value
            End If
        End If
    End With

    If IsEmpty(retdesValue) Then
        DesignValue = IIf(returnColor, vbBlack, 0)
    Else
        DesignValue = retdesValue
    End If
End Function

'Format a range, add lines, or format borders

'Draw lines arround borders
Private Sub DrawLines(rng As Range, _
                      Optional At As String = "All", _
                      Optional weight As Integer = xlHairline, _
                      Optional line As Integer = xlContinuous, _
                      Optional color As Variant = vbBlack)
    Dim BorderPos As Byte
    If At = "All" Then
        With rng
            With .Borders
                .weight = weight
                .LineStyle = line
                .color = color
                .TintAndShade = 0.4
            End With
        End With
    Else
        Select Case At
        Case "Left"
            BorderPos = xlEdgeLeft
        Case "Right"
            BorderPos = xlEdgeRight
        Case "Bottom"
            BorderPos = xlEdgeBottom
        Case "Top"
            BorderPos = xlEdgeTop
        Case Else
            BorderPos = xlEdgeBottom
        End Select
        With rng
            With .Borders(BorderPos)
                .weight = weight
                .LineStyle = line
                .color = color
                .TintAndShade = 0.4
            End With
        End With
    End If
End Sub

Private Sub DrawBorder(rng As Range, Optional weight As Integer = xlThin, Optional color As Variant = vbBlack, Optional line As Long = xlContinuous)

    Dim counter As Integer
    For counter = 7 To 10                        'xltop, left, right and bottom
        With rng.Borders(counter)
            .LineStyle = line
            .color = color
            .TintAndShade = 0.4
            .weight = weight
        End With
    Next
End Sub

Private Sub FormatRange(ByVal rng As Range, _
                        Optional ByVal fillValue As String = "", _
                        Optional ByVal interiorColor As Variant = "", _
                        Optional ByVal fontColor As Variant = "", _
                        Optional ByVal isBold As Boolean = False, _
                        Optional ByVal Horiz As Integer = xlHAlignCenter, _
                        Optional ByVal Verti As Integer = xlVAlignCenter, _
                        Optional ByVal FontSize As Double = 0, _
                        Optional ByVal NumFormat As String = "")

    With rng
        If interiorColor <> vbNullString Then .Interior.color = interiorColor
        If fontColor <> vbNullString Then .Font.color = fontColor
        .Font.Bold = isBold
        If FontSize <> 0 Then .Font.Size = FontSize
        If NumFormat <> vbNullString Then .NumberFormat = NumFormat
        .HorizontalAlignment = Horiz
        .VerticalAlignment = Verti
        If fillValue <> vbNullString Then .Value = fillValue
    End With
End Sub

Private Sub FormatCharacters(ByVal char As Characters, _
                             Optional ByVal fontColor As Variant = "", _
                             Optional ByVal isBold As Boolean = False, _
                             Optional ByVal FontSize As Double = 0)

    With char
        If fontColor <> vbNullString Then .Font.color = fontColor
        .Font.Bold = isBold
        If FontSize <> 0 Then .Font.Size = FontSize
    End With
End Sub

Private Sub FormatShape(ByVal shp As WorksheetShape, _
                        Optional ByVal fontColor As Variant = vbNullString, _
                        Optional ByVal interiorColor As Variant = vbNullString, _
                        Optional ByVal isBold As Boolean = False, _
                        Optional ByVal FontSize As Double = 0)
    With shp
        'Fill color
        If interiorColor <> vbNullString Then
            .Fill.ForeColor = interiorColor
            .Fill.BackColor = interiorColor
        End If

        With .TextFrame2
            'Text alignment
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoTrue

            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter

                'Font Zize and Font Shape
                If FontSize <> 0 Then .Font.Size = FontSize
                .Font.Bold = isBold
                'Text Color
                If fontColor <> vbNullString Then .Font.Fill.ForeColor = fontColor
            End With
        End With
    End With
End Sub

Private Sub ApplyFormat(ByVal formattingObject As Variant, ByVal scope As FormatScopes)

    Dim rng As Range
    Dim char As Characters
    Dim sh As Worksheet
    Dim shp As WorksheetShape

    Dim fontColorValue As Variant
    Dim interiorColorValue As Variant
    Dim bordersColorValue As Variant
    Dim styleValue As Variant
    Dim dimensionValue As Variant
    Dim verticalAlign As Long                    'Vertical Alignment
    Dim horizAlign As Long

    'Base analysis font size
    styleValue = DesignValue("analysis base font size", False)
    If (Not TypeName(styleValue) = "Double") Then styleValue = 11

    'Formatting on analysis  part:
    Select Case scope
    Case AllAnalysisSheet
        If (Not TypeName(sh) = "Worksheet") Then Exit Sub
        Set sh = formattingObject
        sh.Cells.Font = styleValue
        dimensionValue = DesignValue("default analysis column width", False)
        On Error Resume Next
        sh.Cells.EntireColumn.ColumnWidth = 22
        sh.Cells.EntireColumn.ColumnWidth = dimensionValue
        On Error GoTo 0

    Case AnalysisSection

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("table sections font color")
        dimensionValue = DesignValue("default analysis column width", False)

        Set sh = rng.Parent
        With sh
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 7))
        End With
        On Error Resume Next
        rng.Merge
        On Error GoTo 0
        'Section color
        FormatRange rng, FontSize:=styleValue + 5, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, Verti:=xlVAlignBottom, _
                    isBold:=True
        rng.WrapText = True
        'Draw a border arround the section
        DrawLines rng:=rng, weight:=xlMedium, color:=fontColorValue, At:="Bottom"

        'Gap after section
        On Error Resume Next
        rng.Offset(1).EntireRow.RowHeight = 45
        rng.Offset(1).EntireRow.RowHeight = dimensionValue
        On Error GoTo 0

    Case AnalysisTableTitle
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        dimensionValue = DesignValue("gap for graphs", False)
        fontColorValue = DesignValue("table title font color")


        FormatRange rng, FontSize:=styleValue, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, _
                    Verti:=xlVAlignBottom
        rng.WrapText = True
        rng.Font.Italic = True

    Case LinelistSelectionDropdown
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("dropdown font color")
        interiorColorValue = DesignValue("dropdown interior color")
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    isBold:=False, FontSize:=styleValue - 2
        rng.WrapText = True

    Case AnalysisCategoriesLabelRow, AnalysisCategoriesLabelColumn, AnalysisUnivariateHeader
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject

        verticalAlign = IIf(scope = AnalysisCategoriesLabelRow, xlVAlignCenter, xlVAlignBottom)
        horizAlign = IIf(scope = AnalysisUnivariateHeader, xlHAlignCenter, xlHAlignLeft)

        fontColorValue = DesignValue("categories names font color")
        FormatRange rng, fontColor:=fontColorValue, isBold:=True, FontSize:=styleValue - 1, _
                    Horiz:=horizAlign, Verti:=verticalAlign

    Case AnalysisCategoriesUnivariateRow, AnalysisFirstRow, AnalysisCategoriesBivariateRow, AnalysisCategoriesBivariateColumn

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("table categories font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("table categories interior color")

        horizAlign = IIf(scope = AnalysisCategoriesBivariateColumn, xlHAlignCenter, xlHAlignLeft)

        'Format the categories
        If scope <> AnalysisFirstRow Then
            FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue, _
                        Horiz:=horizAlign, Verti:=xlVAlignCenter, interiorColor:=interiorColorValue
            DrawLines rng, color:=bordersColorValue
        Else
            FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue - 2, _
                        Horiz:=xlHAlignCenter, Verti:=xlVAlignCenter
            DrawLines rng, color:=bordersColorValue
        End If

        If scope = AnalysisCategoriesBivariateColumn Or scope = AnalysisFirstRow Then
            DrawLines rng, Ate² = "Left", color:=bordersColorValue, line:=xlDouble
            DrawLines rng, At:="Right", color:=bordersColorValue, line:=xlDouble
        Else
            DrawBorder rng, color:=bordersColorValue
        End If

        'Borders at top and bottom for bivariate
        If scope = AnalysisCategoriesBivariateRow Then
            DrawLines rng, line:=xlDouble, At:="Top", color:=bordersColorValue
            DrawLines rng, line:=xlDouble, At:="Bottom", color:=bordersColorValue
        End If

    Case AnalysisGlobalSummaryRow
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("table categories font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("table categories interior color")

        Set sh = rng.Parent
        FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue, _
                    Horiz:=xlHAlignLeft, Verti:=xlVAlignCenter, interiorColor:=interiorColorValue
        'All the range line
        With sh
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 3))
            DrawLines rng, color:=bordersColorValue, weight:=xlThin
            DrawLines rng, color:=bordersColorValue, At:="Top"
            DrawLines rng, color:=bordersColorValue, At:="Bottom"
            Set rng = .Range(rng.Cells(1, 2), rng.Cells(1, 3))
            FormatRange rng, FontSize:=styleValue - 2, Horiz:=xlHAlignRight
        End With

    Case AnalysisInteriorTable, AnalysisWholeTable
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        bordersColorValue = DesignValue("analysis table borders color")

        If scope = AnalysisInteriorTable Then
            DrawLines rng, color:=bordersColorValue
            DrawLines rng, color:=bordersColorValue, At:="Top", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Bottom", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Right", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Left", line:=xlDouble
        End If

        If scope = AnalysisWholeTable Then DrawBorder rng, color:=bordersColorValue


    Case AnalysisTimeSeriesHeader
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("time series header font")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("time series header interior")

        On Error Resume Next
        rng.Merge
        On Error GoTo 0
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue + 1, Horiz:=xlHAlignCenter, Verti:=xlVAlignCenter, isBold:=True
        DrawLines rng, At:="Top", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Left", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Bottom", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Right", color:=bordersColorValue, line:=xlDouble

    Case AnalysisMissingRow, AnalysisMissingColumn

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("missing font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("missing interior color")

        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue, isBold:=True, Horiz:=xlHAlignRight

        'Draw lines
        DrawLines rng, color:=bordersColorValue
        DrawLines rng, color:=fontColorValue, At:="Top", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Left", line:=xlDouble

        If scope = AnalysisMissingRow Then
            Set rng = rng.Cells(1, 0)
            ApplyFormat rng, scope:=AnalysisOneCellLeft
        End If

    Case AnalysisOneCell, AnalysisOneCellLeft, AnalysisOneCellCenter

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("missing font color")
        interiorColorValue = DesignValue("missing interior color")
        horizAlign = Switch(scope = AnalysisOneCellLeft, xlHAlignLeft, _
                            scope = AnalysisOneCellCenter, xlHAlignCenter, _
                            True, xlHAlignRight)

        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue, isBold:=True, Horiz:=horizAlign

        DrawLines rng, color:=fontColorValue, At:="Bottom", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Top", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Left", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Right", line:=xlDouble

    Case AnalysisPercent

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        FormatRange rng, NumFormat:="0.00%"

    Case LinelistHiddenCell

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        Set rng = formattingObject
        rng.Interior.color = vbWhite
        rng.Font.color = vbWhite
        rng.Locked = True
        rng.FormulaHidden = True

    Case AnalysisEndTable

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        Set rng = formattingObject
        dimensionValue = DesignValue("gap for graphs", False)

        On Error Resume Next

        'Gap for graphs
        rng.EntireRow.RowHeight = 50
        rng.Offset(1).EntireRow.RowHeight = 50

        rng.EntireRow.RowHeight = dimensionValue
        rng.Offset(1).EntireRow.RowHeight = dimensionValue

        On Error GoTo 0

    Case AnalysisInfoCell

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = IIf(IsEmpty(rng), "", DesignValue("missing font color"))
        FormatRange rng, fontColor:=fontColorValue, _
                    FontSize:=8, isBold:=False, Horiz:=xlHAlignCenter

    Case HListMainLab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set char = formattingObject
        fontColorValue = DesignValue("hlist main label font color")
        dimensionValue = DesignValue("hlist main label font size", False)
        FormatCharacters char:=char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=True

    Case HListSublab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set char = formattingObject
        fontColorValue = DesignValue("hlist sub label font color")
        dimensionValue = DesignValue("hlist sub label font size", False)

        FormatCharacters char:=char, FontSize:=dimensionValue, fontColor:=fontColorValue

    Case HListCalculatedFormulaCell
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("calculated formula font color")
        interiorColorValue = DesignValue("calculated formula column interior color")
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue

    Case HListCalculatedFormulaHeader
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        interiorColorValue = DesignValue("calculated formula header color")
        FormatRange rng, interiorColor:=interiorColorValue

    Case HListGeoHeader, HListGeo
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        interiorColorValue = IIf(scope = HListGeoHeader, DesignValue("geo interior color"), "")
        fontColorValue = DesignValue("geo font color")
        FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue

    Case VListMainLab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set char = formattingObject
        fontColorValue = DesignValue("vlist main label font color")
        dimensionValue = DesignValue("vlist main label font size", False)
        FormatCharacters char:=char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=True

    Case VListSublab

        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set char = formattingObject
        fontColorValue = DesignValue("vlist sub label font color")
        dimensionValue = DesignValue("vlist sub label font size", False)
        FormatCharacters char:=char, FontSize:=dimensionValue, fontColor:=fontColorValue

    Case LinelistSmallButton, AdminLargeButton

        If (Not TypeName(formattingObject) = "WorksheetShape") Then Exit Sub
        Set shp = formatingObject

        fontColorValue = DesignValue("button default font color")
        interiorColorValue = DesignValue("button default interior color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue)
        dimensionValue = IIf(scope = AdminLargeButton, dimensionValue + 3, dimensionValue)

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue

    Case LinelistGeoButton

        If (Not TypeName(formattingObject) = "WorksheetShape") Then Exit Sub
        Set shp = formattingObject

        fontColorValue = DesignValue("geo interior color")
        interiorColorValue = DesignValue("geo font color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue)

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue

    Case HListSubSection, VListSubSection

      If (Not TypeName(formattingObject) = "Range") Then Exit Sub
      Set rng = formattingObject

      fontColorValue = DesignValue("sub section interior color")
      interiorColorValue = DesignValue("sub section font color")
      dimensionValue = DesignValue("sub section font size", False)
      dimensionValue = IIf(dimensionValue = 0, 11, dimensionValue)

      On Error Resume Next
        rng.Merge
      On Error GoTo 0

      FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue, _
      isBold:=True

    Case HListSection, VListSection

      If (Not TypeName(formattingObject) = "Range") Then Exit Sub
      Set rng = formattingObject

      fontColorValue = DesignValue("section interior color")
      interiorColorValue = DesignValue("section font color")
      dimensionValue = DesignValue("section font size", False)
      dimensionValue = IIf(dimensionValue = 0, 12, dimensionValue)

      On Error Resume Next
        rng.Merge
      On Error GoTo 0
     
      FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue, _
      isBold:=True

    End Select

End Sub

Private Function RangeExists(ByVal sh As Worksheet, rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = Not (rng Is Nothing)
End Function

'Throw Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

'Exposed methods---------------------------------------------------------------------
Public Sub ILLFormat_ApplyFormat(ByVal formattingObject As Variant, ByVal scope As FormatScopes)
    ApplyFormat formattingObject:=formattingObject, scope:=scope
End Sub

Public Function ILLFormat_DesignValue(ByVal labelName As String, Optional ByVal returnColor As Boolean = True) As Variant
    ILLFormat_DesignValue = DesignValue(labelName, returnColor)
End Function


