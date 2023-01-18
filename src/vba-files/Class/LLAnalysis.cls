VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis Class"


'@Folder("Analysis")
'@ModuleDescription("Analysis Class")
'@IgnoreModule

Option Explicit
Implements ILLAnalysis

Private Type TLLAnalysis
    anash As Worksheet
    lData As ILinelistSpecs
End Type

Private this As TLLAnalysis
Private Const CLASSNAME As String = "LLAnalysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"

'Exposed methods
Public Function Create(ByVal anaWksh As Worksheet) As ILLAnalysis

    If anaWksh Is Nothing Then ThrowErrorUnableInstanciate "Worksheet object not defined"
    CheckTablesRequirements

    With New LLAnalysis
        Set .Wksh = anaWksh
        Set Create = .Self
    End With

End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
    Set this.anash = anash
End Property

Public Property Get Self() As ILLAnalysis
    Set Self = Me
End Property

Private Sub CheckTablesRequirements()
    Dim loTable As BetterArray
    Dim counter As Long
    Dim Lo As ListObject
    Dim loName As String
    
    Set loTable = New BetterArray

    loTable.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, GRAPHTIMESERIES, SPATIAL

    For counter = loTable.LowerBound To loTable.UpperBound
        loName = loTable.Items(counter)
        On Error Resume Next
        Set Lo = Wksh.ListObjects(loName)
        If Lo Is Nothing Then ThrowErrorUnableInstanciate " Listobject " & loName & " does not exist in analysis worksheet "
        On Error GoTo 0
    Next
End Sub

'Find if a column Exists in a listObject of the analysis
Private Function ColumnExists(ByVal Lo As ListObject, ByVal colName As String) As Boolean
    Dim hRng As Range
    Set hRng = Lo.HeaderRowRange

    ColumnExists = (Not hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal Lo As ListObject, ByVal colName) As Long
    Dim index As Long
    Dim hRng As Range

    Set hRng = Lo.HeaderRowRange
    If ColumnExists(Lo, colName) Then
        index = hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False).Column - hRng.Column + 1
    Else
        index = -1
    End If
    ColumnIndex = index
End Function

Private Function DataRange(ByVal Lo As ListObject, ByVal colName As String) As Range
    If ColumnExists(Lo, colName) Then Set DataRange = Lo.ListColumns(ColumnIndex(Lo, colName)).Range
End Function

Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub TransferCodeWksh(ByVal sh As Worksheet, ll As ILinelist)
    
    Const MODULENAME As String = "LinelistAnalysisChange"
                                  
    Dim inpWkb As Workbook                       'Input workbook
    Dim lData As ILinelistSpecs
    Dim codeContent As String                    'a string to contain code to add
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sheetName As String

    Set lData = ll.linelistData()
    Set inpWkb = lData.specswkb()
    sheetName = sh.Name
    

    'save the code module in the string sNouvCode
    With inpWkb.VBProject.VBComponents(MODULENAME).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With sh.Parent
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.sheets(sheetName).codeName)
        Set codeMod = vbComp.CodeModule
    End With

    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
    
End Sub

Private Sub Build(ll As ILinelist)
    
    Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
    Const SHPFILTERNAME As String = "SHP_Filter"
    Const GOTOSECCODE As String = "go_to_section" 'Goto section code
    Const UPDATESPCOMMAND As String = "UpdateSpTables" 'Spatial tables update
    
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim counter As Long
    Dim graphCounter As Long
    Dim Lo As ListObject
    Dim hRng As Range
    Dim tabRng As Range
    Dim cellRng As Range
    Dim specs As ITablesSpecs
    Dim prevSpecs As ITablesSpecs
    Dim grSpecs As IGraphSpecs
    Dim gr As IGraphs
    Dim tabId As String
    Dim lData As ILinelistSpecs
    Dim tabl As ICrossTable
    Dim tabform As ICrossTableFormula
    Dim design As ILLFormat
    Dim tradmsg As ITranslation
    Dim tradshp As ITranslation
    Dim forml As IFormulaData
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim pass As ILLPasswords
    Dim maxFormula As String
    Dim minFormula As String
    Dim rowVar As String
    Dim secId As String
    Dim Valid As Boolean
    Dim but As IButtons
    Dim sectionsList As BetterArray              'Dropdown for sections
    Dim drop As IDropdownLists
    Dim goToSection As String
    Dim grLo As ListObject                       'ListObject of graph for time series
    Dim grLoTitle As ListObject                   'Title of graphs of time series
    Dim grBAinput As BetterArray                 'BetterArray input (list of listobjects) for time series f
    Dim grBAoutput As BetterArray                'BetterArray output (list of graphs) for time series graphs
    Dim temptabl As BetterArray                  'Temporary table to stock specs of graphs for time series
    Dim lastValid As ITablesSpecs
    
    'Those six tables are temporary tables for time series graphs
    Dim tablSeries As BetterArray
    Dim tablType As BetterArray
    Dim tablPos As BetterArray
    Dim tablLabel As BetterArray
    Dim tablColLabel As BetterArray
    Dim tablPrefLabel As BetterArray


    Set sh = Wksh()
    Set lData = ll.linelistData()
    Set tradmsg = lData.TransObject(TranslationOfMessages)
    Set tradshp = lData.TransObject(TranslationOfShapes)
    Set design = lData.DesignFormat
    Set forml = lData.FormulaDataObject
    Set dict = lData.Dictionary()
    Set vars = LLVariables.Create(dict)
    Set sectionsList = New BetterArray
    Set drop = ll.Dropdown()
    Set pass = lData.Password()
    
    
    Set outsh = ll.Wksh(tradmsg.TranslatedValue("LLSHEET_Analysis"))
    BusyApp
    goToSection = tradmsg.TranslatedValue("MSG_GoToSec")
    design.ApplyFormat outsh, AllAnalysisSheet
    
    'Add the update filter button
    Set cellRng = outsh.Cells(2, 1)
    Set but = Buttons.Create(cellRng, tradshp, SHPFILTERNAME, ButtonScopeSmall)
    but.Add UPDATEFILTERCOMMAND
    but.Format design
    
    'Add The Sheet Type
    Set cellRng = outsh.Cells(1, 3)
    cellRng.Value = "Uni-Bi-Analysis"
    design.ApplyFormat cellRng, scope:=LinelistHiddenCell
    
    On Error Resume Next
    'Activating the worksheet restores all the application
    'status to true, remember to recall the busy app just
    'after
    BusyApp
    outsh.Activate
    With Application.ActiveWindow
        .SplitRow = 2
        .FreezePanes = True
    End With
    On Error GoTo 0
    
    
    Set Lo = sh.ListObjects(GLOBALSUMMARY)
    If Not (Lo.DataBodyRange Is Nothing) Then
        Set hRng = Lo.HeaderRowRange
        For counter = 1 To Lo.ListRows.Count
            Set tabRng = Lo.ListRows(counter).Range
            Set specs = TablesSpecs.Create(hRng, tabRng, lData)
            If specs.ValidTable() Then
                'Create the table
                Set tabl = CrossTable.Create(specs, outsh, tradmsg)
                tabl.Build
                'Add formulas and format
                Set tabform = CrossTableFormula.Create(tabl, forml)
                tabform.AddFormulas
                tabl.Format design
            End If
        Next
    End If
    
    
    'Build Univariate Analysis =====================================================
    
    Set Lo = sh.ListObjects(UNIVARIATE)
    If Not (Lo.DataBodyRange Is Nothing) Then
        Set hRng = Lo.HeaderRowRange
        For counter = 1 To Lo.ListRows.Count
            Set tabRng = Lo.ListRows(counter).Range
            Set specs = TablesSpecs.Create(hRng, tabRng, lData)
             
            If specs.ValidTable() Then
                If specs.isNewSection Then sectionsList.Push goToSection & ": " & specs.Value("section")
                'Create the table
                Set tabl = CrossTable.Create(specs, outsh, tradmsg)
                tabl.Build
                'Add formulas and format
                Set tabform = CrossTableFormula.Create(tabl, forml)
                tabform.AddFormulas
                tabl.Format design
                'Add graph
                If specs.Value("graph") = "yes" Then
                    Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
                    Set grSpecs = GraphSpecs.Create(tabl)
                    grSpecs.CreateSeries
                    Set gr = Graphs.Create(outsh, cellRng)
                    gr.Add
                    For graphCounter = 1 To grSpecs.NumberOfSeries 'Number of series for univariate and bivariate
                        gr.AddSeries grSpecs.SeriesName(graphCounter), grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
                        gr.AddLabels grSpecs.SeriesLabel(graphCounter), grSpecs.SeriesColumnLabel(graphCounter)
                    Next
                    gr.Format catTitle:=vars.Value(colName:="main label", varName:=specs.Value("row")), valuesTitle:=specs.Value("label")
                End If
            End If
        Next
    End If
     
     
    'Build Bivariate Analysis ===================================================== 'Can fator this later on
    
    Set Lo = sh.ListObjects(BIVARIATE)
    If Not (Lo.DataBodyRange Is Nothing) Then
        Set hRng = Lo.HeaderRowRange
        For counter = 1 To Lo.ListRows.Count
            Set tabRng = Lo.ListRows(counter).Range
            Set specs = TablesSpecs.Create(hRng, tabRng, lData)
            If specs.ValidTable() Then
                If specs.isNewSection Then sectionsList.Push goToSection & ": " & specs.Value("section")
                tabId = specs.TableId
                'Create the table
                Set tabl = CrossTable.Create(specs, outsh, tradmsg)
                tabl.Build
                'Add formulas and format
                Set tabform = CrossTableFormula.Create(tabl, forml)
                tabform.AddFormulas
                tabl.Format design
               
                'Add graph
                If specs.Value("graph") = "values" Or specs.Value("graph") = "percentage" Or specs.Value("graph") = "both" Then
                    Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
                    Set grSpecs = GraphSpecs.Create(tabl)
                    grSpecs.CreateSeries
                    Set gr = Graphs.Create(outsh, cellRng)
                    gr.Add
                    For graphCounter = 1 To grSpecs.NumberOfSeries
                        gr.AddSeries grSpecs.SeriesName(graphCounter), grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
                        gr.AddLabels grSpecs.SeriesLabel(graphCounter), grSpecs.SeriesColumnLabel(graphCounter)
                    Next
                    gr.Format catTitle:=vars.Value(colName:="main label", varName:=specs.Value("row")), valuesTitle:=specs.Value("label")
                End If
            End If
        Next
    End If
     
    'Add GoTo Section dropdown
    drop.Add sectionsList, "uagotosection"
    outsh.Cells(2, 2).Value = goToSection
    drop.SetValidation outsh.Cells(2, 2), "uagotosection"
    design.ApplyFormat outsh.Cells(2, 2), scope:=LinelistSelectionDropdown
    outsh.Cells(2, 2).Locked = False
    outsh.Cells(2, 2).Name = "ua_" & GOTOSECCODE
    
    'Transfert event change to the univariate analysis worksheet
    TransferCodeWksh outsh, ll
    pass.Protect "_active", True
     
    'Build Time series Analysis =====================================================
     
    Set Lo = sh.ListObjects(TIMESERIES)
    Set outsh = ll.Wksh(tradmsg.TranslatedValue("LLSHEET_TemporalAnalysis"))
    design.ApplyFormat outsh, AllAnalysisSheet
     
    'Add the update filter button
    Set cellRng = outsh.Cells(2, 1)
    Set but = Buttons.Create(cellRng, tradshp, SHPFILTERNAME, ButtonScopeSmall)
    but.Add UPDATEFILTERCOMMAND
    but.Format design
    
    'Add the sheet type
    Set cellRng = outsh.Cells(1, 3)
    cellRng.Value = "TS-Analysis"
    design.ApplyFormat cellRng, LinelistHiddenCell
    
    On Error Resume Next
    BusyApp
    outsh.Activate
    With Application.ActiveWindow
        .SplitRow = 2
        .SplitColumn = 3
        .FreezePanes = True
        .Zoom = 65
    End With
    On Error GoTo 0
    BusyApp
     
    sectionsList.Clear
    
    If Not (Lo.DataBodyRange Is Nothing) Then
        Set hRng = Lo.HeaderRowRange
         
        For counter = 1 To Lo.ListRows.Count
             
            Set tabRng = Lo.ListRows(counter).Range
            Set specs = TablesSpecs.Create(hRng, tabRng, lData)
             
            Valid = specs.ValidTable()
             
            If Valid Then
                 
                'Minimum and maximum formulas
                rowVar = specs.Value("row")
                 
                If specs.isNewSection() Then
                    sectionsList.Push goToSection & ": " & specs.Value("section")
                    Set lastValid = specs
                End If
                 
                On Error Resume Next             'Could generate an error if the table doesn't have a previous table
                Set prevSpecs = specs.Previous()
                On Error GoTo 0

                Set tabl = CrossTable.Create(specs, outsh, tradmsg)
                tabl.Build
                Set tabform = CrossTableFormula.Create(tabl, forml)
                tabform.AddFormulas
                tabl.Format design
                 
                'Minimum and maximum formulas for time series
                If (prevSpecs Is Nothing) Then
                    ' First New section
                    minFormula = "MIN(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"
                    maxFormula = "MAX(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"

                ElseIf specs.isNewSection() And (Not prevSpecs Is Nothing) Then
                    
                    'New section after the first new section
                    outsh.Range("MAX_MAX_DATE_" & prevSpecs.TableSectionId).formula = "= MAX(" & maxFormula & ")"
                    outsh.Range("MIN_MIN_DATE_" & prevSpecs.TableSectionId).formula = "= MIN(" & minFormula & ")"
                    minFormula = "MIN(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"
                    maxFormula = "MAX(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"

                ElseIf Not specs.isNewSection() Then

                    'Another part of the table, on the same section
                    minFormula = minFormula & " , " & "MIN(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"
                    maxFormula = maxFormula & ", " & "MAX(" & vars.Value(colName:="table name", varName:=rowVar) & "[" & rowVar & "]" & ")"

                End If
            End If
        Next
                    
        'Add formulas to the last section
        If Not lastValid Is Nothing Then
            outsh.Range("MAX_MAX_DATE_" & lastValid.TableSectionId).formula = "= MAX(" & maxFormula & ")"
            outsh.Range("MIN_MIN_DATE_" & lastValid.TableSectionId).formula = "= MIN(" & minFormula & ")"
            Set cellRng = outsh.Range("ENDTABLE_" & lastValid.TableId).Offset(5, 2)
        End If
        
        'Add the graphs for time series analysis
        Set grLo = sh.ListObjects(GRAPHTIMESERIES)
        Set grLoTitle = sh.ListObjects(TITLEGRAPHTIMESERIES)
        
        If (Not grLo.DataBodyRange Is Nothing) And (Not grLoTitle.DataBodyRange Is Nothing) Then
        
            Set grBAinput = New BetterArray
            grBAinput.Push grLo, Lo, grLoTitle             'The order of the listobjects in the betterarray is important, first graph, second the table
            Set grSpecs = GraphSpecs.CreateRangeSpecs(grBAinput, outsh, lData)
            
            'Test the validity of the graphs (No empty graph ids) before building thems
            If grSpecs.Valid Then
                Set grBAoutput = grSpecs.SpecsLists()
                
                For graphCounter = 1 To grSpecs.NumberOfGraphs
                    Set temptabl = grBAoutput.Item(graphCounter) 'Temporary data containing specifications for one graph
                    'temptabl is a table of five tables

                    Set tablSeries = temptabl.Item(1) 'List of series for one graph
                    Set tablType = temptabl.Item(2) 'List of types for one graph (line, bar, point)
                    Set tablPos = temptabl.Item(3) 'List of the positions for one graph

                    'Labels and column labels
                    Set tablLabel = temptabl.Item(4)
                    Set tablColLabel = temptabl.Item(5)
                    Set tablPrefLabel = temptabl.Item(6)
    
                    Set gr = Graphs.Create(outsh, cellRng)
                    gr.Add
    
                    For counter = 1 To tablSeries.Length
                        gr.AddSeries tablSeries.Item(counter), tablType.Item(counter), tablPos.Item(counter)
                        gr.AddLabels tablLabel.Item(counter), tablColLabel.Item(counter), tablPrefLabel.Item(counter), prefixOnly := True
                    Next
                    
                    gr.Format catTitle := tradmsg.TranslatedValue("MSG_Period"), valuesTitle := tradmsg.TranslatedValue("MSG_Values"), _ 
                              plotTitle := temptabl.Item(7),  asTimeSeries:=True 'The last item of temptabl is the title
    
                    Set cellRng = cellRng.Offset(50)
                Next
            End If
            
        End If
    End If
     
    drop.Add sectionsList, "tsgotosection"
    drop.SetValidation outsh.Cells(2, 3), "tsgotosection"
    design.ApplyFormat outsh.Cells(2, 3), scope:=LinelistSelectionDropdown
    outsh.Cells(2, 3).Name = "ts_" & GOTOSECCODE
    outsh.Cells(2, 3).Value = goToSection
    outsh.Cells(2, 3).Locked = False

    'Hide the second column (column B)
    outsh.Columns(2).Hidden = True
     
    'Send the LinelistTimeSeries change
    TransferCodeWksh outsh, ll
    pass.Protect "_active", True
     
     
    'Building the spatial Analysis ===================================================================================
     
    ' Set outsh = ll.Wksh(tradmsg.TranslatedValue("LLSHEET_SpatialAnalysis"))
    ' design.ApplyFormat outsh, AllAnalysisSheet
    ' BusyApp
    ' sectionsList.Clear
     
    ' 'Add the update filter button
    ' Set cellRng = outsh.Cells(2, 1)
    ' Set but = Buttons.Create(cellRng, tradshp, SHPFILTERNAME, ButtonScopeSmall)
    ' but.Add UPDATESPCOMMAND
    ' but.Format design
    
    ' 'Add The Sheet Type
    ' Set cellRng = outsh.Cells(1, 3)
    ' cellRng.Value = "SP-Analysis"
    ' design.ApplyFormat cellRng, scope:=LinelistHiddenCell
    
    ' On Error Resume Next
    ' 'Activating the worksheet restores all the application
    ' 'status to true, remember to recall the busy app just
    ' 'after
    ' outsh.Activate
    ' BusyApp
    ' With ActiveWindow
    '     .SplitRow = 2
    '     .FreezePanes = True
    ' End With
    ' On Error GoTo 0
     
    ' 'Same as Bivariate analysis
    ' Set Lo = sh.ListObjects(SPATIAL)
    ' If Not (Lo.DataBodyRange Is Nothing) Then
    '     Set hRng = Lo.HeaderRowRange
    '     For counter = 1 To Lo.ListRows.Count
    '         Set tabRng = Lo.ListRows(counter).Range
    '         Set specs = TablesSpecs.Create(hRng, tabRng, lData)
             
    '         If specs.ValidTable() Then
    '             If specs.isNewSection Then sectionsList.Push goToSection & ": " & specs.Value("section")
    '             tabId = specs.TableId
    '             'Create the table
    '             Set tabl = CrossTable.Create(specs, outsh, tradmsg)
    '             tabl.Build
                 
    '             'Add formulas and format
    '             Set tabform = CrossTableFormula.Create(tabl, forml)
    '             tabform.AddFormulas
    '             tabl.Format design
                
               
    '             'Add graph
    '             If specs.Value("graph") = "values" Or specs.Value("graph") = "percentage" Or specs.Value("graph") = "both" Then
    '                 Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
    '                 Set grSpecs = GraphSpecs.Create(tabl)
    '                 grSpecs.CreateSeries
    '                 Set gr = Graphs.Create(outsh, cellRng)
    '                 gr.Add
    '                 For graphCounter = 1 To grSpecs.NumberOfSeries
    '                     gr.AddSeries grSpecs.SeriesName(graphCounter), grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
    '                     gr.AddLabels grSpecs.SeriesLabel(graphCounter), grSpecs.SeriesColumnLabel(graphCounter)
    '                 Next

    '                 gr.Format catTitle:=vars.Value(colName:="main label", varName:=specs.Value("row")), valuesTitle:=specs.Value("label")
    '             End If
                
    '         End If
    '     Next
    ' End If

    ' 'Add the goto section dropdown

    ' drop.Add sectionsList, "spgotosection"
    ' drop.SetValidation outsh.Cells(2, 2), "spgotosection"
    ' design.ApplyFormat outsh.Cells(2, 2), scope:=LinelistSelectionDropdown
    ' outsh.Cells(2, 2).Name = "sp_" & GOTOSECCODE
    ' outsh.Cells(2, 2).Value = goToSection

End Sub

'Import listobjects in analysis
Private Sub Import(ByVal impWksh As Worksheet)

    Dim anash As Worksheet
    Dim impRng As Range
    Dim actRng As Range
    Dim impAddress As String
    Dim impLo As ListObject

    Set anash = Wksh()

    'Test if the sheet is empty
    If impWksh Is Nothing Then Exit Sub          'Add trace for import

    anash.Cells.Clear

    On Error GoTo LoFail
    Set impLo = impWksh.ListObjects(1)
    On Error GoTo 0


    For Each impLo In impWksh.ListObjects

        Set impRng = impLo.Range
        impAddress = impRng.Address
        Set actRng = anash.Range(impAddress)
        actRng.Value = impRng.Value

        actRng.Cells(-1, 1).Value = impRng.Cells(-1, 1).Value
        actRng.Cells(-1, 1).Font.Size = impRng.Cells(-1, 1).Font.Size
        actRng.Cells(-1, 1).Font.color = impRng.Cells(-1, 1).Font.color
        actRng.Cells(-1, 1).Font.Bold = True

        'Define the listobject
        With anash
            .ListObjects.Add(xlSrcRange, actRng, , xlYes).Name = impLo.Name
            .ListObjects(impLo.Name).TableStyle = impLo.TableStyle
        End With
    Next


    'Be sure all the tables exist after import
    CheckTablesRequirements

    anash.Cells(1, 1).Value = "analysis"
    anash.Cells(1, 1).Font.color = vbBlue

    'Sort the analysis worksheet on each one of the tables
    SortAnalysis

    Exit Sub
LoFail:
    'add trace for no list objects
End Sub


Private Sub SortListOnFirst(ByVal Lo As ListObject, ByVal colIndex As Long)

    Dim numRng As Range 'Number range
    Dim colRng As Range 'Colum range to sort on
    Dim LoRng As Range 'ListObject Range
    Dim hRng As Range 'Header Range of the list Object
    Dim resRng As Range 'Will resize on this range
    Dim counter As Long
    Dim trackData As BetterArray
    Dim sh As Worksheet
    Dim sortedValue As String
    Dim num As Long
    Dim numData As BetterArray

    'Sort on values only If there is data
    If Lo.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(Lo.DataBodyRange) Then Exit Sub

    Set trackData = New BetterArray
    Set numData = New BetterArray
    Set sh = Wksh()

    Set LoRng = Lo.DataBodyRange
    Set hRng = Lo.HeaderRowRange
    Set colRng = Lo.ListColumns(colIndex).DataBodyRange

    'Test values progressively
    'If the trackData does not contains the values,
    'Update the number (where you will sort)

    hRng.Cells(1, hRng.Cells.Count + 1).Value = "__number"

    With sh
        Set resRng = .Range(hRng.Cells(1, 1), LoRng.Cells(LoRng.Rows.Count, LoRng.Columns.Count + 1))
        Lo.Resize resRng
    End With

    Set numRng = Lo.ListColumns(hRng.Cells.Count + 1).DataBodyRange

    'Seems like sometimes the databodyrange of a new column could be empty in that case exit
    If numRng Is Nothing Then
        Lo.ListColumns(hRng.Cells.Count + 1).Delete
       Exit Sub
   End If

    num = 1 'Number to sort later on

    'Resize the listObject to be sure it includes the number added
    For counter = 1 To colRng.Cells.Count

       sortedValue = colRng.Cells(counter, 1).Value

      If Not trackData.Includes(sortedValue) Then
        
        trackData.Push sortedValue
        numData.Push num
        numRng.Cells(counter, 1).Value = num
        num = num + 1
      Else
        numRng.Cells(counter, 1) = numData.Items(trackData.IndexOf(sortedValue))
      End If
    Next

     'Now sort the listObject on the column
     LoRng.Sort key1:=numRng, order1:=xlAscending, Header:=xlYes
     
     'Remove the _number column
     Lo.ListColumns(hRng.Cells.Count + 1).Delete
    
End Sub


Private Sub SortAnalysis()

    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim colIndex As Long
    Dim tabl As BetterArray 'Will contain the list of listObjects to sort
    Dim counter As Long
    Dim loName As String

    Set tabl = New BetterArray
    tabl.LowerBound = 1

    tabl.Push UNIVARIATE, BIVARIATE, TIMESERIES, SPATIAL

    Set sh = Wksh()

    For counter = 1 To tabl.Length
        loName = tabl.Items(counter)
        Set Lo = sh.ListObjects(loName)
        colIndex = ColumnIndex(Lo, "section")
    
        SortListOnFirst Lo, colIndex
    Next
End Sub

'Translate
Private Sub Translate(ByVal TransObject As ITranslation)
    Dim transCols As BetterArray
    Dim anash As Worksheet
    Dim counter As Long
    Dim Lo As ListObject
    Dim colRng As Range

    'TransObject is the translation object (from Translation.cls). It can
    'translate ranges, and formula ranges.

    Set transCols = New BetterArray
    Set anash = Wksh()
    transCols.Push "section", "title", "summary label", "choices" 'choices is for the time series graphs translation

    For Each Lo In anash.ListObjects
        For counter = transCols.LowerBound To transCols.UpperBound
            If ColumnExists(Lo, transCols.Item(counter)) Then
                Set colRng = DataRange(Lo, transCols.Item(counter))
                TransObject.TranslateRange colRng
            End If
        Next
        'Formulas
        If ColumnExists(Lo, "summary function") Then
            Set colRng = DataRange(Lo, "summary function")
            TransObject.TranslateRange colRng, containsFormula:=True
        End If
    Next
End Sub

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Interface
Public Sub ILLAnalysis_Import(ByVal impWksh As Worksheet)
    Import impWksh
End Sub

Public Sub ILLAnalysis_Translate(ByVal TransObject As ITranslation)
    Translate TransObject
End Sub

Public Property Get ILLAnalysis_Wksh() As Worksheet
    Set ILLAnalysis_Wksh = Wksh()
End Property

Public Sub ILLAnalysis_Build(ll As ILinelist)
    Build ll
End Sub

