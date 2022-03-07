Attribute VB_Name = "M_NomVisible"
Option Explicit

Const C_PWD As String = "1234"
Const C_TitleCol As Byte = 5
Public TriggerShowHide As Boolean
' Retrieving the heading of the dictionnary (names of columns)

Function CreateDicTitle() As BetterArray
    Dim T_headers As BetterArray                 'headers: colnames of the dictionary
    Set T_headers = New BetterArray
        
    'loading headers
    T_headers.Clear
    T_headers.FromExcelRange Sheets("Dico").Range("A1"), DetectLastRow:=False, DetectLastColumn:=True
    'Checking the visibility variable
    If Not T_headers.Includes("Visibility") Then
        T_headers.Push "Visibility"
        'add the visibility
        Sheets("Dico").Cells(1, T_headers.UpperBound).value = "Visibility"
    End If
    
    Set CreateDicTitle = T_headers.Clone
    'freeing memory
    Set T_headers = Nothing
End Function

'Function to get some special columns of the dictionary into a betterarray 2D table
Function ExtractDicColumns(T_colnames As BetterArray) As BetterArray
    Dim T_headers As BetterArray                 'headers of the dictionary
    Dim T_data As BetterArray                    'Temporary data, to return
    Dim T_values As BetterArray
    Dim i As Integer                             'counter for indexing the colnames
    Dim iColIndex As Integer
    Dim iRowIndex As Integer
    Dim j As Integer                             'counter for the number of Items of the Data
    
    
    'Get the headers of the dictionary
    Set T_headers = New BetterArray
    Set T_values = New BetterArray
    Set T_headers = CreateDicTitle
    Set T_data = New BetterArray
    
    T_data.LowerBound = 1
    j = 1
    With Sheets("Dico")
        iRowIndex = .Cells(1, 1).End(xlDown).Row
        For i = 1 To T_colnames.UpperBound
            iColIndex = T_headers.IndexOf(T_colnames.Items(i)) 'T_colnames is a 1-D table
            If iColIndex > 0 Then
                T_values.FromExcelRange Range(.Cells(2, iColIndex), _
                                              .Cells(iRowIndex, iColIndex)), _
        DetectLastRow:=False, DetectLastColumn:=False
                T_data.Item(j) = T_values.Items
                T_values.Clear
                j = j + 1
            End If
        Next
    End With
    'Destroy headers and values
    Set T_values = Nothing
    Set T_headers = Nothing
    
    If T_data.UpperBound > 1 Then
        'Convert the jagged to multidimensional
        'Seems like there is something wrong with the dimensions which is imposing to clone
        'the data before transposing: maybe linked to the betterarray class itself
        T_data.ArrayType = BA_MULTIDIMENSION
        Set T_data = T_data.Clone
        T_data.Transpose
        Set ExtractDicColumns = T_data.Clone
        Set T_data = Nothing
    Else
        Set ExtractDicColumns = T_data.Clone
        Set T_data = Nothing
    End If
End Function

'This command loads variables and
'put all of them in the list of the show/hide forms
'only not hidden variables are shown. We need to filtered out
'those variables

Sub ClicCmdVisibleName()

    Dim T_mainlab As BetterArray                 'main label table
    Dim T_varname As BetterArray                 'varname table
    Dim T_status As BetterArray                  'status table
    Dim T_headers As BetterArray                 'headers of the dictionary table
    Dim T_data As BetterArray                    'temporary data for storing the values
    Dim wks As Worksheet                         'Setting a temporary variable for dictionary selection
    Dim i As Integer
    
    'Setting and initializing the tables
    Set T_mainlab = New BetterArray
    Set T_varname = New BetterArray
    Set T_status = New BetterArray
    Set T_data = New BetterArray
    Set T_headers = New BetterArray
    
    T_varname.LowerBound = 1
    T_mainlab.LowerBound = 1
    T_status.LowerBound = 1
    T_headers.LowerBound = 1
    T_data.LowerBound = 1
    
    ActiveSheet.Unprotect (C_PWD)
    Set wks = ThisWorkbook.Worksheets("Dico")
    
    'Get the headers
    Set T_headers = CreateDicTitle
    'Now update the mainlabel, status and variable name
    
    i = 1
    While (i <= wks.Cells(1, 1).End(xlDown).Row)
        If ActiveSheet.Name = wks.Cells(i, T_headers.IndexOf("Sheet")) Then
            'update only on non hidden variables
            If LCase(wks.Cells(i, T_headers.IndexOf("Status")).value) <> "hidden" Then
                T_mainlab.Push wks.Cells(i, T_headers.IndexOf("Main label")).value
                T_varname.Push wks.Cells(i, T_headers.IndexOf("Variable name")).value
                
                If LCase(wks.Cells(i, T_headers.IndexOf("Status")).value) = "mandatory" Then
                    T_status.Push "Mandatory"
                ElseIf LCase(wks.Cells(i, T_headers.IndexOf("Visibility")).value) = "hidden by user" Then
                    T_status.Push ""
                Else
                    T_status.Push "Shown"
                End If
            Else
                wks.Cells(i, T_headers.IndexOf("Visibility")).value = "Hidden by designer"
            End If
        End If
        i = i + 1
    Wend
    Set T_headers = Nothing
    
    T_data.Item(1) = T_mainlab.Items
    Set T_mainlab = Nothing
    T_data.Item(2) = T_varname.Items
    Set T_varname = Nothing
    T_data.Item(3) = T_status.Items
    Set T_status = Nothing
    
    Application.EnableEvents = False
    T_data.ArrayType = BA_MULTIDIMENSION
    Set T_data = T_data.Clone
    T_data.Transpose
    
    F_NomVisible.LST_NomChamp.ColumnCount = 3
    F_NomVisible.LST_NomChamp.BoundColumn = 2
    F_NomVisible.LST_NomChamp.List = T_data.Items
    'Setting objects to nothing

    Set wks = Nothing
    Set T_data = Nothing
    
    Application.EnableEvents = True
    F_NomVisible.FRM_AffMas.Visible = True
    F_NomVisible.FRM_AffMas.Width = 90
    F_NomVisible.Width = 450
    F_NomVisible.Height = 270
    F_NomVisible.CMD_Fermer.SetFocus
    F_NomVisible.show

    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
End Sub

'Find the index of sValue on column iCol of a BetterArray T_table
Function FindIndex(T_table As BetterArray, iCol As Integer, sValue As String) As Integer
    Dim T_data As BetterArray
    Set T_data = New BetterArray
    T_data.Items = T_table.ExtractSegment(ColumnIndex:=iCol)
    FindIndex = T_data.IndexOf(sValue)
End Function

'This sub will works with the logic related to the selection of oneline in the
'Show/hide multibox page
Sub UpdateVisibilityStatus(iIndex As Integer)

    Dim T_formdata As BetterArray                'Actual form data
    Set T_formdata = New BetterArray
    T_formdata.LowerBound = 1
    T_formdata.Items = F_NomVisible.LST_NomChamp.List
    F_NomVisible.FRM_AffMas.Visible = True
    
    Application.ScreenUpdating = False
    
    Select Case LCase(T_formdata.Items(iIndex + 1, 3))
    Case "mandatory"
        TriggerShowHide = False
        F_NomVisible.OPT_Affiche.value = 1
        F_NomVisible.OPT_Affiche.Caption = "Show/Mandatory"
        F_NomVisible.OPT_Affiche.Width = 80
        F_NomVisible.OPT_Affiche.Left = 0
        F_NomVisible.OPT_Affiche.Top = 20
            
        F_NomVisible.OPT_Masque.Visible = False
    Case ""                                'It is hidden, show masking
        TriggerShowHide = False
        F_NomVisible.OPT_Affiche.value = 0
        F_NomVisible.OPT_Affiche.Caption = "Show"
        F_NomVisible.OPT_Affiche.Width = 45
        F_NomVisible.OPT_Affiche.Left = 10
        F_NomVisible.OPT_Affiche.Top = 6
            
        F_NomVisible.OPT_Masque.Visible = True
        F_NomVisible.OPT_Affiche.Visible = True
        F_NomVisible.OPT_Masque.value = 1
    Case Else                                    'It is shown if not
        TriggerShowHide = False
        F_NomVisible.OPT_Affiche.value = 1
        F_NomVisible.OPT_Affiche.Caption = "Show"
        F_NomVisible.OPT_Affiche.Width = 45
        F_NomVisible.OPT_Affiche.Left = 10
        F_NomVisible.OPT_Affiche.Top = 6
            
        F_NomVisible.OPT_Masque.Visible = True
        F_NomVisible.OPT_Affiche.Visible = True
        F_NomVisible.OPT_Masque.value = 0
    End Select
             
    'Freeing the memory
    Set T_formdata = Nothing
    
    'Return the triggering status
    TriggerShowHide = True
    
    Application.ScreenUpdating = True
End Sub

'This procedures hides or shows one column from the One sheet given the variable name selected
'in the visibility form
Sub ShowHideColumnSheet(sSheetname As String, ByVal sVarname As String, Optional bhide As Boolean = True)
    'bhide is a boolean to hide or show one column
    Dim indexCol As Integer                      'Column The index of the column to Hide
    Dim T_headers As BetterArray                 'Temporary data for headers
    Set T_headers = New BetterArray
    Dim T_control As BetterArray                 'Extracting the control label to be sure we can hide all the geos
    Set T_control = New BetterArray
    Dim i As Integer                             'counter index for the columns
    Dim ifSheetRow As Integer                    'first Row index of the sheet in the dicitonary sheet
    Dim ilSheetRow As Integer                    'last row index of the sheet in the dictionary
    Dim bisGeo As Boolean                        'Geo or Not?
    Dim checkLines As Boolean                    'prevent from checking new lines if the variable is present once
    
    T_headers.LowerBound = 1
    T_control.LowerBound = 1
    'First, Get the values of the headers names
    'I will probably get an error if the cells doesn't have a name
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect (C_PWD)
    
    'varname index on the actual sheet
    For i = 1 To Sheets(sSheetname).Cells(C_TitleCol, 1).End(xlToRight).Column
        T_headers.Push Sheets(sSheetname).Cells(C_TitleCol, i).Name.Name
    Next
    indexCol = T_headers.IndexOf(sVarname)
    
    'Actual index of the varname in dictionary sheet
    T_headers.Clear
    T_headers.Push "Sheet"
    Set T_headers = ExtractDicColumns(T_headers)
    T_headers.Flatten
    ifSheetRow = T_headers.IndexOf(sSheetname)
    ilSheetRow = T_headers.LastIndexOf(sSheetname)
    
    'Extract the control column
    T_headers.Clear
    T_headers.Push "Control"
    Set T_control = ExtractDicColumns(T_headers)
    T_control.Flatten
    
    'Extract the Variable names column
    T_headers.Clear
    T_headers.Push "Variable name"
    Set T_headers = ExtractDicColumns(T_headers)
    T_headers.Flatten
    checkLines = (T_headers.IndexOf(sVarname) <> T_headers.LastIndexOf(sVarname))
    
    bisGeo = False
    If checkLines Then
        For i = ifSheetRow To ilSheetRow
            If T_headers.Items(i) = sVarname Then
                If T_control.Items(i) = "geo" Then
                    bisGeo = True
                End If
            End If
        Next i
    Else
        bisGeo = (T_control.Items(T_headers.IndexOf(sVarname)) = "geo")
    End If
    'Destroying
    Set T_headers = Nothing
    Set T_control = Nothing
    
    If indexCol > 0 Then
        'Now hiding
        Sheets(sSheetname).Columns(indexCol).Hidden = bhide
        'Testing if it is a geo column and hide the following
        If bisGeo Then
            Sheets(sSheetname).Columns(indexCol + 1).Hidden = bhide
            Sheets(sSheetname).Columns(indexCol + 2).Hidden = bhide
            Sheets(sSheetname).Columns(indexCol + 3).Hidden = bhide
        End If
    End If
    
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    Application.ScreenUpdating = True
End Sub

'A simple Procedure to update the third column of the next formname to either hidden or Shown before
'moving to the logic

Sub UpdateFormData(ByRef T_table As BetterArray, index As Integer, Optional bhide As Boolean = True)
    'We can't mutate easily on a 2D array. We need first to work on updating the data from a
    '1D array. That is the purpose of the T_values BetterArray
    Dim T_values As BetterArray
    Set T_values = New BetterArray
    
    'Get the current row
    T_values.Items = T_table.Item(index + 1)     'the index starts at -1 on a listbox
    'Update the visibility status
    If bhide Then
        T_values.Item(3) = ""
    Else
        T_values.Item(3) = "Shown"
    End If
    'Mutate in the form table
    T_table.Item(index + 1) = T_values.Items
    Set T_values = Nothing
End Sub

'Logic behind the show/hide click
Sub ShowHideLogic(iIndex As Integer)
    
    If Not TriggerShowHide Or iIndex < 0 Then    'when the form is shown at the begining, nothing is selected and index can be -1
        Exit Sub
    Else
        Application.ScreenUpdating = False
        
        Dim T_formdata As BetterArray
        Set T_formdata = New BetterArray
        T_formdata.Items = F_NomVisible.LST_NomChamp.List
        T_formdata.LowerBound = 1
        
        'Update data in form
        If LCase(T_formdata.Items(iIndex + 1, 2)) <> "mandatory" Then
        
            'For mutating, we can only use the item method. Items with s, read only values,
            'we can't set values with items
            
            If F_NomVisible.OPT_Masque.value Then
                '// --- Here I update the Data to show "Hidden"
                UpdateFormData T_table:=T_formdata, index:=iIndex, bhide:=True
                '//--- Actually hide the column
                ShowHideColumnSheet sSheetname:=ActiveSheet.Name, sVarname:=T_formdata.Items(iIndex + 1, 2), bhide:=True
            Else
                '// --- Here I udpate the data to show "Shown"
                UpdateFormData T_table:=T_formdata, index:=iIndex, bhide:=False
                ShowHideColumnSheet sSheetname:=ActiveSheet.Name, sVarname:=T_formdata.Items(iIndex + 1, 2), bhide:=False
            End If
        End If
    
        'Reload it back
        F_NomVisible.LST_NomChamp.Clear
        F_NomVisible.LST_NomChamp.List = T_formdata.Items
        F_NomVisible.LST_NomChamp.Selected(iIndex) = True
        Set T_formdata = Nothing
        
        Application.ScreenUpdating = True
    End If
    
    
End Sub

'Writes actual values of visibility criteria in the dictionary sheet when you click back
Sub WriteVisibility()
    Dim T_formdata As BetterArray                'Form data
    Dim T_varname As BetterArray                 'varname in the dictionary sheet
    Dim ssheetVarname As String                  'sheet variable name
    
    Dim i As Integer
    Dim T_values As BetterArray                  'Temporary data for controling values in one sheet
    Dim isheetIndex As Integer                   ' Sheet line index in the dictionary data
    Dim ilsheetIndex As Integer                  'last index of the sheet
    Dim ivisIndex As Integer                     'Index of the visibility
    
    Application.ScreenUpdating = False
      
    Set T_values = New BetterArray
    T_values.LowerBound = 1
    Set T_varname = New BetterArray
    
    'Extract the sheet to get first and last index of sheet name
    T_values.Push "Sheet"
    Set T_values = ExtractDicColumns(T_values)
    T_values.Flatten
    
    isheetIndex = T_values.IndexOf(ActiveSheet.Name)
    ilsheetIndex = T_values.LastIndexOf(ActiveSheet.Name)
    
    'index of the visibility to update
    ivisIndex = Sheets("Dico").Cells(1, 1).End(xlToRight).Column
    T_values.Clear
    
    'The variable names are used of searching in the form vs sheet
    T_values.Push "Variable name"
    Set T_varname = ExtractDicColumns(T_values)
    T_varname.Flatten
    T_values.Clear
    
    'Be as lazy as possible when updating, to avoid doing
    'unecessary computation
    If (isheetIndex > 0) Then
        Set T_formdata = New BetterArray
        T_formdata.LowerBound = 1
        T_formdata.Items = F_NomVisible.LST_NomChamp.List
        'here values is the variable name in the form
        T_values.Items = T_formdata.ExtractSegment(ColumnIndex:=2)
        T_values.Flatten
        For i = isheetIndex To (ilsheetIndex - 1)
            'if the variable name is in the form
            ssheetVarname = T_varname.Item(i)
            If T_values.Includes(ssheetVarname) Then
                'if the visibility is different
                If T_formdata.Items(T_values.IndexOf(ssheetVarname), 3) <> _
                                                                        Sheets("Dico").Cells(i + 1, ivisIndex).value Then
                    'Then update the visibility i + 1 to take in account the first index with column names
                    Sheets("Dico").Cells(i + 1, ivisIndex).value = _
                                                             T_formdata.Items(T_values.IndexOf(ssheetVarname), 3)
                ElseIf T_formdata.Items(T_values.IndexOf(ssheetVarname), 3) = "" Then 'hidden
                     Sheets("Dico").Cells(i + 1, ivisIndex).value = "Hidden by user"
                End If
            End If
        Next
    End If
    Set T_formdata = Nothing
    Set T_values = Nothing
    Set T_varname = Nothing
    
    Application.ScreenUpdating = True
End Sub


