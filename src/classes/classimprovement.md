We are still facing issues for the filtering in checkingoutput. They are linked
to the worksheet not beeing




Private Function StripIcons(ByVal txt As String) As String 
   Dim stripped As String 
   'Remove severity glyphs so comparisons remain text based 
   stripped = Replace(txt, ChrW(10060), ) 
   stripped = Replace(stripped, ChrW(9888), ) 
   stripped = Replace(stripped, ChrW(8505), ) 
   stripped = Replace(stripped, ChrW(9998), ) 
   stripped = Replace(stripped, ChrW(10004), ) 
   StripIcons = Trim$(stripped) 
 End Function 

 Private Function ResolveMatchType(ByVal candidate As String) As String 
   'Normalise the requested status so the filtering block can branch quickly 
   Select Case UCase$(Trim$(candidate)) 
       Case , ALL 
           ResolveMatchType = ALL 
       Case ERROR, ERRORS 
           ResolveMatchType = ERROR 
       Case WARNING, WARNINGS 
           ResolveMatchType = WARNING 
       Case NOTE, NOTES 
           ResolveMatchType = NOTE 
       Case INFO, INFOS 
           ResolveMatchType = INFO 
       Case SUCCESS, SUCCESSES 
           ResolveMatchType = SUCCESS 
       Case WITHOUT SUCCESS, WITHOUT SUCCESSES 
           ResolveMatchType = WITHOUT_SUCCESS 
       Case Else 
           ResolveMatchType = ALL 
   End Select 
 End Function 

 Private Function ShouldFilterByTitle(ByVal candidate As String) As Boolean 
   Dim trimmedTitle As String 
   trimmedTitle = Trim$(candidate) 
   'The default item keeps all titles visible 
   ShouldFilterByTitle = LenB(trimmedTitle) > 0 And StrComp(trimmedTitle,  & TITLE_FILTER_DEFAULT & , vbTextCompare) <> 0 
 End Function 

 Private Sub Worksheet_Change(ByVal Target As Range) 
   On Error GoTo ExitHandler 
   Dim statusCell As Range 
   Dim titleCell As Range 
   Dim statusChanged As Boolean 
   Dim titleChanged As Boolean 
   Set statusCell = Me.Range( & dq & FILTER_CELL_ADDRESS & dq & ) 
   Set titleCell = Me.Range( & dq & TITLE_FILTER_CELL_ADDRESS & dq & ) 
   statusChanged = Not Intersect(Target, statusCell) Is Nothing 
   titleChanged = Not Intersect(Target, titleCell) Is Nothing 
   'Ignore edits triggered outside of the filter cells 
   If Not statusChanged And Not titleChanged Then GoTo ExitHandler 
   Application.EnableEvents = False 
   Application.ScreenUpdating = False 
   Application.Calculation = xlCalculationManual 
   If titleChanged And Not statusChanged Then 
       'Status filters are secondary; show all severities for the chosen title 
       statusCell.Value = All 
   End If 
   'Reapply visibility with the refreshed selections 
   FilterCheckingOutputRows statusCell.Value, titleCell.Value, statusChanged, titleChanged 
 ExitHandler: 
   Application.EnableEvents = True 
   Application.ScreenUpdating = True 
   Application.Calculation = xlCalculationAutomatic 
 End Sub 

 Private Sub FilterCheckingOutputRows(ByVal statusValue As String, _ 
                                   ByVal titleValue As String, _ 
                                   ByVal statusChanged As Boolean, _ 
                                   ByVal titleChanged As Boolean) 
   Dim effectiveStatus As String 
   Dim matchType As String 
   Dim applyTitleFilter As Boolean 
   Dim lastRow As Long 
   Dim rowIndex As Long 
   Dim rowType As String 
   Dim rowTitle As String 
   Dim hasLabel As Boolean 

   effectiveStatus = statusValue 
   If titleChanged And Not statusChanged Then 
       'Let the title drive the result set when the severity stayed untouched 
       effectiveStatus = All 
   End If 

   matchType = ResolveMatchType(effectiveStatus) 
   applyTitleFilter = ShouldFilterByTitle(titleValue) 

   lastRow = Me.Cells(Me.Rows.Count,  & FIRST_OUTPUT_COLUMN & ).End(xlUp).Row 
   'Exit early when no printable rows are present 
   If lastRow <  & FIRST_ROW_OUTPUT &  Then Exit Sub 

   For rowIndex =  & FIRST_ROW_OUTPUT &  To lastRow 
       'Start by revealing rows before applying filters 
       Me.Rows(rowIndex).Hidden = False 
       rowTitle = Trim$(CStr(Me.Cells(rowIndex,  & HIDDEN_TITLE_COLUMN & ).Value)) 

       If applyTitleFilter Then 
           'Hide any row whose stored parent title differs from the selection 
           If (StrComp(rowTitle, Trim$(titleValue), vbTextCompare) <> 0) And (LenB(rowTitle) <> 0) Then 
               Me.Rows(rowIndex).Hidden = True 
               GoTo ContinueRow 
           End If 
       End If 

       'Labels differentiate data rows from headers 
       hasLabel = LenB(CStr(Me.Cells(rowIndex,  & (FIRST_OUTPUT_COLUMN + 1) & ).Value)) > 0 

       If hasLabel Then 
           rowType = StripIcons(CStr(Me.Cells(rowIndex,  & FIRST_OUTPUT_COLUMN & ).Value)) 
           'Apply severity matching only to data rows 
           Select Case matchType 
               Case ALL 
                   'No filtering required for ALL 
               Case SUCCESS 
                   If StrComp(rowType, Success, vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True 
               Case WITHOUT_SUCCESS 
                   If StrComp(rowType, Success, vbTextCompare) = 0 Then Me.Rows(rowIndex).Hidden = True 
               Case Else 
                   If StrComp(rowType, matchType, vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True 
           End Select 
       End If 
 ContinueRow: 
   Next rowIndex 
 End Sub 

