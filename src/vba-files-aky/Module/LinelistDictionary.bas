Attribute VB_Name = "LinelistDictionary"


'Helpers functions to work with the dictionary. Most of the output are 
'In BetterArray. Here you can extract one column of the dictionary,
'Get the dictionary headers, find the value of one column for one variable
'and Find the values of one variable of the dictionary given a condition
'on another variable. The Goal is to ease as much as possible the process
'behind acessing values of the dictionary so that we don't border ourselves
'with that. If you also wan to get an array instead of a BetterArray you
'can just convert the betterarray to array by retrieving the items of the
'BetterArray:
'Dim myArray()
'Dim BA as BetterArray
'Set BA = New BetterArray
'myArray = BA.Items 'retrieve the items of the BetterArray if you prefer
'working with arrays.


'Get The Dictionnary Headers From the Dictionary worksheet

Function GetDictionaryHeaders() As BetterArray
    Dim DictHeaders as BetterArray
    Dim wkb as workbook

    Set DictHeaders = New BetterArray
    Set wkb = ThisWorkbook

    DictHeaders.FromExcelRange wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=True, DetectLastRow:= False
    'Set the Array

    Set GetDictionaryHeaders = DictHeaders.Clone
    Set DictHeaders = Nothing

End Function


'Be sure one column is in the dictionary Headers
Function isInDictHeader(sColname) as Boolean
    Dim DictHeaders as BetterArray
    Set DictHeaders = New BetterArray
    Set DicHeaders = GetDictionaryHeaders()

    isInDictHeader = DictHeader.Includes(sColname)
    Set DictHeaders = Nothing
End function

'Get one column from the dictionary

Function GetDictionaryColumn(sColname as String) as BetterArray
    
    Dim ColumnData As BetterArray
    Set ColumnData = New BetterArray
    
    'Then we check if the colname is in the headers, if not, you end up with
    'Empty BetterArray

    If isInDictHeader(sColname) Then
        With ThisWorkbook.Worksheets(C_sParamSheetDict)
            ColumnData.FromExcelRange .ListObjects("o" & ClearString(C_sParamSheetDict)).ListColumns(sColname).DataBodyRange
        End with
    End if

    Set GetDictionaryColumn = ColumnData.Clone()
    Set ColumnData = Nothing
End function

'Retrieve all the dictionnary data, excluding the headers
Function GetDictionaryData() as BetterArray
    Dim  DictData as BetterArray
    Set  DictData = New BetterArray
    With ThisWorkbook.Worksheets(C_sParamSheetDict)
         DictData.FromExcelRange .ListObjects("o" & ClearString(C_sParamSheetDict)).DataBodyRange
    End with
    Set GetDictionaryData = DictData.Clone
    Set DictData = Nothing
End Function


'Get Dictionary index of one variable
Function GetDictionaryIndex(sColname as String) as Integer 
  Dim DictHeaders as BetterArray
  Set DictHeaders = GetDictionaryHeaders()
  GetDictionaryIndex = DictHeaders.indexOf(sColname)
  Set DictHeaders = Nothing
End Function

'Retrieve the variable names of the dictionnary on one condition on a variable
'Here the condition is only equallity (a kind of filter, but for the 
'dictionary only)

Function GetVarNamesFromCondition(sColumnName as String, sCondition as String) As BetterArray
    Dim ColumnData as BetterArray
    Dim Rng as Range
    Dim iColIndex as Integer

    'Be sure the sColumnName is present in the headers
    If isInDictHeader(sColumnName) Then
        iColIndex = GetDictionaryIndex(sColumnName)

        'First be sure the dictionnary is filtered on column name:
        With ThisWorkbook.Worksheets(C_sParamSheetDict)
            With .ListObjects("o" & ClearString(C_sParamSheetDict)).Range
               .AutoFilter Field:=iColIndex, Criteria1:=sCondition
            End With
            Set Rng = .ListObjects("o" & ClearString(C_sParamSheetDict)).Range.SpecialCells(xlCellTypeVisible)
        End With
        'Take the special cells
        With ThisWorkbook.Worksheets(C_sSheetTemp)
            .Cells.Clear
            Rng.Copy Destination:= .Cells(1, 1)
            Set ColumnData = New BetterArray
            ColumnData.FromExcelRange .Cells(2, iColIndex), DetectLastColumn:=False, DetectLastRow:=True
        End With

        Set Rng = Nothing
        Set GetVarNamesFromCondition = ColumnData.Clone()
        Set ColumnData = Nothing
    End If
End Function

'Retrieve two variable names from two conditions

Function Get2VarNamesFromCodition(sColumnName1 as String, sColumnName2 as String, _ 
                                 sCondition1 as String, sConditon2 as String) as BetterArray


End Function