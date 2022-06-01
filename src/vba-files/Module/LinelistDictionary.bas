Attribute VB_Name = "LinelistDictionary"

'Helpers functions to work with the dictionary. Most of the output are
'In BetterArray. Here you can extract one column of the dictionary,
'Get the dictionary headers, find the value of one column for one variable
'and Find the values of one variable of the dictionary given a condition
'on another variable. The Goal is to ease as much as possible the process
'behind acessing values of the dictionary so that we don't border ourselves
'with that. If you also want to get an array instead of a BetterArray you
'can just convert the betterarray to array by retrieving the items of the
'BetterArray:
'Dim myArray()
'Dim BA as BetterArray
'Set BA = New BetterArray
'myArray = BA.Items 'retrieve the items of the BetterArray if you prefer
'working with arrays.


'Get The Dictionnary Headers From the Dictionary worksheet
Option Explicit
Option Base 1

Function GetDictionaryHeaders() As BetterArray
    Dim DictHeaders As BetterArray
    Dim Wkb As Workbook

    Set DictHeaders = New BetterArray
    DictHeaders.LowerBound = 1

    Set Wkb = ThisWorkbook
    DictHeaders.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=True, DetectLastRow:=False
    'Set the Array

    Set GetDictionaryHeaders = DictHeaders.Clone()
    Set DictHeaders = Nothing

End Function

'Be sure one column is in the dictionary Headers
Function isInDictHeaders(sColname As String) As Integer
    Dim DictHeaders As BetterArray
    Set DictHeaders = GetDictionaryHeaders()
    isInDictHeaders = DictHeaders.IndexOf(sColname)
    Set DictHeaders = Nothing
End Function

'Get Dictionary index of one variable
Function GetDictionaryIndex(sColname As String) As Integer
  Dim DictHeaders As BetterArray
  Set DictHeaders = GetDictionaryHeaders()
  DictHeaders.LowerBound = 1
  GetDictionaryIndex = DictHeaders.IndexOf(sColname)
  Set DictHeaders = Nothing
End Function

'Get one column from the dictionary
Function GetDictionaryColumn(sColname As String) As BetterArray
    Dim ColumnData As BetterArray
    Set ColumnData = New BetterArray
    ColumnData.LowerBound = 1

    'Then we check if the colname is in the headers, if not, you end up with
    'Empty BetterArray

    If isInDictHeaders(sColname) Then
        With ThisWorkbook.Worksheets(C_sParamSheetDict)
            ColumnData.FromExcelRange .ListObjects("o" & ClearString(C_sParamSheetDict)).ListColumns(sColname).DataBodyRange
        End With
    End If
    Set GetDictionaryColumn = ColumnData.Clone()
    Set ColumnData = Nothing
End Function

'Retrieve all the dictionnary data, excluding the headers
Function GetDictionaryData() As BetterArray
    Dim DictData As BetterArray
    Set DictData = New BetterArray
    DictData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetDict)
         DictData.FromExcelRange .Cells(2, 1), DetectLastRow:=True, DetectLastColumn:=True
    End With

    Set GetDictionaryData = DictData.Clone
    Set DictData = Nothing
End Function

'Retrieve all the Choices data, excluding the headers
Function GetChoicesData() As BetterArray
    Dim ChoicesData As BetterArray
    Set ChoicesData = New BetterArray
    ChoicesData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetChoices)
        .Visible = xlSheetHidden
         ChoicesData.FromExcelRange .Cells(1, 1), DetectLastRow:=True, DetectLastColumn:=True
         .Visible = xlSheetVeryHidden
    End With

    Set GetChoicesData = ChoicesData.Clone
    Set ChoicesData = Nothing
End Function

'Retrieve all the Translation data, excluding the headers
Function GetTransData() As BetterArray
    Dim TransData As BetterArray
    Set TransData = New BetterArray
    TransData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetTranslation)
        .Visible = xlSheetHidden
         TransData.FromExcelRange .Cells(1, 1), DetectLastRow:=True, DetectLastColumn:=True
         .Visible = xlSheetVeryHidden
    End With

    Set GetTransData = TransData.Clone
    Set TransData = Nothing
End Function

'Retrieve the variable names of the dictionnary on one condition on a variable
'Here the condition is only equallity (a kind of filter, but for the
'dictionary only)

Function GetDictDataFromCondition(sColumnName As String, sCondition As String, Optional bVarNamesOnly As Boolean = False) As BetterArray
    Dim ColumnData As BetterArray
    Dim Rng As Range
    Dim iColIndex As Integer

    'Be sure the sColumnName is present in the headers
    If isInDictHeaders(sColumnName) Then
        iColIndex = GetDictionaryIndex(sColumnName)

        'First be sure the dictionnary is filtered on column name:
        With ThisWorkbook.Worksheets(C_sParamSheetDict)
            With .ListObjects("o" & ClearString(C_sParamSheetDict)).Range
               .AutoFilter Field:=iColIndex, Criteria1:=sCondition
            End With
            Set Rng = .ListObjects("o" & ClearString(C_sParamSheetDict)).Range.SpecialCells(xlCellTypeVisible)
        End With

        'Take the special cells and copy the data
        With ThisWorkbook.Worksheets(C_sSheetTemp)
            .Visible = xlSheetHidden
            .Cells.Clear
            Rng.Copy Destination:=.Cells(1, 1)
            Set ColumnData = New BetterArray
            ColumnData.LowerBound = 1
            If bVarNamesOnly Then
                ColumnData.FromExcelRange .Cells(2, 1), DetectLastColumn:=False, DetectLastRow:=True
            Else
                ColumnData.FromExcelRange .Cells(2, 1), DetectLastColumn:=True, DetectLastRow:=True
            End If

            .Cells.Clear
            .Visible = xlSheetVeryHidden
        End With

        Set Rng = Nothing
        Set GetDictDataFromCondition = ColumnData.Clone()
        Set ColumnData = Nothing
    End If
End Function

'Retrieve the value of one column given one variable name
Function GetDictColumnValue(sVarName As String, sColname As String) As String
    Dim VarNameData As BetterArray
    Dim ColnameData As BetterArray
    GetDictColumnValue = vbNullString

    Set VarNameData = GetDictionaryColumn(C_sDictHeaderVarName)

    If VarNameData.Includes(sVarName) Then
        Set ColnameData = GetDictionaryColumn(sColname)
        If ColnameData.Length > 0 Then
           GetDictColumnValue = ColnameData.Item(VarNameData.IndexOf(sVarName))
        End If
    End If
    Set ColnameData = Nothing
    Set VarNameData = Nothing
End Function


'Change value of one variable and one column in the dictionary

Sub UpdateDictionaryValue(sVarName As String, sColname As String, sNewValue As String)

    Dim VarNameData As BetterArray
    Dim ColnameData As BetterArray
    Dim iRow As Integer
    Dim iColumn As Integer

    Set VarNameData = GetDictionaryColumn(C_sDictHeaderVarName)

    If VarNameData.Includes(sVarName) Then
        If isInDictHeaders(sColname) Then
            iColumn = GetDictionaryIndex(sColname)
            iRow = VarNameData.IndexOf(sVarName) + 1

            If iRow > 0 And iColumn > 0 Then
                With ThisWorkbook.Worksheets(C_sParamSheetDict)
                    .Cells(iRow, iColumn).value = Application.WorksheetFunction.Trim(sNewValue)
                End With
            End If

        End If
    End If

End Sub
