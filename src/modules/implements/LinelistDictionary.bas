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
Option Private Module

Function GetDictionaryHeaders() As BetterArray
    Dim DictHeaders As BetterArray
    Dim Wkb As Workbook

    Set DictHeaders = New BetterArray
    DictHeaders.LowerBound = 1

    Set Wkb = ThisWorkbook
    DictHeaders.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=True, DetectLastRow:=False
    'Set the Array

    Set GetDictionaryHeaders = DictHeaders.Clone()

End Function

'Be sure one column is in the dictionary Headers
Function isInDictHeaders(sColname As String) As Integer
    Dim DictHeaders As BetterArray
    Set DictHeaders = GetDictionaryHeaders()
    isInDictHeaders = DictHeaders.IndexOf(sColname)
End Function

'Get Dictionary index of one variable
Function GetDictionaryIndex(sColname As String) As Integer
    Dim DictHeaders As BetterArray
    Set DictHeaders = GetDictionaryHeaders()
    DictHeaders.LowerBound = 1
    GetDictionaryIndex = DictHeaders.IndexOf(sColname)
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
End Function

'Retrieve all the dictionnary data, excluding the headers
Function GetDictionaryData() As BetterArray
    Dim dictData As BetterArray
    Set dictData = New BetterArray
    dictData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetDict)
        dictData.FromExcelRange .Cells(2, 1), DetectLastRow:=True, DetectLastColumn:=True
    End With

    Set GetDictionaryData = dictData.Clone
End Function

'Retrieve all the Choices data, excluding the headers
Function GetChoicesData() As BetterArray
    Dim ChoicesData As BetterArray
    Set ChoicesData = New BetterArray
    ChoicesData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetChoices)
        ChoicesData.FromExcelRange .Cells(1, 1), DetectLastRow:=True, DetectLastColumn:=True
    End With

    Set GetChoicesData = ChoicesData.Clone
End Function

'Retrieve all the Translation data, excluding the headers
Function GetTransData() As BetterArray
    Dim TransData As BetterArray
    Set TransData = New BetterArray
    TransData.LowerBound = 1

    With ThisWorkbook.Worksheets(C_sParamSheetTranslation)
        TransData.FromExcelRange .ListObjects(1).Range
    End With

    Set GetTransData = TransData.Clone
End Function

