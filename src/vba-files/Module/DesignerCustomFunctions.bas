Attribute VB_Name = "DesignerCustomFunctions"
Option Private Module
Option Explicit
'parsing case_when functions (This is done at the designer)

'Convert case_when to excel function
Function CaseWhen(Values As BetterArray) As String
    Dim i        As Long
    Dim sFormula As String
    Dim sPar     As String                       'Number of parenthesis for the formula

    'Initialisations:
    sFormula = vbNullString
    sPar = vbNullString
    i = 1

    'We convert to nested ifs, up to 64 nested ifs can be used in Excel
    Do While i < Values.Length
        sFormula = sFormula & "IF" & "(" & Values.Items(i) & ", " & _
                                                           Values.Items(i + 1) & ", "
        i = i + 2
        sPar = sPar & ")"
    Loop

    If i = Values.Length Then
        'odd number of conditions, there is a default
        sFormula = sFormula & Values.Items(i) & sPar
    Else
        'even number of conditions, there is no default
        sFormula = sFormula & Chr(34) & Chr(34) & sPar
    End If

    CaseWhen = sFormula
End Function

'Parsing the case_when. The user should use the
'CASE_WHEN in upper case.
Function ParseCaseWhen(sFormula As String) As String

    Dim parsingTable As BetterArray
    Dim sLab As String
    Dim i As Long
    Dim iprev As Long
    Dim NbOpened As Long
    Dim NbClosed As Long
    Dim OpenedQuotes As Boolean
    Dim OpenedParenthesis As Boolean


    If InStr(1, sFormula, "CASE_WHEN") > 0 Then
        Set parsingTable = New BetterArray
        parsingTable.LowerBound = 1
        iprev = 1

        NbOpened = 0
        NbClosed = 0

        sLab = Application.WorksheetFunction.Trim(Replace(sFormula, "CASE_WHEN(", ""))
        sLab = Left(sLab, Len(sLab) - 1)
        For i = 1 To Len(sLab)

            'Manage parenthesis and quotes,
            If Mid(sLab, i, 1) = Chr(34) Then OpenedQuotes = Not OpenedQuotes 'Opened quotes or parenthesis

            'Parenthesis not within quotes are expressions
            If Mid(sLab, i, 1) = Chr(40) And Not OpenedQuotes Then NbOpened = NbOpened + 1
            If Mid(sLab, i, 1) = Chr(41) And Not OpenedQuotes Then NbClosed = NbClosed + 1

            'Opened parenthesis is true or false depending on the number of quotes opened or closed
            OpenedParenthesis = (NbOpened <> NbClosed)

            If Not OpenedQuotes And Not OpenedParenthesis And Mid(sLab, i, 1) = "," Then
                parsingTable.Push Application.WorksheetFunction.Trim(Mid(sLab, iprev, i - iprev))
                iprev = i + 1
            End If
        Next

        parsingTable.Push Application.WorksheetFunction.Trim(Mid(sLab, iprev, i - iprev))

        ParseCaseWhen = CaseWhen(parsingTable)
    End If
End Function


