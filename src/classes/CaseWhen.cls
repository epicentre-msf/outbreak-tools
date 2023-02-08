VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseWhen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit
Implements ICaseWhen

Private Type TCaseWhen
    formula As String
    parsingTable As BetterArray
End Type

'Exposed methods
Private This As TCaseWhen

Public Function Create(ByVal formula As String) As ICaseWhen
    With New CaseWhen
        .casewhenFormula = formula
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ICaseWhen
    Set Self = Me
End Property

Public Property Get casewhenFormula() As String
    casewhenFormula = This.formula
End Property

Public Property Let casewhenFormula(ByVal formula As String)
    This.formula = Application.WorksheetFunction.Trim(formula)
End Property

Private Property Get Valid() As Boolean
    Dim formula As String
    formula = casewhenFormula
    Valid = (InStr(1, formula, "CASE_WHEN") = 1)
End Property

Private Property Get CaseWhenTable() As BetterArray
    Dim validity As Boolean
    Dim parsingTable As BetterArray
    Dim formula As String
    Dim counter As Long
    Dim Previous As Long
    Dim NbOpened As Long
    Dim NbClosed As Long
    Dim OpenedQuotes As Boolean
    Dim OpenedParenthesis As Boolean

    If Not This.parsingTable Is Nothing Then
        Set CaseWhenTable = This.parsingTable.Clone()
        Exit Property
    End If

    Set parsingTable = New BetterArray
    parsingTable.LowerBound = 1

    If Valid() Then
        Previous = 1
        NbOpened = 0
        NbClosed = 0
        OpenedQuotes = False
        OpenedParenthesis = False
        formula = Application.WorksheetFunction.Trim(Replace(casewhenFormula(), "CASE_WHEN(", ""))
        formula = Left(formula, Len(formula) - 1) 'Removed the closing parenthesis
        For counter = 1 To Len(formula)
            'Manage parenthesis and quotes,
            If Mid(formula, counter, 1) = chr(34) Then OpenedQuotes = Not OpenedQuotes 'Opened quotes or parenthesis
            'Parenthesis not within quotes are not expressions
            If Mid(formula, counter, 1) = chr(40) And Not OpenedQuotes Then NbOpened = NbOpened + 1
            If Mid(formula, counter, 1) = chr(41) And Not OpenedQuotes Then NbClosed = NbClosed + 1
            OpenedParenthesis = (NbOpened <> NbClosed)
            'Opened parenthesis will allow excel formulas within a  case_when
            If Not OpenedQuotes And Not OpenedParenthesis And Mid(formula, counter, 1) = "," Then
                parsingTable.Push Application.WorksheetFunction.Trim(Mid(formula, Previous, counter - Previous))
                Previous = counter + 1
            End If
        Next
        'push the last chunk
        parsingTable.Push Application.WorksheetFunction.Trim(Mid(formula, Previous, counter - Previous))
    End If

    Set This.parsingTable = parsingTable.Clone()
    Set CaseWhenTable = This.parsingTable.Clone()
End Property

'pasing the case when
Private Function parsedFormula() As String

    Dim counter As Long
    Dim formula As String
    Dim parenthesis  As String
    Dim parsingTable As BetterArray
    'Number of parenthesis for the formula
    formula = vbNullString
    parenthesis = vbNullString
    counter = 1

    If Valid() Then
        Set parsingTable = CaseWhenTable()
        'We convert to nested ifs, up to 64 nested ifs can be used in Excel
        Do While counter < parsingTable.Length
            formula = formula & "IF" & "(" & parsingTable.Items(counter) & ", " & _
                                                                         parsingTable.Items(counter + 1) & ", "
            counter = counter + 2
            parenthesis = parenthesis & ")"
        Loop
        If counter = parsingTable.Length Then
            'odd number of conditions, there is a default
            formula = formula & parsingTable.Items(counter) & parenthesis
        Else
            'even number of conditions, there is no default
            formula = formula & chr(34) & chr(34) & parenthesis
        End If
    Else
        'Trace log here
    End If
    parsedFormula = formula
End Function

Private Function Categories() As BetterArray
    Dim categoriesTable As BetterArray
    Dim counter As Long
    Dim parsingTable As BetterArray
    Dim Level As String

    Set categoriesTable = New BetterArray
    categoriesTable.LowerBound = 1

    If Valid() Then
        counter = 2
        Set parsingTable = CaseWhenTable()
        Do While counter <= parsingTable.Length

            If Mid(parsingTable.Item(counter), 1, 1) = chr(34) Then
                'Quoted character, is a level
                Level = Replace(parsingTable.Item(counter), chr(34), vbNullString)
            Else
                'Not Quoted character, the user should not use case_when in analysis
                Level = chr(34) & parsingTable.Item(counter) & chr(34)
            End If
            If Level <> vbNullString Then categoriesTable.Push Level
            counter = counter + 2
        Loop
    End If

    Set Categories = categoriesTable.Clone()
End Function

'Interface

Public Property Get ICaseWhen_Valid() As Boolean
    ICaseWhen_Valid = Valid()
End Property

Public Property Get ICaseWhen_ParsedFormula() As String
    ICaseWhen_ParsedFormula = parsedFormula()
End Property

Public Property Get ICaseWhen_Categories() As BetterArray
    Set ICaseWhen_Categories = Categories()
End Property


