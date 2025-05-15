VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'@Folder("Linelist")
'@ModuleDescription("Manage the Spatial analysis tables")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Implements ILLSpatial

Private Type TLLSpatial
    sh As Worksheet
End Type

Private this As TLLSpatial
Private Const CLASSNAME As String = "LLSpatial"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const CONCAT As String = "concat_"
Private Const HF As String = "hf"

Public Function Create(ByVal sh As Worksheet) As ILLSpatial

    'output worksheet to add the spatial tables
    CheckRequirements sh

    With New LLSpatial
        Set .Wksh = sh
        Set Create = .Self
    End With

End Function


Public Property Get Self() As ILLSpatial
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property


Private Property Get CurrentHeaderName(ByVal varName As String, ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

Private Property Get CurrentLoName(ByVal varName As String, ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'Test if a listobject exists in a sheet
Private Function Exists(ByVal varName As String) As Boolean
  Dim listVarsRng As Range
  Dim sh As Worksheet

  Set sh = Wksh()
  Set listVarsRng = sh.ListObjects("listofgeovars").Range

  Exists = (Not (listVarsRng.Find(What:=varName, lookAt:=xlPart, MatchCase:=True) Is Nothing))
End Function

Private Function LoExists(ByVal loName As String) As Boolean
  Dim sh As Worksheet
  Dim Lo As listObject
  Set sh = Wksh()

  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0

  LoExists = (Not Lo Is Nothing)
End Function

'Add conditions for the formulas


Private Sub Update()
  Dim actsh As Worksheet 'Actual worksheet (the spatial worksheet)
  Dim sh As Worksheet 'All worksheets in the workbook
  Dim filtsh As Worksheet
  Dim wb As Workbook
  Dim rng As Range
  Dim tablename As String
  Dim varName As String
  Dim cellRng As Range
  Dim controlValue As String
  Dim Lo As listObject

  Set actsh = Wksh()
  Set wb = actsh.Parent

  For Each sh In wb.Worksheets

    'Test if it is a HList
    If sh.Cells(1, 3) = "HList" Then

      Set filtsh = wb.Worksheets(sh.Cells(1, 5).Value)

      'Exit if there is no data in the filters
      If Not (filtsh.ListObjects(1).DataBodyRange Is Nothing) Then

        tablename = sh.Cells(1, 4).Value
        Set cellRng = sh.Range(tablename & "_START")
        Set Lo = filtsh.ListObjects(1)

          'Offset - 1 is the range with the variable names
          Do While Not IsEmpty(cellRng)
            controlValue = cellRng.Offset(-4).Value

            If controlValue = "geo1" Then

              varName = Replace(cellRng.Value, ADMIN & "1_", "")

              If Exists(varName) Then

                Set rng = Lo.ListColumns(CONCAT & ADMIN & "1_" & varName).DataBodyRange()
                UpdateSpatialValue rng, varName

              End If

              'Skip the columns of geo2, geo3 and geo4 and also skip the 4 columns of pcode
              'and the other 4 columns of concatenate (4 + 4 + 3 = 11)
              Set cellRng = cellRng.Offset(, 11)

            ElseIf controlValue = "hf" Then

              'Update spatial values on health facility
              varName = Replace(cellRng.Value, HF & "_", "")

              If Exists(varName) Then

                Set rng = Lo.ListColumns(HF & "_" & varName).DataBodyRange()
                UpdateSpatialValue rng, varName, isHF:=True

              End If

              'Jump pcode for health facility
              Set cellRng = cellRng.Offset(, 1)
            End If

            Set cellRng = cellRng.Offset(, 1)
          Loop
      End If
    End If
  Next

  actsh.UsedRange.calculate
End Sub

Private Function ListOfExistingVarNames(ByVal varName As String) As BetterArray

  Dim listRng As Range
  Dim cellRng As Range
  Dim sh As Worksheet
  Dim listTab As BetterArray

  Set sh = Wksh()
  Set listTab = New BetterArray
  Set listRng = sh.ListObjects("listofgeovars").Range

  Do While (Not (listRng.Find(What:=varName, lookAt:=xlPart, MatchCase:=True) Is Nothing))
    Set cellRng = listRng.Find(What:=varName, lookAt:=xlPart, MatchCase:=True, _
                               after:=listRng.Cells(listRng.Rows.Count, 1))
    listTab.Push cellRng.Value

    'Exit the loop on those two conditions
    If (listRng.Rows.Count = 1) Then Exit Do
    If cellRng.Cells(2, 1).Row > listRng.Cells(listRng.Rows.Count, 1).Row Then Exit Do

    Set listRng = sh.Range(cellRng.Cells(2, 1), listRng.Cells(listRng.Rows.Count, 1))
  Loop

  Set ListOfExistingVarNames = listTab.Clone()
End Function


Private Sub UpdateSpatialValue(ByVal inputRng As Range, ByVal inputVarName As String, Optional ByVal isHF As Boolean = False)

  Dim pasteRng As Range
  Dim counter As Long
  Dim adminCounter As Long
  Dim sh As Worksheet
  Dim cellRng As Range
  Dim rng As Range
  Dim Lo As listObject
  Dim loName As String
  Dim adminName As String
  Dim sortRng As Range
  Dim nbRows As Long
  Dim LoRng As Range
  Dim listVars As BetterArray
  Dim varName As String
  Dim endCounter As Long

  Set sh = Wksh()

  'Get the list of varNames concerned
  Set listVars = ListOfExistingVarNames(inputVarName)

  'Number of admin tables to update
  endCounter = IIf(isHF, 1, 4)

  For counter = listVars.LowerBound To listVars.UpperBound

    varName = listVars.Item(counter)
    Set rng = inputRng

    For adminCounter = 1 To endCounter
        nbRows = rng.Cells.Count

        loName = IIf(isHF, CurrentLoName(varName, HF), CurrentLoName(varName, ADMIN & adminCounter))
        adminName = IIf(isHF, HF, ADMIN & adminCounter)

        Set Lo = sh.ListObjects(loName)

        'First cell where to paste the values of the range
        Set cellRng = sh.Range("RNG_PastingCol")
        Set pasteRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))

        'copy values to the pasting Rng
        pasteRng.Value = rng.Value

        On Error Resume Next
        'Remove duplicates to the pasting Rng
        pasteRng.RemoveDuplicates Columns:=1, Header:=xlNo

        On Error GoTo 0

        'Resize the pasting Range
        nbRows = sh.Cells(sh.Rows.Count, pasteRng.Column).End(xlUp).Row
        Set pasteRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))

        'Get the range for the list Object column
        Set cellRng = Lo.Range.Cells(2, 1)

        'First clear the Ranges of the listObject (Databodyrange could be nothing)
        On Error Resume Next
          Lo.ListColumns(1).DataBodyRange.Cells.Clear
        On Error GoTo 0

        Set LoRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))
        LoRng.Value = pasteRng.Value

        'Clear the pasting range
        pasteRng.Cells.Clear

        'Resize the listObject (including the Attack Rate)
        Set LoRng = sh.Range(cellRng.Cells(0, 1), cellRng.Cells(nbRows, 4))
        Lo.Resize LoRng

        'Sort on the formula column
        Set sortRng = Lo.ListColumns("formula_" & adminName).DataBodyRange

        'Sort the Calculate range for formulas
        sortRng.calculate

        On Error Resume Next
          Lo.ListColumns(1).DataBodyRange.Sort key1:=sortRng, Order1:=xlDescending
        On Error GoTo 0

        Set rng = rng.Offset(, 1)
    Next
  Next

End Sub


'@Description("Sort the tables on either attack rate or values")
Private Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)

  'onAR is to know if it is needed to sort on attack rate or values. Default
  'is values. Keep in mind we sort only on "geo" tables because health facilities
  'don't have a linked population. No sorting is done in "hf" tables.

  Dim sortRng As Range
  Dim LoRng As Range
  Dim Lo As listObject
  Dim sh As Worksheet
  Dim adminCounter As Long
  Dim sortRngName As String
  Dim varName As String
  Dim listRng As Range
  Dim cellRng As Range


  Set sh = Wksh()

  'Find the corresponding variable name of the table id
   Set listRng = sh.ListObjects("listofgeovars").Range

   'Look for the table in list range
   Set cellRng = listRng.Find(What:=tabId, lookAt:=xlPart, MatchCase:=True)

   'If you can't find the table, something went wrong
   If cellRng Is Nothing Then Exit Sub

   'Table ids are unique, so there is one variable name per table id
   '(variable name as combination of "dictionary variable name" and tabId)
   varName = cellRng.Value

  For adminCounter = 1 To 4

    Set Lo = sh.ListObjects(CurrentLoName(varName, ADMIN & adminCounter))
    If Lo.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(Lo.ListColumns(1).DataBodyRange) Then Exit Sub

    'Define the ranges and sort on them
    Set LoRng = Lo.DataBodyRange
    sortRngName = IIf(onAR, "attack_rate_", "formula_")
    sortRngName = sortRngName & ADMIN & adminCounter
    Set sortRng = Lo.ListColumns(sortRngName).Range

    On Error Resume Next
      Lo.ListColumns(1).Range.Sort key1:=sortRng, Order1:=xlDescending, Header:=xlYes
    On Error GoTo 0
  Next
End Sub


Private Property Get TopGeoValue(ByVal adminName As String, ByVal orderValue As Integer, ByVal inputVarName As String, _
                                 Optional ByVal tabId As String = vbNullString, Optional ByVal concatValue As Byte = 1) As String
  Dim sh As Worksheet
  Dim Lo As listObject
  Dim loName As String
  Dim varName As String
  Dim rng As Range

  Set sh = Wksh()

  varName = inputVarName & "_" & tabId
  loName = CurrentLoName(varName, adminName)

  If Not LoExists(loName) Then Exit Property

  Set Lo = sh.ListObjects(loName)

  If concatValue = 1 Then
    Set rng = Lo.ListColumns(1).DataBodyRange()
  Else
    Set rng = Lo.ListColumns(3).DataBodyRange()
  End If

  TopGeoValue = CStr(rng.Cells(orderValue, 1).Value)

End Property

Private Property Get TopHFValue(ByVal orderValue As Integer, ByVal inputVarName As String, Optional ByVal tabId As String = vbNullString) As String
  Dim sh As Worksheet
  Dim Lo As listObject
  Dim loName As String
  Dim varName As String
  Dim rng As Range
  Set sh = Wksh()

  varName = inputVarName & "_" & tabId
  loName = CurrentLoName(varName, HF)

  If Not LoExists(loName) Then Exit Property
  Set Lo = sh.ListObjects(loName)
  Set rng = Lo.ListColumns(1).DataBodyRange()

  TopHFValue = rng.Cells(orderValue, 1).Value
End Property


Private Sub CheckRequirements(ByVal sh As Worksheet)

  If sh Is Nothing Then ThrowError "Worksheet not found, please provide a worksheet"
  If sh.Name <> SPATIALSHEET Then ThrowError "The worksheet name is not " & SPATIALSHEET

End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Private Property Get ILLSpatial_Exists(ByVal varName As String) As Boolean
    ILLSpatial_Exists = Exists(varName)
End Property

Private Sub ILLSpatial_Update()
  Update
End Sub


Private Property Get ILLSpatial_TopGeoValue(ByVal adminName As String, ByVal levelValue As Integer, ByVal varName As String, _
  Optional ByVal tabId As String = vbNullString, Optional ByVal concatValue As Byte = 1) As String
  ILLSpatial_TopGeoValue = TopGeoValue(adminName, levelValue, varName, tabId, concatValue)
End Property


Private Property Get ILLSpatial_TopHFValue(ByVal orderValue As Integer, ByVal inputVarName As String, Optional ByVal tabId As String = vbNullString) As String
  ILLSpatial_TopHFValue = TopHFValue(orderValue, inputVarName, tabId)
End Property

Private Sub ILLSpatial_Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)
  Sort tabId, onAR
End Sub
