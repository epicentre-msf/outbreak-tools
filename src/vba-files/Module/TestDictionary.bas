




@TestMethod
Private Sub TestPreparation()

    On Error GoTo PreparationFailed
    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim randRng As Range
    Dim endCol As Long

    Set dictWksh = dictObject.Wksh

    If Not dictObject.Prepared Then
        With dictWksh
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
            If Not dictObject.ColumnExists("randnumber") Then
                .Cells(1, endCol) = "randnumber"
                .Cells(2, endCol).Formula = "= RAND()"
                Set randRng = dictObject.DataRange("randnumber")
                .Cells(2, endCol).AutoFill randRng, Type:=xlFillValues
            End If
            Set dictRng = dictObject.DataRange
            Set randRng = dictObject.DataRange("randnumber")
            dictRng.Sort key1:=randRng
            dictObject.Prepare
        End With
    End If

    Assert.IsTrue dictObject.Prepared, "dictionary not prepared for buildlist"
    Exit Sub

PreparationFailed:
    Assert.Fail "Prepared Failed: #" & Err.Number & " : " & Err.Description
End Sub