Sub ReplaceAll()

Dim RValue As String
Dim FValue As String
Dim Index As Long
    Index = 1
        For i = 1 To 893
            FValue = Cells(Index, 2).Value
            RValue = Cells(Index, 1).Value
            ActiveSheet.Cells.Replace What:=FValue, Replacement:=RValue, _
                LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                SearchFormat:=False, ReplaceFormat:=False
            Index = Index + 1
    Next
End Sub