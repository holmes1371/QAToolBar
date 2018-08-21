Public Function trimmer()
    Application.ScreenUpdating = False
    Set startCell = ActiveCell

    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    Dim thisSheet As Range
    Set thisSheet = Range(Cells(1, 1).Address(), Cells(mylastcell.Row, mylastcell.Column).Address())
    
    Dim cel As Range
    
        For Each cel In thisSheet.Cells
            On Error GoTo skipIt
            With cel
                If Trim(.Value) = Empty Then GoTo skipIt
                .Value = Trim(.Value)
            End With
skipIt:
        Next cel
End Function

