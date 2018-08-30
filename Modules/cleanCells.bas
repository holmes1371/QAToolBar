Option Explicit
Public Function trimmer()

    Dim startCell   As Object: Set startCell = ActiveCell
    Dim thisSheet   As Range
    Dim cel         As Range
        
    'Application.ScreenUpdating = False
    
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    
    Set thisSheet = Range(Cells(1, 1).Address(), Cells(mylastcell.row, mylastcell.column).Address())

    
        For Each cel In thisSheet.Cells
            On Error GoTo skipIt
            With cel
                If Trim(.Value) = Empty Then GoTo skipIt
                .Value = Trim(.Value)
            End With
skipIt:
        Next cel
        
        
End Function

