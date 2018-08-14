
Public Function trimmer(scope) 'scope determines how much of the worksheet will be trimmed
                               'scope = 1 will trim ALL the active cells in the worksheet
                               'scope = 2 will only trim the header column names

Application.ScreenUpdating = False
Set startcell = ActiveCell

If scope = "1" Then
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    Dim thisSheet As Range
    Set thisSheet = Range(Cells(1, 1).Address(), Cells(mylastcell.Row, mylastcell.Column).Address())
    Dim cel As Range
    
        For Each cel In thisSheet.Cells
            With cel
                If Trim(.Value) = Empty Then GoTo skipIt
                .Value = Trim(.Value)
            End With
skipIt:
        Next cel
End If
    
If scope = "2" Then
    On Error GoTo niceExit:
    Cells.find(What:="*comment", After:=ActiveCell, LookIn:=xlValues, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False).Activate
    
    Do While ActiveCell.Value <> Empty
        ActiveCell.Value = Trim(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Activate
    Loop
End If
    
niceExit:
    startcell.Activate
    
End Function
