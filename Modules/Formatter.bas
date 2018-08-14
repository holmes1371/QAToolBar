Public Function SheetFixIngestF()
' Developed DTCC 3APR2017
' Nicholas Lopez nlopez@DTCC.com
' edited by Tom Holmes tholmes@dtcc.com
' Googled meaning of ";@" by Frank Castillo fcastilloandino@dtcc.com
    
    startcell = ActiveCell.Address
    Application.ScreenUpdating = False
    trimmer

    'makeBland 'removes any coloring or special formatting to the text

    'Selects the cell on the first row and first column
    Range("A1").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"

    'Clear formatting
    'Application.FindFormat.Clear
    'Application.ReplaceFormat.Clear

    'Set formatting
    Application.FindFormat.NumberFormat = "m/d/yyyy"
    Application.ReplaceFormat.NumberFormat = "yyyy-mm-dd;@"

    'Find and replace date formatting based on above defined formatting
    Cells.Replace What:="", Replacement:="", LookAt:=xlWhole, SearchOrder:= _
    xlByRows, SearchFormat:=True, ReplaceFormat:=True

    'Clear formatting
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear

    'Finds and replaces case for TRUE boolean values
    Cells.Replace What:="TRUE", Replacement:="'true", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Finds and replaces case for FALSE boolean values
    Cells.Replace What:="FALSE", Replacement:="'false", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Reset match case and entire contents
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Clear Formatting
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
    Application.ScreenUpdating = True
    
    Columns.AutoFit
    Range(startcell).Select
    
End Function

Function makeBland()

    ActiveSheet.Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
    Selection.Font.Underline = xlUnderlineStyleNone
    Selection.Font.Italic = False

End Function
