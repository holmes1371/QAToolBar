Option Explicit
Option Compare Text

Public Sub SheetFixIngestF(Optional buttonIndicator As Variant)
' Developed DTCC 3APR2017
' Nicholas Lopez nlopez@DTCC.com
' edited by Tom Holmes tholmes@dtcc.com
' Googled meaning of ";@" by Frank Castillo fcastilloandino@dtcc.com
    
    Dim startCell As Object: Set startCell = ActiveCell
    
    trimmer

    'Selects the cell on the first row and first column
    Range("A1").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"

    'Set formatting
    Application.FindFormat.NumberFormat = "m/d/yyyy"
    Application.ReplaceFormat.NumberFormat = "yyyy-mm-dd;@"

    'Find and replace date formatting based on above defined formatting
    Cells.Replace what:="", Replacement:="", LookAt:=xlWhole, SearchOrder:= _
    xlByRows, SearchFormat:=True, ReplaceFormat:=True

    'Clear formatting
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear

    'Finds and replaces case for TRUE boolean values
    Cells.Replace what:="TRUE", Replacement:="'true", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Finds and replaces case for FALSE boolean values
    Cells.Replace what:="FALSE", Replacement:="'false", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Reset match case and entire contents
    Cells.Replace what:="", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Clear Formatting
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
    
    If IsMissing(buttonIndicator) = False Then
        buildBitBucketLinks
    End If
    
    Columns.AutoFit
    startCell.Activate
    
End Sub

Sub buildBitBucketLinks()
  
  setHeaderVals (1)
  addTheLinks ("input file")
  addTheLinks ("expected file")


 'https://code.dtcc.com:8443/projects/GTR/repos/test-automation-data/raw/esma/regression/csv_inputs/297736_INPUT_REGRESSION.csv
   
End Sub
Sub addTheLinks(headerName As String)

    Dim path            As String:  path = "https://code.dtcc.com:8443/projects/GTR/repos/test-automation-data/raw/esma/regression/"
    Dim testFileType    As String
    Dim thisCol         As Integer

  thisCol = headerSearch(headerName)
  
  If Cells(headerRow, thisCol).Value <> headerName Then Exit Sub

  Cells(headerRow + 1, thisCol).Activate
  
  While ActiveCell.Value <> Empty
    With ActiveCell
        If .Value Like ("*INPUT*") And .Value Like ("*.csv") Then
            testFileType = "csv_inputs"
        ElseIf .Value Like ("*EXPECT*") And .Value Like ("*.csv") Then
            testFileType = "csv_expects"
        ElseIf .Value Like ("*INPUT*") And .Value Like ("*.xml") Then
            testFileType = "fpml_inputs"
        ElseIf .Value Like ("*EXPECT*") And .Value Like ("*.xml") Then
            testFileType = "csv_expects"
        End If
       ActiveSheet.Hyperlinks.Add Cells(.row, .Column), path & testFileType & "/" & .Value
       .Offset(1, 0).Select
    End With
  Wend

End Sub

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
