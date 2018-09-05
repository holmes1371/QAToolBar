Public oCode As String

Option Compare Text
Option Explicit

Public Sub ocodeVal_onChange(control As IRibbonControl, Text As String)
    oCode = Text
End Sub

Public Sub autoHeader2()

' Developed by Tom Holmes tholmes@dtcc.com
' formats header and trailer to be compliant with BigFish requirements

    Dim i                   As Integer
    Dim commentLocation     As Integer
    Dim foundIt             As Boolean
    Dim count               As Integer
    Dim bigFishHeader       As String
    
    'Application.ScreenUpdating = False
 
    foundIt = False
    count = 0
        
    trimmer
    
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
        
    'prevents formatting if more than one cell in row 1 is "*comment"
    
    Range("A1").Select
    For i = 1 To mylastcell.row
        If (ActiveCell.Value = "*comment") Or (ActiveCell.Value = "comment") Then
            count = count + 1
        End If
         ActiveCell.Offset(1, 0).Select
    Next i
    
    If count >= 2 Then
        'Application.ScreenUpdating = True
        MsgBox "More than one 'comment' found in row 1. Please verify", vbInformation, "WARNING!"
        endIt = True
        Range("A1").Select
        Exit Sub
    End If
    
    trimmer
    
    Range("A1").Select
    For i = 1 To mylastcell.row
        If ActiveCell.Value = "*comment" Then    'formatting check. Will only apply the header if "*Comment" is found
        commentLocation = ActiveCell.row
        foundIt = True
        Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    If foundIt = True Then                       'will perform the formatting if "*Comment" was found
        
        If commentLocation <> 1 Then
            commentLocation = commentLocation - 1
            Rows("1:" & commentLocation).Select
            Selection.Delete Shift:=xlUp
        End If
            
        Range("A1").Select
        ActiveCell.Value = "*Comment"           'ensures comment is capitalized
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        bigFishHeader = "*" & getOcode & bigFishDate & "O" & Environ$("Username") & "@dtcc.com"
        
        Cells(1, 1).Value = bigFishHeader
        
        find ("action") 'uses action column (must be populated in all templates) to find the last record
        While ActiveCell.Value <> Empty
            ActiveCell.Offset(1, 0).Activate
        Wend

        Cells(ActiveCell.row, 1).Value = "*" & getOcode & "-END"
         
    Else
        Range("A1").Select
        MsgBox "No '*Comment' box found in Row 1. Please verify this is the correct sheet you want to format", _
        vbInformation, "WARNING!"
        endIt = True
    End If
    
    'Application.ScreenUpdating = True
    Range("A1").Select
    SheetFixIngestF
    
End Sub

Function bigFishDate() As String
    Dim dt As Date
    Dim tdate As String
    Dim fdate As String
    
    dt = Date
    fdate = Format(dt, "yyyy-mm-dd")          'Formats date to yyyy-mm-dd
    tdate = CStr(fdate)                       'Converts Date to string
    bigFishDate = tdate                       'Saves converted date string to function return
    
End Function
Function getOcode()
    
    If oCode <> Empty Then
        getOcode = oCode
    Else
        getOcode = "XXXX"
    End If
End Function


