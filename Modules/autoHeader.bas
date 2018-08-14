Option Compare Text
Public oCode As String
Public Sub ocodeVal_onChange(control As IRibbonControl, text As String)
    oCode = text
End Sub

Public Function autoHeader2()

' Developed by Tom Holmes tholmes@dtcc.com
' Emotional counselor: Frank
' formats header and trainer to be compliant with BigFish requirements
    
    startcell = ActiveCell.Address
    Application.ScreenUpdating = False
    
    Dim commentLocation
    Dim foundIt
    Dim count
    
    foundIt = False
    count = 0
        
    trimmer
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
        
    'prevents formatting if more than one cell in row 1 is "*comment"
    Range("A1").Select
    For i = 1 To mylastcell.Row
        If (ActiveCell.Value = "*comment") Or (ActiveCell.Value = "comment") Then
            count = count + 1
        End If
         ActiveCell.Offset(1, 0).Select
    Next i
    
    If count >= 2 Then
        Application.ScreenUpdating = True
        MsgBox "More than one 'comment' found in row 1. Please verify", vbInformation, "WARNING!"
        endIt = True
        Range("A1").Select
        Exit Function
    End If
    
    trimmer
    
    Range("A1").Select
    For i = 1 To mylastcell.Row
        If ActiveCell.Value = "*comment" Then    'formatting check. Will only apply the header if "*Comment" is found
        commentLocation = ActiveCell.Row
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
        
        bigFishHeader = "*" & getOcode & bigFishDate & Environ$("Username") & "@dtcc.com"
        
        Cells(1, 1).Value = bigFishHeader
        
        Range("A1").Select
        
        While ActiveCell.Value <> Empty
            ActiveCell.Offset(1, 0).Activate
            If ActiveCell.Value Like "*-END" And ActiveCell.Offset(1, 0).Value = Empty Then GoTo bailOut
        Wend

bailOut:
        ActiveCell.Value = "*" & getOcode & "-END"
        
         
'        For i = 2 To mylastcell.Column
'            ActiveCell.Offset(0, 1).Select
'            ActiveCell.Value = i
'        Next i
        
        Range(startcell).Select
    Else
        Range("A1").Select
        MsgBox "No '*Comment' box found in Row 1. Please verify this is the correct sheet you want to format", vbInformation, "WARNING!"
        Application.ScreenUpdating = True
        Range("A1").Select
        endIt = True
        Exit Function
    End If
    SheetFixIngestF
    Range(startcell).Select
    
End Function

Function bigFishDate() As String
    Dim dt As Date
    Dim tdate As String
    dt = Date
    fdate = Format(dt, "yyyy-mm-dd")          'Formats date to yyyy-mm-dd
    tdate = CStr(fdate)                       'Converts Date to string
    bigFishDate = tdate                        'Saves converted date string to function return
End Function
Function getOcode()
    
    If oCode <> Empty Then
        getOcode = oCode
    Else
        getOcode = "XXXX"
    End If
        

End Function


