Public searchPosition As Object
Public foundOne As Boolean
Public usiActive As Boolean

Option Explicit
Option Compare Text

Public Function findAssetClass()
    Cells(1, 1).Select
    foundOne = False
    
    findIt ("Primary Asset Class")
    findIt ("Asset Class")
            
   If foundOne = False Then
        MsgBox "Could not find Asset Class field", vbInformation, "WARNING!"
        'Application.ScreenUpdating = True
        Cells(1, 1).Select
        endIt = True
    resetSearchParameters
  
       Exit Function
   End If
    
    Cells(searchPosition.row, searchPosition.column).Select
 
End Function

Public Sub findID(Optional buttonIndicator As Variant)

    Dim i As Integer
    Columns.AutoFit

    setHeaderVals
    
    For i = LBound(csvHeader) To UBound(csvHeader)
    
        If csvHeader(i) = "UTI" Then
            Set searchPosition = Cells(headerRow, i + 1)
            Exit For
        ElseIf csvHeader(i) = "UTI ID" Then
            Set searchPosition = Cells(headerRow, i + 1)
            Exit For
        ElseIf csvHeader(i) = "Trade ID" Then
            Set searchPosition = Cells(headerRow, i + 1)
            Exit For
        End If
    Next i
    
    'Application.ScreenUpdating = True
    
    If i > UBound(csvHeader) Then
        MsgBox "Could not find UTI ID/UTI or Trade ID field", vbInformation, "WARNING!"
        Cells(1, 1).Select
    Else
        If IsMissing(buttonIndicator) = False Then Application.ScreenUpdating = True
        Cells(searchPosition.row, searchPosition.column).Activate
    End If
    
    resetSearchParameters
    
End Sub
Public Function findIt(findThis)
Cells(1, 1).Select
usiActive = False

On Error GoTo handler
        Cells.find(what:=findThis, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
        Set searchPosition = ActiveCell
        foundOne = True
        
        If findThis = "USI Value" Then
            foundOne = False
            usiActive = True
        End If
        
        Exit Function
handler:
    Exit Function
End Function

Private Function findUSI()

On Error GoTo handler
Cells(1, 1).Select
        Cells.find(what:="USI Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
        Set searchPosition = ActiveCell
        findUSI = searchPosition.column
        
        Exit Function
        
handler:
    Exit Function

End Function

Public Function find(this)
Cells(1, 1).Select
On Error GoTo niceExit:
        Cells(1, 1).Activate
        Cells.find(what:=this, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        find = True
        foundOne = True
        Exit Function
        
niceExit:
    find = False
    Exit Function
      
End Function

Public Function precheck()
    Dim foundComment    As Boolean
    Dim displayMessage  As Boolean
    Dim i               As Integer
    Dim message         As String
    
    'Application.ScreenUpdating = False
    
    foundComment = False
    precheck = True
    displayMessage = False
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)

   message = "Could not find: " & vbCrLf
    
    If find("UTI") = False And find("UTI ID") = False And find("Trade ID") = False Then
        message = message & "UTI/Trade ID" & vbCrLf
        displayMessage = True
    End If
   
    If find("action") = False Then
        message = message & "Action" & vbCrLf
        displayMessage = True
    End If
    
    Range("A1").Select
    For i = 1 To mylastcell.row
        If ActiveCell.Value = "*comment" Then
            foundComment = True
            Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next i

    If foundComment = False Then
        message = message & "*Comment" & vbCrLf
        displayMessage = True
    End If
    
    If find("asset class") = False And find("primary asset class") = False Then
        message = message & "Asset Class" & vbCrLf
        displayMessage = True
    End If

   If displayMessage = True Then
        MsgBox message, vbInformation, "WARNING!"
        'Application.ScreenUpdating = True
        Cells(1, 1).Select
        precheck = False
        resetSearchParameters
       Exit Function
    End If
    
End Function

