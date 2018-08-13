Option Compare Text
Public searchPosition
Public foundOne
Public usiActive

Public Function findAssetClass()
    
    foundOne = False
    
    findIt ("Primary Asset Class")
    findIt ("Asset Class")
            
   If foundOne = False Then
        MsgBox "Could not find Asset Class field", vbInformation, "WARNING!"
        Application.ScreenUpdating = True
        Cells(1, 1).Select
        endIt = True
    resetSearchParameters
  
       Exit Function
   End If
    
    Cells(searchPosition.Row, searchPosition.Column).Select
 
End Function

Public Sub findID()

foundOne = False
Columns.AutoFit

    findIt ("UTI")
    findIt ("UTI ID")
    findIt ("Trade ID")
        
   If foundOne = False Then
        MsgBox "Could not find UTI ID/UTI or Trade ID field", vbInformation, "WARNING!"
        Application.ScreenUpdating = True
        Cells(1, 1).Select
        
    resetSearchParameters
  
       Exit Sub
   End If
    
    Cells(searchPosition.Row, searchPosition.Column).Select

End Sub
Public Function findIt(findThis)

usiActive = False

On Error GoTo handler
        Cells.find(What:=findThis, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
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

        Cells.find(What:="USI Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
        Set searchPosition = ActiveCell
        findUSI = searchPosition.Column
        
        Exit Function
        
handler:
    Exit Function

End Function

Public Function find(this)

On Error GoTo niceExit:
        
        Cells.find(What:=this, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
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
foundComment = False
precheck = True
Dim message As String
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
    For i = 1 To mylastcell.Row
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
        Application.ScreenUpdating = True
        Cells(1, 1).Select
        precheck = False
        resetSearchParameters
       Exit Function
    End If
    
End Function

