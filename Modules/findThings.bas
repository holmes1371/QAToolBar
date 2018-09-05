Public searchPosition As Object
Public foundOne As Boolean
Public usiActive As Boolean

Option Explicit
Option Compare Text

Public Function findAssetClass()
    
    Dim i As Integer
    
    For i = LBound(csvHeader) To UBound(csvHeader)
    
        If csvHeader(i) = "Primary Asset Class" Then
            Exit For
        ElseIf csvHeader(i) = "Asset Class" Then
            Exit For
        End If
        
    Next i
          
    Cells(headerRow, i + 1).Activate
 
End Function

Public Sub findID(Optional buttonIndicator As Variant)

    Dim i As Integer
    Columns.AutoFit
    
    For i = LBound(csvHeader) To UBound(csvHeader)
    
        If csvHeader(i) = "UTI" Then
            Exit For
        ElseIf csvHeader(i) = "UTI ID" Then
            Exit For
        ElseIf csvHeader(i) = "Trade ID" Then
            Exit For
        End If
    Next i
        
    If i > UBound(csvHeader) Then
        MsgBox "Could not find UTI ID/UTI or Trade ID field", vbInformation, "WARNING!"
        Cells(1, 1).Select
    Else
        If IsMissing(buttonIndicator) = False Then Application.ScreenUpdating = True
        Cells(headerRow, i + 1).Activate
    End If
    
End Sub

Public Function find(this As String) As Boolean
    Dim i As Integer
    
    find = False
    
    For i = LBound(csvHeader) To UBound(csvHeader)
        If csvHeader(i) = this Then
            find = True
            foundOne = True
            Exit For
        End If
    Next i
      
      Cells(headerRow, i + 1).Activate
      
End Function

Public Function preCheck() As Boolean

    Dim foundComment    As Boolean
    Dim displayMessage  As Boolean
    Dim i               As Integer
    Dim message         As String
    Dim startCell       As Object: Set startCell = ActiveCell
    
   
    foundComment = False
    preCheck = True
    displayMessage = False
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)

    message = "Could not find: " & vbCrLf
   
    Range("A1").Select
    For i = 1 To mylastcell.row
        If ActiveCell.Value = "*comment" Then
            foundComment = True
            Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next i

    
    If foundComment = False Then
        message = message & "Please ensure '*Comment' is in column 1 of your header row." & vbCrLf
        displayMessage = True
        GoTo endMessage
    End If
        
    setHeaderVals
    
    If find("action") = False Then
        message = message & "Action column" & vbCrLf
        displayMessage = True
        GoTo endMessage
    End If
       
    If find("asset class") = False And find("primary asset class") = False Then
        message = message & "Asset Class column" & vbCrLf
        displayMessage = True
        GoTo endMessage
    End If
    
    If find("UTI") = False And find("UTI ID") = False And find("Trade ID") = False Then
        message = message & "UTI/Trade ID column" & vbCrLf
        displayMessage = True
        GoTo endMessage
    End If
   
    
endMessage:
    startCell.Activate
    
    If displayMessage = True Then
        MsgBox message, vbInformation, "WARNING!"
        Application.ScreenUpdating = True
        Cells(1, 1).Select
        preCheck = False
        resetSearchParameters
       Exit Function
    End If
    
    
End Function

