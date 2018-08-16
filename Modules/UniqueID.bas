Option Compare Text
Public mylastcell
Public gtxString As String
Public Sub gtxValue_onChange(control As IRibbonControl, Text As String)
    gtxString = Text
End Sub

Public Function autoHeaderUniquinizerIngestF()
' Developed by Delfino Ballesteros and Tom Holmes

    SheetFixIngestF
    Application.ScreenUpdating = False
    
    If precheck = False Then Exit Function  'precheck for required fields.
    Set startcell = ActiveCell
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    abortIt = False
    endIt = False
      
    Dim i As Integer
    Dim rCount As Integer
      
    findTradeIdField
    rCount = getRCount
    
    'sets the new tradeID
    For i = 0 To numOfTrades
        ActiveCell.Value = formatTradeId(rCount, ActiveCell.Row)
        rCount = getLastFour
        ActiveCell.Offset(1, 0).Select
        If spillOverCheck = True Then Exit For ' prevents extra ID's due to hidden characters
    Next i

    'does exit trade check.
    exitCheck
    'copies UTI to USI in CORE templates
    usiCheck
    
    'brings the active cell to the bottom of the TradeID column and autoFit's all columns
    findTradeIdField
    Application.ScreenUpdating = True
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    Columns.AutoFit
  
   
End Function
Function spillOverCheck()

'this function prevents a possible error of creating extra UTI ID's
'due to hidden characters in an excel spreadsheet that cannot be deleted.

    If Application.WorksheetFunction.CountA(Rows(ActiveCell.Row)) <= 3 Then
        spillOverCheck = True
    End If
    

End Function

Function getAssClass(currentRow)
 'creating Asset Class portion of the UTI / Trade ID
 
    Set ThisCell = ActiveCell       'setting start position
    findAssetClass
    pacColumn = ActiveCell.Column
    ThisCell.Select                 'returning to active cell after getting the PAC column
    
    'AssetClass  for Harmonized, CORE and EU Lite abbreviations:
    If Trim(Cells(currentRow, pacColumn).Value) = "ForeignExchange" Then
        getAssClass = "FX"
    
    ElseIf Trim(Cells(currentRow, pacColumn).Value) = "CU" Then
        getAssClass = "CU"
    
    ElseIf Trim(Cells(currentRow, pacColumn).Value) = "InterestRate" Or _
        Trim(Cells(currentRow, pacColumn).Value) = "IR" Then
        getAssClass = "IR"
    
    ElseIf Trim(Cells(currentRow, pacColumn).Value) = "Commodity" Or _
        Trim(Cells(currentRow, pacColumn).Value) = "CO" Then
        getAssClass = "CO"
    
    ElseIf Trim(Cells(currentRow, pacColumn).Value) = "Equity" Or _
        Trim(Cells(currentRow, pacColumn).Value) = "EQ" Then
        getAssClass = "EQ"
    
    ElseIf Trim(Cells(currentRow, pacColumn).Value) = "Credit" Or _
        Trim(Cells(currentRow, pacColumn).Value) = "CR" Then
        getAssClass = "CR"
        
    Else
        getAssClass = "??"  'Asset Class not provided or recognized
    End If

    
End Function

Function formatTradeId(count As Integer, currentRow) As String
    
    If endIt = True Then Exit Function
    Dim harn As String
    
    harn = "HARNESS_AUTO_"
    tradeid = harn & getSuffix(count, currentRow)
    formatTradeId = tradeid
    
End Function
Function getTestNumber()
    On Error GoTo notNumber
    If Trim(Len(Cells(ActiveCell.Row, 1).Value)) = 6 Then
        getTestNumber = Int(Cells(ActiveCell.Row, 1).Value)
    End If
    Exit Function
    
notNumber:
    getTestNumber = ""

End Function
Function getSuffix(count, currentRow)
    
    Dim counter As String
    Dim newFour As Integer
    Dim dt As String
    
    dt = todaysDate
    newFour = count + 1                                                     'Adds 1 to the current count
    counter = Format(newFour, "0000")
       
    If getTestNumber = Empty Then
        If gtxString = Empty Then
            getSuffix = getAssClass(currentRow) & "_" & dt & "_" & counter
        Else
            getSuffix = gtxString & "_" & getAssClass(currentRow) & "_" & dt & "_" & counter
        End If
    Else
        getSuffix = getTestNumber & "_" & getAssClass(currentRow)
    End If
    
End Function

Function usiCheck()
    
    Set ThisCell = ActiveCell       'setting start position
    findTradeIdField
    idcolumn = ActiveCell.Column
    ThisCell.Select                 'returning to active cell after getting the PAC column

    findIt ("USI Value")
    
    If usiActive = True Then
        ActiveCell.Offset(1, 0).Select
        Set searchPosition = ActiveCell
        thisRow = ActiveCell.Row
        Range(Cells(thisRow, idcolumn), Cells(lastRow, idcolumn)).Select
        Selection.Copy
        searchPosition.Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    End If
    
End Function

Function exitCheck()

        findTradeIdField
        idcolumn = ActiveCell.Column
       
        Range("A1").Select
        
On Error GoTo handleErrorAction
        
        Cells.find(What:="action", After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        
        actioncolumn = ActiveCell.Column
        headerRow = ActiveCell.Row
        ActiveCell.Offset(1, 0).Select
        
        For i = headerRow To mylastcell.Row
            If Trim(ActiveCell.Value) = "exit" Then
                exitTradeRow = ActiveCell.Row
                tradeName = Cells(exitTradeRow, 1).Value
                Columns("A:A").Select

                Selection.find(What:=tradeName, After:=ActiveCell, _
                LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
                
                If ActiveCell.Row = exitTradeRow Then
                    Selection.find(What:=tradeName, After:=ActiveCell, _
                    LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
                End If
 
                If ActiveCell.Value = tradeName And Trim(Cells(ActiveCell.Row, actioncolumn).Value) = _
                 "new" Then
                    ActiveCell.Offset(0, idcolumn - 1).Select
                    tradeid = ActiveCell.Value
                    Cells(exitTradeRow, idcolumn).Value = tradeid
                End If
            
                Cells(exitTradeRow, actioncolumn).Select
            End If
            ActiveCell.Offset(1, 0).Select
         Next i
      
    resetSearchParameters
        
    Exit Function
    
handleErrorAction:
     MsgBox "No 'Action' field was found", vbInformation, "WARNING!"
    'Reset match case and entire contents
     Cells.Replace What:="", Replacement:="", LookAt:=xlPart, _
     SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
     ReplaceFormat:=False
     endIt = True
     Exit Function
     

End Function

Public Function findTradeIdField()
    Cells(1, 1).Activate
    findID
    If foundOne = True Then
        ActiveCell.Offset(1, 0).Select
    Else
        endIt = True
    End If
        
End Function

Function getRCount()
    Dim runningCounter As Integer
    Dim currCount As Integer
    Set returncell = ActiveCell
    For i = 0 To numOfTrades
        If ActiveCell.Value <> Empty Then
            currCount = getLastFour
            If currCount > runningCounter Then
                runningCounter = currCount
                ActiveCell.Offset(1, 0).Select
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next i
    getRCount = runningCounter
    returncell.Activate
End Function
 
Function todaysDate() As String
    Dim dt As Date
    Dim tdate As String
    dt = Date
    fdate = Format(dt, "yyyymmdd")          'Formats date to yyyymmdd
    tdate = CStr(fdate)                     'Converts Date to string
    todaysDate = tdate                      'Saves converted date string to function return
End Function

Function getLastFour()

On Error GoTo changeFormat:
    tradeid = ActiveCell.Value              'Save the value to the variable tradeId
    counter = Right(tradeid, 4)             'Extract the last four digits
    getLastFour = CInt(counter)             'Convert string to integer and save
    Exit Function
    
changeFormat:
    getLastFour = "0000"
    
End Function

Function headerCount() As Integer
    'Determines the number of header columns by counting the first 5 rows containing
    'an asterisk(*) in column 1.
    headerCount = Application.WorksheetFunction.CountIf(Range("A1:A5"), "~**")
End Function


Function numOfTrades() As Integer
    'Returns the actual number of trades by subtracting the header rows from the
    'total rows and returning the difference
    numOfTrades = lastRow - 1 - headerCount
    'MsgBox numOfTrades
End Function
Function lastRow()
   Set returncell = ActiveCell
    find ("Action")
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    lastRow = ActiveCell.Row
    returncell.Activate
    'MsgBox lastRow
End Function


