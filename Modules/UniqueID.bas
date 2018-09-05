Option Compare Text
Option Explicit

Public mylastcell
Public gtxString As String
Public utiMode As String
Public thisLastRow As Integer

Public Sub gtxValue_onChange(control As IRibbonControl, Text As String)
    gtxString = Text
End Sub

Public Function autoHeaderUniquinizerIngestF(Optional buttonIndicator As Variant)

' Forked by Tom Holmes based on original code by Delfino Ballesteros
          
    Dim i               As Integer
    Dim rCount          As Integer
    Dim startCell       As Object: Set startCell = ActiveCell
    
    SheetFixIngestF
    thisLastRow = getLastRecord(ActiveSheet.Name)
        
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    abortIt = False
    endIt = False

      
    findTradeIdField
    rCount = getRCount
    
    'sets the new tradeID
    For i = 0 To numOfTrades
        ActiveCell.Value = formatTradeId(rCount, ActiveCell.row)
        rCount = rCount + 1
        ActiveCell.Offset(1, 0).Select
        If spillOverCheck = True Then Exit For ' prevents extra ID's due to hidden characters
    Next i

    'does exit trade check.
    exitCheck
    'copies UTI to USI in CORE templates
    usiCheck
    
    'brings the active cell to the bottom of the TradeID column and autoFit's all columns
    findTradeIdField
    
    If IsMissing(buttonIndicator) = False Then Application.ScreenUpdating = True
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    Columns.AutoFit
  
End Function
Function spillOverCheck()

'this function prevents a possible error of creating extra UTI ID's
'due to hidden characters in an excel spreadsheet that cannot be deleted.

    If Application.WorksheetFunction.CountA(Rows(ActiveCell.row)) <= 3 Then
        spillOverCheck = True
    End If
    

End Function

Function getAssClass(currentRow)
 'creating Asset Class portion of the UTI / Trade ID
 
    Dim thisCell       As Object: Set thisCell = ActiveCell       'setting start position
    Dim pacColumn      As Integer

    findAssetClass
    pacColumn = ActiveCell.Column
    thisCell.Select                 'returning to active cell after getting the PAC column
    
    'AssetClass  for Harmonized, CORE and EU Lite abbreviations:
    If Trim(Cells(currentRow, pacColumn).Value) = "ForeignExchange" Or _
        Trim(Cells(currentRow, pacColumn).Value) = "FX" Then
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
    
    Dim tradeid As String
    
    If endIt = True Then Exit Function

    tradeid = getSuffix(count, currentRow)
    formatTradeId = tradeid
    
End Function
Function getTestNumber()
    On Error GoTo notNumber
    If Trim(Len(Cells(ActiveCell.row, 1).Value)) = 6 Then
        getTestNumber = Int(Cells(ActiveCell.row, 1).Value)
    End If
    Exit Function
    
notNumber:
    getTestNumber = ""

End Function
Function getSuffix(count, currentRow)
    
    Dim counter As String
    Dim newFour As Integer
    Dim dt As String
    Dim harn As String
    
    harn = "HARNESS_AUTO_"
    dt = todaysDate
    newFour = count + 1                                                     'Adds 1 to the current count
    counter = Format(newFour, "0000")
    
    If getTestNumber = Empty Then
        If utiMode = "auto" Then
            If gtxString = Empty Then
                getSuffix = harn & getAssClass(currentRow) & "_" & counter
            Else
                getSuffix = harn & UCase(specialStrip(gtxString)) & "_" & getAssClass(currentRow) & "_" & counter
            End If
        End If
        If utiMode = "manual" Then
            If gtxString = Empty Then
                getSuffix = "MANUAL_" & getAssClass(currentRow) & "_" & dt & "_" & counter
            Else
                getSuffix = UCase(specialStrip(gtxString)) & "_" & dt & "_" & getAssClass(currentRow) & "_UTI" & counter
            End If
        End If
    Else
        getSuffix = harn & getTestNumber & "_" & getAssClass(currentRow)
    End If
    
End Function
Function specialStrip(someString As String) As String

    Dim char                As Variant
    Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?,-,~,/,\,:"  'modify as needed
  
    For Each char In Split(SpecialCharacters, ",")
        someString = Replace(someString, char, "")
    Next

    specialStrip = someString

End Function

Function usiCheck()
    
    Dim thisCell    As Object:  Set thisCell = ActiveCell       'setting start position
    Dim idColumn    As Integer: idColumn = ActiveCell.Column
    Dim thisRow     As Integer
        
        
    findTradeIdField
    thisCell.Select                 'returning to active cell after getting the PAC column
 
    If find("USI Value") = True Then
        ActiveCell.Offset(1, 0).Select
        Set searchPosition = ActiveCell
        thisRow = ActiveCell.row
        Range(Cells(thisRow, idColumn), Cells(thisLastRow, idColumn)).Select
        Selection.Copy
        searchPosition.Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    End If
    
End Function

Function exitCheck()

        Dim lastRecord      As Integer: lastRecord = getLastRecord(ActiveSheet.Name)
        Dim idColumn        As Integer
        Dim actionColumn    As Integer
        Dim trade           As Variant
        Dim i, j, k         As Integer: j = 0
        Dim commentColumn() As Variant
        Dim exitTradeRow    As Integer
        Dim tradeName       As String
        
                
        ReDim commentColumn(lastRecord - headerRow - 1, 1) As Variant
        
        findTradeIdField
        idColumn = ActiveCell.Column
        find ("action")
                
        For i = headerRow + 1 To lastRecord
            commentColumn(j, 0) = Cells(i, 1).Value
            commentColumn(j, 1) = i
            j = j + 1
        Next i
                           
        actionColumn = ActiveCell.Column
        ActiveCell.Offset(1, 0).Select
        
        For i = headerRow To lastRecord
            If Trim(ActiveCell.Value) = "exit" Then
                exitTradeRow = ActiveCell.row
                tradeName = Cells(exitTradeRow, 1).Value
                For k = LBound(commentColumn) To UBound(commentColumn)
                    If commentColumn(k, 0) = tradeName And commentColumn(k, 1) <> exitTradeRow Then
                        If Cells(commentColumn(k, 1), actionColumn).Value = "new" Then
                            Cells(exitTradeRow, idColumn).Value = Cells(commentColumn(k, 1), idColumn).Value
                        End If
                    End If
                Next k
                Cells(exitTradeRow, actionColumn).Select
            End If
            ActiveCell.Offset(1, 0).Select
         Next i
    
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

    Dim runningCounter      As Integer
    Dim currCount           As Integer
    Dim returnCell          As Object: Set returnCell = ActiveCell
    Dim i                   As Integer
        
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
    returnCell.Activate
    
End Function
 
Function todaysDate() As String
    Dim dt As Date
    Dim tdate As String
    Dim fdate As Variant
    
    dt = Date
    fdate = Format(dt, "yyyymmdd")          'Formats date to yyyymmdd
    tdate = CStr(fdate)                     'Converts Date to string
    todaysDate = tdate                      'Saves converted date string to function return
End Function

Function getLastFour() As Integer

    Dim tradeid     As String
    Dim counter     As Integer

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
    
    numOfTrades = thisLastRow - headerRow
End Function

Function lastRow() As Integer 'action is a mandatory field for all templates

    Dim returnCell  As Object: Set returnCell = ActiveCell


    find ("Action")
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    lastRow = ActiveCell.row
    returnCell.Activate

End Function
