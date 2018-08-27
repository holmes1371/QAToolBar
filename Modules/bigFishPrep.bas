Public headerRow
Public csvHeader
Public tempSheet
Public fileCount
Public prepMode As Boolean
Public newArr() As String
Option Compare Text

Sub splitFiles(control As IRibbonControl)
    
    setStartVals
    
    Set startCell = ActiveCell
    utiMode = "auto"
    
    setHeaderVals
    
    splitOptions = SplitSelector(csvHeader)
    
    If endIt = True Then
        refreshScreen
        Exit Sub
    End If
    
    setHeaderVals
    autoHeaderUniquinizerIngestF
   ' workingMessage
    ProgressBar.Show
    
    Application.ScreenUpdating = False
    
    startsheet = ActiveSheet.Name

    Call splitThisSheet(ActiveSheet.Name, splitOptions)
    
    verifyFinal (splitOptions)


    Unload ProgressBar
    'Unload Working
    refreshScreen
    startCell.Activate
    resetSearchParameters
    
    If fileCount = 0 Then
        MsgBox "No files have been created", vbInformation, "Complete"
    Else
        MsgBox fileCount & " files have been successfully created", vbInformation, "Complete"
        'open fileExplorer to the location where files are saved
        retVal = Shell("explorer.exe " & getpath, vbNormalFocus)
    End If
    
    prepMode = False
    
End Sub


Function getTotalNumber()

Dim totalNumber As Integer

    totalNumber = 1

    For k = LBound(splitOptions) To UBound(splitOptions)
        tempUnique = getUnique(splitOptions(k))
        totalNumber = totalNumber * (UBound(tempUnique) + 1)
    Next k
    
    getTotalNumber = totalNumber

End Function

Function splitThisSheet(sheetName As String, criteria)

Sheets(sheetName).Activate

    For Each opt In criteria

        tempUnique = getUnique(opt)
        
        If UBound(tempUnique) = 0 Then
            GoTo skipIt:
        Else
            For i = LBound(tempUnique) To UBound(tempUnique)
            tempSheet = tempSheet + 1
            createTempSheet (CStr(tempSheet))
            Call doCopy(sheetName, CStr(tempSheet), headerSearch(opt), tempUnique(i))
            Call splitThisSheet(CStr(tempSheet), criteria)
            Next i
        End If
    
skipIt:
    Next opt

End Function
Function verifyFinal(criteria)
Application.DisplayAlerts = False
    For i = 1 To tempSheet
    
        Sheets(CStr(i)).Activate
        
        For j = LBound(criteria) To UBound(criteria)
            isFinished = False
            tempUnique = getUnique(criteria(j))
            If UBound(tempUnique) <> 0 Then Exit For
            isFinished = True
skipIt:
        Next j

        If isFinished = True Then
            Call makeNewBook(criteria, CStr(i))
            fileCount = fileCount + 1
            
            With ProgressBar
                .text.Caption = (fileCount / getTotalNumber) & "% Complete"
                .Bar.Width = (fileCount / getTotalNumber) * 2
            End With
            DoEvents
            
            Sheets(CStr(i)).Delete
        Else
            Sheets(CStr(i)).Delete
        End If
    Next i

Application.DisplayAlerts = True
End Function
Function makeNewBook(criteria, sheetName)
'
    fileName = getNamePart(criteria, sheetName)

    Worksheets(sheetName).Copy
    
    With ActiveWorkbook
        .SaveAs fileName:=fileName, FileFormat:=xlCSV
         SheetFixIngestF
         Application.ScreenUpdating = False
        .Close SaveChanges:=True
    End With

End Function
Function getNamePart(criteria, sheetName)

    thispath = ActiveWorkbook.Path & "\"
    startsheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    testName = getUnique("*comment")
    testLength = UBound(testName)
    
    If testLength = 0 Then
        If testName(0) = "" Then testName(0) = "BLANK"
        testNumber = testName(0)
    Else
        testNumber = "MTC"
    End If
    
    Sheets(sheetName).Range("A" & getPasteRow(sheetName)).Value = getTrailer
    
    AssetClass = checkAssets(getUnique(Cells(headerRow, getAssClassCol).Value))
    
    Dim endPart As String
    
    For i = 0 To UBound(criteria)
        If criteria(i) = "Asset Class" Or criteria(i) = "Primary Asset Class" Then GoTo skipIt
        appendage = getUnique(criteria(i))
        If appendage(0) = "" Then appendage(0) = "blank"
        endPart = endPart & "_" & UCase(appendage(0))
skipIt:
    Next i
    
    If hasUTI(criteria) = True Then
        getNamePart = thispath & testNumber & "_INPUT" & "_ESMA" & endPart
    Else
        getNamePart = thispath & testNumber & "_INPUT" & assAbbr(AssetClass) & "_ESMA" & endPart
    End If
    
    Sheets(startsheet).Activate
    
End Function
Function hasUTI(criteria) As Boolean

     hasUTI = False
     For i = LBound(criteria) To UBound(criteria)
        If criteria(i) = "UTI" Or criteria(i) = "UTI ID" Or criteria(i) = "Trade ID" Then hasUTI = True
    Next i

End Function

Function doCopy(parentSheet, targetSheet As String, icol, criteria)
    startsheet = ActiveSheet.Name
    Sheets(parentSheet).Activate
    
        For j = headerRow + 1 To getLastRecord
            If CStr(Cells(j, icol).Value) = criteria Then
                Rows(j).Copy
                Sheets(targetSheet).Range("A" & getPasteRow(targetSheet)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False
            End If
        Next j
           
    Sheets(startsheet).Activate
    
End Function
Function createTempSheet(tempSheet As String)
    startsheet = ActiveSheet.Name
    Worksheets.Add
    ActiveSheet.Name = tempSheet
    Sheets(startsheet).Activate
    Rows("1:" & headerRow).Copy
    Sheets(tempSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(startsheet).Activate
End Function

Function setHeaderVals()

    Dim headerVal() As String, size As Integer, i As Integer
    
    Application.ScreenUpdating = False
    Set startCell = ActiveCell
    Cells(1, 1).Activate
    
    While ActiveCell.Value <> "*comment"
        ActiveCell.Offset(1, 0).Activate
    Wend
    headerRow = ActiveCell.Row
    
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(0, 1).Activate
    Wend
    
    size = ActiveCell.column - 2
    
    ReDim headerVal(size)
    
    Cells(headerRow, 1).Activate
    
    For i = 0 To UBound(headerVal)
        headerVal(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Activate
    Next i
    
    csvHeader = headerVal
    
    startCell.Activate
    Application.ScreenUpdating = True

End Function
Function getUnique(x)

    Dim items() As String, size As Integer, i As Integer

    thisCol = headerSearch(x)
    
    Cells(headerRow + 1, thisCol).Activate


    size = getLastRecord - headerRow
    uniqueLast = ActiveCell.Row - 1
    
    ReDim items(size - 1)
    assPosition = 0
    For i = headerRow + 1 To getLastRecord
        items(assPosition) = Cells(i, thisCol).Value
        assPosition = assPosition + 1
    Next i
    
    For i = LBound(items) To UBound(items)
        For j = LBound(items) To UBound(items)
            If j <> i Then
                If items(i) = items(j) Then
                    items(j) = 0
                End If
            End If
        Next j
    Next i
    
    count = 0
        
    For i = LBound(items) To UBound(items)
        If items(i) <> "0" Then
            count = count + 1
        End If
    Next i
    
    ReDim uniqueItems(count - 1) As String
    
    For i = LBound(uniqueItems) To UBound(uniqueItems)
        For j = LBound(items) To UBound(items)
            If items(j) <> "0" Then
                uniqueItems(i) = items(j)
                items(j) = 0
                Exit For
            End If
        Next j
    Next i
            
    getUnique = uniqueItems

End Function

Function checkAssets(theseAssets)
        
    If UBound(theseAssets) = 0 Then
        checkAssets = theseAssets(0)
    Else
        checkAssets = "XA"
    End If


End Function
Function getMessageType()
    Set startCell = ActiveCell
    
    find ("Message Type")
    
    ActiveCell.Offset(1, 0).Activate
    
    If ActiveCell.Value = "Trade State" Then getMessageType = "TRD"
    
    If ActiveCell.Value = "Valuation" Then getMessageType = "VAL"
    
    startCell.Activate

End Function
Function workingMessage()
    
    With Working
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless
        .Repaint
    End With
    
End Function

Function getpath()
    getpath = ActiveWorkbook.Path & "\"
End Function
Function getPasteRow(sheetName)

    startsheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    findID
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getPasteRow = ActiveCell.Row
    
    Sheets(startsheet).Activate

End Function


Function getArrayLength(x)
    getArrayLength = UBound(x) - LBound(x)
End Function
Function assAbbr(AssetClass)

    If AssetClass = "ForeignExchange" Or AssetClass = "FX" Then
        assAbbr = "_FX"
    
    ElseIf AssetClass = "CU" Then
        assAbbr = "_CU"
    
    ElseIf AssetClass = "InterestRate" Or AssetClass = "IR" Then
        assAbbr = "_IR"
    
    ElseIf AssetClass = "Commodity" Or AssetClass = "CO" Then
        assAbbr = "_CO"
    
    ElseIf AssetClass = "Equity" Or AssetClass = "EQ" Then
        assAbbr = "_EQ"
    
    ElseIf AssetClass = "Credit" Or AssetClass = "CR" Then
        assAbbr = "_CR"
        
    ElseIf AssetClass = "XA" Then
        assAbbr = "_XA"
        
    Else
        assAbbr = ""  'Asset Class not provided or recognized
    End If
    
End Function

Function getTrailer()
    getTrailer = Left(Cells(1, 1), 5) & "-END"
End Function

Function deleteTheSheets(tempSheet, newSheet, exitSheet)

On Error GoTo niceExit
    Application.DisplayAlerts = False
    Sheets(tempSheet).Delete
    Sheets(newSheet).Delete
    Sheets(exitSheet).Delete
    Application.DisplayAlerts = True

niceExit:

End Function


Function createNewExit(headerRow, newSheet, exitSheet)
    startsheet = ActiveSheet.Name
    Worksheets.Add
    ActiveSheet.Name = newSheet
    Worksheets.Add
    ActiveSheet.Name = exitSheet
    Sheets(startsheet).Activate
    Rows("1:" & headerRow).Copy
    Sheets(newSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Rows("1:" & headerRow).Copy
    Sheets(exitSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
End Function

Function getAssClassCol()
Set startCell = ActiveCell
findAssetClass
getAssClassCol = ActiveCell.column
startCell.Activate

End Function

Function getActionCol()
Set startCell = ActiveCell
find ("action")
getActionCol = ActiveCell.column

startCell.Activate
End Function

Function getLastRecord()
    Set startCell = ActiveCell
    find ("action")
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getLastRecord = ActiveCell.Row - 1
    startCell.Activate
    

End Function


Function headerSearch(columnName) 'returns column number

For i = 0 To UBound(csvHeader)
    If csvHeader(i) = columnName Then Exit For
Next i

headerSearch = i + 1

End Function

Function SplitSelector(inArrs As Variant) As Variant
    Dim inArr As Variant
    Dim outArr() As String
    Dim sortAr() As String
    Dim i As Integer
    
    For i = 0 To UBound(inArrs)
        ReDim Preserve newArr(i)
        newArr(i) = inArrs(i)
    Next i
On Error GoTo hardExit
    sortAr = sortArray(newArr)
    
    With SplitSelectForm
        For Each inArr In sortAr
            .optionList.AddItem inArr
        Next inArr
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
        ReDim Preserve outArr(.selectedList.ListCount - 1)
        For i = 0 To (.selectedList.ListCount - 1)
            outArr(i) = .selectedList.List(i)
        Next i
    End With
    SplitSelector = outArr
    Exit Function
    
hardExit:
    endIt = True

End Function

Function sortArray(arr As Variant) As String()
    Dim i As Integer
    Dim j As Integer
    Dim tmp
    
    For i = LBound(arr) To UBound(arr)
        For j = i + 1 To UBound(arr)
            If UCase(arr(i)) > UCase(arr(j)) Then
                tmp = arr(j)
                arr(j) = arr(i)
                arr(i) = tmp
            End If
        Next j
    Next i
    
    sortArray = arr
End Function

Function setStartVals()
    
    endIt = False
    prepMode = True
    fileCount = 0
    tempSheet = 0

End Function
