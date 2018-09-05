Public totalFiles   As Integer
Public fileCount    As Integer
Public filesMade    As Integer
Public tempSheet    As String
Public csvHeader    As Variant
Public headerRow    As Integer
Public prepMode     As Boolean
Public newArr()     As String
Public proceedSplit As Boolean
Public randPrefix   As String

Option Explicit
Option Compare Text

Sub splitFiles(control As IRibbonControl)

    Dim startCell       As Object: Set startCell = ActiveCell
    Dim startSheet      As String: startSheet = ActiveSheet.Name
    Dim splitOptions()  As String
    Dim retVal          As Variant
    
    setStartVals
    If preCheck = False Then Exit Sub
    utiMode = "auto"
    
    Application.ScreenUpdating = True
    On Error GoTo softExit
    splitOptions = SplitSelector(csvHeader)
    Application.ScreenUpdating = False
    
    If endIt = True Then
        Exit Sub
    End If
    
    setTotalFiles (splitOptions)
    setHeaderVals
    autoHeaderUniquinizerIngestF
    progressBarMessage
    
    Call splitThisSheet(ActiveSheet.Name, splitOptions)
    
    verifyFinal (splitOptions)

    Unload ProgressBar
    startCell.Activate
    resetSearchParameters
    
    If filesMade = 0 Then
        MsgBox "No files have been created", vbInformation, "Complete"
    Else
        MsgBox filesMade & " files have been successfully created", vbInformation, "Complete"
        'open fileExplorer to the location where files are saved
        retVal = Shell("explorer.exe " & getpath, vbNormalFocus)
    End If
    
    prepMode = False
    
softExit:
    
    Application.ScreenUpdating = True
    
End Sub

Function splitThisSheet(sheetName As String, criteria)
    
    Dim opt             As Variant
    Dim tempUnique()    As String
    Dim i               As Integer
        
    Sheets(sheetName).Activate

    For Each opt In criteria

        tempUnique = getUnique(opt, ActiveSheet.Name)
        
        If UBound(tempUnique) = 0 Then
            GoTo skipIt:
        Else
            For i = LBound(tempUnique) To UBound(tempUnique)
            tempSheet = tempSheet + 1
            createTempSheet (randPrefix & CStr(tempSheet))
            Call doCopy(sheetName, randPrefix & CStr(tempSheet), headerSearch(opt), tempUnique(i))
            Call splitThisSheet(randPrefix & CStr(tempSheet), criteria)
            incrementProgressBar
            Next i
        End If
    
skipIt:
    Next opt

End Function

Function createRandomPrefix() As String

    Dim i           As Integer
    Dim thisPrefix  As String
    
    For i = 0 To 10
        thisPrefix = thisPrefix & CStr(Round(((10 - 1 + 1) * Rnd + 1), 0))
    Next i
    
    createRandomPrefix = thisPrefix
    
End Function
Function verifyFinal(criteria)

    Dim isFinished      As Boolean
    Dim i, j            As Integer
    Dim tempUnique()    As String
    
    
    Application.DisplayAlerts = False
    
    For i = 1 To tempSheet
    
        Sheets(randPrefix & CStr(i)).Activate
        
        For j = LBound(criteria) To UBound(criteria)
            isFinished = False
            tempUnique = getUnique(criteria(j), ActiveSheet.Name)
            If UBound(tempUnique) <> 0 Then Exit For
            isFinished = True
        Next j

        If isFinished = True Then
            Call makeNewBook(criteria, randPrefix & CStr(i))
            incrementProgressBar
            Sheets(randPrefix & CStr(i)).Delete
        Else
            Sheets(randPrefix & CStr(i)).Delete
        End If
    Next i

    Application.DisplayAlerts = True

End Function
Function progressBarMessage()
    With ProgressBar
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless
        .Repaint
    End With
End Function
Sub incrementProgressBar()

    fileCount = fileCount + 1
    If fileCount / totalFiles >= 1 Then
        fileCount = totalFiles
    End If
    With ProgressBar
        .Text.Caption = CStr(Round(((fileCount / totalFiles) * 100), 0)) & "% Complete"
        .Bar.Width = ((fileCount / totalFiles) * 100) * 2
    End With
    
    DoEvents
    
End Sub
Function setTotalFiles(criteria)

    Dim totalNumber As Integer
    Dim i           As Integer

    totalNumber = 1
    
    For i = LBound(criteria) To UBound(criteria)
        totalNumber = totalNumber * (UBound(getUnique(criteria(i), ActiveSheet.Name)) + 1)
    Next i
    
    If UBound(criteria) = 0 Or UBound(criteria) = 1 Then
        totalFiles = totalNumber * 2.5
    Else
        totalFiles = totalNumber / 1.5
    End If
    
End Function
Function makeNewBook(criteria, sheetName As String)
    
    Dim fileName As String
    
    filesMade = filesMade + 1
    fileName = getNamePart(criteria, sheetName)

    Worksheets(sheetName).Copy
    
    With ActiveWorkbook
        .SaveAs fileName:=fileName, FileFormat:=xlCSV
         SheetFixIngestF
         'Application.ScreenUpdating = False
        .Close SaveChanges:=True
    End With

End Function
Function getNamePart(criteria, sheetName As String)
    
    Dim thisPath, startSheet    As String
    Dim AssetClass, testNumber  As String
    Dim testLength, i           As Integer
    Dim testName()              As String
    Dim char                    As Variant
    Dim appendage()             As String
        
    Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?,-,~,/,\,:"
    
    thisPath = ActiveWorkbook.path & "\"
    startSheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    testName = getUnique("*comment", ActiveSheet.Name)
    testLength = UBound(testName)
    
    If testLength = 0 Then
        If testName(0) = "" Then testName(0) = "BLANK"
        testNumber = testName(0)
    Else
        testNumber = "MTC"
    End If
    
    Sheets(sheetName).Range("A" & getPasteRow(sheetName)).Value = getTrailer
    
    AssetClass = checkAssets(getUnique(Cells(headerRow, getAssClassCol).Value, ActiveSheet.Name))
    
    Dim endPart As String
    
    For i = 0 To UBound(criteria)
        If criteria(i) = "Asset Class" Or criteria(i) = "Primary Asset Class" Then GoTo skipIt
        appendage = getUnique(criteria(i), ActiveSheet.Name)
        If appendage(0) = "" Then appendage(0) = "blank"
             
        endPart = endPart & "_" & UCase(appendage(0))
        
        For Each char In Split(SpecialCharacters, ",")
            endPart = Replace(endPart, char, "_")
        Next
skipIt:
    Next i
    
    If hasUTI(criteria) = True Then
        getNamePart = thisPath & testNumber & "_INPUT" & "_ESMA" & endPart
    Else
        getNamePart = thisPath & testNumber & "_INPUT" & assAbbr(AssetClass) & "_ESMA" & endPart
    End If
    
    Sheets(startSheet).Activate
    
End Function
Function hasUTI(criteria) As Boolean
    
    Dim i As Integer
    
    hasUTI = False
    
    For i = LBound(criteria) To UBound(criteria)
       If criteria(i) = "UTI" Or criteria(i) = "UTI ID" Or criteria(i) = "Trade ID" Then hasUTI = True
    Next i

End Function

Sub doCopy(parentSheet, targetSheet As String, icol As Integer, criteria)

    Dim j               As Integer
    Dim thisLastRecord  As Integer

    thisLastRecord = getLastRecord(parentSheet)
    
        For j = headerRow + 1 To thisLastRecord
            If CStr(Sheets(parentSheet).Cells(j, icol).Value) = criteria Then
                Sheets(targetSheet).Rows(getPasteRow(targetSheet)).Value = Sheets(parentSheet).Rows(j).Value
            End If
        Next j
    
End Sub
Sub createTempSheet(tempSheet As String)

    Dim startSheet As String: startSheet = ActiveSheet.Name
    
    Worksheets.Add.Name = tempSheet
    
    Sheets(tempSheet).Rows("1:" & headerRow).Value = Sheets(startSheet).Rows("1:" & headerRow).Value

    Sheets(startSheet).Activate
End Sub

Public Function setHeaderVals(Optional bitBucket As Variant)

    Dim headerVal()  As String
    Dim size         As Integer
    Dim i            As Integer
    
    Cells(1, 1).Activate
    
    If IsMissing(bitBucket) = True Then
        While ActiveCell.Value <> "*comment"
            ActiveCell.Offset(1, 0).Activate
        Wend
    End If
    
    headerRow = ActiveCell.row
       
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(0, 1).Activate
    Wend
    
    size = ActiveCell.Column - 2
    
    ReDim headerVal(size)
    
    Cells(headerRow, 1).Activate
    
    For i = 0 To UBound(headerVal)
        headerVal(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Activate
    Next i
    
    csvHeader = headerVal

End Function
Function getUnique(x, currentSheet)

    Dim items()         As String
    Dim size            As Integer
    Dim i, j, count     As Integer
    Dim thisCol         As Integer: thisCol = headerSearch(x)
    Dim thisLastRecord  As Integer: thisLastRecord = getLastRecord(currentSheet)
    Dim thisPosition    As Integer
    
    Cells(headerRow + 1, thisCol).Activate

    size = thisLastRecord - headerRow
    ReDim items(size - 1)
    
    thisPosition = 0
    For i = headerRow + 1 To thisLastRecord
        items(thisPosition) = Cells(i, thisCol).Value
        thisPosition = thisPosition + 1
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

Function checkAssets(theseAssets) As String
        
    If UBound(theseAssets) = 0 Then
        checkAssets = theseAssets(0)
    Else
        checkAssets = "XA"
    End If


End Function
Sub getMessageType()
    
    Dim startCell As Object: Set startCell = ActiveCell
        
    find ("Message Type")
    
    ActiveCell.Offset(1, 0).Activate
    
    If ActiveCell.Value = "Trade State" Then getMessageType = "TRD"
    
    If ActiveCell.Value = "Valuation" Then getMessageType = "VAL"
    
    startCell.Activate

End Sub

Function getpath() As String

    getpath = ActiveWorkbook.path & "\"
    
End Function
Function getPasteRow(sheetName)

    Dim startSheet As String: startSheet = ActiveSheet.Name
    
    Sheets(sheetName).Activate
    findID
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getPasteRow = ActiveCell.row
    
    Sheets(startSheet).Activate

End Function

Function getArrayLength(x) As Integer

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

Function getTrailer() As String

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

Function getAssClassCol()
    
    Dim startCell As Object: Set startCell = ActiveCell
    
    findAssetClass
    getAssClassCol = ActiveCell.Column
    startCell.Activate

End Function


Function getLastRecord(currentSheet) As Integer
    
    Dim startSheet  As String: startSheet = ActiveSheet.Name
    Dim startCell   As Object: Set startCell = ActiveCell
    
    Sheets(currentSheet).Activate
    
    find ("action")
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getLastRecord = ActiveCell.row - 1
    
    Sheets(startSheet).Activate
    startCell.Activate
    
End Function

Public Function headerSearch(columnName) As Integer 'returns column number
        
    Dim i As Integer
        
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
    proceedSplit = False
    With SplitSelectForm
        For Each inArr In sortAr
            .optionList.AddItem inArr
        Next inArr
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
        If proceedSplit = True Then
            ReDim Preserve outArr(.selectedList.ListCount - 1)
            For i = 0 To (.selectedList.ListCount - 1)
                outArr(i) = .selectedList.List(i)
            Next i
        Else
            GoTo hardExit
        End If
    End With
    SplitSelector = outArr
    Exit Function
    
hardExit:
    proceedSplit = False
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
    
    filesMade = 0
    endIt = False
    prepMode = True
    fileCount = 0
    tempSheet = 0
    randPrefix = createRandomPrefix
    Application.ScreenUpdating = False

End Function
