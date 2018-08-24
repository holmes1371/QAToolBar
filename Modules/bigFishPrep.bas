Public headerRow
Public csvHeader
Public tempSheet
Public counter
Option Compare Text

Sub bigFishPrep(control As IRibbonControl)
Set startCell = ActiveCell
utiMode = "auto"
workingMessage
autoHeaderUniquinizerIngestF
Application.ScreenUpdating = False
startsheet = ActiveSheet.Name
counter = 0
tempSheet = 0
setHeaderVals
Dim splitOptions(2) As String

splitOptions(1) = "Primary Asset Class"
splitOptions(2) = "Action"

Call splitThisSheet(ActiveSheet.Name, splitOptions)

verifyFinal (splitOptions)

    Unload Working
    refreshScreen
    resetSearchParameters
    MsgBox counter & " Files have been successfully created", vbInformation, "Complete"
    retVal = Shell("explorer.exe " & getpath, vbNormalFocus)
    startCell.Activate
    
End Sub

Function splitThisSheet(sheetName As String, criteria)

Sheets(sheetName).Activate

    For Each opt In criteria
        If opt = "" Then
        GoTo skipIt
        End If
        tempUnique = getUnique(opt)
        If getArrayLength(tempUnique) = 1 Then
            GoTo skipIt:
        Else
            For Each element In tempUnique
            If element <> "" Then
            tempSheet = tempSheet + 1
            createTempSheet (CStr(tempSheet))
            Call doCopy(sheetName, CStr(tempSheet), headerSearch(opt), element)
            Call splitThisSheet(CStr(tempSheet), criteria)
            End If
            Next element
        End If
    
skipIt:

    Next opt

End Function
Function verifyFinal(criteria)
Application.DisplayAlerts = False
    For i = 1 To tempSheet
    
        Sheets(CStr(i)).Activate
        
        For Each elem In criteria
            If elem = "" Then GoTo skipIt
            isFinished = False
            tempUnique = getUnique(elem)
            If getArrayLength(tempUnique) <> 1 Then
                Exit For
            End If
            isFinished = True
skipIt:
        Next elem

        If isFinished = True Then
            Call makeNewBook(criteria, CStr(i))
            counter = counter + 1
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
    Debug.Print thispath
    startsheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    testName = getUnique("*comment")
    testLength = UBound(testName) - 1
    
    If testLength = 1 Then
        testNumber = testName(1)
    Else
        testNumber = "MTC"
    End If
    
    Sheets(sheetName).Range("A" & getPasteRow(sheetName)).Value = getTrailer
    
    AssetClass = checkAssets(getUnique(Cells(headerRow, getAssClassCol).Value))
    
    Dim endPart As String
    
    For i = 1 To UBound(criteria)
        If criteria(i) = "Asset Class" Or criteria(i) = "Primary Asset Class" Then GoTo skipIt
        appendage = getUnique(criteria(i))
        endPart = endPart & "_" & UCase(appendage(1))
skipIt:
    Next i
    
    getNamePart = thispath & testNumber & "_INPUT_" & assAbbr(AssetClass) & "_ESMA" & endPart
    'Debug.Print getNamePart
    Sheets(startsheet).Activate
    
End Function

Function doCopy(parentSheet, targetSheet As String, icol, criteria)
    startsheet = ActiveSheet.Name
    Sheets(parentSheet).Activate
    
        For j = headerRow + 1 To getLastRecord
            If Cells(j, icol).Value = criteria Then
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


End Function
Function getUnique(x)
Dim items() As String, size As Integer, i As Integer

thisCol = headerSearch(x)

Cells(headerRow + 1, thisCol).Activate


    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend

    size = ActiveCell.Row - headerRow
    uniqueLast = ActiveCell.Row - 1
    
'Debug.Print lastRecord
'Debug.Print size

ReDim items(size)
assPosition = 1
For i = headerRow + 1 To uniqueLast
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

ReDim uniqueItems(count) As String

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
        
    If UBound(theseAssets) - LBound(theseAssets) - 1 = 1 Then
        checkAssets = theseAssets(1)
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
    Debug.Print getpath

End Function
Function getPasteRow(sheetName)
    startsheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    Range("A1").Activate
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getPasteRow = ActiveCell.Row
    
    Sheets(startsheet).Activate

End Function


Function getArrayLength(x)
    getArrayLength = UBound(x) - LBound(x) - 1
End Function
Function assAbbr(AssetClass)

    If AssetClass = "ForeignExchange" Or AssetClass = "FX" Then
        assAbbr = "FX"
    
    ElseIf AssetClass = "CU" Then
        assAbbr = "CU"
    
    ElseIf AssetClass = "InterestRate" Or AssetClass = "IR" Then
        assAbbr = "IR"
    
    ElseIf AssetClass = "Commodity" Or AssetClass = "CO" Then
        assAbbr = "CO"
    
    ElseIf AssetClass = "Equity" Or AssetClass = "EQ" Then
        assAbbr = "EQ"
    
    ElseIf AssetClass = "Credit" Or AssetClass = "CR" Then
        assAbbr = "CR"
        
    ElseIf AssetClass = "XA" Then
        assAbbr = "XA"
        
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
Debug.Print getAssClassCol
startCell.Activate

End Function

Function getActionCol()
Set startCell = ActiveCell
find ("action")
getActionCol = ActiveCell.column
Debug.Print getActionCol
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
    Debug.Print getLastRecord

End Function


Function headerSearch(columnName) 'returns column number

For i = 0 To UBound(csvHeader)
    If csvHeader(i) = columnName Then Exit For
Next i

headerSearch = i + 1

End Function

