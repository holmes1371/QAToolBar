Public headerRow
Option Compare Text

Sub bigFishPrep(control As IRibbonControl)
workingMessage
autoHeaderUniquinizerIngestF
Application.ScreenUpdating = False
startSheet = ActiveSheet.Name
uniqueAss = getUnique("asset")

For i = LBound(uniqueAss) To UBound(uniqueAss)
    Debug.Print uniqueAss(i)
Next i

'create random names for sheets
tempSheet = "working" & CStr(Int((500000 - 1 + 1) * Rnd + 1))
newSheet = "new" & CStr(Int((500000 - 1 + 1) * Rnd + 1))
exitSheet = "exit" & CStr(Int((500000 - 1 + 1) * Rnd + 1))

For i = 1 To UBound(uniqueAss)
    If uniqueAss(i) = "" Then Exit For
    createTempSheet (tempSheet)
    For j = headerRow + 1 To getLastRecord
        If Cells(j, getAssClassCol).Value = uniqueAss(i) Then
        Rows(j).Copy
        Sheets(tempSheet).Range("A" & getPasteRow(tempSheet)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        End If
    Next j
    Sheets(tempSheet).Range("A" & getPasteRow(tempSheet)).Value = getTrailer
    Sheets(tempSheet).Activate
    uniqueAction = getUnique("action")
    
    If getActionLength(uniqueAction) = 1 Then 'if only 1 action type, new file is created
        Call makeNewBook(uniqueAction(1), uniqueAss(i), tempSheet)
        Call deleteTheSheets(tempSheet, newSheet, exitSheet)
    Else
   
    Call createNewExit(headerRow, newSheet, exitSheet)
    
        For k = headerRow + 1 To getLastRecord
            If Cells(k, getActionCol).Value = "new" Then
                Rows(k).Copy
                Sheets(newSheet).Range("A" & getPasteRow(newSheet)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            End If
            If Cells(k, getActionCol).Value = "exit" Then
                Rows(k).Copy
                Sheets(exitSheet).Range("A" & getPasteRow(exitSheet)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            End If
        Next k
        
        Sheets(newSheet).Range("A" & getPasteRow(newSheet)).Value = getTrailer
        Sheets(exitSheet).Range("A" & getPasteRow(exitSheet)).Value = getTrailer
        
        Call makeNewBook("NEW", uniqueAss(i), newSheet)
        Call makeNewBook("EXIT", uniqueAss(i), exitSheet)
        Call deleteTheSheets(tempSheet, newSheet, exitSheet)
      End If

Next i

    Unload Working
    refreshScreen
     resetSearchParameters
    MsgBox "Files have been successfully created", vbInformation, "Complete"
    retVal = Shell("explorer.exe " & getpath, vbNormalFocus)
    refreshScreen
    Cells(1, 1).Activate
End Sub
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
    startSheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    Range("A1").Activate
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getPasteRow = ActiveCell.Row
    
    Sheets(startSheet).Activate

End Function
Function getNamePart(ActionType, AssetClass, sheetName)

    thispath = ActiveWorkbook.Path & "\"
    Debug.Print thispath
    startSheet = ActiveSheet.Name
    Sheets(sheetName).Activate
    testName = getUnique("*comment")
    testLength = UBound(testName) - LBound(testName) - 2
    If testLength = 1 Then
        testNumber = testName(1)
    Else
        testNumber = "MTC"
    End If
    
    getNamePart = thispath & testNumber & "_INPUT_" & assAbbr(AssetClass) & "_ESMA_" & UCase(ActionType)
    'Debug.Print getNamePart
    Sheets(startSheet).Activate
    
End Function
Function makeNewBook(ActionType, AssetClass, sheetName)
'
    fileName = getNamePart(ActionType, AssetClass, sheetName)

    Worksheets(sheetName).Copy
    
    With ActiveWorkbook
        .SaveAs fileName:=fileName, FileFormat:=xlCSV
         SheetFixIngestF
         Application.ScreenUpdating = False
        .Close SaveChanges:=True
    End With

End Function
Function getActionLength(uniqueAction)
    getActionLength = UBound(uniqueAction) - LBound(uniqueAction) - 1
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
        
    Else
        assAbbr = ""  'Asset Class not provided or recognized
    End If
    
End Function

Function getTrailer()
    getTrailer = Left(Cells(1, 1), 5) & "-END"
End Function
Function getUnique(x)
Dim assets() As String, size As Integer, i As Integer
Cells(1, 1).Activate
If x = "asset" Then
    findAssetClass
Else
    find (x)
End If

headerRow = ActiveCell.Row
thisCol = ActiveCell.Column

    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend

    size = ActiveCell.Row - headerRow - 1
    uniqueLast = ActiveCell.Row - 1
    
'Debug.Print lastRecord
'Debug.Print size

ReDim assets(size)
assPosition = 1
For i = headerRow + 1 To uniqueLast
    assets(assPosition) = Cells(i, thisCol).Value
    assPosition = assPosition + 1
Next i

For i = LBound(assets) To UBound(assets)
    For j = LBound(assets) To UBound(assets)
        If j <> i Then
            If assets(i) = assets(j) Then
                assets(j) = 0
            End If
        End If
    Next j
Next i

count = 0
    
For i = LBound(assets) To UBound(assets)
    If assets(i) <> "0" Then
        count = count + 1
    End If
Next i

ReDim uniqueAssets(count) As String

For i = LBound(uniqueAssets) To UBound(uniqueAssets)
    For j = LBound(assets) To UBound(assets)
        If assets(j) <> "0" Then
            uniqueAssets(i) = assets(j)
            assets(j) = 0
            Exit For
        End If
    Next j
Next i
        
getUnique = uniqueAssets

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
Function createTempSheet(tempSheet)
    startSheet = ActiveSheet.Name
    Worksheets.Add
    ActiveSheet.Name = tempSheet
    Sheets(startSheet).Activate
    Rows("1:" & headerRow).Copy
    Sheets(tempSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Sheets(startSheet).Activate
End Function

Function createNewExit(headerRow, newSheet, exitSheet)
    startSheet = ActiveSheet.Name
    Worksheets.Add
    ActiveSheet.Name = newSheet
    Worksheets.Add
    ActiveSheet.Name = exitSheet
    Sheets(startSheet).Activate
    Rows("1:" & headerRow).Copy
    Sheets(newSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Rows("1:" & headerRow).Copy
    Sheets(exitSheet).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
End Function

Function getAssClassCol()
Set startcell = ActiveCell
findAssetClass
getAssClassCol = ActiveCell.Column
Debug.Print getAssClassCol
startcell.Activate

End Function

Function getActionCol()
Set startcell = ActiveCell
find ("action")
getActionCol = ActiveCell.Column
Debug.Print getActionCol
startcell.Activate
End Function

Function getLastRecord()
    Set startcell = ActiveCell
    find ("action")
    
    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    getLastRecord = ActiveCell.Row - 1
    startcell.Activate
    Debug.Print getLastRecord

End Function