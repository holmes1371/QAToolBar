Sub neSplit()

uniqueAss = getUniqueAssets

For i = LBound(uniqueAss) To UBound(uniqueAss)
    Debug.Print uniqueAss(i)
Next i
        
        
End Sub

Function getUniqueAssets()
Dim assets() As String, size As Integer, i As Integer

findAssetClass
headerRow = ActiveCell.Row
assClassCol = ActiveCell.Column



    While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Activate
    Wend

    size = ActiveCell.Row - headerRow - 1
    lastRecord = ActiveCell.Row - 1
    
'Debug.Print lastRecord
'Debug.Print size

ReDim assets(size)
assPosition = 1
For i = headerRow + 1 To lastRecord
    assets(assPosition) = Cells(i, assClassCol).Value
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
        
getUniqueAssets = uniqueAssets

End Function


Function addNewBook()

thispath = ActiveWorkbook.Path & "\"
thisname = "TestBook"

Set NewBook = Workbooks.Add
    With NewBook
        .Title = "All Sales" 'You can modify this value.
        .Subject = "Sales" 'You can modify this value.
        .SaveAs Filename:=thispath & thisname & ".csv"
    End With


End Function
