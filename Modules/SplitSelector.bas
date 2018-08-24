Function splitSelector(inArrs() As String) As String()
    Dim inArr As Variant
    Dim outArr() As String
    Dim i As Integer
    With SplitSelectForm
        For Each inArr In inArrs
            .selectCombo.AddItem inArr
        Next inArr
        .Show
        ReDim Preserve outArr(.orderList.listCount - 1)
        For i = 0 To (.orderList.listCount - 1)
            outArr(i) = .orderList.List(i)
        Next i
    End With
    splitSelector = outArr
End Function

Sub splitTest()
    Dim oldArr() As String
    Dim newArr() As String
    
    oldArr = Split("one,two,three,four,five", ",")
    
    newArr = splitSelector(oldArr)
    
    Debug.Print UBound(newArr) & ":" & Join(newArr, ",")
End Sub
