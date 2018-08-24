Attribute VB_Name = "SplitSelector"
Function splitSelector(inArrs() As String) As String()
    Dim inArr As Variant
    Dim outArr() As String
    Dim i As Integer
    With SplitSelectForm
        For Each inArr In inArrs
            .selectCombo.AddItem inArr
        Next inArr
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
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
    
    Debug.Print (UBound(newArr) + 1) & ":" & Join(newArr, ",")
End Sub
