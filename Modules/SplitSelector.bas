Attribute VB_Name = "SplitSelector"
Option Explicit

Function splitSelector(inArrs() As String) As String()
    Dim inArr As Variant
    Dim outArr() As String
    Dim sortAr() As String
    Dim i As Integer

    sortAr = sortArray(inArrs)
    
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
    splitSelector = outArr
End Function

Function sortArray(arr As Variant) As String()
    Dim i As Integer
    Dim j As Integer
    Dim tmp
    
    For i = LBound(arr) To UBound(arr) - 1
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

Sub splitTest()
    Dim oldArr() As String
    Dim newArr() As String
    
    oldArr = Split("one,two,three,four,five", ",")
    
    newArr = splitSelector(oldArr)
    
    Debug.Print (UBound(newArr) + 1) & ":" & Join(newArr, ",")
End Sub
