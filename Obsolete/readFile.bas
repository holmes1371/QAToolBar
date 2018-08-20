'Function for a reading file
Function readFile(filePath)
    Dim FSO       As Object
    Dim OTF       As Object
    Dim fileStr   As String
    Dim fileArr() As String
    Dim fileLen   As Integer
    Dim fileIdx   As Integer
    Dim outArr()  As Variant
    
    'Reads file path and loads file into an array
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OTF = FSO.OpenTextFile(filePath, 1)
    fileStr = OTF.readall
    fileArr = Split(fileStr, vbNewLine)
    OTF.Close
    
    'Determines the length of the array (ie. # of lines in file)
    fileLen = UBound(fileArr)
    
    'Reclassifies array in order to fit fileArr
    ReDim outArr(fileLen) As Variant
    
    'Loops through to enter CSV into two-dimensional array
    For fileIdx = 0 To fileLen
        outArr(fileIdx) = Split(fileArr(fileIdx), ",")
    Next fileIdx
    
    'Returns two-dimensional array
    readFile = outArr
End Function
