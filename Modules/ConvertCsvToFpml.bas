Sub csvToFpml(control As IRibbonControl)
    'Invoke CSV To FPML window and place it in the middle of the Excel window
    With csvToFpmlForm
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub

Sub parseFields(taxFilePath, fpmlFldPath)
    Dim i As Integer
    Dim j As Integer
    Dim taxArr() As Variant
    Dim taxLen   As Integer
    Dim taxText  As String
    
    taxArr = readFile(taxFilePath)
    taxLen = UBound(taxArr)
        
    For i = 0 To taxLen
        taxText = taxText & Join(taxArr(i), ",") & vbNewLine
    Next i
    
    MsgBox taxText
End Sub

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

'Invokes file or folder picker, depending on string passed.
'Returns path of object selected
Function selectFile(name As String) As String
    Dim fileDiag As FileDialog
    Dim selItems As String
    Dim diagType As Object
    Dim diagName As String
    
    'Conditional based on whether name is file or fldr
    Select Case name
        Case "file"
            Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
            diagName = "Select Taxonomy File"
        Case "fldr"
            Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
            diagName = "Select FPML Files Destination"
    End Select
    
    'Sets properties for picker
    With fileDiag
        .Title = diagName
        If name = "file" Then
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Comma Separated Values file", "*.csv"
        End If
        If .Show = True Then selItems = .SelectedItems(1)
    End With
    
    'Returns path selected as a string
    selectFile = selItems
End Function
