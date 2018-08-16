Sub csvToFpml(control As IRibbonControl)
    'Invoke CSV To FPML window and place it in the middle of the Excel window
    With csvToFpmlForm
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub

Sub parseFields(taxFilePath, fpmlFldPath)
    Dim taxTable As String
    Dim taxArray() As String
    Dim taxElem As Variant
    Dim taxLine As String
    
    taxArray = readFile(taxFilePath)
    
    For Each taxElem In taxArray
        
    Next taxElem
    
    For Each taxElem In taxArray
        taxTable = taxTable & taxElem & vbNewLine
    Next str
        
    MsgBox taxTable
End Sub

'Function for a reading file
Function readFile(filePath)
    Dim objFSO As Object
    Dim objTF  As Object
    Dim objStr As String
    Dim objArr() As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTF = objFSO.OpenTextFile(filePath, 1)
    objStr = objTF.readall
    objArr = Split(objStr, vbNewLine)
    objTF.Close
    
    'Returns an array with each new file line
    readFile = objArr
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
    loadPicker = selItems
End Function
