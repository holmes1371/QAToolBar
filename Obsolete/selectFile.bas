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
