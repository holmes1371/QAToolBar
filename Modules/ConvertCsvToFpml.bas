Option Explicit


Sub csvToFpml(control As IRibbonControl)
'    Dim taxArr    As Variant
'    Dim trnxArr   As Variant
'    Dim taxonomy  As String
    Dim schemaTxt As String
'
'    taxArr = readFile(selectFile("file", "Select Taxonomy File"))
'
'    taxonomy = arraySearch(taxArr, "FOREIGNEXCHANGE:NDF", 1, 4)
    
    schemaTxt = getSchema("fxSingleLeg")
    
    Debug.Print schemaTxt
    Call endIESess
End Sub

Function getSchema(name As String) As String
    Dim source As IHTMLElement
    Dim compnt As IHTMLElement
    Dim schema As IHTMLElement
    Dim repsum As IHTMLElement
    Dim repArr() As String
    Dim repLine  As Variant
    Dim xmlUrl   As String
    
    xmlUrl = "http://www.fpml.org/spec/fpml-5-6-1-wd-1/html/recordkeeping/schemaDocumentation/schemas/fpml-main-5-6_xsd/schema-overview.html"
    
    For Each source In getUrl(xmlUrl, "t2")
        For Each compnt In source.document.getElementsByClassName("f22")
            For Each schema In getUrl(compnt.href, "f22")
                If schema.innerText = name Then
                    For Each repsum In getUrl(schema.href, "f36")
                        getSchema = repsum.innerHTML
                        Exit Function
'                        repArr = Split(repsum.all.Item(0).innerText, vbNewLine)
'                        For Each repLine In repArr
'                            getSchema = repLine
'                        Next repLine
                    Next repsum
                End If
            Next schema
        Next compnt
    Next source
End Function

Function getUrl(url As String, className As String)
    Dim ie As InternetExplorer
    
    Set ie = New InternetExplorer
    
    With ie
        .Visible = False
        .RegisterAsBrowser = True
        .navigate url
        While .Busy Or .readyState <> 4: DoEvents: Wend
        Set getUrl = .document.getElementsByClassName(className)
    End With
    
    Set ie = Nothing
End Function

Sub endIESess()
    Dim wmSess As Object
    Dim wmProc As Object
    Set wmSess = GetObject("winmgmts://.").ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'iexplore.exe'")
    For Each wmProc In wmSess
        Call wmProc.Terminate
    Next wmProc
    
    Set wmSess = Nothing
    Set wmProc = Nothing
End Sub

Function arrToText(arr As Variant) As String
    Dim obj As Object
    Dim txt As String
    For Each obj In arr
        txt = txt & Join(obj, ",") & vbNewLine
    Next obj
    arrToText = txt
End Function

'Searches array and returns nth value
Function arraySearch(arr As Variant, searchStr As String, isdaIdx As Integer, taxIdx As Integer) As String
    Dim line As Variant
    Dim bool As Boolean
    For Each line In arr
        If StrComp(line(isdaIdx), searchStr, vbTextCompare) = 0 Then
            arraySearch = Replace(line(taxIdx), "/", "")
            Exit Function
        End If
    Next line
End Function

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

Sub writeFile(content As String, path As String)
    Dim FSO As Object
    Dim obj As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set obj = FSO.createTextFile(path)
    obj.WriteLine content
    obj.Close
End Sub

'Invokes file or folder picker, depending on string passed.
'Returns path of object selected
Function selectFile(ftype As String, title As String) As String
    Dim fileDiag As FileDialog
    
    'Conditional based on whether name is file or fldr
    Select Case ftype
        Case "file"
            Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
        Case "fldr"
            Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    End Select
    
    'Sets properties for picker
    With fileDiag
        .title = title
        If ftype = "file" Then
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Comma Separated Values file", "*.csv"
        End If
        'Returns path selected as a string
        If .Show = True Then selectFile = .SelectedItems(1)
    End With
End Function
