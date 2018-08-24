Attribute VB_Name = "FpmlConvert"
Option Explicit

Public Sub fpmlMain()
    WriteFile getSchemas(), "\\SVFILE05B\FCastill$\My Documents\test.txt"
End Sub

Function getSchemas() As String
    Dim uniques  As Scripting.Dictionary
    Dim sources  As Variant
    Dim schemas  As Variant
    Dim unique   As Variant
    Dim source   As Variant
    Dim schema   As Variant
    Dim xmlUrl   As String
    Dim writeOut As String
    Dim htmlText As String
    
    xmlUrl = "http://www.fpml.org/spec/fpml-5-6-1-wd-1/html/recordkeeping/schemaDocumentation/schemas/fpml-main-5-6_xsd/schema-overview.html"
    Set uniques = New Scripting.Dictionary
    
    sources = GetUrl(xmlUrl, "f22")
    If UBound(sources) > 0 Then
        For Each source In sources
            schemas = GetUrl(source, "f22")
            If UBound(schemas) > 0 Then
                For Each schema In schemas
                    If InStr(schema, ".html#") = 0 Then
                        htmlText = GetContent(schema, "t2")
                        If InStr(htmlText, "XML Representation Summary") Then
                            httpText = Trim(Between(htmlText), ">", "<")
                            uniques.item(htmlText) = 1
                        End If
                    End If
                Next schema
            End If
        Next source
    End If
    
    For Each unique In uniques
        
    Next unique
    
    getSchemas = writeOut
End Function

Function Between(str As String, sta As String, stp As String)
    Dim staPos As Integer
    Dim stpPos As Integer
    staPos = InStr(str, sta)
    stpPos = InStr(staPos, str, stp)
    
    
End Function

Function GetContent(ByVal url As String, ByVal cls As String) As String
    Dim ie    As New InternetExplorer
    Dim docs  As Variant
    Dim doc   As Variant
    Dim txt   As String
    
    With ie
        .Visible = False
        .Navigate url
        While .Busy Or .ReadyState <> READYSTATE_COMPLETE: DoEvents: Wend
        Set docs = .Document
        For Each doc In docs.getElementsByClassName(cls)
            txt = txt & doc.innerText & vbNewLine
        Next doc
        .Quit
    End With
    Set ie = Nothing
    GetContent = txt
End Function

Function GetUrl(ByVal url As String, ByVal cls As String) As String()
    Dim ie    As New InternetExplorer
    Dim arr() As String
    Dim docs  As Variant
    Dim doc   As Variant
    Dim i     As Integer
    
    i = 0
    With ie
        .Visible = False
        .Navigate url
        While .Busy Or .ReadyState <> READYSTATE_COMPLETE: DoEvents: Wend
        Set docs = .Document
        For Each doc In docs.getElementsByClassName(cls)
            If doc.href <> "" Then
                ReDim Preserve arr(i)
                arr(i) = doc.href
                i = i + 1
            End If
        Next doc
        .Quit
    End With
    If i = 0 Then ReDim Preserve arr(0)
    Set ie = Nothing
    GetUrl = arr
End Function

Function ArrToText(arr As Variant) As String
    Dim obj As Object
    Dim txt As String
    For Each obj In arr
        txt = txt & Join(obj, ",") & vbNewLine
    Next obj
    ArrToText = txt
End Function

'Searches array and returns nth value
Function ArraySearch(arr As Variant, searchStr As String, isdaIdx As Integer, taxIdx As Integer) As String
    Dim line As Variant
    Dim bool As Boolean
    For Each line In arr
        If StrComp(line(isdaIdx), searchStr, vbTextCompare) = 0 Then
            ArraySearch = Replace(line(taxIdx), "/", "")
            Exit Function
        End If
    Next line
End Function

'Function for a reading file
Function ReadFile(filePath)
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
    ReadFile = outArr
End Function

Sub WriteFile(content As String, path As String)
    Dim FSO As Object
    Dim obj As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set obj = FSO.createTextFile(path)
    obj.WriteLine content
    obj.Close
End Sub

'Invokes file or folder picker, depending on string passed.
'Returns path of object selected
Function SelectFile(ftype As String, title As String) As String
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
        If .Show = True Then SelectFile = .SelectedItems(1)
    End With
End Function


