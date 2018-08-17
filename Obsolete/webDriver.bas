Attribute VB_Name = "Reports"
Option Explicit

Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_MAXIMIZE = 3

Sub pullReports(control As IRibbonControl)
    'IE Object
    Dim dl As Object
    
    'GTR URL
    Dim gtrUrl As String
    gtrUrl = <GTR URL here>
    
    Set dl = newIE(gtrUrl)
    
    Call gtrLogin(dl, <participantID here>)
      
    'Call putFiles(dl)
    Call getFiles(dl)
    
    Set dl = Nothing
End Sub

Sub gtrLogin(ieObj As Object, p As String)
    'Values corresponding to target objects
    Dim optElem  As String
    optElem = "signonWorkflowValue"
    
    'Web Driver(operation, IE object, input value, target object)
    webDriver "optn", ieObj, 3, optElem 'Super User
    webDriver "optn", ieObj, 1, optElem 'Participant
    webDriver "text", ieObj, p, optElem 'Participant ID
    webDriver "optn", ieObj, 3, optElem 'O-Code
End Sub

Sub putFiles(ieObj As Object)
    webDriver "innr", ieObj, "Upload", "a"  'Click Upload button
End Sub

Sub getFiles(ieObj As Object)
    webDriver "link", ieObj, "Download", "a"           'Click Download button
    webDriver "link", ieObj, "Document", "a"           'Click Download (reports tab)
    webDriver "cbox", ieObj, "", "chbksReportsIds_CD"  'Click each checkbox and download
    webDriver "link", ieObj, " Back to Dashboard", "a"  'Returns to dashboard
End Sub

Sub webDriver(oper, ieObj, val, elem)
    Dim obj  As Object
    Dim objs As Object
    Dim count As Integer
    Dim submit As Boolean
    
    While ieObj.readyState <> 4 Or ieObj.Busy: DoEvents: Wend
    
checkElem:
    Do
        Set obj = Nothing
        On Error Resume Next
        If oper = "optn" Then        'choose from drop-down list
            Set obj = ieObj.document.getElementById(elem).Options(val)
            DoEvents
            obj.Selected = True
            On Error GoTo checkElem
            submit = True
        ElseIf oper = "text" Then    'enter string into textbox
            Set obj = ieObj.document.getElementsByName(elem)(0)
            DoEvents
            obj.Value = val
            On Error GoTo checkElem
            submit = True
        ElseIf oper = "name" Then    'click a button or link by finding element name
            Set obj = ieObj.document.getElementsByName(elem)
            DoEvents
            obj.Value = val
            On Error GoTo checkElem
            submit = False
        ElseIf oper = "cbox" Then    'click individual checkboxes in download page
            Dim actionLink  As String
            Dim chkedBatch  As String
            Dim filePath    As String
            Dim token       As String
            Dim tokenName   As String
            Dim assetArr(8) As String
                assetArr(0) = "Document"
                assetArr(1) = "Equity"
                assetArr(2) = "Credit"
                assetArr(3) = "Rates ODRF"
                assetArr(4) = "Commodities"
                assetArr(5) = "Rates"
                assetArr(6) = "Fx"
                assetArr(7) = "RT"

            tokenName = "org.apache.struts.taglib.html.TOKEN"
            token = ieObj.document.getElementsByName(tokenName)(0).Value
                        
            For count = 0 To 7
                webDriver "link", ieObj, assetArr(count), "a"
                Set objs = ieObj.document.getElementsByName(elem)
                On Error GoTo checkElem
                DoEvents
                If objs Is Nothing Then Exit For
                For Each obj In objs
                    actionLink = obj.form.FirstChild.Value
                    chkedBatch = obj.Value
                    filePath = "actionLink=" & actionLink & _
                               "&org.apache.struts.taglib.html.TOKEN=" & token & _
                               "&command=download" & _
                               "&chkedBatchDetails=" & chkedBatch & _
                               "&chbksReportsIds_CD=" & chkedBatch
                Next obj
            Next count
            
            If count = 8 Then GoTo cBoxEnd
            
            submit = False
        ElseIf oper = "link" Then    'btn: click a button or link by finding innerHTML value
            Set objs = ieObj.document.getElementsByTagName(elem)
            On Error GoTo checkElem
            DoEvents
            For Each obj In objs
                If obj.innerText = val Then
                    obj.Click
                    Exit For
                End If
            Next
            submit = False
        End If
    Loop While obj Is Nothing
    
    While ieObj.readyState <> 4 Or ieObj.Busy: DoEvents: Wend
    If submit = True Then ieObj.document.forms(0).submit
    Set obj = Nothing
    Set objs = Nothing
cBoxEnd:
End Sub

Function newIE(url)
    'Create IE object
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
        
    'Set IE properties
    With IE
        .Toolbar = False
        .MenuBar = False
        .Visible = False
        .Navigate url
    End With
    
    'Wait until page finishes loading
    While IE.readyState <> 4 Or IE.Busy: DoEvents: Wend

    apiShowWindow IE.hwnd, SW_MAXIMIZE  'Maximize IE window
    IE.Visible = True                   'Make IE window visible
    Set newIE = IE                      'Return IE object
End Function
