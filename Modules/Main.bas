'QA Toolbar, v. 1.3
'Dtd: 08/15/2018

Option Compare Text
Public endIt               'these are catch variables which allow a function to terminate the sub
Public abortIt
Public MyRibbon As IRibbonUI

Sub OnRibbonLoad(ribbonUI As IRibbonUI)
    Set MyRibbon = ribbonUI
End Sub

Sub clearBox(control As IRibbonControl, ByRef returnVal)
    Select Case (control.ID)
        Case "ocodeVal": returnVal = "": oCode = ""
        Case "gtxValue": returnVal = "": gtxString = ""
    End Select
End Sub

Sub ApplyAllFormatting(control As IRibbonControl)
    autoHeader2
    autoHeaderUniquinizerIngestF
    resetSearchParameters
    'ActiveWorkbook.Save 'uncomment this line to activate the autosave function.
End Sub

Sub autoHeaderIngest(control As IRibbonControl)
    autoHeader2
    If endIt = False Then
        Exit Sub
    End If
    resetSearchParameters
End Sub

Sub SheetFixIngest(control As IRibbonControl)
    SheetFixIngestF
    resetSearchParameters
End Sub

Sub autoHeaderFormatterIngest(control As IRibbonControl)
    autoHeader2
    SheetFixIngestF
    If endIt = False Then
       Exit Sub
    End If
End Sub

Sub autoHeaderUniquinizerIngest(control As IRibbonControl)
    autoHeaderUniquinizerIngestF
    resetSearchParameters
End Sub

Sub findTradeID(control As IRibbonControl)
    SheetFixIngestF
    findID
    If foundOne = True Then
        Application.ScreenUpdating = True
        Cells(searchPosition.Row, searchPosition.Column).Select
    End If
    resetSearchParameters
End Sub

Public Function resetSearchParameters() 'run this function at the end of each "Main" sub
    'Reset match case and entire contents
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function
