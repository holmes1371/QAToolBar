'QA Toolbar, v. 1.3.5
'Developed by Tom Holmes and Frank Castillo
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
    finalReset
    'ActiveWorkbook.Save 'uncomment this line to activate the autosave function.
End Sub

Sub autoHeaderIngest(control As IRibbonControl)
    autoHeader2
    If endIt = False Then
        Exit Sub
    End If
    finalReset
End Sub

Sub SheetFixIngest(control As IRibbonControl)
    SheetFixIngestF
    resetSearchParameters
    finalReset
End Sub

Sub autoHeaderFormatterIngest(control As IRibbonControl)
    autoHeader2
    SheetFixIngestF
    If endIt = False Then
       Exit Sub
    End If
End Sub

Sub manualNewUti(control As IRibbonControl)
    utiMode = "manual"
    autoHeaderUniquinizerIngestF
    finalReset
End Sub
Sub autoNewUti(control As IRibbonControl)
    utiMode = "auto"
    autoHeaderUniquinizerIngestF
    finalReset
End Sub

Sub findTradeID(control As IRibbonControl)
    SheetFixIngestF
    findID
    If foundOne = True Then
        Application.ScreenUpdating = True
        Cells(searchPosition.Row, searchPosition.Column).Select
    End If
    finalReset
End Sub
Public Function finalReset()
    resetSearchParameters
    refreshScreen
End Function
Public Function resetSearchParameters() 'run this function at the end of each "Main" sub
    'Reset match case and entire contents
    Cells.Replace what:="", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function
Public Function refreshScreen()
Application.ScreenUpdating = True
Application.CutCopyMode = False
End Function
