'QA Toolbar, v. 1.4
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
    Set startCell = startPrep
    autoHeader2
    autoHeaderUniquinizerIngestF
    finalReset (startCell)
    'ActiveWorkbook.Save 'uncomment this line to activate the autosave function.
End Sub

Sub autoHeaderIngest(control As IRibbonControl)
    Dim startCell As Object: Set startCell = startPrep
    
    autoHeader2
    Call finalReset(1, 1)
    If endIt = False Then
        Exit Sub
    End If
    
    
End Sub

Sub SheetFixIngest(control As IRibbonControl)
    Dim startCell As Object: Set startCell = startPrep
    
    SheetFixIngestF
    resetSearchParameters
    Call finalReset(startCell.row, startCell.column)
End Sub

Sub autoHeaderFormatterIngest(control As IRibbonControl)
    Dim startCell As Object: Set startCell = startPrep
    
    autoHeader2
    SheetFixIngestF
    If endIt = False Then
       Exit Sub
    End If
    Call finalReset(2, 1)
    
End Sub

Sub manualNewUti(control As IRibbonControl)
    Set startCell = startPrep
    utiMode = "manual"
    autoHeaderUniquinizerIngestF ("clicked")
    finalReset
End Sub
Sub autoNewUti(control As IRibbonControl)
    Set startCell = startPrep
    utiMode = "auto"
    autoHeaderUniquinizerIngestF ("clicked")
    finalReset
End Sub

Sub findTradeID(control As IRibbonControl)
    Dim startCell As Object: Set startCell = startPrep
    SheetFixIngestF
    findID ("clicked")
    Call finalReset
    
End Sub

Public Function resetSearchParameters() 'run this function at the end of each "Main" sub
    'Reset match case and entire contents
    Cells.Replace what:="", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function
Function startPrep() As Object

    Dim startCell As Object: Set startCell = ActiveCell
    Set startPrep = startCell
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
End Function

Sub finalReset(Optional row As Variant, Optional col As Variant)
    
       
    resetSearchParameters
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    
    If IsMissing(row) = False And IsMissing(col) = False Then
        Cells(row, col).Activate
    End If
        
End Sub
