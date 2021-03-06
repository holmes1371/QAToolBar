Public othStr As String
Public tabBool As Boolean, _
       smcBool As Boolean, _
       cmmBool As Boolean, _
       spaBool As Boolean, _
       othBool As Boolean

Option Explicit

Public Sub textToCol(control As IRibbonControl)
    Dim objRange1 As Range
    Dim mylastcell
    
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
 
    'Set up the range
    Set objRange1 = Range("A1:A" & mylastcell.row)
    
   'Do the parse
    objRange1.TextToColumns _
    Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=tabBool, _
        Semicolon:=smcBool, _
        Comma:=cmmBool, _
        Space:=spaBool, _
        Other:=othBool, _
        OtherChar:=othStr
        
    Columns.AutoFit
    
' ActiveWorkbook.Save 'uncomment this line to activate the autosave function.

End Sub

Public Sub boxChecked(control As IRibbonControl, pressed As Boolean)
    If control.ID = "TabCheckBox" Then tabBool = pressed
    If control.ID = "SmcCheckBox" Then smcBool = pressed
    If control.ID = "CmmCheckBox" Then cmmBool = pressed
    If control.ID = "SpaCheckBox" Then spaBool = pressed
    If control.ID = "OthCheckBox" Then othBool = pressed
End Sub

Public Sub OthValue_onChange(control As IRibbonControl, Text As String)
    If (Text <> "" And Text <> Chr(32)) Then
        othBool = True
        othStr = Text & ", " & Asc(Text)
    ElseIf (Text = "") Then
        othBool = False
        othStr = Text
    End If
End Sub


