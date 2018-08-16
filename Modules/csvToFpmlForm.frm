VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} csvToFpmlForm 
   Caption         =   "CSV To FPML Converter"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "csvToFpmlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "csvToFpmlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBrowseFldDest_Click()
    csvToFpmlForm.textBoxFldDest.Value = selectFile("fldr")
End Sub

Private Sub btnBrowseTaxFile_Click()
    csvToFpmlForm.textBoxTaxFile.Value = selectFile("file")
End Sub

Private Sub btnConvCancel_Click()
    Unload Me
End Sub

Private Sub btnConvClear_Click()
    With csvToFpmlForm
        .textBoxFldDest.Value = ""
        .textBoxTaxFile.Value = ""
    End With
    taxFilePath = ""
    cnvFldrPath = ""
End Sub

Private Sub btnConvOk_Click()
    Dim fldValue As String
    Dim taxValue As String
    Dim usrInput As Boolean
    fldValue = csvToFpmlForm.textBoxFldDest.Value
    taxValue = csvToFpmlForm.textBoxTaxFile.Value
    usrInput = fldValue = "" And taxValue = ""
    Unload Me
    parseFields taxValue, fldValue
End Sub

Private Sub labelConvFldDest_Click()

End Sub

Private Sub labelConvTaxFile_Click()

End Sub

Private Sub textBoxFldDest_Change()

End Sub

Private Sub textBoxTaxFile_Change()

End Sub

Private Sub UserForm_Click()

End Sub
