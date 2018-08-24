VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplitSelectForm 
   Caption         =   "Split File by Columns"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   OleObjectBlob   =   "SplitSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplitSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelBtn_Click()
    Unload Me
    End
End Sub

Private Sub clearBtn_Click()
    Me.orderList.Clear
    Erase fileSplitArr
End Sub

Private Sub okBtn_Click()
    Unload Me
End Sub

Private Sub selectCombo_Change()
    Dim uniqBool As Boolean
    Dim listItem As Integer
    Dim lstCount As Integer
    Dim selValue As Integer
    Dim i As Integer
    
    uniqBool = True
    With Me
        If .orderList.TopIndex > -1 Then
            For i = 0 To (.orderList.listCount - 1)
                If .orderList.List(i) = .selectCombo.Value Then
                    uniqBool = False
                End If
            Next i
        End If
        If uniqBool = True And .selectCombo.Value <> "" Then
            .orderList.AddItem .selectCombo.Value
        End If
    End With
End Sub
