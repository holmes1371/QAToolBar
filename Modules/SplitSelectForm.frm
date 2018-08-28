VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplitSelectForm 
   Caption         =   "Split File by Columns"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
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
    Dim i As Integer
    With Me
        For i = 0 To .optionList.ListCount
            .optionList.Selected(i) = False
        Next i
        .selectedList.Clear
    End With
    'Erase fileSplitArr
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub okBtn_Click()
    Unload Me
End Sub

Private Sub selectCombo_Change()

End Sub

Private Sub optionList_Click()

End Sub

Private Sub selBtnAdd_Click()
    Dim i As Integer
    Dim j As Integer
    Dim uniq As Boolean
    
    uniq = True
    With Me
        For i = 0 To (.optionList.ListCount - 1)
            If .optionList.Selected(i) = True Then
                For j = 0 To (.selectedList.ListCount - 1)
                    If .optionList.List(i) = .selectedList.List(j) Then
                        uniq = False
                    End If
                Next j
                If uniq = True Then .selectedList.AddItem .optionList.List(i)
            End If
            uniq = True
        Next i
    End With
End Sub

Private Sub selBtnRem_Click()
    Dim i As Integer
    Dim j As Integer
    Dim remArr() As Integer
    With Me
        For i = (.selectedList.ListCount - 1) To 0 Step -1
            If .selectedList.Selected(i) = True Then
                .selectedList.RemoveItem (i)
            End If
        Next i
    End With
End Sub

Private Sub selListMoveDn_Click()
    Dim i As Integer
    Dim tmp As String
    
    uniq = True
    With Me
        For i = (.selectedList.ListCount - 1) To 0 Step -1
            If i = (.selectedList.ListCount - 1) _
            And .selectedList.Selected(.selectedList.ListCount - 1) = True Then Exit For
            If .selectedList.Selected(i) = True _
            And ((i + 1) <= .selectedList.ListCount - 1) Then
                tmp = .selectedList.List(i)
                .selectedList.List(i) = .selectedList.List(i + 1)
                .selectedList.List(i + 1) = tmp
                .selectedList.Selected(i + 1) = True
                .selectedList.Selected(i) = False
            End If
        Next i
    End With
End Sub

Private Sub selListMoveUp_Click()
    Dim i As Integer
    Dim tmp As String
    
    uniq = True
    With Me
        For i = 0 To (.selectedList.ListCount - 1)
            If i = 0 And .selectedList.Selected(0) = True Then Exit For
            If .selectedList.Selected(i) = True And (i - 1) >= 0 Then
                tmp = .selectedList.List(i - 1)
                .selectedList.List(i - 1) = .selectedList.List(i)
                .selectedList.List(i) = tmp
                .selectedList.Selected(i - 1) = True
                .selectedList.Selected(i) = False
            End If
        Next i
    End With
End Sub
