VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "パスワード認証"
   ClientHeight    =   1716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3444
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'確認ボタン
Private Sub CB1_Click()
    Dim pw As String
    Me.txtID.SetFocus
    pw = worksheetfunction.XLookup(Me.txtID, Sheets("PW").Range("a:a"), Sheets("PW").Range("b:b"))
    
    If pw = Me.txtPW Then
        Unload Me
    Else
        MsgBox "ID又はパスワードが違います"
        txtPW.Value = ""
        
    End If

End Sub

'キャンセルボタン
Private Sub CB2_Click()
    Application.DisplayAlerts = False
    ThisWorkbook.Close False
End Sub

'×ボタン
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = vbFormControlMenu Then
'        Cancel = True
'    End If
'End Sub
