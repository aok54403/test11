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

Private Sub CB1_Click()
    If txtPW.Value = "1234" Then
        Unload Me
    Else
        With txtPW
            .Value = ""
            .SetFocus
        End With
    End If
End Sub
