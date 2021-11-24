VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Configuraçoes"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  If TextBox1 = "vba" Then
    ''Unload Home
    UserForm1.Hide
    ''Home.Show
    Application.Visible = True
  End If
End Sub

Private Sub CommandButton2_Click()
UserForm1.Hide

End Sub

Private Sub CommandButton7_Click()
   Application.Visible = False
End Sub
