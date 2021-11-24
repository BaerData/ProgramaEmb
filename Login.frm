VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Tela de login"
   ClientHeight    =   9420
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   16755
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Form"
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

''If txtLogin = "adm" And txtSenha = "123" Then

Dim usuario As String
Dim senha As String
Dim pasta As String

pasta = "usuarios"
linha = 2
Do While ThisWorkbook.Sheets(pasta).Cells(linha, 1) <> ""
    If UCase(ThisWorkbook.Sheets(pasta).Cells(linha, 2)) = UCase(txtLogin) And UCase(ThisWorkbook.Sheets(pasta).Cells(linha, 3)) = UCase(txtSenha) Then
    Unload Login
    MsgBox "Bem Vindo!", vbInformation
    Home.Show
    UserForm1.Hide
    Application.Visible = False
    Exit Sub
    End If
    linha = linha + 1
    Loop
    
    If usuario = "" Or senha = "" Then
    MsgBox "Usuario ou senha incorretas", vbCritical, "Erro"
    


    
    End If
    
    
    
    
    
    
''Else
    ''
  ''End If
  
End Sub

Private Sub CommandButton2_Click()
Unload Me
ThisWorkbook.Close
End Sub

