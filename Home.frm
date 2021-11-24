VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Home 
   Caption         =   "Home Embraer GPX"
   ClientHeight    =   10140
   ClientLeft      =   90
   ClientTop       =   315
   ClientWidth     =   18015
   OleObjectBlob   =   "Home.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Dim Lg As Single
''Dim Ht As Single
''Dim Fini As Boolean


Private Sub ajuda_Click()
Dim caminhoajuda As String
CallWebPage ("https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44/Delegao/Formulario_de_Registro_dos_Delegados/Guia%20de%20Ajuda%20Monitoramento%20dos%20Delegados.pptx")
End Sub

Private Sub btnCadastroDelegado_Click() ''Tela de cadastro
Call ExibirTela
Home.Hide


End Sub

Private Sub btnOff_Click()
Unload Me
ThisWorkbook.Close
End Sub

Private Sub CommandButton2_Click()
''Application.Visible = False
Home.Hide
Call Exibir2
End Sub



Private Sub CommandButton4_Click()
Dim caminho As String
caminho = "https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44/Delegao/Formulario_de_Registro_dos_Delegados/CDs%20-%20Monitoramento%20dos%20Delegados%20Rev.-.xlsm"
Workbooks.Open caminho
End Sub



Private Sub CommandButton6_Click()
UserForm1.Show
Application.Visible = True


End Sub


Private Sub Image4_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Unload Me
ThisWorkbook.Close
End Sub



Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label12_Click()
CallWebPage2 ("https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44")
End Sub

Private Sub powerbi_Click()
CallWebPage ("https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44/SitePages/Monitoramento-dos-Delegados.aspx")
End Sub

Private Sub UserForm_Initialize()

''InitMaxMin Me.Caption
    ''Ht = Me.Height
    ''Lg = Me.Width
    
   '' Application.WindowState = xlMaximized
    ''Me.Height = Application.Height
   '' Me.Width = Application.Width
   '' Me.Left  = Application.Left
    ''Me.Top = Application.Top
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then Cancel = True

End Sub

'Instrução para redimensionar o formulário dos conteúdos do formulário (textbox, combobox,etc)
''Private Sub UserForm_Resize()
   '' Dim RtL As Single, RtH As Single
      ''  If Me.Width < 200 Or Me.Height < 100 Or Fini Then Exit Sub
      ''  RtL = Me.Width / Lg
       '' RtH = Me.Height / Ht
      ''  Me.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub
Private Sub UserForm_Terminate()
    Fini = True
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Dim powerbi
powerbi = "https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44/SitePages/Monitoramento-dos-Delegados.aspx"
End Sub

Private Sub WebBrowser1_StatusTextChange2(ByVal Text As String)
Dim PortalQA
PortalQA = "https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44"
End Sub

Private Sub WebBrowser12_StatusTextChange(ByVal Text As String)
Dim ajuda
ajuda = "https://embraer.sharepoint.com/sites/QualidadeDefesaEstrutura44/Delegao/Formulario_de_Registro_dos_Delegados/Guia%20de%20Ajuda%20Monitoramento%20dos%20Delegados.pptx"
End Sub
