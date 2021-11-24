VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Filtro"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnFiltrar1_Click()
Planilha7.Range("A2:K2").Clear


Planilha7.Range("A2").Value = usuario.Value
Planilha7.Range("E2").Value = supqa.Value
Planilha7.Range("K2").Value = programa.Value

ListBox3.RowSource = IntervaloDados


End Sub

Private Sub CommandButton1_Click()
Planilha7.Range("A2:K2").Clear

usuario.Value = ""
supqa.Value = ""
programa.Value = ""

ListBox3.RowSource = IntervaloDados
End Sub





Private Sub CommandButton2_Click()
Planilha7.Range("A2:K2").Clear

usuario.Value = ""
supqa.Value = ""
programa.Value = ""

ListBox3.RowSource = IntervaloDados
End Sub



Private Sub UserForm_Initialize()

InitMaxMin Me.Caption
Ht = Me.Height
Lg = Me.Width
    
Application.WindowState = xlMaximized

supqa.AddItem "ALEXANDRE CINTAS URBANO"
supqa.AddItem "ROGÉRIO DONIZETTI PINTO"
supqa.AddItem "RAPHAEL P. PEREIRA"

programa.AddItem "DIVERSOS"
programa.AddItem "ENGENHARIA"
programa.AddItem "LOGISTICA"
programa.AddItem "PRAETOR"
programa.AddItem "PROGRAMAS"
programa.AddItem "SUPER TUCANO"
programa.AddItem "KC-390"
programa.AddItem "TODOS"

End Sub
