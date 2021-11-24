VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TelaDelegadosCad 
   Caption         =   "Cadastro dos Delegados"
   ClientHeight    =   10200
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   18840
   OleObjectBlob   =   "TelaDelegadosCad.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "TelaDelegadosCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lg As Single
Dim Ht As Single
Dim Fini As Boolean

Private Sub btnEditar_Click()
Call Editar
End Sub

Private Sub btnLimparCampos_Click()
Call LimparCampos
End Sub

Private Sub btnSalvar_Click()

Dim nlin        As Integer                               ''Nlin = numeros de linhas
If btnEditar.Value = True Then
    nlin = ListBox1.ListIndex
    If nlin = 0 Then
        MsgBox "Selecione um item para editar"
    Exit Sub
    ElseIf ListBox1.Value = 0 Then
        MsgBox "Selecione um item para editar"
    Exit Sub
    End If
    Call Editar
Else
    Call Cadastrar
End If

End Sub


Private Sub btnExcluir_Click()                           ''Nlin = numeros de linhas

Call Excluir

End Sub


Private Sub CommandButton2_Click()                       ''Nlin = numeros de linhas
TelaDelegadosCad.Hide
Home.Show
UserForm2.Hide


End Sub


Private Sub CommandButton3_Click()
UserForm2.Show
End Sub


Private Sub CommandButton4_Click()
TelaInspecao.Show
TelaDelegadosCad.Hide

End Sub


Private Sub CommandButton5_Click()

End Sub

Private Sub ListBox1_Change()

Dim nlin As Integer                                        ''Nlin = numeros de linhas

nlin = ListBox1.ListIndex

'Campos para listbox identificar os valores e limpar os campos

If nlin = -1 Then Exit Sub
If bloqueado = True Then Exit Sub
If ListBox1.Value = 0 Then
    txtLogin.Value = ""
    txtNome.Value = ""
    txtArea.Value = ""
    txtSupProd.Value = ""
    txtSupQa.Value = ""
    txtIdCu.Value = ""
    txtTituloCu.Value = ""
    txtStatus.Value = ""
    ''txtDateAtribuicao.Value = ""
    ''txtDateVenc.Value = ""
    txtPrograma.Value = ""
    
Else
    txtLogin.Value = ListBox1.List(nlin, 1)
    txtNome.Value = ListBox1.List(nlin, 2)
    txtArea.Value = ListBox1.List(nlin, 3)
    txtSupProd.Value = ListBox1.List(nlin, 4)
    txtSupQa.Value = ListBox1.List(nlin, 5)
    txtIdCu.Value = ListBox1.List(nlin, 6)
    txtTituloCu.Value = ListBox1.List(nlin, 7)
    txtStatus.Value = ListBox1.List(nlin, 8)
    ''txtDateAtribuicao.Value = ListBox1.List(nlin, 9)
    ''txtDateVenc.Value = ListBox1.List(nlin, 10)
    txtPrograma.Value = ListBox1.List(nlin, 11)
    
End If

End Sub




Private Sub UserForm_Initialize()

InitMaxMin Me.Caption
Ht = Me.Height
Lg = Me.Width
    
Application.WindowState = xlMaximized
    
'REGISTROS DOS CAMPOS DE LISTAGEM
''campo area
txtArea.AddItem "ENGENHARIA"
txtArea.AddItem "ENSAIOS"
txtArea.AddItem "LOGISTICA"
txtArea.AddItem "LOGISTICA MIP"
txtArea.AddItem "MONTAGEM ESTRUTURAL"
txtArea.AddItem "MONTAGEM Final"
txtArea.AddItem "MONTAGEM Final(ELETRICA)"
txtArea.AddItem "MONTAGEM Final(Interior)"
txtArea.AddItem "MONTAGEM Final(MECÂNICO)"
txtArea.AddItem "PINTURA"
txtArea.AddItem "PINTURA MOVEIS"
txtArea.AddItem "PPV PREPARAÇÃO"
txtArea.AddItem "PPV PREPARAÇÃO PARA VOO"
txtArea.AddItem "PPV PREPARAÇÃO PARA VOO ENSAIOS"
txtArea.AddItem "PROGRAMAS"
txtArea.AddItem "QUALIDADE LOGISTICA"
txtArea.AddItem "QUALIDADE QUARENTENA"
txtArea.AddItem "QUARENTENA"
txtArea.AddItem "SELAGEM"

''campo supervisor de produçao
txtSupProd.AddItem "ALESSANDRO PIRES DA SILVA"
txtSupProd.AddItem "ALEXANDRE CINTAS URBANO"
txtSupProd.AddItem "ANDRE DEMAMBRO"
txtSupProd.AddItem "ANDRÉ L.CORREA DO NASCIMENTO"
txtSupProd.AddItem "BRUNO GOMES RIBEIRO"
txtSupProd.AddItem "CLAUDINEI DA SILVA PEREIRA"
txtSupProd.AddItem "CLEITON DE OLIVEIRA"
txtSupProd.AddItem "DENY WALACE PAGANINI"
txtSupProd.AddItem "EDMILSON DA SILVA"
txtSupProd.AddItem "EDUARDO AP.BARNABE"
txtSupProd.AddItem "EVANIR RAMOS"
txtSupProd.AddItem "FABIO ARRUDA CAMARGO"
txtSupProd.AddItem "GIOVANI RODOLFO DA SILVA"
txtSupProd.AddItem "IDEMAURO BERTTI"
txtSupProd.AddItem "JULIO GRANZOTO JUNIOR"
txtSupProd.AddItem "LUCIANO DA CRUZ DUARTE"
txtSupProd.AddItem "MARCIO DE CASTRO YUKINO"
txtSupProd.AddItem "MARCOS PAULO GUIMARAES"
txtSupProd.AddItem "MARIO APARECIDO RIBEIRO"
txtSupProd.AddItem "ORLANDO HENRIQUE DE OLIVEIRA"
txtSupProd.AddItem "PAULO HENRIQUE GONZAGA"
txtSupProd.AddItem "PAULO VITOR REGASSO"
txtSupProd.AddItem "RAFAEL FERNANDES CAVALARO"
txtSupProd.AddItem "RAPHAEL RODRIGUES LEITE"
txtSupProd.AddItem "RENATO RODRIGUES"
txtSupProd.AddItem "ROBSON EDUARDO DOS SANTOS"
txtSupProd.AddItem "ROGERIO SIQUEIRA RAMOS DE OLIVEIRA"
txtSupProd.AddItem "THIAGO VENANCIO DE MATOS"
txtSupProd.AddItem "VANDERSON DE OLIVEIRA BARBOSA"

''campo Supervisor de qualidade
txtSupQa.AddItem "ALEXANDRE CINTAS URBANO"
txtSupQa.AddItem "RAPHAEL PINTO MOREIRA"
txtSupQa.AddItem "ROGERIO DONIZETTI PINTO"

''campo titulo do curriculo
txtTituloCu.AddItem "Conf Ensaios Voo (Delegado)-236"
txtTituloCu.AddItem "Conf Proc Produtivo (Delegado)-236"
txtTituloCu.AddItem "Conformidade Recebimento - Inspeção Básica (Delegado) - 236"
txtTituloCu.AddItem "Conformidade Quarentena - Inspeção(Delegado) - 236"
txtTituloCu.AddItem "Conf Ensaios Voo (Delegado)-236"
txtTituloCu.AddItem "Conformidade Ensaios Estruturais (Delegado) - 236"
txtTituloCu.AddItem "Conformidade Expedição - Inspeção(Delegado) - 236"
txtTituloCu.AddItem "Conformidade Set Up Instalacões de Ensaio (Delegado) - 236"
txtTituloCu.AddItem "Conformidade Revalidação de Estoque (Delegado) - 236"

''campo status
txtStatus.AddItem "Ativo"
txtStatus.AddItem "Em Qualificação"

''campo programas
txtPrograma.AddItem "DIVERSOS"
txtPrograma.AddItem "ENGENHARIA"
txtPrograma.AddItem "LOGISTICA"
txtPrograma.AddItem "PRAETOR"
txtPrograma.AddItem "PROGRAMAS"
txtPrograma.AddItem "SUPER TUCANO"
txtPrograma.AddItem "KC-390"
txtPrograma.AddItem "TODOS"

 
Call Atualizar_ListBox

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Cancel = True
End Sub

'Instrução para redimensionar o formulário dos conteúdos do formulário (textbox, combobox,etc)
Private Sub UserForm_Resize()
    Dim RtL As Single, RtH As Single
        If Me.Width < 300 Or Me.Height < 200 Or Fini Then Exit Sub
        RtL = Me.Width / Lg
        RtH = Me.Height / Ht
        Me.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub
'============================================================================================================================

Private Sub UserForm_Terminate()
    Fini = True
    Planilha3.Range("A2:L2").Clear

txtFiltroNome.Value = ""
txtFiltroArea.Value = ""
txtFiltroStatus.Value = ""

End Sub

Private Sub btnFiltrar_Click()

''DataIni = txtFiltroDataAT1.Value


Planilha3.Range("A2:L2").Clear

Planilha3.Range("C2").Value = txtFiltroNome.Value
Planilha3.Range("D2").Value = txtFiltroArea.Value
Planilha3.Range("I2").Value = txtFiltroStatus.Value
Planilha3.Range("F2").Value = txtFiltroSupQA.Value
Planilha3.Range("L2").Value = txtFiltroProg.Value
''Planilha3.Range("J2").Value = txtFiltroDataAT1


ListBox1.RowSource = IntervaloDados

End Sub

Private Sub CommandButton1_Click()
Planilha3.Range("A2:L2").Clear

txtFiltroNome.Value = ""
txtFiltroArea.Value = ""
txtFiltroStatus.Value = ""
txtFiltroSupQA.Value = ""
ListBox1.RowSource = IntervaloDados

End Sub
