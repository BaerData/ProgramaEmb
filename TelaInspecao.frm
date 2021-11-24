VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TelaInspecao 
   Caption         =   "Registro do Problema"
   ClientHeight    =   10155
   ClientLeft      =   135
   ClientTop       =   555
   ClientWidth     =   18690
   OleObjectBlob   =   "TelaInspecao.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TelaInspecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Lg As Single
Dim Ht As Single
Dim Fini As Boolean

 


Private Sub btnLimparCampos_Click()
Call LimparCampos2
End Sub

Private Sub btnSalvar_Click()

Call Cadastrar2

End Sub


Private Sub btnExcluir_Click()

Call Excluir2
    
End Sub


Private Sub CommandButton1_Click()
TelaInspecao.Hide
Home.Show
UserForm2.Hide
End Sub


Private Sub CommandButton2_Click()
Call Editar2
End Sub

Private Sub CommandButton3_Click()
UserForm2.Show
End Sub

Private Sub CommandButton5_Click()
TelaDelegadosCad.Show
TelaInspecao.Hide
End Sub

Private Sub ListBox2_Change()
Dim nlin As Integer

nlin = ListBox2.ListIndex

'Campos para a listbox identificar os valores e limpar os campos
If nlin = -1 Then Exit Sub
If bloqueado2 = True Then Exit Sub
If ListBox2.Value = 0 Then
    txtNota.Value = ""
    txtGravidade.Value = ""
    txtNome.Value = ""
    txtNomeResp.Value = ""
    txtObs.Value = ""
    txtSupProd.Value = ""
    txtSupQa.Value = ""
    txtCTAtual.Value = ""
    txtArea.Value = ""
    txtDoc.Value = ""
    txtAplic.Value = ""
    txtProblema.Value = ""
    txtCargoResp.Value = ""
    txtChapaResp.Value = ""
    ''txtDateEncResp.Value = ""
    txtProgramas.Value = ""
    'Dados do delegado monitorado
Else
    
    txtNota.Value = ListBox2.List(nlin, 1)
    txtGravidade.Value = ListBox2.List(nlin, 2)
    txtNome.Value = ListBox2.List(nlin, 3)
    txtNomeResp.Value = ListBox2.List(nlin, 4)
    txtObs.Value = ListBox2.List(nlin, 5)
    txtSupProd.Value = ListBox2.List(nlin, 6)
    txtSupQa.Value = ListBox2.List(nlin, 7)
    txtCTAtual.Value = ListBox2.List(nlin, 8)
    txtArea.Value = ListBox2.List(nlin, 9)
    txtDoc.Value = ListBox2.List(nlin, 10)
    txtAplic.Value = ListBox2.List(nlin, 11)
    txtProblema.Value = ListBox2.List(nlin, 12)
    txtCargoResp.Value = ListBox2.List(nlin, 13)
    txtChapaResp.Value = ListBox2.List(nlin, 14)
    ''txtDateEncResp.Value = ""
    txtProgramas.Value = ListBox2.List(nlin, 16)
    
End If
End Sub




Private Sub txtNome_Change()

End Sub



Private Sub UserForm_Initialize()

InitMaxMin Me.Caption                                                   ''variaveis para redimencionar o tamanho da janela do app
    Ht = Me.Height
    Lg = Me.Width
    
Application.WindowState = xlMaximized
    ''Me.Height = Application.Height
   '' Me.Width = Application.Width
   '' Me.Left  = Application.Left
    ''Me.Top = Application.Top
    
'REGISTROS DOS CAMPOS DE LISTAGEM
    
''campo supervisor de qualidade
txtSupQa.AddItem "ALEXANDRE CINTAS URBANO"
txtSupQa.AddItem "RAPHAEL PINTO OLIVEIRA"
txtSupQa.AddItem "ROGERIO DONIZETTI PINTO"

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

''campo tipos de documentos
txtDoc.AddItem "Nota CD"
txtDoc.AddItem "Gate"
txtDoc.AddItem "OP/OM"

''campo problemas
txtProblema.AddItem "Falta participante na análise preliminar"
txtProblema.AddItem "Falta de informação"
txtProblema.AddItem "Falta requisito de projeto"
txtProblema.AddItem "Erro na informação"
txtProblema.AddItem "Aplic em desacordo com a não conformidade"
txtProblema.AddItem "Operador não qualificado"
txtProblema.AddItem "Refluxo ou fluxo incompleto"
txtProblema.AddItem "Anexo incompleto"
txtProblema.AddItem "Duas Não conformidade distintas no mesma CD"
txtProblema.AddItem "Erro na atribuição de CTO"
txtProblema.AddItem "CD fornecedor estrangeiro idioma PT"
txtProblema.AddItem "NFF - CD cancelada não conforme inexistente"
txtProblema.AddItem "Preenchimento desacordo 5529"
txtProblema.AddItem "CD Engenharia MLB idioma ingles."
txtProblema.AddItem "Não existe anexo"
txtProblema.AddItem "Erro Área responsável"
txtProblema.AddItem "Nenhum Problema Encontrado"

''campo programas
txtProgramas.AddItem "DIVERSOS"
txtProgramas.AddItem "ENGENHARIA"
txtProgramas.AddItem "LOGISTICA"
txtProgramas.AddItem "PRAETOR"
txtProgramas.AddItem "PROGRAMAS"
txtProgramas.AddItem "SUPER TUCANO"
txtProgramas.AddItem "KC-390"
txtProgramas.AddItem "TODOS"

''campo gravidades
txtGravidade.AddItem "Alta"
txtGravidade.AddItem "Médio"
txtGravidade.AddItem "Baixo"

''campo aplic
txtAplic.AddItem "APLIC001"
txtAplic.AddItem "APLIC002"
txtAplic.AddItem "APLIC003"
txtAplic.AddItem "APLIC009"
txtAplic.AddItem "APLIC018"
txtAplic.AddItem "APLIC020"
txtAplic.AddItem "APLIC023"

''campo responsavel pela analise
txtCargoResp.AddItem "CRM"
txtCargoResp.AddItem "Técnico Qualidade"

''campo desc CT atual
txtCTAtual.AddItem "GPX KC390 - Mont.Final - MF1"
txtCTAtual.AddItem "GPX050-Bordo de Fuga Asa EMB190"
txtCTAtual.AddItem "GPX173 -Manutenção / Config.Prototipos"
txtCTAtual.AddItem "GPX217 - Logística Ind. de  Distribuição"
txtCTAtual.AddItem "GPX260-Logística Central"
txtCTAtual.AddItem "GPX572 - Modernização de Aeronaves"
txtCTAtual.AddItem "GPX790- Mont. Wing Stub Legacy 450/ 500"
txtCTAtual.AddItem "GPX-Equipagem / Praetor Greenish"
txtCTAtual.AddItem "GPX-Interior / Praetor Greenish"
txtCTAtual.AddItem "GPX-Mont. Final e Testes / Praetor Green"
txtCTAtual.AddItem "GPX-PPV Praetor Greenish"
txtCTAtual.AddItem "SJC_PPV Entrega/Praetor Greenish"

Call Atualizar_ListBox2

End Sub

''desabilitar o fechamento da janela pelo icone "X"
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Cancel = True
End Sub

'Instrução para redimensionar o formulário dos conteúdos do formulário (textbox, combobox,etc)
Private Sub UserForm_Resize()
    Dim RtL As Single, RtH As Single                                        ''RtL = resize left , RtH = resize right
        If Me.Width < 300 Or Me.Height < 200 Or Fini Then Exit Sub
        RtL = Me.Width / Lg
        RtH = Me.Height / Ht
        Me.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub
'============================================================================================================================

Private Sub UserForm_Terminate()
    Fini = True
End Sub
