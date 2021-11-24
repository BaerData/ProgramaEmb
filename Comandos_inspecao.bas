Attribute VB_Name = "Comandos_inspecao"
Global bloqueado2 As Boolean

''Variavel para salvar toda vez que algo for inserido no txtbox
Sub Cadastrar2()
bloqueado2 = True
Dim tabela  As ListObject
Dim n       As Integer, id      As Integer
Set tabela = Planilha6.ListObjects(1)

id = Range("IIDD").Value

n = tabela.Range.Rows.Count
tabela.Range(n, 1).Value = id

'Dados do delegado monitorado
tabela.Range(n, 2).Value = TelaInspecao.txtNota.Value
tabela.Range(n, 3).Value = TelaInspecao.txtGravidade.Value
tabela.Range(n, 4).Value = TelaInspecao.txtNome.Value
tabela.Range(n, 5).Value = TelaInspecao.txtNomeResp.Value
tabela.Range(n, 6).Value = TelaInspecao.txtObs.Value
tabela.Range(n, 7).Value = TelaInspecao.txtSupProd.Value
tabela.Range(n, 8).Value = TelaInspecao.txtSupQa.Value
tabela.Range(n, 9).Value = TelaInspecao.txtCTAtual.Value
tabela.Range(n, 10).Value = TelaInspecao.txtArea.Value

'Descrição da análise
tabela.Range(n, 11).Value = TelaInspecao.txtDoc.Value
tabela.Range(n, 12).Value = TelaInspecao.txtAplic.Value

tabela.Range(n, 13).Value = TelaInspecao.txtProblema.Value



'Responsável pela análise
tabela.Range(n, 14).Value = TelaInspecao.txtCargoResp.Value

tabela.Range(n, 15).Value = TelaInspecao.txtChapaResp.Value
tabela.Range(n, 16).Value = TelaInspecao.txtDateEncResp.Value
tabela.Range(n, 17).Value = TelaInspecao.txtProgramas.Value

TelaInspecao.ListBox2.RowSource = ""
tabela.ListRows.Add
Range("IIDD").Value = id + 1
    
Call Atualizar_ListBox2
MsgBox "Cadastro Realizado com sucesso", vbInformation, "Novo Delegado"
Call LimparCampos2



bloqueado2 = False

End Sub

''Variavel para editar toda vez que algo for inserido no txtbox
Sub Editar2()
bloqueado2 = True
Dim Resp As VbMsgBoxResult
Dim tabela As ListObject
Dim n As Integer, linha As Integer
Set tabela = Planilha6.ListObjects(1)

n = TelaInspecao.ListBox2.Value
L = tabela.Range.Columns().Find(n, , , xlWhole).Row

msgResp = MsgBox("Voce deseja editar?", vbYesNo)
If msgResp = vbYes Then
MsgBox "Cadastro editado"

'Dados do delegado monitorado
'Dados do delegado monitorado
tabela.Range(L, 2).Value = TelaInspecao.txtNota.Value
tabela.Range(L, 3).Value = TelaInspecao.txtGravidade.Value
tabela.Range(L, 4).Value = TelaInspecao.txtNome.Value
tabela.Range(L, 5).Value = TelaInspecao.txtNomeResp.Value
tabela.Range(L, 6).Value = TelaInspecao.txtObs.Value
tabela.Range(L, 7).Value = TelaInspecao.txtSupProd.Value
tabela.Range(L, 8).Value = TelaInspecao.txtSupQa.Value
tabela.Range(L, 9).Value = TelaInspecao.txtCTAtual.Value
tabela.Range(L, 10).Value = TelaInspecao.txtArea.Value

'Descrição da análise
tabela.Range(L, 11).Value = TelaInspecao.txtDoc.Value
tabela.Range(L, 12).Value = TelaInspecao.txtAplic.Value
tabela.Range(L, 13).Value = TelaInspecao.txtProblema.Value

'Responsável pela análise
tabela.Range(L, 14).Value = TelaInspecao.txtCargoResp.Value

tabela.Range(L, 15).Value = TelaInspecao.txtChapaResp.Value
tabela.Range(L, 16).Value = TelaInspecao.txtDateEncResp.Value
tabela.Range(L, 17).Value = TelaInspecao.txtProgramas.Value
Else

End If

bloqueado2 = False
Call Atualizar_ListBox2
    
Call LimparCampos2

End Sub

'Variavel para atualizar a caixa de texto com as colunas exibidas
Sub Atualizar_ListBox2()
bloqueado2 = True
Dim tabela As ListObject
Set tabela = Planilha6.ListObjects(1)

TelaInspecao.ListBox2.RowSource = tabela.DataBodyRange.Address(, , , True)


bloqueado2 = False
End Sub

''Variavel para excluir toda vez que algo for inserido no txtbox
Sub Excluir2()
bloqueado2 = True
Dim tabela As ListObject
Dim n As Integer, L As Integer
Dim Resp2 As VbMsgBoxResult
Set tabela = Planilha6.ListObjects(1)

n = TelaInspecao.ListBox2.Value
L = tabela.Range.Columns().Find(n, , , xlWhole).Row

msgResp2 = MsgBox("Voce deseja excluir?", vbYesNo)
If msgResp2 = vbYes Then
TelaInspecao.ListBox2.RowSource = ""
tabela.Range.Rows(L).Delete

Else

End If

Call Atualizar_ListBox2
MsgBox "Item excluído com sucesso"
bloqueado2 = False
End Sub

''Limpar campos todas vez que houver um registro
Sub LimparCampos2()

TelaDelegadosCad.txtLogin.Value = ""

'Dados do delegado monitorado
TelaInspecao.txtNome.Value = ""
TelaInspecao.txtSupProd.Value = ""
TelaInspecao.txtSupQa.Value = ""
TelaInspecao.txtCTAtual.Value = ""
TelaInspecao.txtArea.Value = ""

'Descrição da análise
TelaInspecao.txtDoc.Value = ""
TelaInspecao.txtAplic.Value = ""
TelaInspecao.txtNota.Value = ""
TelaInspecao.txtProblema.Value = ""
TelaInspecao.txtGravidade.Value = ""
TelaInspecao.txtObs.Value = ""

'Responsável pela análise
TelaInspecao.txtNomeResp.Value = ""
TelaInspecao.txtCargoResp.Value = ""
TelaInspecao.txtChapaResp.Value = ""
TelaInspecao.txtProblema.Value = ""
TelaInspecao.txtProgramas.Value = ""

End Sub


Sub Exibir2()
TelaInspecao.Show
End Sub

Sub ExibirHome()
Home.Show

End Sub
