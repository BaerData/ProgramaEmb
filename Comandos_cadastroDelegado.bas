Attribute VB_Name = "Comandos_cadastroDelegado"
Global bloqueado As Boolean

''Botao para salvar toda vez que algo for inserido no txtbox
Sub Cadastrar()
bloqueado = True
Dim tabela  As ListObject                       ''tabela = planilha em questao
Dim n       As Integer, id      As Integer      ''n = numero , id = chave de identificaçao
Set tabela = Planilha4.ListObjects(1)           '' tabela é atribuida a planilha4

id = Range("ID").Value

''Codigo para inserir os valores dentros das textbox e enviar para a tabela
n = tabela.Range.Rows.Count
tabela.Range(n, 1).Value = id
tabela.Range(n, 2).Value = TelaDelegadosCad.txtLogin.Value
tabela.Range(n, 3).Value = TelaDelegadosCad.txtNome.Value
tabela.Range(n, 4).Value = TelaDelegadosCad.txtArea.Value
tabela.Range(n, 5).Value = TelaDelegadosCad.txtSupProd.Value
tabela.Range(n, 6).Value = TelaDelegadosCad.txtSupQa.Value
tabela.Range(n, 7).Value = TelaDelegadosCad.txtIdCu.Value
tabela.Range(n, 8).Value = TelaDelegadosCad.txtTituloCu.Value
tabela.Range(n, 9).Value = TelaDelegadosCad.txtStatus.Value
tabela.Range(n, 10).Value = TelaDelegadosCad.txtDateAtribuicao.Value
tabela.Range(n, 11).Value = TelaDelegadosCad.txtDateVenc.Value
tabela.Range(n, 12).Value = TelaDelegadosCad.txtPrograma.Value

''tabela procura a coluna id e depois de inserir os valores acrescenta mais uma linha
TelaDelegadosCad.ListBox1.RowSource = ""
tabela.ListRows.Add
Range("ID").Value = id + 1
    
Call Atualizar_ListBox
MsgBox "Cadastro Realizado com sucesso", vbInformation, "Novo Delegado"
Call LimparCampos

bloqueado = False

End Sub


Sub Editar() ''Botao para editar toda vez que algo for inserido no txtbox
    
bloqueado = True
Dim Resp As VbMsgBoxResult
Dim tabela As ListObject                            ''tabela = planilha em questao
Dim n As Integer, L As Integer                      ''n = numero , L = linha
Set tabela = Planilha4.ListObjects(1)

n = TelaDelegadosCad.ListBox1.Value
L = tabela.Range.Columns().Find(n, , , xlWhole).Row

msgResp = MsgBox("Voce deseja editar?", vbYesNo)
If msgResp = vbYes Then
MsgBox "Cadastro editado"

''mesma funçao para acrescentar porem so faz a subtituiçao dos valores
tabela.Range(L, 2).Value = TelaDelegadosCad.txtLogin.Value
tabela.Range(L, 3).Value = TelaDelegadosCad.txtNome.Value
tabela.Range(L, 4).Value = TelaDelegadosCad.txtArea.Value
tabela.Range(L, 5).Value = TelaDelegadosCad.txtSupProd.Value
tabela.Range(L, 6).Value = TelaDelegadosCad.txtSupQa.Value
tabela.Range(L, 7).Value = TelaDelegadosCad.txtIdCu.Value
tabela.Range(L, 8).Value = TelaDelegadosCad.txtTituloCu.Value
tabela.Range(L, 9).Value = TelaDelegadosCad.txtStatus.Value
tabela.Range(L, 10).Value = TelaDelegadosCad.txtDateAtribuicao.Value
tabela.Range(L, 11).Value = TelaDelegadosCad.txtDateVenc.Value
tabela.Range(L, 12).Value = TelaDelegadosCad.txtPrograma.Value
Else

End If

bloqueado = False
Call Atualizar_ListBox
    
Call LimparCampos
    
End Sub

''Variavel para atualizar a caixa de texto com as colunas exibidas
Sub Atualizar_ListBox()
bloqueado = True
Dim tabela As ListObject
Set tabela = Planilha4.ListObjects(1)

TelaDelegadosCad.ListBox1.RowSource = tabela.DataBodyRange.Address(, , , True)      ''listbox atualiza visualmente atraves da busca que o rowsource faz na planilha

bloqueado = False
End Sub

''Variavel para excluir toda vez que algo for inserido no txtbox
Sub Excluir()
bloqueado = True

Dim tabela As ListObject                            ''tabela = planilha em questao
Dim n As Integer, L As Integer                      ''n = numero , L = linha
Dim Resp As VbMsgBoxResult
Set tabela = Planilha4.ListObjects(1)


n = TelaDelegadosCad.ListBox1.Value
L = tabela.Range.Columns().Find(n, , , xlWhole).Row

msgResp = MsgBox("Voce deseja excluir?", vbYesNo)
If msgResp = vbYes Then

TelaDelegadosCad.ListBox1.RowSource = ""
tabela.Range.Rows(L).Delete

Else

End If
Call Atualizar_ListBox
MsgBox "Item excluído com sucesso"
bloqueado = False
End Sub

''Variavel criada para quando terminar de registrar as txtbox limpar automaticamente os campos
Sub LimparCampos()

TelaDelegadosCad.txtLogin.Value = ""
TelaDelegadosCad.txtNome.Value = ""
TelaDelegadosCad.txtArea.Value = ""
TelaDelegadosCad.txtSupProd.Value = ""
TelaDelegadosCad.txtSupQa.Value = ""
TelaDelegadosCad.txtIdCu.Value = ""
TelaDelegadosCad.txtTituloCu.Value = ""
TelaDelegadosCad.txtStatus.Value = ""
TelaDelegadosCad.txtPrograma.Value = ""
''TelaDelegadosCad.txtDateAtribuicao.Value = ""
''TelaDelegadosCad.txtDateVenc.Value = ""


End Sub

''variavel para exibir tela de cadastro de delegados
Sub ExibirTela()
TelaDelegadosCad.Show
End Sub

''Filtro das colunas que estao na planilha3
Sub Filtro()

Dim base As Range               ''base = atribuida a planilha4
Dim crt As Range                ''crt  = colunas range tabela (range max. de colunas da tabela filtrada)
Dim filtrada

Set base = Planilha4.Range("A1").CurrentRegion
Set crt = Planilha4.Range("Q2:AA2")

base.AdvancedFilter xlFilterCopy, crt, Planilha3.Range("A1:k1")

Set filtrada = Planilha3.Range("A1").CurrentRegion
TelaDelegadosCad.ListBox1.RowSource = filtrada.Address


End Sub
