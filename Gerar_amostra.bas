Attribute VB_Name = "Gerar_amostra"
Sub Gerar_Amostra()
'
' Gerar_Amostra

Qtd_nota = Worksheets("Dados").Range("D8")
If Qtd_nota < 1 Then GoTo nãoamostra
Data_Amostra = Worksheets("Amostra").Range("N1")
Data_RelNovo = Worksheets("CDs").Range("K1")
If Data_RelNovo > Data_Amostra Then GoTo Iniciar


Hora_Amostra = Worksheets("Amostra").Range("O1")
Hora_RelNovo = Worksheets("CDs").Range("L1")
If Hora_Amostra > Hora_RelNovo Then GoTo Atualize


Iniciar:

Sheets("Amostra").Select
Range("A4:I4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.EntireRow.Delete
Range("A4").Select

Sheets("CDs").Select
Range("A4:I4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'Inserir Formulas
Sheets("Amostra").Select
Range("A4:I4").Select
ActiveSheet.Paste

On Error Resume Next
''ActiveSheet.ShowAllData
Range("J4").Select
ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,(Dados!R8C4)+(1/Dados!R8C4))"
Range("K4").Select
ActiveCell.FormulaR1C1 = "=IF(RC[-1]<=Dados!R9C4,""Extrair"",""Não"")"
Range("L3").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[-1],""Extrair"")"
Range("J4:K4").Select
Selection.Copy
Range("A60000").End(xlUp).Offset(0, 9).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste
Application.CutCopyMode = False

'Rodar a logica até ter o numero de amostra necessário
Valor_real = Worksheets("Dados").Range("D9")
TESTE = Worksheets("Amostra").Range("L3")

Do While TESTE <> Valor_real
If TESTE = Valor_real Then GoTo Continue
Range("I4").Select
Calculate

Valor_real = Worksheets("Dados").Range("D9")
TESTE = Worksheets("Amostra").Range("L3")
Loop

Continue:
Range("J4:K4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveSheet.Range("$A$3:$K$60000").AutoFilter Field:=11, Criteria1:="Não"
 
'Sub LocalizarPrimeira()
    Dim Excluir As String
    Dim Rng As Range
    Dim Nome As String
    Não = "Não"
    NomeProcurado = Não
    If Trim(NomeProcurado) <> "NÃO" Then
    With Sheets("Amostra").Range("K:K") 'Nome da planilha e área de procura
            Set Rng = .Find(What:=NomeProcurado, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
            
                Application.GoTo Rng, True
            Else
                GoTo Continue1:
            End If
        End With
    End If

Continue1:
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
On Error Resume Next
ActiveSheet.ShowAllData
Range("L3").Select
Selection.ClearContents
Range("J4:K4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A60000").End(xlUp).Offset(1, 0).Select
Range("N1") = Date
Range("O1") = Time
MsgBox ("Relatório gerado!")

Exit Sub

Atualize:
MsgBox ("Atualize o Relatório de CDs para depois gerar nova Amostra!")

Exit Sub
nãoamostra:
MsgBox ("Não tem a quantidade minima de dados para gerar nova Amostra!")

End Sub

