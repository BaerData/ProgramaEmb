Attribute VB_Name = "Funcao"
''Funçao criada para realizar encontrar um intervalo das celulas para fazer o filtro
Function IntervaloDados() As String

Dim base As Range, intC As Range, destino As Range

Set base = Planilha4.Range("A1").CurrentRegion
Set intC = Planilha7.Range("A1:K2")
Set destino = Planilha7.Range("A4:K4")

base.AdvancedFilter xlFilterCopy, intC, destino
    
IntervaloDados = destino.CurrentRegion.Offset(1, 0).Address(, , , True)

End Function



