Attribute VB_Name = "Iniciar_busca"
Sub CDs()
Attribute CDs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CDs Macro
      
MsgBox ("Você esta logado no SAP?")

Data_I = Worksheets("Dados").Range("D4")
Data_F = Worksheets("Dados").Range("D5")
login = Worksheets("Dados").Range("A5")

If login = "" Then GoTo Preencha
If Data_I = "" Then GoTo Preencha
If Data_F = "" Then GoTo Preencha

'Exibir planilhas de dados
Sheets("Dados").Select
Sheets("Temp").Visible = True

Dim y, X As String
    Dim a As Integer
        y = ActiveWorkbook.Name
        a = Len(y)
            X = Left(y, (Len(y) - 4))

'Abre nova janela do SAP
Set SapGuiAuto = GetObject("SAPGUI")          'Utiliza o objeto da interface gráfica do SAP
Set SAPApp = SapGuiAuto.GetScriptingEngine    'Conecta ao SAP que está rodando no momento
On Error GoTo SAP_OFF
Set SAPCon = SAPApp.Children(0)               'Encontra o primeiro sistema que está conectado
Set session = SAPCon.Children(0)              'Encontra a primeira sessão (janela) dessa conexão

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
End If
If Not IsObject(Connection) Then
End If
If Not IsObject(session) Then
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

' Limpeza_Inicial_Ordens
Sheets("Amostra").Select
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A4:I4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A4").Select
Sheets("Resumo").Select
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A5:D5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A5").Select
Sheets("CDs").Select
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A4:I4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
On Error Resume Next
''ActiveSheet.ShowAllData
Range("A4").Select
Sheets("Temp").Select
On Error Resume Next
''ActiveSheet.ShowAllData
Columns("A:H").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select



'Realizar busca das CDs no SAP e verificar CDs abertas
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZQLRQM150"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtS_QMART-LOW").Text = "CD"
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtS_QMDAT-LOW").Text = Data_I
session.findById("wnd[0]/usr/ctxtS_QMDAT-HIGH").Text = Data_F
session.findById("wnd[0]/usr/txtS1_SEARK-LOW").SetFocus
session.findById("wnd[0]/usr/txtS1_SEARK-LOW").caretPosition = 0
Sheets("Dados").Select
Range("A5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
session.findById("wnd[0]/usr/btn%_S1_SEARK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_VARIA").Text = "/M_Delegados"
On Error GoTo Não_CD
session.findById("wnd[0]/tbar[1]/btn[8]").press


'Exportar dados para Excel
session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Temp.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


' Abrir excel com relatório de CDs e copiar dados
Workbooks.Open ("C:\temp\Temp.xls")
Windows("Temp.xls").Activate
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:11").Select
Selection.Delete Shift:=xlUp
Columns("C:C").Select
Selection.Delete Shift:=xlToLeft
Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'Inserir os dados das CDs abertas para planilha aba temp
Windows(y).Activate
Sheets("Temp").Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("A100000").End(xlUp).Offset(1, 0).Select
    
'Fecha planilha Temp
Windows("Temp.xls").Activate
Cells.Select
Selection.ClearContents
ActiveWorkbook.Save
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True

'Realizar busca das CDs no SAP e verificar CDs fechadas
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZQLRQM150"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtS_QMART-LOW").Text = "CD"
session.findById("wnd[0]/usr/radP_ST_ENC").Select
session.findById("wnd[0]/usr/ctxtS_QMDAT-LOW").Text = Data_I
session.findById("wnd[0]/usr/ctxtS_QMDAT-HIGH").Text = Data_F
session.findById("wnd[0]/usr/txtS1_SEARK-LOW").SetFocus
session.findById("wnd[0]/usr/txtS1_SEARK-LOW").caretPosition = 0
Sheets("Dados").Select
Range("A5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
session.findById("wnd[0]/usr/btn%_S1_SEARK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_VARIA").Text = "/M_Delegados"
On Error GoTo Não_CD
session.findById("wnd[0]/tbar[1]/btn[8]").press


'Exportar dados para Excel
session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Temp.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

' Abrir excel com relatório de CDs e copiar dados
Workbooks.Open ("C:\temp\Temp.xls")
Windows("Temp.xls").Activate
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:11").Select
Selection.Delete Shift:=xlUp
Columns("C:C").Select
Selection.Delete Shift:=xlToLeft
Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'Inserir os dados para planilha aba temp e ajustar
Windows(y).Activate
Sheets("Temp").Select
Range("A100000").End(xlUp).Offset(1, 0).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
  With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("A:I").Select
    Columns("A:I").EntireColumn.AutoFit
Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    Sheets("CDs").Select
        Range("A60000").End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
Range("L1") = Date
Range("M1") = Time
Range("A100000").End(xlUp).Offset(1, 0).Select


'Fecha planilha Temp
Windows("Temp.xls").Activate
Cells.Select
Selection.ClearContents
ActiveWorkbook.Save
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True


Sheets("Temp").Select
Columns("A:I").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select
Sheets("CDs").Select

'Atualizar planilha resumo
Sheets("Resumo").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Sheets("Dados").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Resumo").Select
    Range("A5").Select
    ActiveSheet.Paste
    Range("B5").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=COUNTIF(CDs!R3C1:R60000C1,Resumo!RC[-1])"
    Range("B5").Select
    Selection.Copy
    Range("A60000").End(xlUp).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Resumo").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Resumo").Sort.SortFields.Add2 Key:=Range("B60"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Resumo").Sort
        .SetRange Range("A5:B60000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:R[60000]C[-2])"
    Range("D5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C5").Select
 
'Ocultar planilhas
Sheets("Temp").Select
ActiveWindow.SelectedSheets.Visible = False
Sheets("Dados").Select
Range("D8").Select
ActiveCell.FormulaR1C1 = "=COUNTA(CDs!R[-4]C[-1]:R[59992]C[-1])"
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("D9").Select
ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,""0"",Tamanho_amostra!R[-7]C[5])"
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("C15").Select
Range("D11") = Date
Range("D12") = Environ("USERNAME")
ActiveWorkbook.Save
           
MsgBox ("ATUALIZAÇÃO REALIZADA")

Exit Sub

Não_CD:
session.findById("wnd[0]/tbar[0]/btn[3]").press

'Ocultar planilhas temp
Sheets("Dados").Select
Sheets("Temp").Select
ActiveWindow.SelectedSheets.Visible = False
Sheets("Dados").Select
Range("D11") = Date
Range("D12") = Environ("USERNAME")

Range("D8") = "0"
Range("D9") = "0"

MsgBox ("Não nota CDs emitidas para o usuários no periodo!")

Exit Sub
SAP_OFF:
Sheets("Temp").Select
ActiveWindow.SelectedSheets.Visible = False
Sheets("Dados").Select
MsgBox ("ATENÇÃO - Você não esta logado no SAP!")
    
Exit Sub
Preencha:
MsgBox ("ATENÇÃO - Você não informou os dados necessários para analise, login, data inicial ou data final!")

End Sub
