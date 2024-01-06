# Macros

-- Excel
-- Word

## Excel Workbook_Open

### Onde Salvar Workbook_Open

```hash
Personal.xlsb  
    |- Microsoft Excel Objetos   
          |- Esta Pasta de Trabalho   
    |- Módulos   
```

### Macro Workbook_Open
```vba
Private Sub Workbook_Open()
    ' Coloque o código que deseja executar aqui
    MsgBox "Bem-vindo ao Excel!"
End Sub
```

## Excel Mymacros

```vba
' Declaração das variáveis como globais em um módulo geral
Public AmareloPastel As Long
Public RosaPastel As Long
Public LaranjaPastel As Long
Public VerdePastel As Long
Public AzulPastel As Long
Public RoxoPastel As Long
Public DarkPinkPastel As Long
Public TealColor As Long
Public PurpleColor As Long
Public BlackColor As Long
Public WhiteColor As Long
Public CorBordaPadrao As Long
Public LinhaClara As Long
Public LinhaEscura As Long

Sub Auto_Open()
    MsgBox "André." & Chr(13) & Chr(10) & "Carregados atalhos e macros." & Chr(13) & Chr(10) & "Bom trabalho!"
    Call Inicializar
End Sub
Public Sub Inicializar()
    ' Inicialização das variáveis globais
    AmareloPastel = RGB(255, 246, 155)
    RosaPastel = RGB(251, 218, 219)
    LaranjaPastel = RGB(255, 218, 158)
    VerdePastel = RGB(189, 236, 182)
    AzulPastel = RGB(173, 216, 230)
    RoxoPastel = RGB(236, 199, 238)
    DarkPinkPastel = RGB(204, 153, 204)
    TealColor = RGB(0, 128, 128)
    PurpleColor = RGB(128, 0, 96)
    BlackColor = RGB(0, 0, 0)
    WhiteColor = RGB(255, 255, 255)
    CorBordaPadrao = RGB(100, 100, 100)
    LinhaClara = RGB(255, 255, 255)
    LinhaEscura = RGB(245, 245, 245)
    ' Chama a sub-rotina para definir os atalhos
    Call MeusAtalhos
End Sub

Private Sub MeusAtalhos()
    ' Define atalhos de teclado
    Application.OnKey "^+1", "FundoAmareloPastel"
    Application.OnKey "^+2", "FundoLaranjaPastel"
    Application.OnKey "^+5", "FundoRosaPastel"
    Application.OnKey "^+4", "FundoVerdePastel"
    Application.OnKey "^+3", "FundoAzulPastel"
    Application.OnKey "^+6", "FundoRoxoPastel"
    Application.OnKey "^+7", "FundoDarkPinkPastel"
    Application.OnKey "^+8", "FundoTeal"
    Application.OnKey "^+9", "FundoPurple"
    Application.OnKey "^+{F8}", "TableNoHeader"
    Application.OnKey "^+{F11}", "TableHeaderTeal"
End Sub
Private Sub TableNoHeader()

    Const evenFormula As String = "=PAR(LIN())=LIN()"
    Const oddFormula As String = "=ÍMPAR(LIN())=LIN()"
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection
        
        ' Set aligment and border to Selected Range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CorBordaPadrao
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        
        Dim oFormula As String, eFormula As String
        If .Row Mod 2 = 0 Then
            eFormula = evenFormula
            oFormula = oddFormula
        Else
            eFormula = oddFormula
            oFormula = evenFormula
        End If
        
        .FormatConditions.Delete
        
        'Apply colors for ROW = EVEN
        .FormatConditions.Add Type:=xlExpression, Formula1:=eFormula
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            .Interior.Color = LinhaClara
            .Font.Color = BlackColor
            .StopIfTrue = False
        End With
 
        ' Apply colors for ROW = ODD
        .FormatConditions.Add Type:=xlExpression, Formula1:=oFormula
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            .Interior.Color = LinhaEscura
            .Font.Color = BlackColor
            .StopIfTrue = False
        End With

    End With

End Sub

Private Sub TableHeaderTeal()

    Const evenFormula As String = "=PAR(LIN())=LIN()"
    Const oddFormula As String = "=ÍMPAR(LIN())=LIN()"
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection
        
        ' Set aligment and border to Selected Range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CorBordaPadrao
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        
        Dim oFormula As String, eFormula As String
        If .Row Mod 2 = 0 Then
            eFormula = evenFormula
            oFormula = oddFormula
        Else
            eFormula = oddFormula
            oFormula = evenFormula
        End If
        
        .FormatConditions.Delete
        
        'Apply colors for ROW = EVEN
        .FormatConditions.Add Type:=xlExpression, Formula1:=eFormula
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            .Interior.Color = LinhaClara
            .Font.Color = BlackColor
            .StopIfTrue = False
        End With
 
        ' Apply colors for ROW = ODD
        .FormatConditions.Add Type:=xlExpression, Formula1:=oFormula
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            .Interior.Color = LinhaEscura
            .Font.Color = BlackColor
            .StopIfTrue = False
        End With

        ' Apply colors to HEADER
        .FormatConditions.Add Type:=xlExpression, Formula1:="=LIN()=" & .Row
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            .Interior.Color = TealColor
            .Font.Color = WhiteColor
            .Font.Bold = True
            .StopIfTrue = False
        End With

    End With

End Sub


Private Sub FundoAmareloPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = AmareloPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).


End Sub


Private Sub FundoVerdePastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = VerdePastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub


Private Sub FundoRosaPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = RosaPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub

Private Sub FundoLaranjaPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = LaranjaPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub

Private Sub FundoAzulPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = AzulPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub
Private Sub FundoRoxoPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = RoxoPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub

Private Sub FundoDarkPinkPastel()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = DarkPinkPastel
    corFonte = BlackColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub
Private Sub FundoTeal()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = TealColor
    corFonte = WhiteColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub

Private Sub FundoPurple()
    
    Dim corFundo As Long
    Dim corFonte As Long
    corFundo = PurpleColor
    corFonte = WhiteColor
    Selection.FormatConditions.Delete 'apaga a formatação condicional
    Selection.Interior.Color = corFundo ' cor de fundo
    Selection.Font.Color = corFonte 'cor da fonte
    Selection.Interior.TintAndShade = 0 ' Define a tonalidade e sombra do padrão de preenchimento como zero (sem alteração).

End Sub

Sub PreencherTracinho()
    Dim selecao As Range, coluna As Range
    Dim ws As Worksheet
    Dim primeiraLinha As Long, ultimaLinha As Long

    If Not TypeOf Selection Is Range Then
        MsgBox "Por favor, selecione um intervalo de células."
        Exit Sub
    End If

    Set selecao = Selection
    Set ws = selecao.Worksheet

    ' Determinar a última linha preenchida na aba usando uma abordagem mais eficiente
    ultimaLinha = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    Application.ScreenUpdating = False ' Desativa a atualização da tela para acelerar o script
    Application.Calculation = xlCalculationManual ' Desativa o cálculo automático para acelerar o script

    For Each coluna In selecao.Columns
        primeiraLinha = ws.Cells(coluna.Column).End(xlDown).Row

        ' Selecionar o intervalo e substituir células vazias
        With ws.Range(ws.Cells(primeiraLinha, coluna.Column), ws.Cells(ultimaLinha, coluna.Column))
            .Replace What:="", Replacement:="-", LookAt:=xlWhole
        End With
    Next coluna

    Application.Calculation = xlCalculationAutomatic ' Reativa o cálculo automático
    Application.ScreenUpdating = True ' Reativa a atualização da tela

    MsgBox "Células vazias preenchidas com '-' nas colunas selecionadas."
End Sub

Sub PreencherZero()
    Dim selecao As Range, coluna As Range
    Dim ws As Worksheet
    Dim primeiraLinha As Long, ultimaLinha As Long

    If Not TypeOf Selection Is Range Then
        MsgBox "Por favor, selecione um intervalo de células."
        Exit Sub
    End If

    Set selecao = Selection
    Set ws = selecao.Worksheet

    ' Determinar a última linha preenchida na aba usando uma abordagem mais eficiente
    ultimaLinha = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    Application.ScreenUpdating = False ' Desativa a atualização da tela para acelerar o script
    Application.Calculation = xlCalculationManual ' Desativa o cálculo automático para acelerar o script

    For Each coluna In selecao.Columns
        primeiraLinha = ws.Cells(coluna.Column).End(xlDown).Row

        ' Selecionar o intervalo e substituir células vazias
        With ws.Range(ws.Cells(primeiraLinha, coluna.Column), ws.Cells(ultimaLinha, coluna.Column))
            .Replace What:="", Replacement:=0, LookAt:=xlWhole
        End With
    Next coluna

    Application.Calculation = xlCalculationAutomatic ' Reativa o cálculo automático
    Application.ScreenUpdating = True ' Reativa a atualização da tela

    MsgBox "Células vazias preenchidas com '-' nas colunas selecionadas."
End Sub

```



