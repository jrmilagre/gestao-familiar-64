VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMovimentacoes 
   Caption         =   ":: Movimentações ::"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   OleObjectBlob   =   "fMovimentacoes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMovimentacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta          As New cConta
Private oContaPara      As New cContaPara
Private oFornecedor     As New cFornecedor
Private oCategoria      As New cCategoria
Private oSubcategoria   As New cSubcategoria
Private oMovimentacao   As New cMovimentacao
Private oTransferencia  As New cTransferencia
Private oAgendamento    As New cAgendamento
Private sDecisao        As String
Private iDias           As Integer

Private Sub UserForm_Initialize()
    
    Call lstContasPopular
    Call cbbGrupoPopular
    Call cbbContaPopular
    Call cbbFornecedorPopular
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnConfirmar.Visible = False
    btnCancelar.Visible = False
    lblContaPara.Visible = False
    cbbContaPara.Visible = False
    
    iDias = 90
    btn7dias.Enabled = False
    btn15dias.Enabled = False
    btn30dias.Enabled = False
    btn90dias.Enabled = False
    
    btnIncluir.SetFocus

End Sub

Private Sub btn7dias_Click()
    oMovimentacao.DiasExtrato = 7
    Call lstMovimentacoesPopular
End Sub
Private Sub btn15dias_Click()
    oMovimentacao.DiasExtrato = 15
    Call lstMovimentacoesPopular
End Sub
Private Sub btn30dias_Click()
    oMovimentacao.DiasExtrato = 30
    Call lstMovimentacoesPopular
End Sub
Private Sub btn90dias_Click()
    oMovimentacao.DiasExtrato = 90
    Call lstMovimentacoesPopular
End Sub
Private Sub btnLiquidado_Click()
    dtDate = IIf(txbLiquidado.Text = Empty, Date, txbLiquidado.Text)
    txbLiquidado.Text = GetCalendario
End Sub
Private Sub lstContas_Change()
    If lstContas.ListIndex > -1 Then
        lstRegistros.ListIndex = -1
        oConta.Carrega (CLng(lstContas.List(lstContas.ListIndex, 1)))
        
        Call lstMovimentacoesPopular
        Call Campos("Limpar")
        btnAlterar.Enabled = False
        btnExcluir.Enabled = False
    End If
    
    btn7dias.Enabled = True
    btn15dias.Enabled = True
    btn30dias.Enabled = True
    btn90dias.Enabled = True
End Sub
Private Sub lstRegistros_Change()
    If lstRegistros.ListIndex > 0 Then
        oMovimentacao.Carrega CLng(lstRegistros.List(lstRegistros.ListIndex, 0))
        oConta.Carrega oMovimentacao.ContaID
        
        If oConta.ID <> 0 Then
            btnAlterar.Enabled = True
            btnExcluir.Enabled = True
            
            lblRegistro.Caption = Format(oMovimentacao.ID, "0000000000")
            
            If oMovimentacao.Grupo = "T" Then
                oTransferencia.Carrega oMovimentacao.TransferenciaID
                oConta.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoDeID)
                oContaPara.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoParaID)
            Else
                oFornecedor.Carrega oMovimentacao.FornecedorID
                oSubcategoria.Carrega oMovimentacao.SubcategoriaID
                oCategoria.Carrega oSubcategoria.CategoriaID
            End If
        End If
        
        Call InformacoesCarregar
    Else
        lblRegistro.Caption = ""
    End If

End Sub

Private Sub chbTransferencia_Click()
    If chbTransferencia.Value = True Then
        lblFornecedor.Visible = False: cbbFornecedor.Visible = False
        lblGrupo.Visible = False: cbbGrupo.Visible = False
        lblCategoria.Visible = False: cbbCategoria.Visible = False
        lblSubcategoria.Visible = False: cbbSubcategoria.Visible = False
        lblContaPara.Visible = True: cbbContaPara.Visible = True
        lblContaDe.Caption = "Conta origem"
    Else
        lblFornecedor.Visible = True: cbbFornecedor.Visible = True
        lblGrupo.Visible = True: cbbGrupo.Visible = True
        lblCategoria.Visible = True: cbbCategoria.Visible = True
        lblSubcategoria.Visible = True: cbbSubcategoria.Visible = True
        lblContaPara.Visible = False: cbbContaPara.Visible = False
        lblContaDe.Caption = "Conta"
    End If
End Sub
Private Sub cbbContaDe_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    
    If cbbContaDe.ListIndex > -1 Then
        oConta.ID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
    Else
        If cbbContaDe.Text <> "" Then
            vbResposta = MsgBox("Esta Conta não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                oConta.Conta = cbbContaDe.Text
                oConta.Inclui
                Call ComboBoxCarregarContas
            Else
                cbbContaDe.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub cbbContaDe_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then
        If chbTransferencia.Value = False Then
            cbbFornecedor.SetFocus
            cbbFornecedor.DropDown
        End If
    End If

End Sub
Private Sub cbbContaPara_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
    
    If cbbContaPara.ListIndex > -1 Then
        oContaPara.ID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
    Else
        If cbbContaPara.Text <> "" Then
            vbResposta = MsgBox("Esta Conta não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                oContaPara.Conta = cbbContaPara.Text
                oContaPara.Inclui
                Call ComboBoxCarregarContas
            Else
                cbbContaPara.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub cbbContaPara_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If chbTransferencia.Value = True Then
            txbValor.SetFocus
        End If
    End If
End Sub
Private Sub cbbFornecedor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Ao pressionar enter, foca no campo Valor
    If KeyCode = 13 Then
        txbValor.SetFocus
    End If
End Sub
Private Sub cbbFornecedor_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
        
    If cbbFornecedor.ListIndex = -1 Then
        If cbbFornecedor.Text <> "" Then
            vbResposta = MsgBox("Este Fornecedor não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                oFornecedor.NomeFantasia = cbbFornecedor.Text
                oFornecedor.Inclui
                Call cbbFornecedorPopular
                cbbFornecedor.Text = oFornecedor.NomeFantasia
            Else
                cbbFornecedor.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub txbValor_AfterUpdate()
    If IsNumeric(txbValor) Then
        txbValor.Text = Format(txbValor.Text, "#,##0.00")
        Exit Sub
    Else
        txbValor.Text = Empty
    End If
End Sub
Private Sub txbValor_Enter()
    ' Seleciona todos os caracteres do campo
    txbValor.SelStart = 0
    txbValor.SelLength = Len(txbValor.Text)
End Sub
Private Sub txbValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8          ' Backspace (seta de apagar)
        Case 48 To 57   ' Números de 0 a 9
        Case 44         ' Vírgula
        If InStr(txbValor.Text, ",") Then ' Se o campo já tiver vírgula então ele não adiciona
            KeyAscii = 0 ' Não adiciona a vírgula caso ja tenha
        Else
            KeyAscii = 44 ' Adiciona uma vírgula
        End If
        Case Else
            KeyAscii = 0 ' Não deixa nenhuma outra caractere ser escrito
    End Select
End Sub
Private Sub txbLiquidado_AfterUpdate()
    If IsDate(txbLiquidado.Text) Then
        txbLiquidado.Text = Format(txbLiquidado.Text, "dd/mm/yyyy")
        Exit Sub
    Else
        txbLiquidado.Text = Empty
    End If
End Sub
Private Sub txbLiquidado_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '---se a tecla F4 for pressionada
    If KeyCode = 115 Then
        dtDate = IIf(txbLiquidado.Text = "", Date, txbLiquidado.Text)
        txbLiquidado.Text = GetCalendario
    ElseIf KeyCode = 13 Then
        If chbTransferencia.Value = False Then
            cbbGrupo.SetFocus
            cbbGrupo.DropDown
        Else
            btnConfirmar.SetFocus
        End If
    End If
End Sub
Private Sub txbLiquidado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With txbLiquidado
        Select Case KeyAscii
            Case 8                      ' Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}"   ' Emula o TAB
            Case 48 To 57
                If .SelStart = 2 Then .SelText = "/"
                If .SelStart = 5 Then .SelText = "/"
            Case Else: KeyAscii = 0     ' Ignora os outros caracteres
        End Select
    End With
End Sub
Private Sub btnVencimento_Click()
    dtDatabase = IIf(txbLiquidado.Text = Empty, Date, txbLiquidado.Text)
    txbLiquidado.Text = GetCalendario
    oMovimentacao.Liquidado = CDate(txbLiquidado.Text)
End Sub

Private Sub cbbGrupo_AfterUpdate()
    
    If cbbGrupo.ListIndex > -1 Then
        If oMovimentacao.Grupo <> "" Then
            oMovimentacao.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
            If oMovimentacao.Grupo = "D" And oMovimentacao.Valor > 0 Then
                oMovimentacao.Valor = oMovimentacao.Valor * -1
            ElseIf oMovimentacao.Grupo = "R" And oMovimentacao.Valor < 0 Then
                oMovimentacao.Valor = oMovimentacao.Valor * -1
            End If
        Else
            oMovimentacao.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
            If oMovimentacao.Grupo = "D" And oMovimentacao.Valor > 0 Then
                oMovimentacao.Valor = oMovimentacao.Valor * -1
            ElseIf oMovimentacao.Grupo = "R" And oMovimentacao.Valor < 0 Then
                oMovimentacao.Valor = oMovimentacao.Valor * -1
            End If
        End If
        
        If cbbGrupo.ListIndex > -1 Then
            cbbCategoria.Clear
            cbbCategoria.ListIndex = -1
            cbbSubcategoria.Clear
            cbbSubcategoria.ListIndex = -1
            Call cbbCategoriaPopular
            cbbCategoria.Style = fmStyleDropDownCombo
        End If
    End If
End Sub
Private Sub cbbGrupo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        cbbCategoria.SetFocus
        cbbCategoria.DropDown
    End If
End Sub
Private Sub cbbCategoria_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim oCategoria As cCategoria

    If cbbCategoria.ListIndex > -1 Then
        oSubcategoria.CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
        
        cbbSubcategoria.Clear
        cbbSubcategoria.ListIndex = -1
        
        If oMovimentacao.Grupo <> "" And cbbCategoria.Text <> "" Then
            Call cbbSubcategoriaPopular
        End If
        cbbSubcategoria.Style = fmStyleDropDownCombo
    Else
        
        If cbbCategoria.Text <> "" Then
            cbbSubcategoria.ListIndex = -1
            
            vbResposta = MsgBox("Esta Categoria não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            
            If vbResposta = vbYes Then
                
                Set oCategoria = New cCategoria
                oCategoria.Categoria = cbbCategoria.Text
                oCategoria.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                oCategoria.Inclui
                Call cbbCategoriaPopular
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                Call cbbSubcategoriaPopular
                cbbSubcategoria.Style = fmStyleDropDownCombo
                
            Else
                cbbCategoria.ListIndex = -1
                'cbbSubcategoria.Style = fmStyleDropDownList
            End If
        End If
    End If
    
End Sub
Private Sub cbbCategoria_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        cbbSubcategoria.SetFocus
        cbbSubcategoria.DropDown
    End If
End Sub
Private Sub cbbSubcategoria_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
    
    If cbbSubcategoria.ListIndex > -1 Then
        oMovimentacao.SubcategoriaID = cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1)
    Else
        If cbbSubcategoria.Text <> "" Then
            vbResposta = MsgBox("Esta Subcategoria não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                oSubcategoria.Subcategoria = cbbSubcategoria.Text
                oSubcategoria.Inclui
                Call cbbSubcategoriaPopular
                oMovimentacao.SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
            Else
                cbbSubcategoria.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub txbObservacao_AfterUpdate()
    oMovimentacao.Observacao = txbObservacao.Text
End Sub
Private Sub txbObservacao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnConfirmar.SetFocus
End Sub
Private Sub btnIncluir_Click()
    
    sDecisao = "Inclusão"
    
    lstContas.Enabled = False
    lstRegistros.Enabled = False

    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & vbNewLine & sDecisao
    btnCancelar.Caption = "Cancelar " & vbNewLine & sDecisao
    
    btnIncluir.Visible = False
    btnAlterar.Visible = False
    btnExcluir.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Habilitar")
    
    txbLiquidado.Text = Format(Date, "dd/mm/yyyy")
    txbValor.Text = Format(0, "#,##0.00")
    
    cbbContaDe.SetFocus
    
End Sub
Private Sub btnAlterar_Click()
    
    sDecisao = "Alteração"
    
    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
    
    btnConfirmar.SetFocus
    
    btnIncluir.Visible = False
    btnAlterar.Visible = False
    btnExcluir.Visible = False
    
    Call Campos("Habilitar")
    
    cbbCategoria.Enabled = True
    cbbSubcategoria.Enabled = True
    
    lstRegistros.Enabled = False
    lstContas.Enabled = False
    
    cbbContaDe.SetFocus
End Sub
Private Sub btnExcluir_Click()
    
    sDecisao = "Exclusão"
    
    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
    
    btnConfirmar.SetFocus
    
    btnIncluir.Visible = False
    btnAlterar.Visible = False
    btnExcluir.Visible = False
    
    lvRegistros.Enabled = False
    
    lstContas.Enabled = False

End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VbMsgBoxResult
    Dim sContaSelecionada As String
    
    If Valida = True Then
        
        If sDecisao = "Inclusão" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then oMovimentacao.Inclui IsTransferencia:=chbTransferencia.Value, IsAgendamento:=False
        ElseIf sDecisao = "Alteração" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then
                oMovimentacao.Altera chbTransferencia.Value, oMovimentacao.Valor
                MsgBox sDecisao & " realizada com sucesso!", vbInformation, sDecisao
            End If
        ElseIf sDecisao = "Exclusão" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then
                oMovimentacao.Exclui IIf(oMovimentacao.Grupo = "T", True, False), oMovimentacao.Valor, IIf(oMovimentacao.Origem = "Agendamento", True, False)
                MsgBox sDecisao & " realizada com sucesso!", vbInformation, sDecisao
            End If

        End If
        
        Call Campos("Limpar")
        Call Campos("Desabilitar")
        
        ' Se houver uma conta selecionada
        If lstContas.ListIndex > -1 Then
            sContaSelecionada = lstContas.List(lstContas.ListIndex, 0)
            Call lstContasPopular
            lstContas.Text = sContaSelecionada
            Call lstMovimentacoesPopular
        Else
            Call lstContasPopular
            Call lstMovimentacoesPopular
        End If
        
        btnConfirmar.Visible = False
        btnCancelar.Visible = False
        
        btnIncluir.Visible = True
        btnAlterar.Visible = True
        btnExcluir.Visible = True
        Call Campos("Desabilitar")
        Call Campos("Limpar")
        
        btnAlterar.Enabled = False
        btnExcluir.Enabled = False
        btnIncluir.SetFocus
        
        lstRegistros.Enabled = True
        lstContas.Enabled = True
            
    End If

End Sub
Private Sub btnCancelar_Click()

    btnConfirmar.Visible = False
    btnCancelar.Visible = False
    
    btnIncluir.Visible = True
    btnAlterar.Visible = True
    btnExcluir.Visible = True
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    lstRegistros.Enabled = True
    lstContas.Enabled = True
    
    btnIncluir.SetFocus
    
End Sub


'+-------------------------------------------------------+
'|                                                       |
'| ROTINAS E FUNÇÕES                                     |
'|                                                       |
'+-------------------------------------------------------+

Private Function Valida() As Boolean
    
    Valida = False
    
    If cbbContaDe.Text = Empty Then
        MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigatório"
    ElseIf txbLiquidado.Text = Empty Then
        MsgBox "Informe a 'Data'", vbCritical, "Campo obrigatório"
    Else

        ' Se for uma Receita ou Despesa
        If chbTransferencia.Value = False Then
            
            If txbValor.Text = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
            ElseIf cbbFornecedor.Text = Empty Then
                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigatório"
            ElseIf cbbCategoria.Text = Empty Then
                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigatório"
            ElseIf cbbSubcategoria.Text = Empty Then
                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigatório"
            Else
                With oMovimentacao
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .Valor = CCur(txbValor.Text)
                    .Liquidado = CDate(txbLiquidado.Text)
                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                    .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                    .Observacao = txbObservacao.Text
                End With
                
                Valida = True
                
            End If
            
        ' Se for uma transferência
        Else
            If txbValor.Text = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
            ElseIf cbbContaPara.Text = Empty Then
                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigatório":
            Else
                With oMovimentacao
                    .Grupo = "T"
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    .Liquidado = CDate(txbLiquidado.Text)
                    '.ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Observacao = txbObservacao.Text
                End With
                
                Valida = True
                
            End If
        
        End If
        
    End If
    
    
End Function
Private Sub InformacoesCarregar()

    cbbContaDe.Text = oConta.Conta
    cbbFornecedor.Text = oFornecedor.NomeFantasia
    txbLiquidado.Text = oMovimentacao.Liquidado
    txbValor.Text = IIf(oMovimentacao.Valor < 0, Format(oMovimentacao.Valor * -1, "#,##0.00"), Format(oMovimentacao.Valor, "#,##0.00"))
    txbObservacao.Text = oMovimentacao.Observacao
    
    If oMovimentacao.Grupo = "T" Then
        chbTransferencia.Value = True
        lblFornecedor.Visible = False: cbbFornecedor.Visible = False
        lblContaPara.Visible = True: cbbContaPara.Visible = True
        lblGrupo.Visible = False: cbbGrupo.Visible = False
        lblCategoria.Visible = False: cbbCategoria.Visible = False
        lblSubcategoria.Visible = False: cbbSubcategoria.Visible = False
        
        cbbContaDe.Text = oConta.Conta: lblContaDe.Caption = "Conta de origem"
        cbbContaPara.Text = oContaPara.Conta: lblContaPara.Caption = "Conta de destino"
    Else
        chbTransferencia.Value = False
        
        Select Case oMovimentacao.Grupo
            Case "R": cbbGrupo.ListIndex = 0
            Case "D": cbbGrupo.ListIndex = 1
            Case "I": cbbGrupo.ListIndex = 2
        End Select
        
        'cbbGrupo.ListIndex = IIf(oMovimentacao.Grupo = "D", 1, 0)
        
        Call cbbGrupo_AfterUpdate
        cbbCategoria.Text = oCategoria.Categoria
        Call cbbCategoria_AfterUpdate
        lblContaDe.Caption = "Conta"
        cbbSubcategoria.Text = oSubcategoria.Subcategoria
        lblFornecedor.Visible = True: cbbFornecedor.Visible = True
        lblContaPara.Visible = False: cbbContaPara.Visible = False
        lblGrupo.Visible = True: cbbGrupo.Visible = True
        lblCategoria.Visible = True: cbbCategoria.Visible = True
        lblSubcategoria.Visible = True: cbbSubcategoria.Visible = True
    End If
End Sub

Private Sub lstContasPopular()

    Dim col As New Collection
    
    Set col = oConta.Listar("conta")
    
    ' Configura ListBox
    With lstContas
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 2     ' Determina número de colunas
        .ColumnWidths = "80 pt; 0pt"       'Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            
            oConta.Carrega CLng(n)
        
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
            
        Next n
        
    End With
    
End Sub
Private Sub lstMovimentacoesPopular()

    Dim col         As New Collection
    Dim curSaldo    As Currency
    
    oConta.Carrega CLng(lstContas.List(lstContas.ListIndex, 1))
    
    Set col = oMovimentacao.ListaMovimentacoes(oConta.ID)
    curSaldo = oMovimentacao.SaldoAnteriorExtrato(oConta.ID) + oConta.SaldoInicial
    
    With lstRegistros
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 7   ' Determina número de colunas
        .ColumnWidths = "0pt; 55pt; 120pt; 220pt; 75pt; 90pt; 55pt"   'Configura largura das colunas
        .Font = "Consolas"
        
        .AddItem
        .List(.ListCount - 1, 3) = "Saldo anterior"
        .List(.ListCount - 1, 5) = Space(12 - Len(Format(curSaldo, "#,##0.00"))) & Format(curSaldo, "#,##0.00")
        
        Dim n As Variant
        
        For Each n In col
        
            oMovimentacao.Carrega CLng(n)
            
        
            .AddItem
            .List(.ListCount - 1, 0) = oMovimentacao.ID
            .List(.ListCount - 1, 1) = oMovimentacao.Liquidado
        
            If oMovimentacao.Grupo = "T" Then
            
                oTransferencia.Carrega oMovimentacao.TransferenciaID
                oConta.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoDeID)
                oContaPara.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoParaID)
            
                
                
                If oMovimentacao.Valor < 0 Then
                    .List(.ListCount - 1, 2) = "---Transferência-->"
                    .List(.ListCount - 1, 3) = "Foi para a conta: " & oContaPara.Conta
                Else
                    .List(.ListCount - 1, 2) = "<--Transferência---"
                    .List(.ListCount - 1, 3) = "Veio da conta : " & oConta.Conta
                End If

            ' Se não for uma transferência entre contas...
            Else
                oFornecedor.Carrega oMovimentacao.FornecedorID
                oSubcategoria.Carrega oMovimentacao.SubcategoriaID
                oCategoria.Carrega oSubcategoria.CategoriaID
                
                .List(.ListCount - 1, 2) = oFornecedor.NomeFantasia
                .List(.ListCount - 1, 3) = oCategoria.Categoria & " : " & oSubcategoria.Subcategoria
            End If

            .List(.ListCount - 1, 4) = Space(10 - Len(Format(oMovimentacao.Valor, "#,##0.00"))) & Format(oMovimentacao.Valor, "#,##0.00")

            curSaldo = curSaldo + oMovimentacao.Valor

            .List(.ListCount - 1, 5) = Space(12 - Len(Format(curSaldo, "#,##0.00"))) & Format(curSaldo, "#,##0.00")

            If Not oMovimentacao.AgendamentoID = 0 Then
            
                oAgendamento.Carrega oMovimentacao.AgendamentoID

                If oAgendamento.Recorrente = False Then
                    .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-ú"
                Else
                    If oAgendamento.Infinito = False Then
                        .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-" & oMovimentacao.Parcela
                    Else
                        .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-i"
                    End If

                End If
            Else
                .List(.ListCount - 1, 6) = ""
            End If
        Next n
    End With
    
End Sub
Private Sub cbbGrupoPopular()
       
    Dim col As Collection
    Dim n   As Variant
    Dim s() As String
    
    Set col = oCategoria.ListarGrupos
    
    For Each n In col
    
        s() = Split(n, ",")
        
        With cbbGrupo
            .AddItem
            .List(.ListCount - 1, 0) = s(0)
            .List(.ListCount - 1, 1) = s(1)
        End With
    Next n
    
End Sub
Private Sub cbbContaPopular()

    Dim col As New Collection
    Dim n   As Variant
    
    Set col = oConta.Listar("conta")
       
    For Each n In col
        
        oConta.Carrega CLng(n)
    
        With cbbContaDe
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
        With cbbContaPara
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
    Next n
        
End Sub
Private Sub cbbFornecedorPopular()
    
    Dim col As New Collection
    Dim n   As Variant

    Set col = oFornecedor.Listar("nome_fantasia")
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.NomeFantasia
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
End Sub
Private Sub cbbCategoriaPopular()

    Dim col As Collection
    Dim n As Variant
    
    Set col = oCategoria.Listar("categoria", oMovimentacao.Grupo)
    
    cbbCategoria.Clear
    
    For Each n In col
    
        oCategoria.Carrega CLng(n)
        
        With cbbCategoria
            .AddItem
            .List(.ListCount - 1, 0) = oCategoria.Categoria
            .List(.ListCount - 1, 1) = oCategoria.ID
        End With
    
    Next n
    
End Sub
Private Sub cbbSubcategoriaPopular()

    Dim col As Collection
    Dim n As Variant
    
    Set col = oSubcategoria.Listar("subcategoria", oSubcategoria.CategoriaID)
    
    cbbSubcategoria.Clear
    
    For Each n In col
    
        oSubcategoria.Carrega CLng(n)
    
        With cbbSubcategoria
            .AddItem
            .List(.ListCount - 1, 0) = oSubcategoria.Subcategoria
            .List(.ListCount - 1, 1) = oSubcategoria.ID
        End With
    
    Next n

End Sub
Private Sub Campos(Acao As String)
    If Acao = "Limpar" Then
        lblRegistro.Caption = ""
        cbbContaDe.ListIndex = -1
        cbbContaPara.ListIndex = -1
        cbbContaDe.ListIndex = -1
        cbbFornecedor.ListIndex = -1
        cbbGrupo.ListIndex = -1
        cbbCategoria.ListIndex = -1
        cbbSubcategoria.ListIndex = -1
        txbValor.Text = ""
        txbLiquidado.Text = ""
        txbObservacao.Text = ""
    ElseIf Acao = "Habilitar" Then
        chbTransferencia.Enabled = IIf(sDecisao = "Inclusão", True, False)
        cbbContaDe.Enabled = True: lblContaDe.Enabled = True
        cbbContaPara.Enabled = True: lblContaPara.Enabled = True
        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        cbbGrupo.Enabled = True: lblGrupo.Enabled = True
        cbbCategoria.Enabled = True: lblCategoria.Enabled = True
        cbbSubcategoria.Enabled = True: lblSubcategoria.Enabled = True
        txbValor.Enabled = True: lblValor.Enabled = True
        txbLiquidado.Enabled = True: lblLiquidado.Enabled = True: btnLiquidado.Enabled = True
        txbObservacao.Enabled = True: lblObservacao.Enabled = True
    ElseIf Acao = "Desabilitar" Then
        chbTransferencia.Enabled = False
        cbbContaDe.Enabled = False: lblContaDe.Enabled = False
        cbbContaPara.Enabled = False: lblContaPara.Enabled = False
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        cbbGrupo.Enabled = False: lblGrupo.Enabled = False
        cbbCategoria.Enabled = False: lblCategoria.Enabled = False
        cbbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False
        txbValor.Enabled = False: lblValor.Enabled = False
        txbLiquidado.Enabled = False: lblLiquidado.Enabled = False: btnLiquidado.Enabled = False
        txbObservacao.Enabled = False: lblObservacao.Enabled = False
    End If
End Sub
Private Sub UserForm_Terminate()
    Call Desconecta
End Sub
