VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fRegistrar 
   Caption         =   ":: Registrar ::"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "fRegistrar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fRegistrar"
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
Private oTransferencia  As New cTransferencia
Private colControles    As New Collection
Private Const tbl       As String = "tbl_movimentacoes"

Private Sub UserForm_Initialize()

    Call EventosCampos("tbl_movimentacoes")
    
    ' Se for uma MOVIMENTAÇÃO de origem de agendamento, então ...
    If oMovimentacao.IsAgendamento = True Then
    
        oAgendamento.Carrega oMovimentacao.AgendamentoID
        
        If oAgendamento.Grupo = "T" Then
            oMovimentacao.IsTransferencia = True
        Else
            oMovimentacao.IsTransferencia = False
        End If
        
        ' Se o registro for um AGENDAMENTO e também um RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
        If oMovimentacao.IsTransferencia = False Then
            
            ' Deixa invisível o checkbox de transferência
            chbTransferencia.Visible = False
            
            ' Coloca o número do agendamento
            lblTitulo.Caption = "Registro do agendamento nº: " & Format(oMovimentacao.AgendamentoID, "00000000")
            
            ' Parece redundante carregar o objeto oMovimentação agora, sendo que ele já é carregado
            ' na rotina Valida, mas faço isso porque o registro pode não ser oriundo de agendamento
'            With oMovimentacao
'                .ContaID = oAgendamento.ContaID
'                .SubcategoriaID = oAgendamento.SubcategoriaID
'                .FornecedorID = oAgendamento.FornecedorID
'                .Liquidado = oAgendamento.Vencimento
'                .Grupo = oAgendamento.Grupo
'                .Valor = oAgendamento.Valor
'                .AgendamentoID = oAgendamento.ID
'            End With
            
            ' Carrega detalhes necessários dos cadastros
            Call ComboBoxCarregar
            Call ComboBoxCarregarContas
            Call ComboBoxCarregarFornecedores
            
            oConta.Carrega oAgendamento.ContaID 'oMovimentacao.ContaID
            oSubcategoria.Carrega oAgendamento.SubcategoriaID
            oCategoria.Carrega oSubcategoria.CategoriaID
            oFornecedor.Carrega oAgendamento.FornecedorID 'oMovimentacao.FornecedorID
    
            txbVencimento.Text = oAgendamento.Vencimento
            cbbContaDe.Text = oConta.Conta
            cbbFornecedor.Text = oFornecedor.NomeFantasia
            cbbGrupo.Text = IIf(oAgendamento.Grupo = "R", "Receitas", "Despesas")
            txbValor.Text = Format(IIf(oAgendamento.Grupo = "R", oAgendamento.Valor, oAgendamento.Valor * -1), "#,##0.00")
            cbbCategoria.Text = oCategoria.Categoria
            cbbSubcategoria.Text = oSubcategoria.Subcategoria
            txbObservacao.Text = oAgendamento.Observacao
            
            cbbGrupo.Enabled = False
        
            lblContaPara.Visible = False: cbbContaPara.Visible = False
            
            
        
        ' Se o AGENDAMENTO a ser registrado for uma TRANSFERÊNCIA ENTRE CONTAS, então ...
        Else
        
            lblTitulo.Caption = "Registro do agendamento nº: " & Format(oAgendamento.ID, "00000000")
            
            chbTransferencia.Value = oMovimentacao.IsTransferencia
            chbTransferencia.Visible = True
            chbTransferencia.Enabled = False
            
            Call ComboBoxCarregarContas
            
            oConta.Carrega oAgendamento.ContaID
            oContaPara.Carrega oAgendamento.ContaParaID
            
            cbbContaDe.Text = oConta.Conta
            cbbContaPara.Text = oContaPara.Conta
            lblContaPara.Visible = True: cbbContaPara.Visible = True
            txbValor.Text = Format(IIf(oAgendamento.Grupo = "R", oAgendamento.Valor, oAgendamento.Valor * -1), "#,##0.00")
            txbVencimento.Text = oAgendamento.Vencimento
            txbObservacao.Text = oAgendamento.Observacao
            'oAgendamento.Vencimento = txbVencimento.Text
            
        End If
    
        btnRegistrar.SetFocus
    
    ' Se for um REGISTRO DIRETO, então
    Else
        
        chbTransferencia.Visible = True
        
        Call ComboBoxCarregar
        Call ComboBoxCarregarContas
        Call ComboBoxCarregarFornecedores
        
        lblContaPara.Visible = False: cbbContaPara.Visible = False
        txbValor.Text = Format(0, "#,##0.00")
        txbVencimento.Text = Date
        
        chbTransferencia.SetFocus

    End If
    
End Sub
Private Sub EventosCampos(Tabela As String)

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
            Set oEvento = New c_EventoCampo
            Set oEvento = oEvento.Evento(oControle, Tabela)
            colControles.Add oEvento
        End If
    Next

End Sub
Private Sub chbTransferencia_AfterUpdate()
    cbbContaDe.SetFocus
End Sub

'+-------------------------------------------------------+
'|                                                       |
'| CONTROLES DO FORMULÁRIO                               |
'|                                                       |
'+-------------------------------------------------------+

Private Sub chbTransferencia_Click()
    If chbTransferencia.Value = True Then
        lblFornecedor.Visible = False
        cbbFornecedor.Visible = False
        lblGrupo.Visible = False
        cbbGrupo.Visible = False
        lblCategoria.Visible = False
        cbbCategoria.Visible = False
        lblSubcategoria.Visible = False
        cbbSubcategoria.Visible = False
        lblContaPara.Visible = True
        cbbContaPara.Visible = True
        lblContaDe.Caption = "Conta origem"
        lblTitulo.Caption = "Transferência entre contas"
    Else
        lblFornecedor.Visible = True
        cbbFornecedor.Visible = True
        lblGrupo.Visible = True
        cbbGrupo.Visible = True
        lblCategoria.Visible = True
        cbbCategoria.Visible = True
        lblSubcategoria.Visible = True
        cbbSubcategoria.Visible = True
        lblContaPara.Visible = False
        cbbContaPara.Visible = False
        lblContaDe.Caption = "Conta"
        lblTitulo.Caption = "Registro direto"
    End If
End Sub
Private Sub chbTransferencia_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        cbbContaDe.SetFocus
        cbbContaDe.DropDown
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
    
    If cbbFornecedor.ListIndex > -1 Then
        oMovimentacao.FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
    Else
        vbResposta = MsgBox("Este Fornecedor não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        If vbResposta = vbYes Then
            oFornecedor.NomeFantasia = cbbFornecedor.Text
            oFornecedor.Inclui
            Call ComboBoxCarregarFornecedores
            cbbFornecedor.Text = oFornecedor.NomeFantasia
        Else
            cbbFornecedor.ListIndex = -1
        End If
        
    End If
End Sub

Private Sub txbVencimento_AfterUpdate()
    If IsDate(txbVencimento.Text) Then
        txbVencimento.Text = Format(txbVencimento.Text, "dd/mm/yyyy")
        Exit Sub
    Else
        txbVencimento.Text = Empty
    End If
End Sub
Private Sub btnVencimento_Click()
    dtDatabase = IIf(txbVencimento.Text = Empty, Date, txbVencimento.Text)
    txbVencimento.Text = GetCalendario
End Sub
Private Sub cbbGrupo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then
        cbbCategoria.SetFocus
        cbbCategoria.DropDown
    End If

End Sub
Private Sub cbbGrupo_AfterUpdate()
    
    If cbbGrupo.ListIndex > -1 Then
        oMovimentacao.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
        If cbbGrupo.ListIndex > -1 Then
            cbbCategoria.Clear
            cbbCategoria.ListIndex = -1
            cbbSubcategoria.Clear
            cbbSubcategoria.ListIndex = -1
            Call ComboBoxCarregarCategorias
            cbbCategoria.Style = fmStyleDropDownCombo
            
        End If
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
            Call ComboBoxCarregarSubcategorias
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
                Call ComboBoxCarregarCategorias
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                Call ComboBoxCarregarSubcategorias
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
                Call ComboBoxCarregarSubcategorias
                oMovimentacao.SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
            Else
                cbbSubcategoria.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub btnCancelar_Click()
    
    Dim i As Integer
    
    If oMovimentacao.IsAgendamento = True Then
    
        If colAgendamentos.Count = 1 Then
            colAgendamentos.Remove 1
        End If
    End If
    
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    '---se for 'Registro de agendamento não precisa desconectar do banco
    If oMovimentacao.IsAgendamento = True Then
        'não desconecta o banco pois o registro é oriundo de agendamento ou registros
    Else
        Call Desconecta
    End If
End Sub
'+-------------------------------------------------------+
'|                                                       |
'| ROTINAS E FUNÇÕES                                     |
'|                                                       |
'+-------------------------------------------------------+

Private Sub ComboBoxCarregarCategorias()
    
    Dim col         As Collection
    Dim n           As Variant
    Dim sCategoria  As String
    
    Set col = oCategoria.Listar("categoria", oMovimentacao.Grupo)
    
    sCategoria = cbbCategoria.Text
    cbbCategoria.Clear
    
    For Each n In col
    
        oCategoria.Carrega CLng(n)
        
        With cbbCategoria
            .AddItem
            .List(.ListCount - 1, 0) = oCategoria.Categoria
            .List(.ListCount - 1, 1) = oCategoria.ID
        End With
    
    Next n

    If sCategoria = "" Then cbbCategoria.ListIndex = -1 Else cbbCategoria.Text = sCategoria
    
End Sub

Private Sub ComboBoxCarregarSubcategorias()
    
    Dim col             As Collection
    Dim n               As Variant
    Dim sSubcategoria   As String
    
    Set col = oSubcategoria.Listar("subcategoria", oSubcategoria.CategoriaID)
    
    sSubcategoria = cbbSubcategoria.Text
    cbbSubcategoria.Clear
    
    For Each n In col
    
        oSubcategoria.Carrega CLng(n)
    
        With cbbSubcategoria
            .AddItem
            .List(.ListCount - 1, 0) = oSubcategoria.Subcategoria
            .List(.ListCount - 1, 1) = oSubcategoria.ID
        End With
    
    Next n
    
    If sSubcategoria = "" Or cbbSubcategoria.ListIndex = -1 Then
        cbbSubcategoria.ListIndex = -1
    Else
        cbbSubcategoria.Text = sSubcategoria
    End If

End Sub

Private Sub btnRegistrar_Click()
    
    ' Se um ou mais campos obrigatórios não foram preenchidos, não registra
    If Valida = True Then
           
        ' Chama a rotina de incluir registro passando 2 parâmetros:
        ' 1 - Parâmetro que informa se o registro é uma transferência entre contas,
        '     um recebimento, um pagamento ou um investimento;
        ' 2 - Parâmetro para informar se o registro é de origem de um agendamento.
        oMovimentacao.Inclui chbTransferencia.Value, oMovimentacao.IsAgendamento
        
        
        ' Se for um registro de agendamento, atualiza a próxima data
        ' de vencimento do agendamento.
        If oMovimentacao.IsAgendamento = True Then
            
            ' Chama a rotina para atualizar o agendamento
            oAgendamento.AtualizaPosRegistro
            
            'If colAgendamentos.Count = 1 Then
            '    colAgendamentos.Remove Format(oAgendamento.ID, "00000000")
            'End If
        End If
        
        MsgBox "Registrado com sucesso!", vbInformation
        
        Unload Me
    End If

End Sub
Private Function Valida() As Boolean
    
    Valida = False
    
    ' Se NÃO for 'Registro de agendamento'
    If oMovimentacao.IsAgendamento = False Then
        
        ' Se for REGISTRO DIRETO e um RECEBIMENTO ou DESPESA, entao ...
        If chbTransferencia.Value = False Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbFornecedor = Empty Then
                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbCategoria = Empty Then
                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbSubcategoria = Empty Then
                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emissão'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbValor = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
                Exit Function
            Else
                With oMovimentacao
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Liquidado = CDate(txbVencimento.Text)
                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                    .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                    .Observacao = txbObservacao.Text
                End With
                
                oCategoria.Categoria = cbbCategoria.Text
                oCategoria.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                oSubcategoria.Subcategoria = cbbSubcategoria.Text
                
                Valida = True
            End If
            
        ' Se FOR 'Transferência entre contas'
        ElseIf chbTransferencia.Value = True Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta origem'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbContaPara.Text = Empty Then
                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emissão'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbValor.Text = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbContaDe.Text = cbbContaPara.Text Then
                MsgBox "Conta Origem e Conta Destino não podem ser iguais!", vbCritical, "Campo obrigatório"
            Else
                With oMovimentacao
                    .Grupo = "T"
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Liquidado = CDate(txbVencimento.Text)
                    .Observacao = txbObservacao.Text
                End With
                Valida = True
            End If
        End If
        
    ' Se for REGISTRO DE AGENDAMENTO, então ...
    Else
    
        ' E se for RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
        If chbTransferencia.Value = False Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbFornecedor.Text = Empty Then
                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbCategoria.Text = Empty Then
                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbSubcategoria.Text = Empty Then
                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emissão'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbValor.Text = Empty Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
                Exit Function
            Else
                With oMovimentacao
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    '.ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Liquidado = CDate(txbVencimento.Text)
                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                    .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                    .Observacao = txbObservacao.Text
                End With
                
                'oCategoria.Categoria = cbbCategoria.Text
                'oCategoria.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                'oSubcategoria.CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
                'oSubcategoria.Subcategoria = cbbSubcategoria.Text
                
                Valida = True
            End If
        Else
            ' Se for REGISTRO DE AGENDAMENTO e for TRANSFERÊNCIA ENTRE CONTAS, então ...
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta origem'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf cbbContaPara.Text = Empty Then
                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Data'", vbCritical, "Campo obrigatório"
                Exit Function
            ElseIf txbValor.Text = Empty Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
                Exit Function
            Else
            
                With oMovimentacao
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Liquidado = CDate(txbVencimento.Text)
                    .Grupo = "T"
                    .Observacao = txbObservacao.Text
                End With
                
                Valida = True
            End If
        
        End If
    End If
End Function
Private Sub ComboBoxCarregarContas()

    Dim col As New Collection
    Dim n   As Variant
    Dim sContaDe As String
    Dim sContaPara As String
    
    Set col = oConta.Listar("conta")
    
    If oMovimentacao.ContaID = 0 Then
        sContaDe = cbbContaDe.Text
        sContaPara = cbbContaPara.Text
    Else
        sContaDe = oConta.Conta
        sContaPara = oContaPara.Conta
    End If
    
    cbbContaDe.Clear
    cbbContaPara.Clear
       
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
    
    If sContaDe = "" Then cbbContaDe.ListIndex = -1 Else cbbContaDe.Text = sContaDe
    If sContaPara = "" Then cbbContaPara.ListIndex = -1 Else cbbContaPara.Text = sContaPara
    
End Sub
    
Private Sub ComboBoxCarregar()

    Dim col As Collection
    Dim n   As Variant
    Dim s() As String
    
    ' Se for um REGISTRO DIRETO, então sai da rotina de popular comboboxes
    If oMovimentacao.IsAgendamento = False Then
        
        ' Carrega combo Grupos
        Set col = oCategoria.ListarGrupos
        
        For Each n In col
        
            s() = Split(n, ",")
            
            With cbbGrupo
                .AddItem
                .List(.ListCount - 1, 0) = s(0)
                .List(.ListCount - 1, 1) = s(1)
            End With
        Next n
        
        
    ' Se for um AGENDAMENTO, então ...
    Else
    
        If oMovimentacao.IsTransferencia = False Then
        
            ' Carrega combo Grupos
            Set col = oCategoria.ListarGrupos
            
            For Each n In col
            
                s() = Split(n, ",")
                
                With cbbGrupo
                    .AddItem
                    .List(.ListCount - 1, 0) = s(0)
                    .List(.ListCount - 1, 1) = s(1)
                End With
            Next n
            
            
        
            ' Carrega combo Categoria
            Set col = oCategoria.Listar("categoria", oAgendamento.Grupo)
            
            cbbCategoria.Clear
            
            For Each n In col
            
                oCategoria.Carrega CLng(n)
                
                With cbbCategoria
                    .AddItem
                    .List(.ListCount - 1, 0) = oCategoria.Categoria
                    .List(.ListCount - 1, 1) = oCategoria.ID
                End With
            
            Next n
            
            cbbCategoria.Text = oCategoria.Categoria
            
            ' Carrega o combobox Subcategoria
            Set col = oSubcategoria.Listar("subcategoria", oAgendamento.SubcategoriaID)
    
            cbbSubcategoria.Clear
            
            For Each n In col
            
                oSubcategoria.Carrega CLng(n)
            
                With cbbSubcategoria
                    .AddItem
                    .List(.ListCount - 1, 0) = oSubcategoria.Subcategoria
                    .List(.ListCount - 1, 1) = oSubcategoria.ID
                End With
            
            Next n
            
            cbbSubcategoria.Text = oSubcategoria.Subcategoria
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        ' Se AGENDAMENTO for uma TRANSFERÊNCIA ENTRE CONTAS
        Else
            
            ' Programar ...
            
        End If
    End If
    
    Set rst = Nothing

End Sub
Private Sub ComboBoxCarregarFornecedores()

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

