VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fRegistrar 
   Caption         =   ":: Registrar ::"
   ClientHeight    =   5640
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


Private Sub UserForm_Initialize()
    
    ' Se for uma MOVIMENTA��O de origem de agendamento, ent�o ...
    If oMovimentacao.IsAgendamento = True Then
    
        ' Se o registro for um AGENDAMENTO e tamb�m um RECEBIMENTO, DESPESA ou INVESTIMENTO, ent�o ...
        If oMovimentacao.IsTransferencia = False Then
            
            ' Deixa invis�vel o checkbox de transfer�ncia
            chbTransferencia.Visible = False
            
            ' Coloca o n�mero do agendamento
            lblTitulo.Caption = "Registro do agendamento n�: " & Format(oMovimentacao.AgendamentoID, "00000000")
            
            ' Parece redundante carregar o objeto oMovimenta��o agora, sendo que ele j� � carregado
            ' na rotina Valida, mas fa�o isso porque o registro pode n�o ser oriundo de agendamento
'            With oMovimentacao
'                .ContaID = oAgendamento.ContaID
'                .SubcategoriaID = oAgendamento.SubcategoriaID
'                .FornecedorID = oAgendamento.FornecedorID
'                .Liquidado = oAgendamento.Vencimento
'                .Grupo = oAgendamento.Grupo
'                .Valor = oAgendamento.Valor
'                .AgendamentoID = oAgendamento.ID
'            End With
            
            ' Carrega detalhes necess�rios dos cadastros
            oAgendamento.Carrega oMovimentacao.AgendamentoID
            oConta.Carrega oAgendamento.ContaID 'oMovimentacao.ContaID
            oSubcategoria.Carrega oAgendamento.SubcategoriaID
            oCategoria.Carrega oSubcategoria.CategoriaID
            oFornecedor.Carrega oAgendamento.FornecedorID 'oMovimentacao.FornecedorID
            
            Call ComboBoxCarregar
            Call ComboBoxCarregarContas
            Call ComboBoxCarregarFornecedores
            
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
            
            
        
        ' Se o AGENDAMENTO a ser registrado for uma TRANSFER�NCIA ENTRE CONTAS, ent�o ...
        Else
        
            Set oConta = New cConta
            Set oContaPara = New cContaPara
            Set oFornecedor = New cFornecedor
            Set oCategoria = New cCategoria
            Set oSubcategoria = New cSubcategoria
            'Set oMovimentacao = New cMovimentacao
            Set oTransferencia = New cTransferencia
        
            lblTitulo.Caption = "Registro do agendamento n�: " & Format(oAgendamento.ID, "00000000")
            
            chbTransferencia.Value = oMovimentacao.IsTransferencia
            chbTransferencia.Visible = True
            chbTransferencia.Enabled = False
            
            oAgendamento.Carrega oAgendamento.ID
            oConta.Carrega oAgendamento.ContaID
            oContaPara.Carrega oAgendamento.ContaParaID
            
            'Call ComboBoxCarregar
            Call ComboBoxCarregarContas
            'Call ComboBoxCarregarFornecedores
            
            'cbbContaDe.Text = oConta.Conta
            'cbbContaPara.Text = oContaPara.Conta
            lblContaPara.Visible = True: cbbContaPara.Visible = True
            txbValor.Text = Format(IIf(oAgendamento.Grupo = "R", oAgendamento.Valor, oAgendamento.Valor * -1), "#,##0.00")
            txbVencimento.Text = oAgendamento.Vencimento
            txbObservacao.Text = oAgendamento.Observacao
            'oAgendamento.Vencimento = txbVencimento.Text
            
        End If
    
        btnRegistrar.SetFocus
    
    ' Se for um REGISTRO DIRETO, ent�o
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
Private Sub chbTransferencia_AfterUpdate()
    cbbContaDe.SetFocus
End Sub

'+-------------------------------------------------------+
'|                                                       |
'| CONTROLES DO FORMUL�RIO                               |
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
        lblTitulo.Caption = "Transfer�ncia entre contas"
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
            vbResposta = MsgBox("Esta Conta n�o existe. Deseja cadastr�-la?", vbQuestion + vbYesNo)
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
            vbResposta = MsgBox("Esta Conta n�o existe. Deseja cadastr�-la?", vbQuestion + vbYesNo)
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
        vbResposta = MsgBox("Este Fornecedor n�o existe, deseja cadastr�-lo?", vbQuestion + vbYesNo)
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
Private Sub txbValor_AfterUpdate()

    txbValor.Text = oValidaCampo.CampoValor(txbValor.Text)

End Sub
Private Sub txbValor_Enter()
    ' Seleciona todos os caracteres do campo
    txbValor.SelStart = 0
    txbValor.SelLength = Len(txbValor.Text)
End Sub
Private Sub txbValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8          ' Backspace (seta de apagar)
        Case 48 To 57   ' N�meros de 0 a 9
        Case 44         ' V�rgula
        If InStr(txbValor.Text, ",") Then ' Se o campo j� tiver v�rgula ent�o ele n�o adiciona
            KeyAscii = 0 ' N�o adiciona a v�rgula caso ja tenha
        Else
            KeyAscii = 44 ' Adiciona uma v�rgula
        End If
        Case Else
            KeyAscii = 0 ' N�o deixa nenhuma outra caractere ser escrito
    End Select
End Sub
Private Sub txbVencimento_AfterUpdate()
    If IsDate(txbVencimento.Text) Then
        txbVencimento.Text = Format(txbVencimento.Text, "dd/mm/yyyy")
        Exit Sub
    Else
        txbVencimento.Text = Empty
    End If
End Sub
Private Sub txbVencimento_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '---se a tecla F4 for pressionada
    If KeyCode = 115 Then
        dtDatabase = IIf(txbVencimento.Text = "", Date, txbVencimento.Text)
        txbVencimento.Text = GetCalendario
    ElseIf KeyCode = 13 Then
        If chbTransferencia.Value = False Then
            cbbGrupo.SetFocus
            cbbGrupo.DropDown
        Else
            btnRegistrar.SetFocus
        End If
    End If
End Sub
Private Sub txbVencimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With txbVencimento
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
            
            vbResposta = MsgBox("Esta Categoria n�o existe. Deseja cadastr�-la?", vbQuestion + vbYesNo)
            
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
            vbResposta = MsgBox("Esta Subcategoria n�o existe. Deseja cadastr�-la?", vbQuestion + vbYesNo)
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
Private Sub txbObservacao_AfterUpdate()
    oMovimentacao.Observacao = txbObservacao.Text
End Sub
Private Sub btnCancelar_Click()
    
    Dim i As Integer
    
    If oAgendamento.RegistrandoAgendamento = True Then
    
        If colAgendamentos.Count = 1 Then
            colAgendamentos.Remove 1
        End If
    End If
    
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    '---se for 'Registro de agendamento n�o precisa desconectar do banco
    If oMovimentacao.IsAgendamento = True Then
        'n�o desconecta o banco pois o registro � oriundo de agendamento ou registros
    Else
        Call Desconecta
    End If
End Sub
'+-------------------------------------------------------+
'|                                                       |
'| ROTINAS E FUN��ES                                     |
'|                                                       |
'+-------------------------------------------------------+

Private Sub ComboBoxCarregarCategorias()
    
    Dim sCategoria As String
    
    ' Preenche combo Categoria
    sSQL = "SELECT id, categoria "
    sSQL = sSQL & "FROM tbl_categorias "
    sSQL = sSQL & "WHERE grupo = '" & oMovimentacao.Grupo & "' "
    sSQL = sSQL & "ORDER BY categoria"
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.Recordset
    
    ' Atribui resultado da consulta SQL ao recordset
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    sCategoria = cbbCategoria.Text
    
    With cbbCategoria
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("categoria").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
    
    If sCategoria = "" Then cbbCategoria.ListIndex = -1 Else cbbCategoria.Text = sCategoria
End Sub

Private Sub ComboBoxCarregarSubcategorias()
    
    Dim sSubcategoria As String
    
    ' Carrega combo Subcategorias
    sSQL = "SELECT id, subcategoria FROM tbl_subcategorias "
    sSQL = sSQL & "WHERE categoria_id = " & oSubcategoria.CategoriaID & " "
    sSQL = sSQL & "ORDER BY subcategoria"
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.Recordset
    
    ' Atribui resultado da consulta SQL ao recordset
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    sSubcategoria = cbbSubcategoria.Text
    
    With cbbSubcategoria
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("subcategoria")
            .List(.ListCount - 1, 1) = rst.Fields("id")
    
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
    
    If sSubcategoria = "" Or cbbSubcategoria.ListIndex = -1 Then
        cbbSubcategoria.ListIndex = -1
    Else
        cbbSubcategoria.Text = sSubcategoria
    End If
End Sub

Private Sub btnRegistrar_Click()
    
    ' Se um ou mais campos obrigat�rios n�o foram preenchidos, n�o registra
    If Valida = True Then
           
        ' Chama a rotina de incluir registro passando 2 par�metros:
        ' 1 - Par�metro que informa se o registro � uma transfer�ncia entre contas,
        '     um recebimento, um pagamento ou um investimento;
        ' 2 - Par�metro para informar se o registro � de origem de um agendamento.
        oMovimentacao.Inclui chbTransferencia.Value, oMovimentacao.IsAgendamento
        
        
        ' Se for um registro de agendamento, atualiza a pr�xima data
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
    
    ' Se N�O for 'Registro de agendamento'
    If oMovimentacao.IsAgendamento = False Then
        
        ' Se for REGISTRO DIRETO e um RECEBIMENTO ou DESPESA, entao ...
        If chbTransferencia.Value = False Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbFornecedor = Empty Then
                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbCategoria = Empty Then
                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbSubcategoria = Empty Then
                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emiss�o'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbValor = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigat�rio"
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
            
        ' Se FOR 'Transfer�ncia entre contas'
        ElseIf chbTransferencia.Value = True Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta origem'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbContaPara.Text = Empty Then
                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emiss�o'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbValor.Text = Empty Or txbValor.Text = 0 Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbContaDe.Text = cbbContaPara.Text Then
                MsgBox "Conta Origem e Conta Destino n�o podem ser iguais!", vbCritical, "Campo obrigat�rio"
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
        
    ' Se for REGISTRO DE AGENDAMENTO, ent�o ...
    Else
    
        ' E se for RECEBIMENTO, DESPESA ou INVESTIMENTO, ent�o ...
        If chbTransferencia.Value = False Then
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbFornecedor.Text = Empty Then
                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbCategoria.Text = Empty Then
                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbSubcategoria.Text = Empty Then
                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Emiss�o'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbValor.Text = Empty Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigat�rio"
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
            ' Se for REGISTRO DE AGENDAMENTO e for TRANSFER�NCIA ENTRE CONTAS, ent�o ...
            If cbbContaDe.Text = Empty Then
                MsgBox "Informe a 'Conta origem'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf cbbContaPara.Text = Empty Then
                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbVencimento.Text = Empty Then
                MsgBox "Informe a 'Data'", vbCritical, "Campo obrigat�rio"
                Exit Function
            ElseIf txbValor.Text = Empty Then
                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigat�rio"
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

    ' Declara vari�veis
    Dim rstContas As ADODB.Recordset
    Dim sContaDe As String
    Dim sContaPara As String

    ' Cria novo objeto recordset
    Set rstContas = New ADODB.Recordset
    
    ' Carrega combo Contas De
    sSQL = "SELECT id, conta FROM tbl_contas ORDER BY conta"
    
    ' Atribui resultado da consulta SQL ao recordset
    With rstContas
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    ' Atribui conte�do dos Textbox as vari�veis
    If oAgendamento.ContaID = 0 Then
        sContaDe = cbbContaDe.Text
        sContaPara = cbbContaPara.Text
    Else
        sContaDe = oConta.Conta
        sContaPara = oContaPara.Conta
        
        ' oConta.Carrega oAgendamento.ContaID
        ' oContaPara.Carrega oAgendamento.ContaParaID
    End If
    
    
    
    ' Limpa os Combobox
    cbbContaDe.Clear
    cbbContaPara.Clear
    
    ' La�o para popular as Combobox
    Do Until rstContas.EOF
    
        With cbbContaDe
            .AddItem
            .List(.ListCount - 1, 0) = rstContas.Fields("conta").Value
            .List(.ListCount - 1, 1) = rstContas.Fields("id").Value
        End With
        
        With cbbContaPara
            .AddItem
            .List(.ListCount - 1, 0) = rstContas.Fields("conta").Value
            .List(.ListCount - 1, 1) = rstContas.Fields("id").Value
        End With
        
        rstContas.MoveNext
    Loop
    
    ' Destr�i recordset
    Set rstContas = Nothing
    
    ' Trata a Combobox quando o conte�do for branco
    If sContaDe = "" Then cbbContaDe.ListIndex = -1 Else cbbContaDe.Text = sContaDe
    If sContaPara = "" Then cbbContaPara.ListIndex = -1 Else cbbContaPara.Text = sContaPara
    
End Sub
    
Private Sub ComboBoxCarregar()
    
    ' Se for um REGISTRO DIRETO, ent�o sai da rotina de popular comboboxes
    If oMovimentacao.IsAgendamento = False Then
        
        ' Carrega combo Grupos
        With cbbGrupo
            .AddItem
            .List(.ListCount - 1, 0) = "Receitas"
            .List(.ListCount - 1, 1) = "R"
            .AddItem
            .List(.ListCount - 1, 0) = "Despesas"
            .List(.ListCount - 1, 1) = "D"
            .AddItem
            .List(.ListCount - 1, 0) = "Investimento"
            .List(.ListCount - 1, 1) = "I"
        End With
        
        
    ' Se for um AGENDAMENTO, ent�o ...
    Else
    
        If oMovimentacao.IsTransferencia = False Then
        
            ' Carrega combo Grupos
            With cbbGrupo
                .AddItem
                .List(.ListCount - 1, 0) = "Receitas"
                .List(.ListCount - 1, 1) = "R"
                .AddItem
                .List(.ListCount - 1, 0) = "Despesas"
                .List(.ListCount - 1, 1) = "D"
                .AddItem
                .List(.ListCount - 1, 0) = "Investimento"
                .List(.ListCount - 1, 1) = "I"
            End With
        
            ' Carrega combo Categoria
            sSQL = "SELECT id, categoria "
            sSQL = sSQL & "FROM tbl_categorias "
            sSQL = sSQL & "WHERE grupo = '" & oAgendamento.Grupo & "' "
            sSQL = sSQL & "ORDER BY categoria"
            
            ' Cria novo objeto recordset
            Set rst = New ADODB.Recordset
            
            ' Atribui resultado da consulta SQL ao recordset
            With rst
                .CursorLocation = adUseServer
                .Open Source:=sSQL, _
                      ActiveConnection:=cnn, _
                      CursorType:=adOpenDynamic, _
                      LockType:=adLockOptimistic, _
                      Options:=adCmdText
            End With
            
            With cbbCategoria
                .Clear
                Do Until rst.EOF
                    .AddItem
                    .List(.ListCount - 1, 0) = rst.Fields("categoria").Value
                    .List(.ListCount - 1, 1) = rst.Fields("id").Value
                    
                    rst.MoveNext
                Loop
                
                
                .Text = oCategoria.Categoria
            End With
            
            ' Carrega o combobox Subcategoria
            sSQL = "SELECT id, subcategoria FROM tbl_subcategorias "
            sSQL = sSQL & "WHERE categoria_id = " & oSubcategoria.CategoriaID & " "
            sSQL = sSQL & "ORDER BY subcategoria"
            
            ' Cria novo objeto recordset
            Set rst = New ADODB.Recordset
            
            ' Atribui resultado da consulta SQL ao recordset
            With rst
                .CursorLocation = adUseServer
                .Open Source:=sSQL, _
                      ActiveConnection:=cnn, _
                      CursorType:=adOpenDynamic, _
                      LockType:=adLockOptimistic, _
                      Options:=adCmdText
            End With
            
            With cbbSubcategoria
                .Clear
                Do Until rst.EOF
                    .AddItem
                    .List(.ListCount - 1, 0) = rst.Fields("subcategoria").Value
                    .List(.ListCount - 1, 1) = rst.Fields("id").Value
                    rst.MoveNext
                Loop
                .Text = oSubcategoria.Subcategoria
            End With
        
        ' Se AGENDAMENTO for uma TRANSFER�NCIA ENTRE CONTAS
        Else
            
            ' Programar ...
            
        End If
    End If
    
    Set rst = Nothing

End Sub
Private Sub ComboBoxCarregarFornecedores()

    ' Carrega combo Fornecedores
    sSQL = "SELECT id, nome_fantasia FROM tbl_fornecedores ORDER BY nome_fantasia"
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.Recordset
    
    ' Atribui resultado da consulta SQL ao recordset
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    With cbbFornecedor
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("nome_fantasia").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
End Sub

