VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type TMovimentacao
    ID                  As Long
    AgendamentoID       As Long
    ContaID             As Long
    ContaParaID         As Long
    CategoriaID         As Long
    SubcategoriaID      As Long
    FornecedorID        As Long
    Grupo               As String
    Liquidado           As Date
    Valor               As Currency
    Observacao          As String
    Origem              As String
    Parcela             As Integer
    TransferenciaID     As Long
    IsAgendamento       As Boolean
    IsTransferencia     As Boolean
    DiasExtrato         As Long
End Type

Private This As TMovimentacao
Private Const tbl As String = "tbl_movimentacoes"

' Propriedade: ID
Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: ID do agendamento
Public Property Get AgendamentoID() As Long: AgendamentoID = This.AgendamentoID: End Property
Public Property Let AgendamentoID(ByVal Value As Long): This.AgendamentoID = Value: End Property

' Propriedade: ID da conta
Public Property Get ContaID() As Long: ContaID = This.ContaID: End Property
Public Property Let ContaID(ByVal Value As Long): This.ContaID = Value: End Property

' Propriedade: ID da conta
Public Property Get ContaParaID() As Long: ContaParaID = This.ContaParaID: End Property
Public Property Let ContaParaID(ByVal Value As Long): This.ContaParaID = Value: End Property

' Propriedade: ID da categoria
Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

' Propriedade: ID da subcategoria
Public Property Get SubcategoriaID() As Long: SubcategoriaID = This.SubcategoriaID: End Property
Public Property Let SubcategoriaID(ByVal Value As Long): This.SubcategoriaID = Value: End Property

' Propriedade: ID da fornecedor
Public Property Get FornecedorID() As Long: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Long): This.FornecedorID = Value: End Property

' Propriedade: Tipo
Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

' Propriedade: Vencimento
Public Property Get Liquidado() As Date: Liquidado = This.Liquidado: End Property
Public Property Let Liquidado(ByVal Value As Date): This.Liquidado = Value: End Property

' Propriedade: Valor
Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

' Propriedade: Observação
Public Property Get Observacao() As String: Observacao = This.Observacao: End Property
Public Property Let Observacao(ByVal Value As String): This.Observacao = Value: End Property

' Propriedade: Origem
Public Property Get Origem() As String: Origem = This.Origem: End Property
Public Property Let Origem(ByVal Value As String): This.Origem = Value: End Property

' Propriedade: Parcelas Quitadas
Public Property Get Parcela() As Integer: Parcela = This.Parcela: End Property
Public Property Let Parcela(ByVal Value As Integer): This.Parcela = Value: End Property

' Propriedade: ID da transferência
Public Property Get TransferenciaID() As Long: TransferenciaID = This.TransferenciaID: End Property
Public Property Let TransferenciaID(ByVal Value As Long): This.TransferenciaID = Value: End Property

Public Property Get IsAgendamento() As Boolean: IsAgendamento = This.IsAgendamento: End Property
Public Property Let IsAgendamento(ByVal Value As Boolean): This.IsAgendamento = Value: End Property

Public Property Get IsTransferencia() As Boolean: IsTransferencia = This.IsTransferencia: End Property
Public Property Let IsTransferencia(ByVal Value As Boolean): This.IsTransferencia = Value: End Property

Public Property Get DiasExtrato() As Long: DiasExtrato = This.DiasExtrato: End Property
Public Property Let DiasExtrato(ByVal Value As Long): This.DiasExtrato = Value: End Property

Public Sub Carrega(ID As Long)

    Dim rstContaPara As ADODB.Recordset

    ' Consulta SQL
    sSQL = "SELECT * FROM tbl_movimentacoes WHERE id = " & ID
    
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
    
    '---Propriedades
    With rst
        Me.ID = .Fields("id").Value
        Me.ContaID = .Fields("conta_id").Value
        Me.SubcategoriaID = IIf(IsNull(.Fields("subcategoria_id").Value), 0, .Fields("subcategoria_id").Value)
        Me.FornecedorID = IIf(IsNull(.Fields("fornecedor_id").Value), 0, .Fields("fornecedor_id").Value)
        Me.Liquidado = .Fields("liquidado").Value
        Me.Valor = .Fields("valor").Value
        Me.Grupo = .Fields("grupo").Value
        Me.Observacao = IIf(IsNull(.Fields("observacao").Value), "", .Fields("observacao").Value)
        Me.Origem = .Fields("origem").Value
        Me.AgendamentoID = IIf(IsNull(.Fields("agendamento_id").Value), 0, .Fields("agendamento_id").Value)
        Me.Parcela = IIf(IsNull(.Fields("parcela").Value), 0, .Fields("parcela").Value)
        
        ' Se for uma MOVIMENTAÇÃO de TRANSFERÊNCIA ENTRE CONTAS
        If .Fields("grupo").Value = "T" Then
        
            ' Coleta o número da transferência na "tbl_transferencias"
            Me.TransferenciaID = .Fields("transferencia_id").Value
            
            ' Consulta SQL para coletar a Conta de contra-partida
            If rst.Fields("valor").Value < 0 Then
                sSQL = "SELECT tbl_contas.id, tbl_contas.conta FROM tbl_movimentacoes LEFT JOIN tbl_contas "
                sSQL = sSQL & "ON tbl_movimentacoes.conta_id = tbl_contas.id "
                sSQL = sSQL & "WHERE tbl_movimentacoes.id = " & Me.ID + 1
            Else
                sSQL = "SELECT tbl_contas.id, tbl_contas.conta FROM tbl_movimentacoes LEFT JOIN tbl_contas "
                sSQL = sSQL & "ON tbl_movimentacoes.conta_id = tbl_contas.id "
                sSQL = sSQL & "WHERE tbl_movimentacoes.id = " & Me.ID - 1
            End If
                    
            ' Atribui resultado da consulta SQL ao recordset
            Set rstContaPara = New ADODB.Recordset
                
            With rstContaPara
                .CursorLocation = adUseServer
                .Open Source:=sSQL, _
                      ActiveConnection:=cnn, _
                      CursorType:=adOpenDynamic, _
                      LockType:=adLockOptimistic, _
                      Options:=adCmdText
            End With
                
            If Me.Valor > 0 Then
                Me.ContaID = rstContaPara.Fields("id").Value
                Me.ContaParaID = rst.Fields("conta_id").Value
            Else
                Me.ContaID = rst.Fields("conta_id").Value
                Me.ContaParaID = rstContaPara.Fields("id").Value
            End If
                
        End If

    End With
    
    Set rst = Nothing
    
End Sub

Public Sub Inclui(IsTransferencia As Boolean, IsAgendamento As Boolean)

    Dim i As Integer
    Dim oTransferencia As cTransferencia
    Dim rstTransferencia As ADODB.Recordset
    
    Set rst = New ADODB.Recordset
    
    ' Se for TRANSFERÊNCIA ENTRE CONTAS, inclui 2 registros na tabela de movimentações (tbl_movimentacoes)
    If IsTransferencia = True Then
    
    
        ' Registra transferência na tabela de controle de transferências (tbl_transferencias)
        
        ' Cria novo objeto
        Set oTransferencia = New cTransferencia
        Set rstTransferencia = New ADODB.Recordset
        
        ' Instancia tabela "tbl_transferencias" no Recordset
        rstTransferencia.Open Source:="tbl_transferencias", ActiveConnection:=cnn, CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, Options:=adCmdTable
        
        ' Inclui um registro na "tbl_transferencias"
        rstTransferencia.AddNew
            
        ' Armazena número do novo id
        oTransferencia.ID = rstTransferencia.Fields("id").Value
        
        With rst
        
            .Open Source:="tbl_movimentacoes", _
                  ActiveConnection:=cnn, _
                  CursorType:=adOpenDynamic, _
                  LockType:=adLockOptimistic, _
                  Options:=adCmdTable
        
            For i = 1 To 2
            
                .AddNew
                
                Me.ID = .Fields("id").Value
                .Fields("conta_id").Value = IIf(i = 1, Me.ContaID, Me.ContaParaID)
                .Fields("grupo").Value = "T"
                .Fields("liquidado").Value = Me.Liquidado
                .Fields("valor").Value = IIf(i = 1, Me.Valor * -1, Me.Valor)
                .Fields("origem").Value = IIf(IsAgendamento = True, "Agendamento", Me.Origem)
                .Fields("observacao").Value = Me.Observacao
                .Fields("transferencia_id").Value = oTransferencia.ID
                
                If IsAgendamento = True Then
                    .Fields("agendamento_id").Value = oMovimentacao.AgendamentoID
                    
                    If oAgendamento.Infinito = False Then
                        .Fields("parcela").Value = oAgendamento.ParcelasQuitadas + 1
                    End If
                    
                End If
                
                Select Case i
                    Case 1: oTransferencia.MovimentacaoDeID = Me.ID
                    Case 2: oTransferencia.MovimentacaoParaID = Me.ID
                End Select
                
                ' Inclui registro na "tbl_movimentacoes"
                .Update
            
            Next i
            
            
            ' Inclui registro na "tbl_transferencias"
            rstTransferencia.Fields("data").Value = Me.Liquidado
            rstTransferencia.Fields("valor").Value = Abs(Me.Valor)
            rstTransferencia.Fields("movimentacaode_id").Value = oTransferencia.MovimentacaoDeID
            rstTransferencia.Fields("movimentacaopara_id").Value = oTransferencia.MovimentacaoParaID
            rstTransferencia.Update
            
            Set rstTransferencia = Nothing
            Set oTransferencia = Nothing
            
        End With
    
    ' Se for um RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
    Else
    
        With rst
            .Open Source:="tbl_movimentacoes", _
                  ActiveConnection:=cnn, _
                  CursorType:=adOpenDynamic, _
                  LockType:=adLockOptimistic, _
                  Options:=adCmdTable
            
            .AddNew
            
            Me.ID = .Fields("id").Value
            .Fields("conta_id").Value = Me.ContaID
            .Fields("subcategoria_id").Value = Me.SubcategoriaID
            .Fields("fornecedor_id").Value = Me.FornecedorID
            .Fields("grupo").Value = Me.Grupo
            .Fields("liquidado").Value = Me.Liquidado
            
            
            
            ' Se for o registro de um AGENDAMENTO, então ...
            If IsAgendamento = True Then
                .Fields("agendamento_id").Value = oMovimentacao.AgendamentoID
                .Fields("origem").Value = "Agendamento"
                
                If ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor > 0) Or (Me.Grupo = "R" And Me.Valor < 0) Then
                    .Fields("valor").Value = Me.Valor * -1
                ElseIf ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor < 0) Or (Me.Grupo = "R" And Me.Valor > 0) Then
                    .Fields("valor").Value = Me.Valor
                End If
                
                If oAgendamento.Recorrente = False Then
                    .Fields("observacao").Value = Me.Observacao & " -> Agendamento: " & Format(oAgendamento.ID, "00000000") & " (Parcela única)"
                Else
                    If oAgendamento.Infinito = False Then
                        .Fields("parcela").Value = oAgendamento.ParcelasQuitadas + 1
                        .Fields("observacao").Value = Me.Observacao & " -> Agendamento: " & Format(oAgendamento.ID, "00000000") & " (Parcela " & oAgendamento.ParcelasQuitadas + 1 & " de " & oAgendamento.Parcelas & ")"
                    Else
                        .Fields("observacao").Value = Me.Observacao & " -> Agendamento: " & Format(oAgendamento.ID, "00000000") & " (Infinito)"
                    End If
                End If
                
            ' Se for um REGISTRO DIRETO, então ...
            Else
                .Fields("observacao").Value = Me.Observacao
                .Fields("origem").Value = Me.Origem
                If ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor > 0) Or (Me.Grupo = "R" And Me.Valor < 0) Then
                    .Fields("valor").Value = Me.Valor * -1
                ElseIf ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor < 0) Or (Me.Grupo = "R" And Me.Valor > 0) Then
                    .Fields("valor").Value = Me.Valor
                End If
            End If
            
            .Update
        End With
    End If
    
    Set rst = Nothing
       
End Sub
Public Sub Altera(IsTransferencia As Boolean, Valor As Double)
    
    Set rst = New ADODB.Recordset
    
    ' Se for RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
    If IsTransferencia = False Then
    
        ' Comando SQL
        sSQL = "SELECT * FROM tbl_movimentacoes WHERE id = " & Me.ID
    
        'Armazena na variável o comando que fará a consulta SQL no BD Access
        With rst
            .Open Source:=sSQL, _
                  ActiveConnection:=cnn, _
                  CursorType:=adOpenDynamic, _
                  LockType:=adLockOptimistic
            
            ' Atribui novos valores aos campos
            
            .Fields("conta_id").Value = Me.ContaID
            .Fields("subcategoria_id").Value = Me.SubcategoriaID
            .Fields("fornecedor_id").Value = Me.FornecedorID
            .Fields("grupo").Value = Me.Grupo
            .Fields("liquidado").Value = Me.Liquidado
            .Fields("observacao").Value = Me.Observacao
            
            If ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor > 0) Or (Me.Grupo = "R" And Me.Valor < 0) Then
                .Fields("valor").Value = Me.Valor * -1
            ElseIf ((Me.Grupo = "D" Or Me.Grupo = "I") And Me.Valor < 0) Or (Me.Grupo = "R" And Me.Valor > 0) Then
                .Fields("valor").Value = Me.Valor
            End If
            
            ' Grava alterações
            .Update
        End With
        
    ' Se for TRANSFERÊNCIA ENTRE CONTAS, então ...
    Else
        
        Dim i As Integer
        Dim oTransferencia As cTransferencia
        
        Set oTransferencia = New cTransferencia
        
        oTransferencia.Altera Me.TransferenciaID, Me.Liquidado, Me.Valor
        
        oTransferencia.Carrega Me.TransferenciaID
        
        
        
        For i = 1 To 2
            
            Select Case i
                Case 1
                    sSQL = "SELECT * FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = " & oTransferencia.MovimentacaoDeID
                Case 2
                    sSQL = "SELECT * FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = " & oTransferencia.MovimentacaoParaID
            End Select
            
            ' Instancia novo Recordset
            Set rst = New ADODB.Recordset
            
            
            With rst
                .Open Source:=sSQL, _
                      ActiveConnection:=cnn, _
                      CursorType:=adOpenDynamic, _
                      LockType:=adLockOptimistic
                  
                .Fields("conta_id").Value = (IIf(i = 1, Me.ContaID, Me.ContaParaID))
                .Fields("liquidado").Value = oTransferencia.Data
                .Fields("valor").Value = IIf(i = 1, oTransferencia.Valor * -1, oTransferencia.Valor)
                .Fields("observacao").Value = Me.Observacao
                .Update
            End With
            
        Next i
    
    End If
    
    Set rst = Nothing
End Sub
Public Sub Exclui(Transferencia As Boolean, Valor As Currency, Agendamento As Boolean)

    Dim vbResposta As VbMsgBoxResult
    Dim i As Integer
    Dim oTransferencia As cTransferencia

    Set rst = New ADODB.Recordset   ' Cria novo objeto recordset
    
    ' Se o registro não for oriundo de agendamento...
    If Agendamento = False Then
        
        ' Se for um RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
        If Transferencia = False Then
            
            ' Consulta SQL
            sSQL = "DELETE FROM tbl_movimentacoes WHERE id = " & Me.ID
            cnn.Execute sSQL
                
        ' Se for TRANSFERÊNCIA ENTRE CONTAS, então ...
        ElseIf Transferencia = True Then
            
            Set oTransferencia = New cTransferencia
            
            oTransferencia.Carrega Me.TransferenciaID
            
            ' Laço para excluir as 2 movimentações relacionadas a transferência
            For i = 1 To 2
            
                sSQL = "DELETE FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = "
            
                Select Case i
                    Case 1
                        sSQL = sSQL & oTransferencia.MovimentacaoDeID
                    Case 2
                        sSQL = sSQL & oTransferencia.MovimentacaoParaID
                End Select
                
                cnn.Execute sSQL
                
            Next i
            
            oTransferencia.Exclui Me.TransferenciaID
            
        End If
        
        Set rst = Nothing
        
        MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusão"
    
    ' Se o registro for de origem de AGENDAMENTO, fará as tratativas para
    ' alterar ou estornar o agendamento
    Else
        ' Verifica se o registro vinculado ao agendamento corresponde ao mais recente
        ' e verificar se o registro corresponde a última parcela
        Set oAgendamento = New cAgendamento
        oAgendamento.Carrega (Me.AgendamentoID)
        
        ' Se o agendamento for de parcela única, pergunta ao usuário se ele deseja
        ' recuperar o agendamento
        If oAgendamento.Recorrente = False Then
        
'            vbResposta = MsgBox("Esse registro é vinculado a um AGENDAMENTO de PARCELA ÚNICA." & _
'                         VBA.vbNewLine & _
'                         VBA.vbNewLine & _
'                         "Deseja recuperar o AGENDAMENTO?", vbYesNo, "Pergunta")
'
'            ' Se a resposta for SIM, estorna o agendamento
'            If vbResposta = vbYes Then
                oAgendamento.Estorna Me.AgendamentoID, oAgendamento.Recorrente, oAgendamento.Infinito
'            End If
            
            ' Exclui as MOVIMENTAÇÕES
            ' Se for RECEBIMENTO, DESPESA ou INVESTIMENTO, então ...
            If Transferencia = False Then
            
                ' Depois de fazer as tratativas com o agendamento, exclui a movimentação
                sSQL = "DELETE FROM tbl_movimentacoes WHERE id = " & Me.ID
                cnn.Execute sSQL
                
            ' e se for TRANSFERÊNCIA ENTRE CONTAS, então ...
            Else

                Set oTransferencia = New cTransferencia
                
                oTransferencia.Carrega Me.TransferenciaID
                
                ' Laço para excluir as 2 movimentações relacionadas a transferência
                For i = 1 To 2
                
                    sSQL = "DELETE FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = "
                
                    Select Case i
                        Case 1
                            sSQL = sSQL & oTransferencia.MovimentacaoDeID
                        Case 2
                            sSQL = sSQL & oTransferencia.MovimentacaoParaID
                    End Select
                    
                    cnn.Execute sSQL
                    
                Next i
                
                oTransferencia.Exclui Me.TransferenciaID
            
            End If
            
            Set rst = Nothing
            
            MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusão"
            
        ' Se o agendamento for recorrente...
        ElseIf oAgendamento.Recorrente = True Then
            
            ' Se o agendamento for recorrente e for infinito, regride a data em uma unidade
            ' de acordo com a periodicidade e o intervalo do agendamento
            If oAgendamento.Infinito = True Then
            
                vbResposta = MsgBox("Esse registro é vinculado ao Agendamento nº: " & Format(Me.AgendamentoID, "00000000") & "." & _
                             VBA.vbNewLine & _
                             VBA.vbNewLine & _
                             "Deseja estornar o AGENDAMENTO?", vbYesNo, "Pergunta")
                             
                If vbResposta = vbYes Then
                    oAgendamento.Estorna Me.AgendamentoID, oAgendamento.Recorrente, oAgendamento.Infinito
                End If
                
                ' Depois de fazer as tratativas com o agendamento, exclui o registro
                If Transferencia = False Then
                
                    ' Consulta SQL
                    sSQL = "DELETE FROM tbl_movimentacoes WHERE id = " & Me.ID
                    cnn.Execute sSQL
                
                ' Se não for um agendamento, e, se for transferência e o valor for menor que zero
                ElseIf Transferencia = True Then
                
                    Set oTransferencia = New cTransferencia
                    
                    oTransferencia.Carrega Me.TransferenciaID
                    
                    ' Laço para excluir as 2 movimentações relacionadas a transferência
                    For i = 1 To 2
                    
                        sSQL = "DELETE FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = "
                        
                        Select Case i
                            Case 1
                                sSQL = sSQL & oTransferencia.MovimentacaoDeID
                            Case 2
                                sSQL = sSQL & oTransferencia.MovimentacaoParaID
                        End Select
                        
                        cnn.Execute sSQL
                        
                    Next i
                    
                    oTransferencia.Exclui Me.TransferenciaID
                    
                End If
                
                Set rst = Nothing
                
                MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusão"
                
            ' Se o agendamento for recorrente e tiver término
            ElseIf oAgendamento.Infinito = False Then
                ' Verifica se é a parcela mais recente registrada
                If Me.Parcela = oAgendamento.ParcelasQuitadas Then

                    oAgendamento.Estorna Me.AgendamentoID, oAgendamento.Recorrente, oAgendamento.Infinito

                    ' Depois de fazer as tratativas com o agendamento, exclui o registro
                    sSQL = "DELETE FROM tbl_movimentacoes WHERE id = " & Me.ID
                    cnn.Execute sSQL
                    
                    If Transferencia = False Then
                        sSQL = "DELETE FROM tbl_movimentacoes WHERE id = " & Me.ID
                        cnn.Execute sSQL
                    Else
                        Set oTransferencia = New cTransferencia
                        oTransferencia.Carrega Me.TransferenciaID
                        ' Laço para excluir as 2 movimentações relacionadas a transferência
                        For i = 1 To 2
                        
                            sSQL = "DELETE FROM tbl_movimentacoes WHERE tbl_movimentacoes.id = "
                            
                            Select Case i
                                Case 1
                                    sSQL = sSQL & oTransferencia.MovimentacaoDeID
                                Case 2
                                    sSQL = sSQL & oTransferencia.MovimentacaoParaID
                            End Select
                            
                            cnn.Execute sSQL
                            
                        Next i
                        
                        oTransferencia.Exclui Me.TransferenciaID
                    End If
                    
                    MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusão"
                Else
                
                    MsgBox "O registro não corresponde ao último registro do agendamento." & vbNewLine & _
                    "Não posso excluir." & vbNewLine & _
                    "Localize o última parcela registrada do agendamento para fazer a exclusão", vbInformation
                    
                End If
            End If
        End If
    End If

End Sub
Public Function PreencheListBoxContas(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM tbl_contas "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set PreencheListBoxContas = col

End Function
Public Function SaldoAnteriorExtrato(ContaID As Integer) As Currency

    Dim Dias As Long
    
    Dias = IIf(Me.DiasExtrato = 0, 30, Me.DiasExtrato)

    sSQL = "SELECT SUM(valor) as saldo_anterior "
    sSQL = sSQL & "FROM tbl_movimentacoes "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "conta_id = " & ContaID & " and "
    sSQL = sSQL & "liquidado < DATE() - " & Dias & " "
    
    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn
    
    SaldoAnteriorExtrato = IIf(IsNull(rst.Fields("saldo_anterior").Value), 0, rst.Fields("saldo_anterior").Value)
    
    Set rst = Nothing

End Function
Public Function ListaMovimentacoes(ContaID As Integer) As Collection

    Dim col As New Collection
    Dim Dias As Long
    
    Dias = IIf(Me.DiasExtrato = 0, 30, Me.DiasExtrato)

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & tbl & " "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "conta_id = " & ContaID & " and "
    sSQL = sSQL & "liquidado >= DATE() - " & Dias & " "
    sSQL = sSQL & "ORDER BY liquidado"

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set ListaMovimentacoes = col

End Function

Public Function CarregaContaID(ID As Long) As Long

    sSQL = "SELECT conta_id FROM tbl_movimentacoes WHERE id = " & ID
    
    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn

    CarregaContaID = rst.Fields("conta_id").Value
    
    Set rst = Nothing
    
End Function
