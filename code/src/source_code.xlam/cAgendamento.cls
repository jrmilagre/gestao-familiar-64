VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAgendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TAgendamento
    ID                      As Long
    ContaID                 As Long
    ContaParaID             As Long
    CategoriaID             As Long
    SubcategoriaID          As Long
    FornecedorID            As Long
    Grupo                   As String
    Recorrente              As Boolean
    Infinito                As Boolean
    Periodicidade           As String
    Parcelas                As Integer
    Vencimento              As Date
    Valor                   As Currency
    Observacao              As String
    ParcelasQuitadas        As Integer
    ParcelasRestantes       As Integer
    Intervalo               As Integer
    Registro                As Long
    Deletado                As Boolean
    RegistrandoAgendamento  As Boolean
    Transferencia           As Boolean
End Type

Private This As TAgendamento
Private Const tbl As String = "tbl_agendamentos"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get ContaID() As Long: ContaID = This.ContaID: End Property
Public Property Let ContaID(ByVal Value As Long): This.ContaID = Value: End Property

Public Property Get ContaParaID() As Long: ContaParaID = This.ContaParaID: End Property
Public Property Let ContaParaID(ByVal Value As Long): This.ContaParaID = Value: End Property

Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

Public Property Get SubcategoriaID() As Long: SubcategoriaID = This.SubcategoriaID: End Property
Public Property Let SubcategoriaID(ByVal Value As Long): This.SubcategoriaID = Value: End Property

Public Property Get FornecedorID() As Long: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Long): This.FornecedorID = Value: End Property

Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

Public Property Get Recorrente() As Boolean: Recorrente = This.Recorrente: End Property
Public Property Let Recorrente(ByVal Value As Boolean): This.Recorrente = Value: End Property

Public Property Get Infinito() As Boolean: Infinito = This.Infinito: End Property
Public Property Let Infinito(ByVal Value As Boolean): This.Infinito = Value: End Property

Public Property Get Periodicidade() As String: Periodicidade = This.Periodicidade: End Property
Public Property Let Periodicidade(ByVal Value As String): This.Periodicidade = Value: End Property

Public Property Get Parcelas() As Integer: Parcelas = This.Parcelas: End Property
Public Property Let Parcelas(ByVal Value As Integer): This.Parcelas = Value: End Property

Public Property Get Vencimento() As Date: Vencimento = This.Vencimento: End Property
Public Property Let Vencimento(ByVal Value As Date): This.Vencimento = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get Observacao() As String: Observacao = This.Observacao: End Property
Public Property Let Observacao(ByVal Value As String): This.Observacao = Value: End Property

Public Property Get ParcelasQuitadas() As Integer: ParcelasQuitadas = This.ParcelasQuitadas: End Property
Public Property Let ParcelasQuitadas(ByVal Value As Integer): This.ParcelasQuitadas = Value: End Property

Public Property Get ParcelasRestantes() As Integer: ParcelasRestantes = This.ParcelasRestantes: End Property
Public Property Let ParcelasRestantes(ByVal Value As Integer): This.ParcelasRestantes = Value: End Property

Public Property Get Intervalo() As Integer: Intervalo = This.Intervalo: End Property
Public Property Let Intervalo(ByVal Value As Integer): This.Intervalo = Value: End Property

Public Property Get Registro() As Long: Registro = This.Registro: End Property
Public Property Let Registro(ByVal Value As Long): This.Registro = Value: End Property

Public Property Get Deletado() As Boolean: Deletado = This.Deletado: End Property
Public Property Let Deletado(ByVal Value As Boolean): This.Deletado = Value: End Property

Public Property Get RegistrandoAgendamento() As Boolean: RegistrandoAgendamento = This.RegistrandoAgendamento: End Property
Public Property Let RegistrandoAgendamento(ByVal Value As Boolean): This.RegistrandoAgendamento = Value: End Property

Public Property Get Transferencia() As Boolean: Transferencia = This.Transferencia: End Property
Public Property Let Transferencia(ByVal Value As Boolean): This.Transferencia = Value: End Property

Public Sub Inclui(IsTransferencia As Boolean)
    
    ' Instancia Recordset
    Set rst = New ADODB.Recordset
    
    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=tbl, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        
        ' Inclui uma nova linha no banco de dados
        .AddNew
        
        ' Atribui valores aos campos da nova linha no banco de dados
        Me.ID = .Fields("id").Value
        .Fields("conta_id").Value = Me.ContaID
        .Fields("grupo").Value = Me.Grupo
        .Fields("recorrente").Value = Me.Recorrente
        .Fields("periodicidade").Value = Me.Periodicidade
        .Fields("infinito").Value = Me.Infinito
        .Fields("intervalo").Value = Me.Intervalo
        .Fields("parcelas").Value = Me.Parcelas
        .Fields("parcelas_quitadas").Value = 0
        .Fields("parcelas_restantes").Value = Me.Parcelas
        .Fields("vencimento").Value = Me.Vencimento
        .Fields("observacao").Value = Me.Observacao
        
        If IsTransferencia = False Then
            .Fields("subcategoria_id").Value = Me.SubcategoriaID
            .Fields("fornecedor_id").Value = Me.FornecedorID
            
            If (Me.Grupo <> "R" And Me.Valor > 0) Or (Me.Grupo = "R" And Me.Valor < 0) Then
                .Fields("valor").Value = Me.Valor * -1
            ElseIf (Me.Grupo <> "R" And Me.Valor < 0) Or (Me.Grupo = "R" And Me.Valor > 0) Then
                .Fields("valor").Value = Me.Valor
            End If
        Else
            .Fields("contapara_id").Value = Me.ContaParaID
            .Fields("valor").Value = Me.Valor * -1
        End If
        
        ' Grava registro no banco de dados
        .Update
        
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Altera(IsTransferencia As Boolean)
    
    ' Instancia Recordset
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_agendamentos WHERE id = " & Me.ID
    
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("conta_id").Value = Me.ContaID
        .Fields("grupo").Value = Me.Grupo
        .Fields("recorrente").Value = Me.Recorrente
        .Fields("periodicidade").Value = Me.Periodicidade
        .Fields("infinito").Value = Me.Infinito
        .Fields("intervalo").Value = Me.Intervalo
        .Fields("parcelas").Value = Me.Parcelas
        .Fields("parcelas_restantes").Value = Me.Parcelas
        .Fields("vencimento").Value = Me.Vencimento
        .Fields("observacao").Value = Me.Observacao
        
        If IsTransferencia = False Then
            .Fields("subcategoria_id").Value = Me.SubcategoriaID
            .Fields("fornecedor_id").Value = Me.FornecedorID
            
            If (Me.Grupo <> "R" And Me.Valor > 0) Or (Me.Grupo = "R" And Me.Valor < 0) Then
                .Fields("valor").Value = Me.Valor * -1
            ElseIf (Me.Grupo <> "R" And Me.Valor < 0) Or (Me.Grupo = "R" And Me.Valor > 0) Then
                .Fields("valor").Value = Me.Valor
            End If
        Else
            .Fields("contapara_id").Value = Me.ContaParaID
            .Fields("valor").Value = Me.Valor * -1
        End If
        
        ' Grava alterações
        .Update
    End With
    
    ' Destrói objeto Recordset
    Set rst = Nothing
End Sub
Public Sub Exclui(IsTransferencia As Boolean)
        
    sSQL = "UPDATE " & tbl & " SET deletado = True WHERE id = " & Me.ID
    
    cnn.Execute sSQL
    
End Sub
Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM tbl_agendamentos "
    sSQL = sSQL & "WHERE id = " & ID
    
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
        Me.ContaParaID = IIf(IsNull(.Fields("contapara_id").Value), 0, .Fields("contapara_id").Value)
        'Me.CategoriaID = .Fields("categoria_id").Value
        Me.SubcategoriaID = IIf(IsNull(.Fields("subcategoria_id").Value), 0, .Fields("subcategoria_id").Value)
        Me.FornecedorID = IIf(IsNull(.Fields("fornecedor_id").Value), 0, .Fields("fornecedor_id").Value)
        Me.Grupo = .Fields("grupo").Value
        Me.Recorrente = .Fields("recorrente").Value
        Me.Infinito = .Fields("infinito").Value
        Me.Periodicidade = .Fields("periodicidade").Value
        Me.Parcelas = .Fields("parcelas").Value
        Me.Vencimento = .Fields("vencimento").Value
        Me.Valor = .Fields("valor").Value
        Me.Observacao = IIf(IsNull(.Fields("observacao").Value), "", .Fields("observacao").Value)
        Me.ParcelasQuitadas = .Fields("parcelas_quitadas").Value
        Me.ParcelasRestantes = IIf(IsNull(.Fields("parcelas_restantes").Value), 0, .Fields("parcelas_restantes").Value)
        Me.Intervalo = .Fields("intervalo").Value
        Me.Deletado = .Fields("deletado").Value
        Me.Transferencia = IIf(Me.Grupo = "T", True, False)
        
    End With
    
    Set rst = Nothing
    
End Sub
Public Function PreencheListBox(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id, grupo, valor "
    sSQL = sSQL & "FROM "
    sSQL = sSQL & "tbl_agendamentos "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
    
        If (rst.Fields("grupo") = "T" And rst.Fields("valor") < 0) Or rst.Fields("grupo").Value <> "T" Then
            col.Add rst.Fields("id").Value
        End If
        
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set PreencheListBox = col

End Function
Public Sub AtualizaPosRegistro()

    ' Cria objeto recordset
    Set rst = New ADODB.Recordset
        
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_agendamentos WHERE id = " & Me.ID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
              
        ' Se o agendamento NÃO for recorrente...
        If oAgendamento.Recorrente = False Then
            .Fields("deletado").Value = True
            
        ' Se o agendamento FOR recorrente
        Else
            ' Se o agendamento for recorrente e não for infinito
            If oAgendamento.Infinito = False Then
    
                ' Se não for a última parcela do agendamento
                If (oAgendamento.Parcelas - oAgendamento.ParcelasQuitadas) > 1 Then
                    .Fields("parcelas_quitadas").Value = oAgendamento.ParcelasQuitadas + 1
                    .Fields("parcelas_restantes").Value = (oAgendamento.Parcelas - (oAgendamento.ParcelasQuitadas + 1))
                    .Fields("vencimento").Value = DateAdd(oAgendamento.Periodicidade, oAgendamento.Intervalo, oAgendamento.Vencimento)
                    
                ' Se for a última parcela do agendamento
                Else
                    .Fields("parcelas_quitadas").Value = oAgendamento.ParcelasQuitadas + 1
                    .Fields("parcelas_restantes").Value = (oAgendamento.Parcelas - oAgendamento.ParcelasQuitadas) + 1
                    .Fields("deletado").Value = True
                End If
            
            ' Se o agendamento for infinito, só atualiza a data de vencimento
            Else
                .Fields("vencimento").Value = DateAdd(oAgendamento.Periodicidade, oAgendamento.Intervalo, oAgendamento.Vencimento)
            End If
            
        End If
        
        .Update
    End With
    
    ' Destrói objeto recordset
    Set rst = Nothing
End Sub
Public Sub Estorna(ID As Long, Recorrente As Boolean, Infinito As Boolean)
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_agendamentos WHERE id = " & ID

    If Recorrente = False Then
        
        'Armazena na variável o comando que fará a consulta SQL no BD Access
        With rst
            .Open Source:=sSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
            
            ' Atribui novos valores aos campos
            .Fields("deletado").Value = False
            
            
            ' Grava alterações
            .Update
        End With
        
    ' Se o agendamento for recorrente e infinito, regride a data em uma unidade
    ' de acordo com a periodicidade e o intervalo do agendamento
    ElseIf Recorrente = True And Infinito = True Then
        
        'Armazena na variável o comando que fará a consulta SQL no BD Access
        With rst
            .Open Source:=sSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
            
            ' Atribui novos valores aos campos
            .Fields("vencimento").Value = DateAdd(Me.Periodicidade, Me.Intervalo * -1, Me.Vencimento)
            
            If Me.ParcelasQuitadas = 1 Then
                Me.Vencimento = Me.Vencimento
            Else
                Me.Vencimento = DateAdd(Me.Periodicidade, Me.Intervalo * -1, Me.Vencimento)
            End If
            
            ' Grava alterações
            .Update
        End With
        
    ' Se o agendamento for recorrente e tiver término
    Else
    
        'Armazena na variável o comando que fará a consulta SQL no BD Access
        With rst
            .Open Source:=sSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
            
            ' Atribui novos valores aos campos
            .Fields("vencimento").Value = DateAdd(Me.Periodicidade, Me.Intervalo * -1, Me.Vencimento)
            .Fields("parcelas_quitadas").Value = Me.ParcelasQuitadas - 1
            .Fields("parcelas_restantes").Value = Me.ParcelasRestantes + 1
            
            Me.Vencimento = DateAdd(Me.Periodicidade, Me.Intervalo * -1, Me.Vencimento)
            Me.ParcelasQuitadas = Me.ParcelasQuitadas - 1
            Me.ParcelasRestantes = Me.ParcelasRestantes + 1
            
            If .Fields("deletado") = True Then
                .Fields("deletado").Value = False
                Me.Deletado = False
            End If
            
            
            ' Grava alterações
            .Update
        End With
        
        
    
    End If
    
    Set rst = Nothing
    
End Sub

