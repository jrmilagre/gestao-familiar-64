VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pID                     As Long
Private pData                   As Date
Private pValor                  As Currency
Private pMovimentacaoDeID       As Long
Private pMovimentacaoParaID     As Long

' Propriedade: ID
Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(l As Long)
    pID = l
End Property

' Propriedade: Data
Public Property Get Data() As Date
    Data = pData
End Property
Public Property Let Data(d As Date)
    pData = d
End Property

' Propriedade: Valor
Public Property Get Valor() As Currency
    Valor = pValor
End Property
Public Property Let Valor(c As Currency)
    pValor = c
End Property

' Propriedade: ID da movimentação de origem
Public Property Get MovimentacaoDeID() As Long
    MovimentacaoDeID = pMovimentacaoDeID
End Property
Public Property Let MovimentacaoDeID(l As Long)
    pMovimentacaoDeID = l
End Property

' Propriedade: ID da movimentação de destino
Public Property Get MovimentacaoParaID() As Long
    MovimentacaoParaID = pMovimentacaoParaID
End Property
Public Property Let MovimentacaoParaID(l As Long)
    pMovimentacaoParaID = l
End Property

Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM tbl_transferencias WHERE id = " & ID
    
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
        Me.Data = .Fields("data").Value
        Me.Valor = .Fields("valor").Value
        Me.MovimentacaoDeID = .Fields("movimentacaode_id").Value
        Me.MovimentacaoParaID = .Fields("movimentacaopara_id").Value
    End With
    
    Set rst = Nothing
    
End Sub

' ROTINA INCLUSÃO
Public Sub Inclui()

    Set rst = New ADODB.Recordset

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:="tbl_transferencias", _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.ID = .Fields("id").Value
        .Fields("conta").Value = Me.Conta
        .Fields("saldo_inicial").Value = Me.SaldoInicial
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Altera(ID As Long, Data As Date, Valor As Currency)
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_transferencias WHERE id = " & ID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("data").Value = Data
        .Fields("valor").Value = Abs(Valor)
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(ID As Long)
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_transferencias WHERE id = " & ID
    
    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Grava alterações
        .Delete
    End With
    
    Set rst = Nothing
    
End Sub
