VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContaPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type TContaPara
    ID              As Long
    Conta           As String
    SaldoInicial    As String
    Endereco        As String
    DataCadastro    As Date
    ContaDestino    As String
End Type

Private This As TContaPara
Private Const tbl As String = "tbl_contas"


Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: Conta
Public Property Get Conta() As String: Conta = This.Conta: End Property
Public Property Let Conta(ByVal Value As String): This.Conta = Value: End Property

' Propriedade: Saldo inicial
Public Property Get SaldoInicial() As Currency: SaldoInicial = This.SaldoInicial: End Property
Public Property Let SaldoInicial(ByVal Value As Currency): This.SaldoInicial = Value: End Property

Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM tbl_contas WHERE id = " & ID
    
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
        Me.Conta = .Fields("conta").Value
        Me.SaldoInicial = .Fields("saldo_inicial").Value
    End With
    
    Set rst = Nothing
    
End Sub
Public Sub Inclui()

    Set rst = New ADODB.Recordset

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:="tbl_contas", _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.ID = .Fields("id").Value
        .Fields("conta").Value = Me.Conta
        .Fields("saldo_inicial").Value = Me.SaldoInicial
        .Fields("data_cadastro").Value = Now
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Function Existe(Conta As String) As Boolean

    Existe = False
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_contas WHERE conta = '" & Conta & "'"
    
    ' Atribui resultado da consulta SQL ao recordset
    Set rst = cnn.Execute(sSQL)
    
    ' (EOF = True) = Fornecedor não existe
    ' (EOF = False) = Fornecedor existe
    If rst.EOF = False Then
        Me.ID = rst.Fields("id").Value
        Existe = True
    End If
    
    ' Destrói Recordset
    Set rst = Nothing
        
End Function

