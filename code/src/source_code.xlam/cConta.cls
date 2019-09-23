VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TConta
    ID              As Long
    Conta           As String
    SaldoInicial    As String
    Endereco        As String
    DataCadastro    As Date
End Type

Private This As TConta
Private Const tbl As String = "tbl_contas"

' Propriedade ID
Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: Conta
Public Property Get Conta() As String: Conta = This.Conta: End Property
Public Property Let Conta(ByVal Value As String): This.Conta = Value: End Property

' Propriedade: Saldo inicial
Public Property Get SaldoInicial() As Currency: SaldoInicial = This.SaldoInicial: End Property
Public Property Let SaldoInicial(ByVal Value As Currency): This.SaldoInicial = Value: End Property

' Propriedade: Data do cadastro
Public Property Get DataCadastro() As Date: DataCadastro = This.DataCadastro: End Property
Public Property Let DataCadastro(ByVal Value As Date): This.DataCadastro = Value: End Property

' Rotina para carregar dados
Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM " & tbl & " WHERE id = " & ID
    
    ' Cria objeto Recordset
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
        Me.DataCadastro = .Fields("data_cadastro").Value
    End With
    
End Sub

' Rotina para incluir registro no banco de dados
Public Sub Inclui()

    Set rst = New ADODB.Recordset

    'Armazena na vari�vel o comando que far� a consulta SQL no BD Access
    With rst
        .Open Source:=tbl, _
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

' Rotina para alterar registro no banco de dados
Public Sub Altera()
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & tbl & " WHERE id = " & Me.ID

    'Armazena na vari�vel o comando que far� a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("conta").Value = Me.Conta
        .Fields("saldo_inicial").Value = Me.SaldoInicial
        
        ' Grava altera��es
        .Update
    End With
    
    Set rst = Nothing
End Sub

' Rotina para Excluir registro do banco de dados
Public Sub Exclui()
    
    sSQL = "UPDATE " & tbl & " SET deletado = True WHERE id = " & Me.ID
    
    cnn.Execute sSQL
    
End Sub
Public Function PreencheListBox(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & tbl & " "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set PreencheListBox = col

End Function
