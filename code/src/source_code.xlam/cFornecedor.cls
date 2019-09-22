VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TFornecedor
    ID              As Long
    NomeFantasia    As String
    RazaoSocial     As String
    Endereco        As String
    Numero          As String
    Bairro          As String
    Cidade          As String
    Estado          As String
    Pais            As String
    DataCadastro    As Date
End Type

Private This As TFornecedor
Private Const tbl As String = "tbl_fornecedores"

' Propriedade ID
Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: Nome fantasia
Public Property Get NomeFantasia() As String: NomeFantasia = This.NomeFantasia: End Property
Public Property Let NomeFantasia(ByVal Value As String): This.NomeFantasia = Value: End Property

' Propriedade: Razao social
Public Property Get RazaoSocial() As String: RazaoSocial = This.RazaoSocial: End Property
Public Property Let RazaoSocial(ByVal Value As String): This.RazaoSocial = Value: End Property

' Propriedade: Endereço
Public Property Get Endereco() As String: Endereco = This.Endereco: End Property
Public Property Let Endereco(ByVal Value As String): This.Endereco = Value: End Property

' Propriedade: Número
Public Property Get Numero() As String: Numero = This.Numero: End Property
Public Property Let Numero(ByVal Value As String): This.Numero = Value: End Property

' Propriedade: Bairro
Public Property Get Bairro() As String: Bairro = This.Bairro: End Property
Public Property Let Bairro(ByVal Value As String): This.Bairro = Value: End Property

' Propriedade: Cidade
Public Property Get Cidade() As String: Cidade = This.Cidade: End Property
Public Property Let Cidade(ByVal Value As String): This.Cidade = Value: End Property

' Propriedade: Estado (União Federativa)
Public Property Get Estado() As String: Estado = This.Estado: End Property
Public Property Let Estado(ByVal Value As String): This.Estado = Value: End Property

' Propriedade: País
Public Property Get Pais() As String: Pais = This.Pais: End Property
Public Property Let Pais(ByVal Value As String): This.Pais = Value: End Property

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
        Me.NomeFantasia = .Fields("nome_fantasia").Value
        Me.RazaoSocial = IIf(IsNull(.Fields("razao_social").Value), "", .Fields("razao_social").Value)
        Me.DataCadastro = .Fields("data_cadastro").Value
        Me.Endereco = .Fields("endereco").Value
        Me.Numero = .Fields("numero").Value
        Me.Bairro = .Fields("bairro").Value
        Me.Cidade = .Fields("cidade").Value
        Me.Estado = .Fields("estado").Value
        Me.Pais = .Fields("pais").Value
        Me.DataCadastro = .Fields("data_cadastro").Value
    End With
    
End Sub

' Rotina para incluir registro no banco de dados
Public Sub Inclui()

    Set rst = New ADODB.Recordset

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=tbl, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.ID = .Fields("id").Value
        .Fields("nome_fantasia").Value = Me.NomeFantasia
        .Fields("razao_social").Value = Me.RazaoSocial
        .Fields("data_cadastro").Value = Me.DataCadastro
        .Fields("endereco").Value = Me.Endereco
        .Fields("numero").Value = Me.Numero
        .Fields("bairro").Value = Me.Bairro
        .Fields("cidade").Value = Me.Cidade
        .Fields("estado").Value = Me.Estado
        .Fields("pais").Value = Me.Pais
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

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("nome_fantasia").Value = Me.NomeFantasia
        .Fields("razao_social").Value = Me.RazaoSocial
        .Fields("data_cadastro").Value = Me.DataCadastro
        .Fields("endereco").Value = Me.Endereco
        .Fields("numero").Value = Me.Numero
        .Fields("bairro").Value = Me.Bairro
        .Fields("cidade").Value = Me.Cidade
        .Fields("estado").Value = Me.Estado
        .Fields("pais").Value = Me.Pais
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub

' Rotina para Excluir registro do banco de dados
Public Sub Exclui()
    
    sSQL = "UPDATE " & tbl & " SET deletado = True WHERE id = " & Me.ID
    
    cnn.Execute sSQL
    
End Sub
Public Function PreencheListBox() As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM tbl_fornecedores "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "ORDER BY nome_fantasia "

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set PreencheListBox = col

End Function
