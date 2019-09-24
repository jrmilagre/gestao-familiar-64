VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TCategoria
    ID          As Long
    Grupo       As String
    Categoria   As String
End Type

Private This As TCategoria
Private Const tbl As String = "tbl_categorias"

' Propriedade: ID
Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: Grupo
Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

' Propriedade: Categoria
Public Property Get Categoria() As String: Categoria = This.Categoria: End Property
Public Property Let Categoria(ByVal Value As String): This.Categoria = Value: End Property
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
        Me.Categoria = .Fields("categoria").Value
    End With
    
    Set rst = Nothing
    
End Sub
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
        .Fields("grupo").Value = Me.Grupo
        .Fields("categoria").Value = Me.Categoria
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Altera(ID As Long)
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & tbl & " WHERE id = " & ID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("categoria").Value = Me.Categoria
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Function Existe(Categoria As String, Grupo As String) As Boolean

    Existe = False
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & tbl & " WHERE categoria = '" & Categoria & "' "
    sSQL = sSQL & "and grupo = '" & Grupo & "' "
    
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
Public Sub Exclui(ID As Long)
    
    sSQL = "UPDATE " & tbl & " SET deletado = True WHERE id = " & ID
    cnn.Execute sSQL
    sSQL = "UPDATE tbl_subcategorias SET deletado = True WHERE categoria_id = " & ID
    cnn.Execute sSQL
End Sub
Public Function PreencheListBox(OrderBy As String, Grupo As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & tbl & " "
    sSQL = sSQL & "WHERE deletado = False and "
    sSQL = sSQL & "grupo = '" & Grupo & "' "
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
