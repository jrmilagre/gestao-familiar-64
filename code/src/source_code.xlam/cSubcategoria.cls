VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSubcategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TSubcategoria
    ID As Long
    CategoriaID As Long
    Subcategoria As String
End Type

Private This As TSubcategoria
Private Const tbl As String = "tbl_subcategorias"

' Propriedade: ID
Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

' Propriedade: ID da categoria
Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

' Propriedade: Subcategoria
Public Property Get Subcategoria() As String: Subcategoria = This.Subcategoria: End Property
Public Property Let Subcategoria(ByVal Value As String): This.Subcategoria = Value: End Property



' Rotina para incluir dados no banco de dados
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
        .Fields("subcategoria").Value = Me.Subcategoria
        .Fields("categoria_id").Value = Me.CategoriaID
                
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
' Rotina para incluir dados no banco de dados
Public Sub Altera(ID As Long)

    ' Consula SQL
    sSQL = "SELECT * FROM " & tbl & " WHERE id = " & ID

    ' Cria objeto recordset
    Set rst = New ADODB.Recordset
    
    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        .Fields("subcategoria").Value = Me.Subcategoria
                
        .Update
    End With
    
    Set rst = Nothing
       
End Sub

Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & tbl & " "
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
    If rst.EOF = False Then
        With rst
            Me.ID = .Fields("id").Value
            Me.Subcategoria = .Fields("subcategoria").Value
            Me.CategoriaID = .Fields("categoria_id").Value
        End With
    End If
    
    Set rst = Nothing
    
End Sub
Public Function Existe(CategoriaID As Long, Subcategoria As String) As Boolean

    Existe = False
    
    ' Comando SQL
    sSQL = "SELECT * FROM tbl_subcategorias WHERE "
    sSQL = sSQL & "categoria_id = " & CategoriaID & " and "
    sSQL = sSQL & "subcategoria = '" & Subcategoria & "' "
    
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
    
End Sub
Public Function Listar(OrderBy As String, CategoriaID As Long) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & tbl & " "
    sSQL = sSQL & "WHERE deletado = False and "
    sSQL = sSQL & "categoria_id = " & CategoriaID & " "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set Listar = col

End Function


