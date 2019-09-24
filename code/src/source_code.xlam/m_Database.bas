Attribute VB_Name = "m_Database"
Option Explicit         ' Obriga a declaração de variáveis
Option Private Module   ' Deixa o módulo privado (invisível)

Public cnn  As ADODB.Connection  ' Objeto de conexão com o banco de dados
Public rst  As ADODB.Recordset   ' Objeto de armazenamento de dados
Public cat  As ADOX.Catalog
Public sSQL As String

' Função para efetuar conexão com o banco de dados
' ---É necessário habilitar a biblioteca Microsoft ActiveX Data Objects 2.8 Library
' ---para o funcionamento desta função
Public Function Conecta() As Boolean
    
    ' Declara varíavel
    Dim sCaminho As String
    Dim vbResultado As VBA.VbMsgBoxResult
    
    ' Define o caminho do banco de dados
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & "database.mdb"
    
    ' Cria objeto de conexão com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    ' Inicia status da conexão como falso (desconectado)
    Conecta = False
    
    ' Se a conexão der erro, desvia para o rótulo Sair
    On Error GoTo Sair
    
    ' Com o objeto conexão, escolhe o provedor e abre o banco de dados
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set cat.ActiveConnection = cnn
    End With
    
    ' Se a conexão estiver funcionando, retorna verdadeiro
    Conecta = True
    
    ' Sai da função
    Exit Function

' Rótulo Sair
Sair:
    ' Mensagem caso a conexão com o banco de dados der problema
    vbResultado = MsgBox("Banco de dados não existe ou não está acessível:" & vbNewLine & _
           vbNewLine & "Caminho do banco procurado: " & vbNewLine & _
           vbNewLine & sCaminho & vbNewLine & vbNewLine & _
           "Deseja criar o arquivo de banco de dados?", vbInformation + vbYesNo)
    
    If vbResultado = vbYes Then
        Call CriaBancoDeDados(sCaminho)
    Else
        Exit Function
    End If

           

End Function

' Função para efetuar a desconexão com o banco de dados
' --- É necessário habilitar a biblioteca "Microsoft ActiveX Data Objects 2.8 Library"
' --- para o funcionamento desta função.
Public Sub Desconecta()

    ' Fecha conexão com o banco de dados
    cnn.Close
    Set cat = Nothing

End Sub


' Procedimento para criar o banco de dados
' --- É necessário habilitar a biblioteca "Microsoft ADO Ext. 2.8 for DDL and Security"
' --- para o funcionamento deste procedimento.
Private Sub CriaBancoDeDados(Caminho As String)
     
    ' Declara variável
    Dim oCatalogo As New ADOX.Catalog
     
    ' Cria o banco de dados
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho
    
    ' Rotina para criar tabelas
    Call CriaTabelas(Caminho)
    
    ' Mensagem de conclusão
    MsgBox "Banco de dados criado com sucesso!", vbInformation
    
End Sub


' +------------------------------------------+
' |Tipos de Dados SQL |Tipos de dados do JET |
' +------------------------------------------+
' | BIT               | YES/NO               |
' | BYTE              | NUMERIC - BYTE       |
' | COUNTER           | COUNTER -contador    |
' | CURRENCY          | CURRENCY - Moeda     |
' | DateTime          | DATE/TIME            |
' | SINGLE            | NUMERIC - SINGLE     |
' | DOUBLE            | NUMERIC - DOUBLE     |
' | SHORT             | NUMERIC - INTEGER    |
' | LONG              | NUMERIC - LONG       |
' | LONGTEXT          | MEMO                 |
' | LONGBINARY        | OLE OBJECTS          |
' | Text              | Text                 |
' +------------------------------------------+

' Rotina para criar tabelas no banco de dados
Private Sub CriaTabelas(Caminho As String)

    Dim tbl As New ADOX.Table
    
    ' Abre catálogo
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open Caminho
        Set cat.ActiveConnection = cnn
    End With
    
    With tbl
        .name = "tbl_fornecedores"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "nome_fantasia", adVarWChar, 60
            .Item("nome_fantasia").Properties("Description") = "Informe o nome do fornecedor."
            .Item("nome_fantasia").Properties("Nullable") = False
            .Append "razao_social", adVarWChar, 120
            .Append "endereco", adVarWChar, 120
            .Append "numero", adVarWChar, 15
            .Append "bairro", adVarWChar, 60
            .Append "cidade", adVarWChar, 60
            .Append "estado", adVarWChar, 2
            .Append "pais", adVarWChar, 60
            .Append "data_cadastro", adDate
            .Append "deletado", adBoolean
        End With
    End With
    
    cat.Tables.Append tbl
    
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_contas"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "conta", adVarWChar, 50
            .Item("conta").Properties("Description") = "Informe o nome da conta."
            .Item("conta").Properties("Nullable") = False
            .Append "saldo_inicial", adCurrency
            .Item("saldo_inicial").Properties("Description") = "Informe o saldo inicial da conta."
            .Item("saldo_inicial").Properties("Nullable") = False
            .Append "data_cadastro", adDate
            .Append "deletado", adBoolean
        End With
    End With
    
    cat.Tables.Append tbl
    
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_subcategorias"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "subcategoria", adVarWChar, 70
            .Item("subcategoria").Properties("Description") = "Informe o nome da subcategoria."
            .Item("subcategoria").Properties("Nullable") = False
            .Append "categoria_id", adInteger
            .Append "deletado", adBoolean
        End With
    End With
    
    cat.Tables.Append tbl
    
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_categorias"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "grupo", adVarWChar, 1
            .Item("grupo").Properties("Description") = "Informe o grupo da categoria."
            .Item("grupo").Properties("Nullable") = False
            .Append "categoria", adVarWChar, 50
            .Item("categoria").Properties("Description") = "Informe o nome da categoria."
            .Item("categoria").Properties("Nullable") = False
            .Append "deletado", adBoolean
        End With
    End With
    
    cat.Tables.Append tbl
    
    Set cat = Nothing
    Call Desconecta
    
End Sub
Private Sub CriaTabelasOld()

    Dim sNomeTabela As String
    Dim sSQL As String
    
    If Conecta = True Then
            
        ' Se não existir tabela, cria
        If ExisteTabela("tbl_contas") = False Then
            sSQL = "CREATE TABLE tbl_contas (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "conta TEXT (50), saldo_inicial CURRENCY)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_fornecedores") = False Then
            sSQL = "CREATE TABLE tbl_fornecedores (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "nome_fantasia TEXT (60), razao_social TEXT (120), endereco TEXT (120), "
            sSQL = sSQL & "numero TEXT (15), bairro TEXT (60), cidade TEXT (60), estado TEXT (2), "
            sSQL = sSQL & "pais TEXT (60), data_cadastro DATETIME)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_movimentacoes") = False Then
            sSQL = "CREATE TABLE tbl_movimentacoes (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "agendamento_id LONG, conta_id SHORT, subcategoria_id SHORT, fornecedor_id SHORT, "
            sSQL = sSQL & "grupo TEXT(1), liquidado DATETIME, valor CURRENCY, origem TEXT(15), "
            sSQL = sSQL & "observacao LONGTEXT, parcela SHORT, transferencia_id LONG)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_categorias") = False Then
            sSQL = "CREATE TABLE tbl_categorias (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "grupo TEXT (1), categoria TEXT (50))"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_subcategorias") = False Then
            sSQL = "CREATE TABLE tbl_subcategorias (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "categoria_id SHORT, subcategoria TEXT (70))"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_agendamentos") = False Then
            sSQL = "CREATE TABLE tbl_agendamentos (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "conta_id SHORT, contapara_id SHORT, subcategoria_id SHORT, fornecedor_id SHORT, grupo TEXT(1), "
            sSQL = sSQL & "recorrente BIT, infinito BIT, periodicidade TEXT(10), parcelas SHORT, "
            sSQL = sSQL & "vencimento DATETIME, valor CURRENCY, observacao LONGTEXT, parcelas_quitadas SHORT, "
            sSQL = sSQL & "parcelas_restantes SHORT, intervalo SHORT, deletado BIT)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_transferencias") = False Then
            sSQL = "CREATE TABLE tbl_transferencias (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "data DATETIME, valor CURRENCY, movimentacaode_id LONG, movimentacaopara_id LONG)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_bens") = False Then
            sSQL = "CREATE TABLE tbl_bens (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "tipo TEXT(30), bem TEXT(50), aquisicao DATETIME, valor CURRENCY)"
            cnn.Execute (sSQL)
        End If
    
        If ExisteTabela("tbl_cartoes") = False Then
            sSQL = "CREATE TABLE tbl_cartoes (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "conta_id SHORT, titular TEXT(30), numero TEXT(16), desde TEXT(5), "
            sSQL = sSQL & "vencimento TEXT(5), seguranca TEXT(3), limite CURRENCY)"
            cnn.Execute (sSQL)
        End If
        
        If ExisteTabela("tbl_faturas") = False Then
            sSQL = "CREATE TABLE tbl_faturas (id COUNTER CONSTRAINT primarykey PRIMARY KEY, "
            sSQL = sSQL & "cartao_id SHORT, data DATETIME, historico TEXT(100), valor CURRENCY, "
            sSQL = sSQL & "conta_id SHORT, subcategoria_id SHORT, fornecedor_id LONG,"
            sSQL = sSQL & "movimentacao_id LONG, agendamento_id LONG, vencimento DATETIME,"
            sSQL = sSQL & "status TEXT(1))"
            cnn.Execute (sSQL)
        End If
        
    End If
    
    Call Desconecta
    
End Sub
Private Function ExisteTabela(Tabela As String) As Boolean
    
    ' Inicia retorno da função como tabela não existente
    ExisteTabela = False
    
    ' Armazena esquema de tabelas no Recordset
    Set rst = cnn.OpenSchema(adSchemaTables)
    
    ' Laço para percorrer todas as tabelas
    Do Until rst.EOF
    
        ' Se a tabela do laço for igual a tabela verificada
        ' significa que existe, então muda o retorno da função
        ' para True e sai da função
        If rst!Table_Name = Tabela Then
            ExisteTabela = True
            GoTo Sair
        End If
        
        ' Move para a próxima tabela
        rst.MoveNext
    Loop
Sair:
    ' Destrói objeto Recordset
    Set rst = Nothing
End Function


