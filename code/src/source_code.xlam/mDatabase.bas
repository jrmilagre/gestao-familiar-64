Attribute VB_Name = "mDatabase"
Option Explicit         ' Obriga a declara��o de vari�veis
Option Private Module   ' Deixa o m�dulo privado (invis�vel)

Public cnn As ADODB.Connection  ' Objeto de conex�o com o banco de dados
Public rst As ADODB.Recordset   ' Objeto de armazenamento de dados
Public sSQL As String

' Fun��o para efetuar conex�o com o banco de dados
' ---� necess�rio habilitar a biblioteca Microsoft ActiveX Data Objects 2.8 Library
' ---para o funcionamento desta fun��o
Public Function Conecta() As Boolean
    
    ' Declara var�avel
    Dim sCaminho As String
    Dim vbResultado As VBA.VbMsgBoxResult
    
    ' Define o caminho do banco de dados
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & "database.mdb"
    
    ' Cria objeto de conex�o com o banco de dados
    Set cnn = New ADODB.Connection
    
    ' Inicia status da conex�o como falso (desconectado)
    Conecta = False
    
    ' Se a conex�o der erro, desvia para o r�tulo Sair
    On Error GoTo Sair
    
    ' Com o objeto conex�o, escolhe o provedor e abre o banco de dados
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
    End With
    
    ' Se a conex�o estiver funcionando, retorna verdadeiro
    Conecta = True
    
    ' Sai da fun��o
    Exit Function

' R�tulo Sair
Sair:
    ' Mensagem caso a conex�o com o banco de dados der problema
    vbResultado = MsgBox("Banco de dados n�o existe ou n�o est� acess�vel:" & vbNewLine & _
           vbNewLine & "Caminho do banco procurado: " & vbNewLine & _
           vbNewLine & sCaminho & vbNewLine & vbNewLine & _
           "Deseja criar o arquivo de banco de dados?", vbInformation + vbYesNo)
    
    If vbResultado = vbYes Then
        Call CriaBancoDeDados(sCaminho)
    Else
        Exit Function
    End If

           

End Function

' Fun��o para efetuar a desconex�o com o banco de dados
' --- � necess�rio habilitar a biblioteca "Microsoft ActiveX Data Objects 2.8 Library"
' --- para o funcionamento desta fun��o.
Public Sub Desconecta()

    ' Fecha conex�o com o banco de dados
    cnn.Close

End Sub


' Procedimento para criar o banco de dados
' --- � necess�rio habilitar a biblioteca "Microsoft ADO Ext. 2.8 for DDL and Security"
' --- para o funcionamento deste procedimento.
Private Sub CriaBancoDeDados(Caminho As String)
     
    ' Declara vari�vel
    Dim oCatalogo As New ADOX.Catalog
     
    ' Cria o banco de dados
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho
    
    ' Rotina para criar tabelas
    Call CriaTabelas
    
    ' Mensagem de conclus�o
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
Private Sub CriaTabelas()

    Dim sNomeTabela As String
    Dim sSQL As String
    
    If Conecta = True Then
            
        ' Se n�o existir tabela, cria
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
    
    ' Inicia retorno da fun��o como tabela n�o existente
    ExisteTabela = False
    
    ' Armazena esquema de tabelas no Recordset
    Set rst = cnn.OpenSchema(adSchemaTables)
    
    ' La�o para percorrer todas as tabelas
    Do Until rst.EOF
    
        ' Se a tabela do la�o for igual a tabela verificada
        ' significa que existe, ent�o muda o retorno da fun��o
        ' para True e sai da fun��o
        If rst!Table_Name = Tabela Then
            ExisteTabela = True
            GoTo Sair
        End If
        
        ' Move para a pr�xima tabela
        rst.MoveNext
    Loop
Sair:
    ' Destr�i objeto Recordset
    Set rst = Nothing
End Function


