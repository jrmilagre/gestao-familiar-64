Attribute VB_Name = "m_Database"
Option Explicit         ' Obriga a declaração de variáveis
Option Private Module   ' Deixa o módulo privado (invisível)

Public cnn  As ADODB.Connection  ' Objeto de conexão com o banco de dados
Public rst  As ADODB.Recordset   ' Objeto de armazenamento de dados
Public cat  As ADOX.Catalog
Public sSQL As String
Private Const sBanco As String = "database_teste.mdb"
Private sCaminho As String

' Função para efetuar conexão com o banco de dados
' ---É necessário habilitar a biblioteca Microsoft ActiveX Data Objects 2.8 Library
' ---para o funcionamento desta função
Public Function Conecta() As Boolean
    
    ' Declara varíavel
    
    Dim vbResultado As VBA.VbMsgBoxResult
    Dim sCaminho As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
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
    Call AtualizaBD
    
    ' Mensagem de conclusão
    MsgBox "Banco de dados criado com sucesso!", vbInformation
    
End Sub

Private Sub AtualizaBD()

    ' Declara variáveis
    Dim oCatalogo       As New ADOX.Catalog
    Dim sCaminho        As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    ' Cria o banco de dados se não existir
    On Error GoTo Conecta
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCaminho

Conecta:
    Set cnn = New ADODB.Connection
    
    ' Abre catálogo
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set oCatalogo.ActiveConnection = cnn        ' Instancia o catálogo
    End With
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '

    Dim FilePath As String
    Dim sText As String
    Dim myArray() As String
    Dim sTableName As String
    
    FilePath = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & "date_dictionary.csv"
    
    Open FilePath For Input As #1
    
    ' Laço para percorrer o arquivo csv que contém o dicionário de dados
    Do Until EOF(1)
    
        Line Input #1, sText
        
        ' Ignora o cabeçalho
        If Trim(sText) <> "table;field;type;size;nullable;autoincrement;description" Then
            
            myArray = Split(sText, ";")
                        
            ' VERIFICA SE EXISTE TABELA
            If sTableName <> myArray(0) Then
            
                Dim oTabela         As New ADOX.Table
                Dim bExisteTabela   As Boolean
                
                bExisteTabela = False
                
                For Each oTabela In oCatalogo.Tables
                    If oTabela.Type = "TABLE" Then
                        If oTabela.name = myArray(0) Then
                            bExisteTabela = True
                            Exit For
                        End If
                    End If
                Next oTabela
            Else
                bExisteTabela = True
            End If
            
            sTableName = myArray(0)
            
            ' Se tabela não existir, cria tabela no banco de dados
            If bExisteTabela = False Then
        
                With oTabela
                    .name = myArray(0)
                    Set .ParentCatalog = oCatalogo
                End With
            
                oCatalogo.Tables.Append oTabela
            End If
            
            ' VERIFICA SE EXISTE CAMPO
            Dim oCampo          As ADOX.Column
            Dim bExisteCampo    As Boolean
            
            Set oCampo = New ADOX.Column
            bExisteCampo = False
            
            For Each oCampo In oCatalogo.Tables(myArray(0)).Columns
                
                If oCampo.name = myArray(1) Then
                    bExisteCampo = True
                    Exit For
                End If
                
            Next oCampo
            
            Set oCampo = Nothing
            
            ' Cria o campo na tabela, caso não exista
            If bExisteCampo = False Then
            
                Set oCampo = New ADOX.Column
                
                With oCampo
                    Set .ParentCatalog = oCatalogo
                    .name = myArray(1)
                    .Type = CInt(myArray(2))
                    
                    If CInt(myArray(2)) = 202 Then
                        .DefinedSize = CInt(myArray(3))
                    End If
                    
                    If CInt(myArray(3)) <> 13 Then
                        .Properties("Nullable").Value = CBool(myArray(4))
                        .Properties("Autoincrement").Value = CBool(myArray(5))
                        .Properties("Description").Value = CStr(myArray(6))
                    End If
                    
                End With
                
                oCatalogo.Tables(myArray(0)).Columns.Append oCampo
                
                Set oCampo = Nothing
                
            End If
        
        End If
    
    Loop
    
    Close #1
    
    cnn.Close
    Set oCatalogo = Nothing

End Sub


' Rotina para criar tabelas no banco de dados
Private Sub CriaTabelas(Caminho As String)
        
    cat.Tables.Append tbl
    
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_agendamentos"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "conta_id", adInteger
            .Item("conta_id").Properties("Description") = "Informe a conta."
            .Item("conta_id").Properties("Nullable") = False
            .Append "contapara_id", adInteger
            .Item("contapara_id").Properties("Description") = "Informe a conta."
            .Append "subcategoria_id", adInteger
            .Item("subcategoria_id").Properties("Nullable") = True
            .Append "fornecedor_id", adInteger
            .Item("fornecedor_id").Properties("Description") = "Informe o fornecedor."
            .Item("fornecedor_id").Properties("Nullable") = True
            .Append "grupo", adVarWChar, 1
            .Item("grupo").Properties("Description") = "Informe o grupo."
            .Append "recorrente", adBoolean
            .Item("recorrente").Properties("Description") = "Informe a recorrencia."
            .Append "infinito", adBoolean
            .Item("infinito").Properties("Description") = "Informe se o agendamento é finito."
            .Append "periodicidade", adVarWChar, 10
            .Item("periodicidade").Properties("Description") = "Informe a periocididade."
            .Append "parcelas", adInteger
            .Item("parcelas").Properties("Description") = "Informe o número de parcelas."
            .Append "vencimento", adDate
            .Item("vencimento").Properties("Description") = "Informe a data de vencimento."
            .Item("vencimento").Properties("Nullable") = False
            .Append "valor", adCurrency
            .Item("valor").Properties("Description") = "Informe o valor do agendamento."
            .Item("valor").Properties("Nullable") = False
            .Append "observacao", adLongVarWChar
            .Item("observacao").Properties("Description") = "Acrescente alguma observação se desejar."
            .Append "parcelas_quitadas", adInteger
            .Append "parcelas_restantes", adInteger
            .Append "intervalo", adInteger
            .Append "deletado", adBoolean
        End With
    End With
    
    cat.Tables.Append tbl
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_transferencias"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "data", adDate
            .Append "valor", adCurrency
            .Append "movimentacaode_id", adInteger
            .Append "movimentacaopara_id", adInteger
        End With
    End With
    
    cat.Tables.Append tbl
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set tbl = New ADOX.Table
    
    With tbl
        .name = "tbl_movimentacoes"
        Set .ParentCatalog = cat
        With .Columns
            .Append "id", adInteger
            .Item("id").Properties("Autoincrement") = True
            .Append "agendamento_id", adInteger: .Item("agendamento_id").Properties("Nullable") = True
            .Append "conta_id", adInteger
            .Append "subcategoria_id", adInteger: .Item("subcategoria_id").Properties("Nullable") = True
            .Append "fornecedor_id", adInteger: .Item("fornecedor_id").Properties("Nullable") = True
            .Append "grupo", adVarWChar, 1: .Item("grupo").Properties("Nullable") = True
            .Append "liquidado", adDate
            .Append "valor", adCurrency
            .Append "origem", adVarWChar, 15
            .Append "observacao", adLongVarWChar: .Item("observacao").Properties("Nullable") = True
            .Append "parcela", adInteger: .Item("parcela").Properties("Nullable") = True
            .Append "transferencia_id", adInteger: .Item("transferencia_id").Properties("Nullable") = True
        End With
    End With
    
    cat.Tables.Append tbl
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
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


