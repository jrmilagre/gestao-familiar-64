VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fImportaTransacoesM99 
   Caption         =   ":: Assistente de Importação ::"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18060
   OleObjectBlob   =   "fImportaTransacoesM99.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fImportaTransacoesM99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscarArquivo_Click()

    Dim lngCount As Long
    Dim sCaminho As String
 
    ' Abre o File Dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        '.Filters.Add "Arquivos de texto", "*.txt;", 1
 
        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            sCaminho = .SelectedItems(lngCount)
        Next lngCount
 
    End With
    
    ' Popula ListView
    Dim iArquivo             As Integer
    Dim sCaminhoArquivo      As String
    Dim sTextoArquivo        As String
    Dim sTextoProximaLinha   As String
    Dim lContadorLinha       As Long
    Dim arr()               As String
    
    With lstPrincipal
        .Clear
        .ColumnCount = 7
        .ColumnWidths = "40pt; 55pt; 180pt; 170pt; 180pt; 65pt;"
        .Font = "Consolas"
        
        iArquivo = FreeFile
        sCaminhoArquivo = sCaminho
    
        ' Abre o arquivo para leitura
        Open sCaminhoArquivo For Input As iArquivo
        
        lContadorLinha = 1
        
        ' Lê o conteúdo do arquivo linha a linha
        Do While Not EOF(iArquivo)
            Line Input #iArquivo, sTextoProximaLinha
            
            If sTextoProximaLinha <> "" Then
                arr() = Split(sTextoProximaLinha, vbTab)
                
                If UBound(arr) = 7 Then
                
                    If IsDate(arr(1)) Then
                        .AddItem
                        .List(.ListCount - 1, 0) = Format(lContadorLinha, "000000")
                        .List(.ListCount - 1, 1) = CDate(arr(1))
                        .List(.ListCount - 1, 2) = IIf(arr(2) = "", "Outros", arr(2))
                        .List(.ListCount - 1, 3) = arr(3)
                        .List(.ListCount - 1, 4) = arr(5)
                        .List(.ListCount - 1, 5) = Space(12 - Len(Format(arr(7), "#,##0.00"))) & Format(arr(7), "#,##0.00")
                        .List(.ListCount - 1, 6) = arr(4)
                        
                        lContadorLinha = lContadorLinha + 1
                        
                    End If
                End If
                
                
            End If
            
            sTextoProximaLinha = sTextoProximaLinha & vbCrLf
            sTextoArquivo = sTextoArquivo & sTextoProximaLinha
        Loop
        
        Close iArquivo
        
    End With
    
    lblResumo.Caption = "Total de " & Format(lContadorLinha - 1, "#,##0") & " transações"
    
End Sub

Private Sub btnImportar_Click()

    Dim vbResposta      As VBA.VbMsgBoxResult
    Dim c               As Long
    Dim oFornecedor     As cFornecedor
    Dim oConta          As cConta
    Dim oContaPara      As cContaPara
    Dim oMovimentacao   As cMovimentacao
    Dim oCategoria      As cCategoria
    Dim oSubcategoria   As cSubcategoria
    Dim arr()           As String
    
    Set oFornecedor = New cFornecedor
    Set oConta = New cConta
    Set oContaPara = New cContaPara
    Set oMovimentacao = New cMovimentacao
    Set oCategoria = New cCategoria
    Set oSubcategoria = New cSubcategoria
    
    vbResposta = MsgBox("Deseja gravar as transações no sistema?", vbYesNo + vbQuestion)
    
    If vbResposta = vbYes Then
    
        For c = 0 To lstPrincipal.ListCount - 1
            
            With lstPrincipal
            
                ' Quebra categoria em "Categoria e Subcategoria"
                If InStr(.List(c, 4), " :") > 0 Then
                    arr() = Split(.List(c, 4), " : ")
                Else
                    arr(0) = .List(c, 4)
                    arr(1) = "Não-atribuído"
                End If
                
                ' Quando a categoria for "Transferir de" significa que é uma TRANSFERÊNCIA
                If arr(0) = "Transferir de" Or arr(0) = "Transferir para" Then
                    
                    ' Verifica se existe a conta de origem para cadastrar ou não
                    oConta.Conta = IIf(arr(0) = "Transferir de", arr(1), .List(c, 3))
                    If oConta.Existe(oConta.Conta) = False Then
                        oConta.SaldoInicial = 0
                        oConta.Inclui
                    End If
                    
                    ' Verifica se existe a conta de destino para cadastrar ou não
                    oContaPara.Conta = IIf(arr(0) = "Transferir de", .List(c, 3), arr(1))
                    If oContaPara.Existe(oContaPara.Conta) = False Then
                        oContaPara.SaldoInicial = 0
                        oContaPara.Inclui
                    End If
                      
                    'oMovimentacao.CategoriaID = oCategoria.ID
                    oMovimentacao.Grupo = "T"
                    oMovimentacao.ContaID = oConta.ID
                    oMovimentacao.ContaParaID = oContaPara.ID
                    oMovimentacao.Valor = CDbl(Abs(.List(c, 5)))
                    oMovimentacao.Liquidado = CDate(.List(c, 1))
                    oMovimentacao.Observacao = .List(c, 6)
                    oMovimentacao.Inclui True, False
                    
                ' É um RECEBIMENTO ou DESPESA
                Else
                
                    ' Verifica se existe Fornecedor para cadastrar ou não
                    oFornecedor.NomeFantasia = CStr(.List(c, 2))
                    If oFornecedor.Existe(oFornecedor.NomeFantasia) = False Then oFornecedor.Inclui
                    
                    ' Verifica se existe Conta para cadastrar ou não
                    oConta.Conta = .List(c, 3)
                    If oConta.Existe(oConta.Conta) = False Then
                        oConta.SaldoInicial = 0
                        oConta.Inclui
                    End If
                    
                    ' Grupo
                    If CDbl(.List(c, 5)) < 0 Then
                        oMovimentacao.Grupo = "D"
                    Else
                        oMovimentacao.Grupo = "R"
                    End If
                    
                    ' Verifica se existe Categoria para cadastrar ou não
                    oCategoria.Grupo = oMovimentacao.Grupo
                    oCategoria.Categoria = arr(0)
                    If (oCategoria.Existe(arr(0), oMovimentacao.Grupo)) = False Then oCategoria.Inclui
                
                    ' Verifica se existe Subcategoria
                    oSubcategoria.Subcategoria = arr(1)
                    If oSubcategoria.Existe(oCategoria.ID, oSubcategoria.Subcategoria) = False Then
                        oSubcategoria.CategoriaID = oCategoria.ID
                        oSubcategoria.Inclui
                    End If
                    
                    
                    oMovimentacao.CategoriaID = oCategoria.ID
                    oMovimentacao.ContaID = oConta.ID
                    oMovimentacao.FornecedorID = oFornecedor.ID
                    oMovimentacao.Liquidado = CDate(.List(c, 1))
                    oMovimentacao.Origem = "Importação txt"
                    oMovimentacao.SubcategoriaID = oSubcategoria.ID
                    oMovimentacao.Valor = CDbl(.List(c, 5))
                    oMovimentacao.Observacao = .List(c, 6)
                    oMovimentacao.Inclui False, False
                
                End If
            End With
        Next c
        MsgBox "Transações gravadas com sucesso!", vbInformation
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub
