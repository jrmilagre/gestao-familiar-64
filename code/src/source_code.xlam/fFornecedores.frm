VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fFornecedores 
   Caption         =   ":: Cadastro de Fornecedores ::"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   OleObjectBlob   =   "fFornecedores.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oFornecedor         As New cFornecedor
Private bListBoxOrdenando   As Boolean
Private colControles        As New Collection



Private Sub UserForm_Initialize()
     
    Call lstPrincipalPopular("nome_fantasia")
    Call cbbEstadoPopular
    Call EventosCampos
    Call Campos("Desabilitar")
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    MultiPage1.Value = 0

End Sub
Private Sub lblHdFornecedor_Click():
    Call lstPrincipalPopular("nome_fantasia")
End Sub

Private Sub lblHdEndereco_Click()
    Call lstPrincipalPopular("endereco")
End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage1.Value = 1
End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_EventoCampo
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
        
            If TypeName(oControle) = "TextBox" Then
                
                Set oEvento = New c_EventoCampo
                
                With oEvento
                    
                    If oControle.Tag = "cpf" Then
                        Stop
                    End If
                    oControle.ControlTipText = cat.Tables("tbl_fornecedores").Columns(oControle.Tag).Properties("Description").Value
                    
                    .FieldType = cat.Tables("tbl_fornecedores").Columns(oControle.Tag).Type
                    .MaxLength = cat.Tables("tbl_fornecedores").Columns(oControle.Tag).DefinedSize
                    .Nullable = cat.Tables("tbl_fornecedores").Columns(oControle.Tag).Properties("Nullable")
                    
                    Set .cGeneric = oControle
                    
                End With
                    
                colControles.Add oEvento
                
            End If
        End If
    Next

End Sub


' Botão confirmar
Private Sub btnConfirmar_Click()
    
    Dim vbResposta  As VbMsgBoxResult
    Dim sDecisao    As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida = True Then
        
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            If sDecisao = vbNewLine & "Inclusão" Then
            
                ' Chama método para incluir registro no banco de dados
                oFornecedor.Inclui
                Call lstPrincipalPopular("nome_fantasia")
                ' Inclui registro na ListBox
'                With lstPrincipal
'                    .AddItem
'                    .List(.ListCount - 1, 0) = oFornecedor.NomeFantasia
'                    .List(.ListCount - 1, 1) = oFornecedor.ID
'                    .List(.ListCount - 1, 2) = oFornecedor.Endereco
'                End With
                    
                
                'Call ListBoxOrdenar
                
            ElseIf sDecisao = vbNewLine & "Alteração" Then
                
                ' Chama método para alterar dados no banco de dados
                oFornecedor.Altera
                Call lstPrincipalPopular("nome_fantasia")
                ' Replica as alterações na ListBox
'                With lstPrincipal
'                    .List(.ListIndex, 0) = oFornecedor.NomeFantasia
'                    .List(.ListIndex, 2) = oFornecedor.Endereco
'                End With
                
                
                'Call ListBoxOrdenar
                    
            ElseIf sDecisao = vbNewLine & "Exclusão" Then
                        
                ' Chama método para deletar registro do banco de dados
                oFornecedor.Exclui
                Call lstPrincipalPopular("nome_fantasia")
                ' Remove item da ListBox
                'lstPrincipal.RemoveItem (lstPrincipal.ListIndex)
            End If
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
                
        ElseIf vbResposta = vbNo Then
        
            ' Se a resposta for não, executa a rotina atribuída ao clique do botão cancelar
            Call btnCancelar_Click
            
        End If
        
        Call Campos("Limpar")                   ' Chama sub-rotina para limpar campos e objeto
        lstPrincipal.Enabled = True      ' Habilita ListBox
        Call Campos("Desabilitar")     ' Chama sub-rotina para desabilitar campos
        
        btnConfirmar.Visible = False: btnCancelar.Visible = False
        
        btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
        
        
        btnAlterar.Enabled = False          ' Desabilita botão alterar
        btnExcluir.Enabled = False          ' Desabilita botão excluir
        btnIncluir.SetFocus                 ' Coloca o foco no botão incluir
        
        ' Tira a seleção
        If lstPrincipal.ListIndex >= 0 Then lstPrincipal.Selected(lstPrincipal.ListIndex) = False
        
        MultiPage1.Value = 0
        
    End If
End Sub

'Private Sub txtPesquisa_Change()
'
'    bPesquisando = True
'    Call PopulaListBox
'    bPesquisando = False
'End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
End Sub
Private Sub btnAlterar_Click()
    Call PosDecisaoTomada("Alteração")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclusão")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnConfirmar.Visible = True: btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & VBA.vbNewLine & Decisao
    btnCancelar.Caption = "Cancelar " & VBA.vbNewLine & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    MultiPage1.Value = 1
    
    If Decisao = "Inclusão" Then
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclusão" Then
        Call Campos("Habilitar")
        txbNomeFantasia.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    If bListBoxOrdenando = False Then
    
        If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
        If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
        
        If lstPrincipal.ListIndex >= 0 Then
            oFornecedor.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
        End If
        
        lblID.Caption = Format(IIf(oFornecedor.ID = 0, "", oFornecedor.ID), "00000")
        lblHdNomeFantasia.Caption = oFornecedor.NomeFantasia
        txbNomeFantasia.Text = oFornecedor.NomeFantasia
        txbRazaoSocial.Text = oFornecedor.RazaoSocial
        txbEndereco.Text = oFornecedor.Endereco
        txbNumero.Text = oFornecedor.Numero
        txbBairro.Text = oFornecedor.Bairro
        txbCidade.Text = oFornecedor.Cidade
        If oFornecedor.Estado = "" Then
            cbbEstado.ListIndex = 0
        Else
            cbbEstado.Text = oFornecedor.Estado
        End If
        txbPais.Text = oFornecedor.Pais
        txbDataCadastro.Text = oFornecedor.DataCadastro
    
    End If

End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    'txtPesquisa.Enabled = True
    btnIncluir.SetFocus
    
    lstPrincipal.Enabled = True
   
    MultiPage1.Value = 0
    
    ' Tira a seleção
    lstPrincipal.ListIndex = -1

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbNomeFantasia.Enabled = False: lblNomeFantasia.Enabled = False
        txbRazaoSocial.Enabled = False: lblRazaoSocial.Enabled = False
        txbEndereco.Enabled = False: lblEndereco.Enabled = False
        txbNumero.Enabled = False: lblNumero.Enabled = False
        txbBairro.Enabled = False: lblBairro.Enabled = False
        txbCidade.Enabled = False: lblCidade.Enabled = False
        cbbEstado.Enabled = False: lblEstado.Enabled = False
        txbPais.Enabled = False: lblPais.Enabled = False
        txbDataCadastro.Enabled = False: lblDataCadastro.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbNomeFantasia.Enabled = True: lblNomeFantasia.Enabled = True
        txbRazaoSocial.Enabled = True: lblRazaoSocial.Enabled = True
        txbEndereco.Enabled = True: lblEndereco.Enabled = True
        txbNumero.Enabled = True: lblNumero.Enabled = True
        txbBairro.Enabled = True: lblBairro.Enabled = True
        txbCidade.Enabled = True: lblCidade.Enabled = True
        cbbEstado.Enabled = True: lblEstado.Enabled = True
        txbPais.Enabled = True: lblPais.Enabled = True
        lblDataCadastro.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblHdNomeFantasia.Caption = ""
        txbNomeFantasia.Text = ""
        txbRazaoSocial.Text = ""
        txbEndereco.Text = ""
        txbNumero.Text = ""
        txbBairro.Text = ""
        txbCidade.Text = ""
        cbbEstado.ListIndex = -1
        txbPais.Text = ""
        txbDataCadastro.Text = ""
        lstPrincipal.ListIndex = -1
    End If

End Sub
Private Sub ListBoxOrdenar()
    
    Dim ini, fim, i, j  As Long
    Dim sCol01          As String
    Dim sCol02          As String
    
    bListBoxOrdenando = True
    
    With lstPrincipal
        
        ini = 0
        fim = .ListCount - 1 '4 itens(0 - 3)
        
        For i = ini To fim - 1  ' Laço para comparar cada item com todos os outros itens
            For j = i + 1 To fim    ' Laço para comparar item com o próximo item
                If .List(i) > .List(j) Then
                    sCol01 = .List(j, 0)
                    sCol02 = .List(j, 1)
                    .List(j, 0) = .List(i, 0)
                    .List(j, 1) = .List(i, 1)
                    .List(i, 0) = sCol01
                    .List(i, 1) = sCol02
                End If
            Next j
        Next i
    End With
    
    bListBoxOrdenando = False
    
End Sub
Private Sub lstPrincipalPopular(OrderBy As String)

    Dim col As New Collection
    
    Set col = oFornecedor.PreencheListBox(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina número de colunas
        .ColumnWidths = "170 pt; 0pt; 180pt;"      ' Configura largura das colunas
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oFornecedor.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oFornecedor.NomeFantasia
            .List(.ListCount - 1, 1) = oFornecedor.ID
            .List(.ListCount - 1, 2) = oFornecedor.Endereco
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub
Private Sub cbbEstadoPopular()
    With cbbEstado
        .Clear
        .AddItem ""
        .AddItem "SP"
        .AddItem "MG"
    End With
End Sub
Private Function Valida() As Boolean
    
    Valida = False
    
    If txbNomeFantasia.Text = Empty Then
        MsgBox "Nome Fantasia é um campo obrigatório", vbInformation: txbNomeFantasia.SetFocus
    Else
        ' Envia valores preenchidos no formulário para o objeto
        With oFornecedor
            .NomeFantasia = txbNomeFantasia.Text
            .RazaoSocial = txbRazaoSocial.Text
            .Endereco = txbEndereco.Text
            .Numero = txbNumero.Text
            .Bairro = txbBairro.Text
            .Cidade = txbCidade.Text
            .Estado = cbbEstado.Text
            .Pais = txbPais.Text
        End With
        
        Valida = True
    End If
    
End Function
Private Function CampoObrigatorio() As Boolean
    'cat.Tables("tbl_forneced
    
End Function
Private Function FormatoCampo() As Variant

End Function
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oFornecedor = Nothing
    Call Desconecta
End Sub
