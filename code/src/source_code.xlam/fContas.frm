VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fContas 
   Caption         =   ":: Cadastro de Contas ::"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   OleObjectBlob   =   "fContas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta              As New cConta
Private colControles        As New Collection

Private Sub UserForm_Initialize()
     
    Call lstPrincipalPopular("conta")
    Call EventosCampos("tbl_contas")
    Call Campos("Desabilitar")
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    MultiPage1.Value = 0

End Sub
Private Sub lstPrincipalPopular(OrderBy As String)

    Dim col As New Collection
    
    Set col = oConta.Listar(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina número de colunas
        .ColumnWidths = "170 pt; 0pt; 55pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oConta.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
            .List(.ListCount - 1, 2) = Space(12 - Len(Format(oConta.SaldoInicial, "#,##0.00"))) & Format(oConta.SaldoInicial, "#,##0.00")
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbConta.Enabled = False: lblConta.Enabled = False
        txbSaldoInicial.Enabled = False: lblSaldoInicial.Enabled = False
        txbDataCadastro.Enabled = False: lblDataCadastro.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbConta.Enabled = True: lblConta.Enabled = True
        txbSaldoInicial.Enabled = True: lblSaldoInicial.Enabled = True
        lblDataCadastro.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabConta.Caption = ""
        txbConta.Text = ""
        txbSaldoInicial.Text = ""
        txbDataCadastro.Text = ""
        lstPrincipal.ListIndex = -1
    End If

End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage1.Value = 1
End Sub
Private Sub EventosCampos(Tabela As String)

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        
    
        If Len(oControle.Tag) > 0 Then
        
            '~~~~~~~~~~~~
            Set oEvento = New c_EventoCampo
            Set oEvento = oEvento.Evento(oControle, Tabela)
            colControles.Add oEvento
            '~~~~~~~~~~~~
        
'            If TypeName(oControle) = "TextBox" Then
'
'                Set oEvento = New c_EventoCampo
'
'                With oEvento
'                    oControle.ControlTipText = cat.Tables(tbl).Columns(oControle.Tag).Properties("Description").Value
'
'                    .FieldType = cat.Tables(tbl).Columns(oControle.Tag).Type
'
'                    If .FieldType = 6 Then
'                        oControle.TextAlign = fmTextAlignRight
'                    End If
'
'                    .MaxLength = cat.Tables(tbl).Columns(oControle.Tag).DefinedSize
'                    .Nullable = cat.Tables(tbl).Columns(oControle.Tag).Properties("Nullable")
'
'                    Set .cGeneric = oControle
'
'                End With
'
'                oControle.Add oEvento
'
'            End If
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
            
                oConta.Inclui
                Call lstPrincipalPopular("conta")
                
            ElseIf sDecisao = vbNewLine & "Alteração" Then
                
                oConta.Altera
                Call lstPrincipalPopular("conta")
                    
            ElseIf sDecisao = vbNewLine & "Exclusão" Then
                        
                oConta.Exclui
                Call lstPrincipalPopular("conta")

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
        lstPrincipal.ListIndex = -1
        
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
        lstPrincipal.ListIndex = -1
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclusão" Then
        Call Campos("Habilitar")
        txbConta.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
End Sub

Private Sub lstPrincipal_Change()

    If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
    If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
    
    If lstPrincipal.ListIndex >= 0 Then
        oConta.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
    End If
    
    lblID.Caption = Format(IIf(oConta.ID = 0, "", oConta.ID), "00000")
    lblCabConta.Caption = oConta.Conta
    txbConta.Text = oConta.Conta
    txbSaldoInicial.Text = Format(oConta.SaldoInicial, "#,##0.00")
    txbDataCadastro.Text = oConta.DataCadastro
    
End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    lstPrincipal.Enabled = True
   
    MultiPage1.Value = 0
    
    lstPrincipal.ListIndex = -1

End Sub
Private Function Valida() As Boolean
    
    Valida = False
    
    If txbConta.Text = Empty Then
        MsgBox "Conta é um campo obrigatório", vbInformation: txbConta.SetFocus
    ElseIf txbSaldoInicial.Text = "" Then
        MsgBox "Saldo inicial é um campo obrigatório", vbInformation: txbSaldoInicial.SetFocus
    Else
        ' Envia valores preenchidos no formulário para o objeto
        With oConta
            .Conta = txbConta.Text
            .SaldoInicial = txbSaldoInicial.Text
        End With
        
        Valida = True
    End If
    
End Function
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oConta = Nothing
    Call Desconecta
End Sub
