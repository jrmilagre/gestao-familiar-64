VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fAgendamentos 
   Caption         =   ":: Agendamentos ::"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13770
   OleObjectBlob   =   "fAgendamentos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fAgendamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta                  As New cConta
Private oContaPara              As New cContaPara
Private oFornecedor             As New cFornecedor
Private oCategoria              As New cCategoria
Private oSubcategoria           As New cSubcategoria
Private colControles            As New Collection

Private sDecisao                As String
Private iRegistrosSelecionados  As Integer

Private Sub UserForm_Initialize()

    Set oAgendamento = New cAgendamento
    
    Call lstPrincipalPopular("vencimento")
    
    Call Campos("Desabilitar")
    
    Call cbbDiversosPopular
    Call cbbContasPopular
    Call cbbFornecedoresPopular
    Call EventosCampos("tbl_agendamentos")
    
    lblAgendamento.Visible = False

    btnAlterar.Enabled = False
    btnConfirmar.Visible = False
    btnCancelar.Visible = False
    btnExcluir.Enabled = False
    btnRegistrar.Enabled = False
    btnIncluir.SetFocus
    optSimples.Value = True
    lblContaPara.Visible = False
    cbbContaPara.Visible = False
    
End Sub
Private Sub cbbContaPara_AfterUpdate()

    Dim vbResposta  As VbMsgBoxResult
    Dim sConta      As String
    
    If cbbContaPara.ListIndex = -1 Then
        If cbbContaPara.Text <> "" Then
            vbResposta = MsgBox("Esta Conta não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                sConta = cbbContaPara.Text
                oConta.Conta = cbbContaPara.Text
                oConta.SaldoInicial = 0
                oConta.Inclui
                Call cbbContasPopular
                cbbContaPara.Text = sConta
            Else
                cbbContaPara.ListIndex = -1
            End If
        End If
    End If
    
End Sub
Private Sub btnValor_Click()
    ccurVisor = IIf(txbValor.Text = "", 0, CCur(txbValor.Text))
    txbValor.Text = Format(GetCalculadora, "#,##0.00")
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
            Set oEvento = New c_EventoCampo
            Set oEvento = oEvento.Evento(oControle, Tabela)
            colControles.Add oEvento
        End If
    Next

End Sub

Private Sub optSimples_Click()
    With lstPrincipal
        .MultiSelect = fmMultiSelectSingle
    End With
    Call Campos("Limpar")
End Sub
Private Sub optMultiplo_Click()
    With lstPrincipal
        .MultiSelect = fmMultiSelectMulti
    End With
End Sub
Private Sub lstPrincipal_Change()
    
    '---declara variáveis
    Dim dSoma As Double
    Dim idTransacao As Long
    Dim i As Integer
    
    ' Seleciona a primeira aba do objeto Multipáginas
    MultiPage1.Value = 0
    
    ' Se houver algum item selecionado na ListBox
    If lstPrincipal.ListIndex > -1 Then
    
        iRegistrosSelecionados = 0
        
        '---laço para contar quantos registros estão selecionados
        For i = 1 To lstPrincipal.ListCount
            If lstPrincipal.Selected(i - 1) = True Then
                dSoma = dSoma + CDbl(lstPrincipal.List(i - 1, 5))
                iRegistrosSelecionados = iRegistrosSelecionados + 1
            End If
        Next i
        
        If iRegistrosSelecionados > 1 Then
            btnAlterar.Enabled = False
            btnExcluir.Enabled = False
            btnRegistrar.Enabled = True
            Call Campos("Limpar")
            
        ' Se só 1 registro tiver selecionado, carrega as informações
        ElseIf iRegistrosSelecionados = 1 Then
            
            btnAlterar.Enabled = True
            btnExcluir.Enabled = True
            btnRegistrar.Enabled = True
            
            For i = 1 To lstPrincipal.ListCount
                If lstPrincipal.Selected(i - 1) = True Then
                    oAgendamento.Carrega (CLng(lstPrincipal.List(i - 1, 0)))
                    oConta.Carrega oAgendamento.ContaID
                    
                    If oAgendamento.Grupo = "T" Then
                        oContaPara.Carrega oAgendamento.ContaParaID
                    Else
                        oFornecedor.Carrega oAgendamento.FornecedorID
                        oSubcategoria.Carrega oAgendamento.SubcategoriaID
                        oCategoria.Carrega oSubcategoria.CategoriaID
                    End If
                End If
            Next i
                   
        ElseIf iRegistrosSelecionados = 0 Then
            btnRegistrar.Enabled = False
            btnAlterar.Enabled = False
            btnExcluir.Enabled = False
            Call Campos("Limpar")
        End If
        
        If dSoma > 0 Then lblTotal.ForeColor = &HFF0000 Else lblTotal.ForeColor = &HFF
        
        lblTotal.Caption = Format(dSoma, "#,##0.00")
        
        If iRegistrosSelecionados <> 1 Then
            lblAgendamento.Visible = False
        ElseIf iRegistrosSelecionados = 1 Then
            Call InformacoesCarregar
        End If
        
        cbbCategoria.Enabled = False
        cbbSubcategoria.Enabled = False
    
    End If
    
End Sub
Private Sub cbbRecorrencia_Change()
    
    If cbbRecorrencia.ListIndex > -1 Then
        oAgendamento.Recorrente = cbbRecorrencia.List(cbbRecorrencia.ListIndex, 1)
        
        If cbbRecorrencia = "Recorrente" Then
            lblPeriodicidade.Visible = True
            cbbPeriodicidade.Visible = True
            chbTermina.Visible = True
            Call chbTermina_Click
        ElseIf cbbRecorrencia = "Uma única vez" Or cbbRecorrencia = "" Then
            lblPeriodicidade.Visible = False
            cbbPeriodicidade.Visible = False
            chbTermina.Visible = False
            txbParcelas.Visible = False
            lblParcelas.Visible = False
        End If
    End If
End Sub
Private Sub cbbRecorrencia_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If cbbRecorrencia <> Empty Then
            If cbbRecorrencia = "Uma única vez" Then
                cbbContaDe.SetFocus
            Else
                cbbPeriodicidade = "Mensal"
                cbbPeriodicidade.SetFocus
            End If
        Else
            cbbRecorrencia.SetFocus
        End If
    End If
End Sub

Private Sub cbbPeriodicidade_Change()
  
    If cbbPeriodicidade.Text <> "" Then
        chbTermina.Value = False
        chbTermina.Visible = True
        oAgendamento.Periodicidade = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 1)
        oAgendamento.Intervalo = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 2)
    Else
        chbTermina.Value = False
        chbTermina.Visible = False
    End If
End Sub
Private Sub cbbPeriodicidade_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If cbbPeriodicidade <> "" Then
        If KeyCode = 13 Then
            chbTermina.SetFocus
        End If
    End If
End Sub

Private Sub chbTermina_Click()
    
    If chbTermina.Value = True Then
        oAgendamento.Infinito = False
        lblParcelas.Visible = True
        txbParcelas.Visible = True
    Else
        oAgendamento.Infinito = True
        lblParcelas.Visible = False
        txbParcelas.Visible = False
    End If
End Sub
Private Sub chbTermina_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If chbTermina.Value = True Then
            txbParcelas.SetFocus
        Else
            cbbContaDe.SetFocus
            cbbContaDe.DropDown
        End If
    End If
End Sub

Private Sub txbParcelas_AfterUpdate()
    oAgendamento.Parcelas = IIf(txbParcelas.Text = "", 0, txbParcelas.Text)
End Sub

Private Sub cbbContaDe_AfterUpdate()
    
    Dim vbResposta  As VbMsgBoxResult
    Dim sConta      As String
    
    If cbbContaDe.ListIndex = -1 Then
        If cbbContaDe.Text <> "" Then
            vbResposta = MsgBox("Esta Conta não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                sConta = cbbContaDe.Text
                oConta.Conta = cbbContaDe.Text
                oConta.SaldoInicial = 0
                oConta.Inclui
                Call cbbContasPopular
                cbbContaDe.Text = sConta
            Else
                cbbContaDe.ListIndex = -1
            End If
        End If
    End If
    
End Sub
Private Sub cbbContaDe_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then cbbFornecedor.SetFocus
End Sub
Private Sub cbbFornecedor_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
    
    If cbbFornecedor.ListIndex > -1 Then
        oAgendamento.FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
    Else
        If cbbFornecedor.Text <> "" Then
            vbResposta = MsgBox("Este Fornecedor não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                oFornecedor.NomeFantasia = cbbFornecedor.Text
                oFornecedor.Inclui
                Call cbbFornecedoresPopular
                oAgendamento.FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
            Else
                cbbFornecedor.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub cbbFornecedor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then txbValor.SetFocus
End Sub

Private Sub txbValor_AfterUpdate()
    
    If IsNumeric(txbValor.Text) Then
        If cbbGrupo.ListIndex > -1 Then
            If oAgendamento.Grupo = "R" Then
                oAgendamento.Valor = CDbl(txbValor.Text)
            Else
                oAgendamento.Valor = CDbl(txbValor.Text) * -1
            End If
        Else
            oAgendamento.Valor = txbValor.Text
        End If
        txbValor.Text = Format(txbValor.Text, "#,##0.00")
        
        Exit Sub
    Else
        txbValor.Text = Empty
    End If
End Sub
Private Sub txbVencimento_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '---se a tecla F4 for pressionada
    If KeyCode = 115 Then
        dtDate = IIf(txbVencimento.Text = "", Date, txbVencimento.Text)
        txbVencimento.Text = GetCalendario
    ElseIf KeyCode = 13 Then
        If chbTransferencia.Value = False Then
            If cbbGrupo.Enabled = True Then
                cbbGrupo.SetFocus
                cbbGrupo.DropDown
            End If
        End If
            
    End If
End Sub

Private Sub btnVencimento_Click()
    dtDate = IIf(txbVencimento.Text = Empty, Date, txbVencimento.Text)
    txbVencimento.Text = GetCalendario
    oAgendamento.Vencimento = CDate(txbVencimento.Text)
End Sub

Private Sub cbbGrupo_AfterUpdate()

    If cbbGrupo.ListIndex > -1 Then
        
        
        If oAgendamento.Grupo <> "" Then
            oAgendamento.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
            If oAgendamento.Grupo = "D" And oAgendamento.Valor > 0 Then
                oAgendamento.Valor = oAgendamento.Valor * -1
            ElseIf oAgendamento.Grupo = "R" And oAgendamento.Valor < 0 Then
                oAgendamento.Valor = oAgendamento.Valor * -1
            End If
        Else
            oAgendamento.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
            If oAgendamento.Grupo = "D" And oAgendamento.Valor > 0 Then
                oAgendamento.Valor = oAgendamento.Valor * -1
            ElseIf oAgendamento.Grupo = "R" And oAgendamento.Valor < 0 Then
                oAgendamento.Valor = oAgendamento.Valor * -1
            End If
                
        End If
        
        oAgendamento.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
        
        If cbbGrupo.ListIndex > -1 Then
            cbbCategoria.Clear
            cbbCategoria.ListIndex = -1
            cbbSubcategoria.Clear
            cbbSubcategoria.ListIndex = -1
            Call cbbCategoriaPopular
            cbbCategoria.Style = fmStyleDropDownCombo
        End If
    End If
End Sub
Private Sub cbbGrupo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        cbbCategoria.SetFocus
        cbbCategoria.DropDown
    End If
End Sub

Private Sub cbbCategoria_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim oCategoria As cCategoria

    If cbbCategoria.ListIndex > -1 Then
        oSubcategoria.CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
        
        cbbSubcategoria.Clear
        cbbSubcategoria.ListIndex = -1
        
        If oAgendamento.Grupo <> "" And cbbCategoria.Text <> "" Then
            Call cbbSubcategoriaPopular
        End If
        cbbSubcategoria.Style = fmStyleDropDownCombo
    Else
        
        If cbbCategoria.Text <> "" Then
            cbbSubcategoria.ListIndex = -1
            
            vbResposta = MsgBox("Esta Categoria não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            
            If vbResposta = vbYes Then
                
                Set oCategoria = New cCategoria
                oCategoria.Categoria = cbbCategoria.Text
                oCategoria.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                oCategoria.Inclui
                Call cbbCategoriaPopular
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                Call cbbSubcategoriaPopular
                cbbSubcategoria.Style = fmStyleDropDownCombo
                
            Else
                cbbCategoria.ListIndex = -1
                'cbbSubcategoria.Style = fmStyleDropDownList
            End If
        End If
    End If
    
End Sub
Private Sub cbbCategoria_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        cbbSubcategoria.SetFocus
        cbbSubcategoria.DropDown
    End If
End Sub

Private Sub cbbSubcategoria_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
    
    If cbbSubcategoria.ListIndex > -1 Then
        oAgendamento.SubcategoriaID = cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1)
    Else
        If cbbSubcategoria.Text <> "" Then
            vbResposta = MsgBox("Esta Subcategoria não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                
                oSubcategoria.CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                oSubcategoria.Subcategoria = cbbSubcategoria.Text
                oSubcategoria.Inclui
                Call cbbSubcategoriaPopular
                oAgendamento.SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
            Else
                cbbSubcategoria.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub cbbSubcategoria_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnConfirmar.SetFocus
End Sub

Private Sub btnIncluir_Click()
    
    Dim i As Integer
    
    sDecisao = "Inclusão"

    ' Tira seleção(ões) de todas as linhas selecionadas na ListBox
    With lstPrincipal
        For i = 0 To .ListCount
            If .Selected(i) = True Then .Selected(i) = False
        Next i
        .Enabled = True
    End With
    
    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
    
    Call BotoesDecisaoEsconder
    Call Campos("Limpar")
    Call Campos("Habilitar")
    
    lblPeriodicidade.Visible = False
    cbbPeriodicidade.Visible = False
    chbTermina.Visible = False
    lblParcelas.Visible = False
    txbParcelas.Visible = False
    
    cbbRecorrencia.SetFocus
    
    txbVencimento.Text = Format(Date, "dd/mm/yyyy"): oAgendamento.Vencimento = CDate(txbVencimento.Text)
    txbValor.Text = Format(0, "#,##0.00"): oAgendamento.Valor = CDbl(txbValor.Text)
    
    lstPrincipal.Enabled = False
    lblCabConta.Enabled = False
    lblCabFornecedor.Enabled = False
    lblCabCatSubcat.Enabled = False
    lblCabVencimento.Enabled = False
    lblCabValor.Enabled = False

End Sub
Private Sub btnAlterar_Click()
    
    sDecisao = "Alteração"
    
    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
    
    btnConfirmar.SetFocus
    
    Call BotoesDecisaoEsconder
    Call Campos("Habilitar")
    
    cbbCategoria.Enabled = True
    cbbSubcategoria.Enabled = True
    
    lstPrincipal.Enabled = False
    lblCabConta.Enabled = False
    lblCabFornecedor.Enabled = False
    lblCabCatSubcat.Enabled = False
    lblCabVencimento.Enabled = False
    lblCabValor.Enabled = False
    
    cbbRecorrencia.SetFocus
End Sub
Private Sub btnExcluir_Click()
    sDecisao = "Exclusão"
    
    btnConfirmar.Visible = True
    btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
    
    btnConfirmar.SetFocus
    
    Call BotoesDecisaoEsconder
    
    lstPrincipal.Enabled = False
    lblCabConta.Enabled = False
    lblCabFornecedor.Enabled = False
    lblCabCatSubcat.Enabled = False
    lblCabVencimento.Enabled = False
    lblCabValor.Enabled = False
End Sub
Private Sub btnCancelar_Click()
    
    Dim i As Integer
    
    btnConfirmar.Visible = False
    btnCancelar.Visible = False
    
    Call BotoesDecisaoExibir
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnRegistrar.Enabled = False
    lstPrincipal.Enabled = True
    lblCabConta.Enabled = True
    lblCabFornecedor.Enabled = True
    lblCabCatSubcat.Enabled = True
    lblCabVencimento.Enabled = True
    lblCabValor.Enabled = True
    btnIncluir.SetFocus
    
    With lstPrincipal
        For i = 0 To .ListCount
            If .Selected(i) = True Then .Selected(i) = False
        Next i
        .Enabled = True
    End With
    
End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VbMsgBoxResult
    
    If Valida = True Then
        
        If sDecisao = "Inclusão" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do agendamento?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then
                oAgendamento.Inclui chbTransferencia.Value
            End If
        ElseIf sDecisao = "Alteração" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do agendamento?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then
                oAgendamento.Altera chbTransferencia.Value
            End If
        ElseIf sDecisao = "Exclusão" Then
            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do agendamento [" & Format(lblAgendamento.Caption, "00000000") & "] ?", vbYesNo, sDecisao & " do registro")
            If vbResposta = VBA.vbYes Then
                oAgendamento.Exclui
            End If

        End If
        
        MsgBox sDecisao & " realizada com sucesso!", vbInformation, sDecisao
        
        Call Campos("Desabilitar")
        Call lstPrincipalPopular("vencimento")
        
        btnConfirmar.Visible = False
        btnCancelar.Visible = False
        
        Call BotoesDecisaoExibir
        
        btnAlterar.Enabled = False
        btnExcluir.Enabled = False
        btnRegistrar.Enabled = False
        btnIncluir.SetFocus
            
    End If

End Sub
Private Sub chbTransferencia_Click()
    If chbTransferencia.Value = True Then
        lblFornecedor.Visible = False: cbbFornecedor.Visible = False
        lblGrupo.Visible = False: cbbGrupo.Visible = False
        lblCategoria.Visible = False: cbbCategoria.Visible = False
        lblSubcategoria.Visible = False: cbbSubcategoria.Visible = False
        lblContaPara.Visible = True: cbbContaPara.Visible = True
        lblContaDe.Caption = "Conta origem"
    Else
        lblFornecedor.Visible = True: cbbFornecedor.Visible = True
        lblGrupo.Visible = True: cbbGrupo.Visible = True
        lblCategoria.Visible = True: cbbCategoria.Visible = True
        lblSubcategoria.Visible = True: cbbSubcategoria.Visible = True
        lblContaPara.Visible = False: cbbContaPara.Visible = False
        lblContaDe.Caption = "Conta"
    End If
End Sub

'+-------------------+
'| ROTINAS E FUNÇÕES |
'+-------------------+
Private Sub cbbDiversosPopular()
    
    With cbbRecorrencia
        .AddItem
        .List(.ListCount - 1, 0) = "Uma única vez"
        .List(.ListCount - 1, 1) = False
        .AddItem
        .List(.ListCount - 1, 0) = "Recorrente"
        .List(.ListCount - 1, 1) = True
    End With
    
    With cbbPeriodicidade
        .AddItem
        .List(.ListCount - 1, 0) = "Mensal"
        .List(.ListCount - 1, 1) = "m"
        .List(.ListCount - 1, 2) = 1
        .AddItem
        .List(.ListCount - 1, 0) = "Anual"
        .List(.ListCount - 1, 1) = "yyyy"
        .List(.ListCount - 1, 2) = 1
        .AddItem
        .List(.ListCount - 1, 0) = "Quinzenal"
        .List(.ListCount - 1, 1) = "d"
        .List(.ListCount - 1, 2) = 15
    End With
    
    ' Carrega combo Grupos
    Dim col As Collection
    Dim n   As Variant
    Dim s() As String
    
    Set col = oCategoria.ListarGrupos
    
    For Each n In col
    
        s() = Split(n, ",")
        
        With cbbGrupo
            .AddItem
            .List(.ListCount - 1, 0) = s(0)
            .List(.ListCount - 1, 1) = s(1)
        End With
    Next n
    
End Sub
Private Sub cbbContasPopular()

    ' Declara variáveis
    Dim col         As New Collection
    Dim n           As Variant
    Dim sContaDe    As String
    Dim sContaPara  As String
     
    Set col = oConta.Listar("conta")
       
    ' Atribui conteúdo dos Textbox as variáveis
    sContaDe = cbbContaDe.Text
    sContaPara = cbbContaPara.Text
    
    ' Limpa o Combobox
    cbbContaDe.Clear
    cbbContaPara.Clear
    
    For Each n In col
        
        oConta.Carrega CLng(n)
    
        With cbbContaDe
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
        With cbbContaPara
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Conta
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
    Next n
    
    ' Trata a Combobox quando o conteúdo for branco
    If sContaDe = "" Then cbbContaDe.ListIndex = -1 Else cbbContaDe.Text = sContaDe
    If sContaPara = "" Then cbbContaPara.ListIndex = -1 Else cbbContaPara.Text = sContaPara
End Sub
Private Sub cbbFornecedoresPopular()

    Dim sFornecedor As String
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome_fantasia")
    
    sFornecedor = cbbFornecedor.Text
    cbbFornecedor.Clear
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.NomeFantasia
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
    If sFornecedor = "" Then cbbFornecedor.ListIndex = -1 Else cbbFornecedor.Text = sFornecedor
End Sub
Private Sub Campos(Acao As String)
    If Acao = "Limpar" Then
        chbTransferencia.Value = False
        MultiPage1.Value = 0
        cbbRecorrencia.ListIndex = -1
        cbbPeriodicidade.ListIndex = -1
        cbbContaDe.ListIndex = -1
        cbbContaPara.ListIndex = -1
        cbbFornecedor.ListIndex = -1
        txbParcelas.Text = ""
        cbbGrupo.ListIndex = -1
        cbbCategoria.ListIndex = -1
        cbbCategoria.Clear
        cbbSubcategoria.ListIndex = -1
        cbbSubcategoria.Clear
        txbVencimento.Text = ""
        txbValor.Text = ""
        lblAgendamento.Caption = ""
        lblTotal.Caption = ""
    ElseIf Acao = "Desabilitar" Then
        fraTipoSelecao.Enabled = True: optSimples.Enabled = True: optMultiplo.Enabled = True
        MultiPage1.Value = 0
        chbTransferencia.Enabled = False
        cbbRecorrencia.Enabled = False: lblRecorencia.Enabled = False
        cbbPeriodicidade.Enabled = False: lblPeriodicidade.Enabled = False
        chbTermina.Enabled = False
        txbParcelas.Enabled = False: lblParcelas.Enabled = False
        cbbContaDe.Enabled = False: lblContaDe.Enabled = False
        lblContaPara.Enabled = False: cbbContaPara.Enabled = False
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        txbVencimento.Enabled = False: lblVencimento.Enabled = False
        btnVencimento.Enabled = False
        txbValor.Enabled = False: lblValor.Enabled = False
        txbObservacao.Enabled = False: lblObservacao.Enabled = False
        cbbGrupo.Enabled = False: lblGrupo.Enabled = False
        cbbCategoria.Enabled = False: lblCategoria.Enabled = False
        cbbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False
    ElseIf Acao = "Habilitar" Then
        fraTipoSelecao.Enabled = False: optSimples.Enabled = False: optMultiplo.Enabled = False
        chbTransferencia.Enabled = True
        cbbRecorrencia.Enabled = True: lblRecorencia.Enabled = True
        cbbPeriodicidade.Enabled = True: lblPeriodicidade.Enabled = True
        chbTermina.Enabled = True
        txbParcelas.Enabled = True: lblParcelas.Enabled = True
        cbbContaDe.Enabled = True: lblContaDe.Enabled = True
        lblContaPara.Enabled = True: cbbContaPara.Enabled = True
        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        txbVencimento.Enabled = True: lblVencimento.Enabled = True
        btnVencimento.Enabled = True: lblVencimento.Enabled = True
        txbValor.Enabled = True: lblValor.Enabled = True
        txbObservacao.Enabled = True: lblObservacao.Enabled = True
        'cbbGrupo.Enabled = True
        cbbGrupo.Enabled = IIf(sDecisao = "Inclusão", True, False): lblGrupo.Enabled = True
        cbbCategoria.Enabled = True: lblCategoria.Enabled = True
        cbbSubcategoria.Enabled = True: lblSubcategoria.Enabled = True
    End If
End Sub

Private Sub lstPrincipalPopular(OrderBy As String)

    Dim col As New Collection
    
    Set col = oAgendamento.PreencheListBox(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 7                    ' Determina número de colunas
        .ColumnWidths = "0 pt; 55 pt; 98 pt; 202 pt; 144 pt; 95 pt; 40pt;"      ' Configura largura das colunas"
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oAgendamento.Carrega CLng(n)
            oConta.Carrega oAgendamento.ContaID
            
            .List(.ListCount - 1, 0) = oAgendamento.ID
            .List(.ListCount - 1, 1) = oAgendamento.Vencimento
            .List(.ListCount - 1, 2) = oConta.Conta
            
            If oAgendamento.Grupo = "T" Then
                oContaPara.Carrega oAgendamento.ContaParaID
                .List(.ListCount - 1, 3) = "Vai para a conta: " & oContaPara.Conta
                .List(.ListCount - 1, 4) = "<Transferência entre contas>"
            Else
                oSubcategoria.Carrega oAgendamento.SubcategoriaID
                oCategoria.Carrega oSubcategoria.CategoriaID
                oFornecedor.Carrega oAgendamento.FornecedorID
                .List(.ListCount - 1, 3) = oCategoria.Categoria & " : " & oSubcategoria.Subcategoria
                .List(.ListCount - 1, 4) = oFornecedor.NomeFantasia
            End If
            
            .List(.ListCount - 1, 5) = Space(12 - Len(Format(oAgendamento.Valor, "#,##0.00"))) & Format(oAgendamento.Valor, "#,##0.00")
            .List(.ListCount - 1, 6) = IIf(oAgendamento.Recorrente = 0, "Parcela única", IIf(oAgendamento.Infinito = 0, Format(oAgendamento.ParcelasQuitadas + 1, "000") & " de " & Format(oAgendamento.Parcelas, "000"), "Infinito"))
            
        Next n
        
    End With
    
    lstPrincipal.Enabled = True
    lblCabConta.Enabled = True
    lblCabFornecedor.Enabled = True
    lblCabCatSubcat.Enabled = True
    lblCabVencimento.Enabled = True
    lblCabValor.Enabled = True
    
    Call Campos("Limpar")
    
End Sub


Private Sub InformacoesCarregar()
    
    lblAgendamento.Visible = True
    
    lblAgendamento.Caption = Format(oAgendamento.ID, "00000000")
    cbbContaDe.Text = oConta.Conta
    
    ' Se for um Recebimento ou Despesa, então ...
    If oAgendamento.Grupo <> "T" Then
        chbTransferencia.Value = False
        cbbFornecedor.Text = oFornecedor.NomeFantasia
        txbVencimento.Text = oAgendamento.Vencimento
        
        'cbbGrupo.List(cbbGrupo.ListIndex, 1) = oAgendamento.Grupo
        
        If oAgendamento.Grupo = "R" Then
            cbbGrupo.Text = "Receitas"
        Else
            cbbGrupo.Text = "Despesas"
        End If
        
        cbbGrupo_AfterUpdate
        cbbCategoria.Text = oCategoria.Categoria
        cbbCategoria_AfterUpdate
        cbbSubcategoria.Text = oSubcategoria.Subcategoria
    

    
        With oAgendamento
        
            If .Valor >= 0 Then txbValor.Text = Format(.Valor, "#,##0.00") Else txbValor.Text = Format(.Valor * -1, "#,##0.00")
            
            If .Recorrente = False Then
                cbbRecorrencia.Text = "Uma única vez"
                lblPeriodicidade.Visible = False: cbbPeriodicidade.Visible = False
                chbTermina.Visible = False
                lblParcelas.Visible = False: txbParcelas.Visible = False
            ElseIf .Recorrente = True And .Infinito = False Then
                cbbRecorrencia.Text = "Recorrente"
        
                If .Periodicidade = "m" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Mensal"
                ElseIf .Periodicidade = "yyyy" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Anual"
                ElseIf .Periodicidade = "d" And .Intervalo = 15 Then
                    cbbPeriodicidade.Text = "Quinzenal"
                End If
        
                chbTermina.Value = True
                txbParcelas.Visible = True
                txbParcelas.Text = .Parcelas
            ElseIf .Recorrente = True And .Infinito = True Then
                cbbRecorrencia = "Recorrente"
                If .Periodicidade = "m" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Mensal"
                ElseIf .Periodicidade = "yyyy" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Anual"
                ElseIf .Periodicidade = "d" And .Intervalo = 15 Then
                    cbbPeriodicidade.Text = "Quinzenal"
                End If
        
                chbTermina.Value = False
                txbParcelas.Visible = False
            End If
        End With
        
        txbObservacao.Text = oAgendamento.Observacao
    
    ' Se for transferência, então ...
    Else
        chbTransferencia.Value = True
        cbbContaPara.Text = oContaPara.Conta
        txbVencimento.Text = oAgendamento.Vencimento
        'If oAgendamento.Grupo = "R" Then cbbGrupo.ListIndex = 0 Else cbbGrupo.ListIndex = 1: Call cbbGrupo_AfterUpdate
        'cbbCategoria.Text = oSubcategoria.Categoria: Call cbbCategoria_AfterUpdate
        'cbbSubcategoria.Text = oSubcategoria.Subcategoria
    
        With oAgendamento
        
            If .Valor >= 0 Then
                txbValor.Text = Format(.Valor, "#,##0.00")
            Else
                txbValor.Text = Format(.Valor * -1, "#,##0.00")
            End If
            
            If .Recorrente = False Then
                cbbRecorrencia.Text = "Uma única vez"
                lblPeriodicidade.Visible = False: cbbPeriodicidade.Visible = False
                chbTermina.Visible = False
                lblParcelas.Visible = False: txbParcelas.Visible = False
            ElseIf .Recorrente = True And .Infinito = False Then
                cbbRecorrencia.Text = "Recorrente"
        
                If .Periodicidade = "m" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Mensal"
                ElseIf .Periodicidade = "yyyy" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Anual"
                ElseIf .Periodicidade = "d" And .Intervalo = 15 Then
                    cbbPeriodicidade.Text = "Quinzenal"
                End If
        
                chbTermina.Value = True
                txbParcelas.Visible = True
                txbParcelas.Text = .Parcelas
            ElseIf .Recorrente = True And .Infinito = True Then
                cbbRecorrencia = "Recorrente"
                If .Periodicidade = "m" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Mensal"
                ElseIf .Periodicidade = "yyyy" And .Intervalo = 1 Then
                    cbbPeriodicidade.Text = "Anual"
                ElseIf .Periodicidade = "d" And .Intervalo = 15 Then
                    cbbPeriodicidade.Text = "Quinzenal"
                End If
        
                chbTermina.Value = False
                txbParcelas.Visible = False
            End If
        End With
        
        txbObservacao.Text = oAgendamento.Observacao
        
    
    End If
    
End Sub
Private Sub cbbCategoriaPopular()
    
    Dim col As Collection
    Dim n As Variant
    
    Set col = oCategoria.Listar("categoria", cbbGrupo.List(cbbGrupo.ListIndex, 1))
    
    cbbCategoria.Clear
    
    For Each n In col
    
        oCategoria.Carrega CLng(n)
        
        With cbbCategoria
            .AddItem
            .List(.ListCount - 1, 0) = oCategoria.Categoria
            .List(.ListCount - 1, 1) = oCategoria.ID
        End With
    
    Next n
    
End Sub
Private Sub cbbSubcategoriaPopular()
    
    Dim col As Collection
    Dim n As Variant
    
    Set col = oSubcategoria.Listar("subcategoria", CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1)))
    
    cbbSubcategoria.Clear
    
    For Each n In col
    
        oSubcategoria.Carrega CLng(n)
    
        With cbbSubcategoria
            .AddItem
            .List(.ListCount - 1, 0) = oSubcategoria.Subcategoria
            .List(.ListCount - 1, 1) = oSubcategoria.ID
        End With
    
    Next n
End Sub

Private Sub btnRegistrar_Click()

    Dim i As Integer
'    Dim arr() As Long
    
    Set oMovimentacao = New cMovimentacao

    'Diz para formulário de registro de agendamento que o registro é oriundo de agendamento
    oMovimentacao.IsAgendamento = True
    
    
    ' Se a lista estiver com modo simples de seleção...
    If optSimples.Value = True Then
        
        oMovimentacao.IsTransferencia = chbTransferencia.Value
    
        ' ...e se houver um agendamento selecionado...
        If lblAgendamento.Caption <> Empty Then
        
            ' ...laço para encontrar linha selecionada na lista
            For i = 0 To lstPrincipal.ListCount - 1
            
                ' Quando a linha estiver selecionada
                If lstPrincipal.Selected(i) = True Then
                    
                    ' Passa o ID do agendamento para o objeto
                    oMovimentacao.AgendamentoID = (CLng(lstPrincipal.List(i, 0)))
                    fRegistrar.Show
                    Call lstPrincipalPopular("vencimento")
                    
                End If
            Next i
            
            Call lstPrincipalPopular("vencimento")
            
            lblTotal = Format(0, "#,##0.00")
            
            btnConfirmar.Visible = False
            btnCancelar.Visible = False
            
            Call BotoesDecisaoExibir
            Call Campos("Desabilitar")
            Call Campos("Limpar")
            
            btnAlterar.Enabled = False
            btnExcluir.Enabled = False
            btnRegistrar.Enabled = False
            btnIncluir.SetFocus
        Else
            MsgBox "É necessário selecionar no mínimo 1 agendamento para efetuar o registro!"
        End If
        
    ' Se vários agendamentos estiverem selecionados
    ElseIf optMultiplo.Value = True Then
        
        Dim n As Variant
        
        Set colAgendamentos = New Collection
        
        For i = 0 To lstPrincipal.ListCount - 1
            If lstPrincipal.Selected(i) = True Then
                colAgendamentos.Add lstPrincipal.List(i, 0)
            End If
        Next i
        
        For Each n In colAgendamentos
            
            oMovimentacao.AgendamentoID = CLng(n)
            
            fRegistrar.Show
        Next n
        
        Call lstPrincipalPopular("vencimento")
        
        lblTotal = Format(0, "#,##0.00")
        
        btnConfirmar.Visible = False
        btnCancelar.Visible = False
        
        Call BotoesDecisaoExibir
        Call Campos("Desabilitar")
        
        btnAlterar.Enabled = False
        btnExcluir.Enabled = False
        btnRegistrar.Enabled = False
        btnIncluir.SetFocus
        
    End If
    
End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    ' Se for uma receita ou despesa, então ...
    If chbTransferencia.Value = False Then
        If cbbContaDe.Text = Empty Then
            MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigatório"
        ElseIf cbbFornecedor.Text = Empty Then
            MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigatório"
        ElseIf txbValor.Text = Empty Or txbValor.Text = 0 Then
            MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
        ElseIf txbVencimento.Text = Empty Or Not IsDate(CDate(txbVencimento.Text)) Then
            MsgBox "Verifique o 'Vencimento'", vbCritical, "Campo obrigatório"
        ElseIf cbbCategoria.Text = Empty Then
            MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigatório"
        ElseIf cbbSubcategoria.Text = Empty Then
            MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigatório"
        ElseIf cbbRecorrencia.Text = Empty Then
            MsgBox "Informe o Tipo de 'Recorrência'", vbCritical, "Campo obrigatório"
        Else
            If cbbRecorrencia.Text = "Recorrente" And cbbPeriodicidade.Text = Empty Then
                 MsgBox "Informe a 'Periodicidade'", vbCritical, "Campo obrigatório"
            ElseIf chbTermina.Value = True Then
                If txbParcelas.Text = Empty Then
                    MsgBox "Informe o nº de 'Parcelas'.", vbCritical, "Campo obrigatório"
                ElseIf txbParcelas.Text = 0 Then
                    MsgBox "O nº de 'Parcelas' não pode ser zero.", vbCritical, "Campo obrigatório"
                ElseIf Not IsNumeric(txbParcelas.Text) Then
                    MsgBox "O nº de 'Parcelas' deve ser um número inteiro.", vbCritical, "Campo obrigatório"
                Else
                    With oAgendamento
                        .Recorrente = cbbRecorrencia.List(cbbRecorrencia.ListIndex, 1)
                        
                        If cbbPeriodicidade.ListIndex > -1 Then
                            .Periodicidade = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 1)
                            .Intervalo = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 2)
                        End If
                        
                        .Infinito = IIf(chbTermina.Value = True, False, True)
                        If chbTermina.Value = True Then .Parcelas = CInt(txbParcelas.Text) Else .Parcelas = 0
                        .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                        .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                        .Valor = CDbl(txbValor.Text)
                        .Vencimento = CDate(txbVencimento.Text)
                        .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                        .CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
                        .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                        .Observacao = txbObservacao.Text
                    End With
                
                    Valida = True
                End If
            Else
                With oAgendamento
                    .Recorrente = cbbRecorrencia.List(cbbRecorrencia.ListIndex, 1)
                    
                    If cbbPeriodicidade.ListIndex > -1 Then
                        .Periodicidade = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 1)
                        .Intervalo = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 2)
                    End If
                    
                    .Infinito = IIf(chbTermina.Value = True, False, True)
                    If chbTermina.Value = True Then .Parcelas = CInt(txbParcelas.Text) Else .Parcelas = 0
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Vencimento = CDate(txbVencimento.Text)
                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                    .CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
                    .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                    .Observacao = txbObservacao.Text
                End With
            
                Valida = True
    
            End If
        
        End If
    ' Se for uma transferência entre contas, então ...
    Else
        If cbbContaDe.Text = Empty Then
            MsgBox "Informe a 'Conta origem'", vbCritical, "Campo obrigatório"
            Exit Function
        ElseIf cbbContaPara.Text = Empty Then
            MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigatório"
            Exit Function
        ElseIf txbVencimento.Text = Empty Then
            MsgBox "Informe a 'Emissão'", vbCritical, "Campo obrigatório"
            Exit Function
        ElseIf txbValor.Text = Empty Or txbValor.Text = 0 Then
            MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
            Exit Function
        ElseIf cbbContaDe.Text = cbbContaPara.Text Then
            MsgBox "Conta Origem e Conta Destino não podem ser iguais!", vbCritical, "Campo obrigatório"
        Else
        
            If cbbRecorrencia.Text = "Recorrente" And cbbPeriodicidade.Text = Empty Then
                 MsgBox "Informe a 'Periodicidade'", vbCritical, "Campo obrigatório"
            ElseIf chbTermina.Value = True Then
                If txbParcelas.Text = Empty Then
                    MsgBox "Informe o nº de 'Parcelas'.", vbCritical, "Campo obrigatório"
                ElseIf txbParcelas.Text = 0 Then
                    MsgBox "O nº de 'Parcelas' não pode ser zero.", vbCritical, "Campo obrigatório"
                ElseIf Not IsNumeric(txbParcelas.Text) Then
                    MsgBox "O nº de 'Parcelas' deve ser um número inteiro.", vbCritical, "Campo obrigatório"
                Else
                    With oAgendamento
                        .Recorrente = cbbRecorrencia.List(cbbRecorrencia.ListIndex, 1)
                        
                        If cbbPeriodicidade.ListIndex > -1 Then
                            .Periodicidade = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 1)
                            .Intervalo = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 2)
                        End If
                        
                        .Infinito = IIf(chbTermina.Value = True, False, True)
                        If chbTermina.Value = True Then .Parcelas = CInt(txbParcelas.Text) Else .Parcelas = 0
                        .Grupo = "T"
                        .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                        .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                        '.FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                        .Valor = CDbl(txbValor.Text)
                        .Vencimento = CDate(txbVencimento.Text)
                        '.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                        '.CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
                        '.SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                        .Observacao = txbObservacao.Text
                    End With
                
                    Valida = True
                End If
            Else
                With oAgendamento
                    .Recorrente = cbbRecorrencia.List(cbbRecorrencia.ListIndex, 1)
                    
                    If cbbPeriodicidade.ListIndex > -1 Then
                        .Periodicidade = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 1)
                        .Intervalo = cbbPeriodicidade.List(cbbPeriodicidade.ListIndex, 2)
                    End If
                    
                    .Infinito = IIf(chbTermina.Value = True, False, True)
                    
                    If chbTermina.Value = True Then
                        .Parcelas = CInt(txbParcelas.Text)
                    Else
                        .Parcelas = 0
                    End If
                    .Grupo = "T"
                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
                    .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
                    '.FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .Valor = CDbl(txbValor.Text)
                    .Vencimento = CDate(txbVencimento.Text)
                    '.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
                    '.CategoriaID = cbbCategoria.List(cbbCategoria.ListIndex, 1)
                    '.SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
                    .Observacao = txbObservacao.Text
                End With
            
                Valida = True
    
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' APOIO
'            With oMovimentacao
'                .Grupo = "T"
'                .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
'                .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
'                .Valor = CDbl(txbValor.Text)
'                .Liquidado = CDate(txbVencimento.Text)
'                .Observacao = txbObservacao.Text
'            End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FIM DO APOIO
        End If
    End If
End Function

Private Sub BotoesDecisaoExibir()
    btnIncluir.Visible = True
    btnIncluir.Enabled = True
    btnAlterar.Visible = True
    btnExcluir.Visible = True
    btnRegistrar.Visible = True
End Sub
Private Sub BotoesDecisaoEsconder()
    MultiPage1.Value = 0
    btnIncluir.Visible = False
    btnAlterar.Visible = False
    btnExcluir.Visible = False
    btnRegistrar.Visible = False
End Sub
Private Sub ListBoxTiraSelecoes()
    
    Dim i As Integer
    
    With lstPrincipal
        For i = 0 To .ListCount
            If .Selected(i) = True Then .Selected(i) = False
        Next i
        .Enabled = True
    End With
    
End Sub
Private Sub UserForm_Terminate()
    Set oAgendamento = Nothing
    Call Desconecta
End Sub

