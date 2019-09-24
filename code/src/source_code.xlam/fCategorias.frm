VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCategorias 
   Caption         =   ":: Cadastro de Categorias ::"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885
   OleObjectBlob   =   "fCategorias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oCategoria    As New cCategoria
Private oSubcategoria As New cSubcategoria

Private sDecisaoCat As String
Private sDecisaoSub As String



Private Sub UserForm_Initialize()

    Call cbbGrupoPopular
    
    btnCatInclui.Enabled = False
    btnCatAltera.Enabled = False
    btnCatExclui.Enabled = False
    btnCatConfirma.Visible = False
    btnCatCancela.Visible = False
    
    btnSubInclui.Enabled = False
    btnSubAltera.Enabled = False
    btnSubExclui.Enabled = False
    btnSubConfirma.Visible = False
    btnSubCancela.Visible = False
    
    txbCategoria.Enabled = False: lblCategoria.Enabled = False
    txbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False

End Sub
Private Sub btnCatExclui_Click()

    sDecisaoCat = "Exclusão"
    txbCategoria.Enabled = True: lblCategoria.Enabled = True
    txbCategoria.SetFocus
    Call BotoesDecisaoCategoriaEsconde
    
End Sub
Private Sub cbbGrupoPopular()
    With cbbGrupo
        .AddItem
        .List(.ListCount - 1, 0) = "Receitas"
        .List(.ListCount - 1, 1) = "R"
        .AddItem
        .List(.ListCount - 1, 0) = "Despesas"
        .List(.ListCount - 1, 1) = "D"
    End With
End Sub
Private Sub txbSubcategoria_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSubConfirma.SetFocus
End Sub
Private Sub lstCategorias_Change()

    If lstCategorias.ListIndex > -1 Then
    
        txbCategoria.Text = lstCategorias.List(lstCategorias.ListIndex, 0)
        
        Call lstSubcategoriasCarregar
        
        btnCatAltera.Enabled = True
        btnCatExclui.Enabled = True
        btnSubInclui.Enabled = True
        
        txbSubcategoria.Text = ""
    Else
        lstSubcategorias.Clear
    End If

End Sub
Private Sub lstSubcategoriasCarregar()

    Dim n As Variant
    Dim col As New Collection
    
    Set col = oSubcategoria.PreencheListBox("subcategoria", CLng(lstCategorias.List(lstCategorias.ListIndex, 1)))
    
    With lstSubcategorias
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "120 pt; 0pt;"
        .Font = "Consolas"
        
        For Each n In col
            .AddItem
            oSubcategoria.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oSubcategoria.Subcategoria
            .List(.ListCount - 1, 1) = oSubcategoria.ID
        Next n
        
    End With
End Sub

Private Sub btnCatInclui_Click()

    Call btnCatCancela_Click
    
    sDecisaoCat = "Inclusão"
    
    txbCategoria.Text = ""
    txbCategoria.Enabled = True: lblCategoria.Enabled = True
    txbCategoria.SetFocus
    
    Call BotoesDecisaoCategoriaEsconde
    
End Sub
Private Sub btnCatAltera_Click()
    sDecisaoCat = "Alteração"
    txbCategoria.Enabled = True: lblCategoria.Enabled = True
    txbCategoria.SetFocus
    Call BotoesDecisaoCategoriaEsconde
End Sub
Private Sub btnCatConfirma_Click()

    If txbCategoria.Text <> Empty Then
    
        oCategoria.Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
        oCategoria.Categoria = txbCategoria.Text
        
        If sDecisaoCat = "Inclusão" Then
            
            oCategoria.Inclui
            
        ElseIf sDecisaoCat = "Alteração" Then
            
            oCategoria.Altera CLng(lstCategorias.List(lstCategorias.ListIndex, 1))
            
        ElseIf sDecisaoCat = "Exclusão" Then
        
            oCategoria.Exclui CLng(lstCategorias.List(lstCategorias.ListIndex, 1))
            
        End If
        
        lstCategoriasCarregar
        lstSubcategorias.Clear
        
        ' Caso tenha algum item selecionado na ListBox, tira a seleção
        lstCategorias.ListIndex = -1
        
        Call BotoesDecisaoCategoriaExibe
        txbCategoria.Text = ""
        txbCategoria.Enabled = False: lblCategoria.Enabled = False
    Else
        MsgBox "Campo 'Categoria' é obrigatório.", vbInformation: txbCategoria.SetFocus
    End If
End Sub
Private Sub btnCatCancela_Click()
    
    Call BotoesDecisaoCategoriaExibe
    
    txbCategoria.Text = ""
    txbCategoria.Enabled = False: lblCategoria.Enabled = False
    btnCatAltera.Enabled = False
    btnCatExclui.Enabled = False
    btnCatInclui.SetFocus
    
    With lstCategorias
        .Enabled = True ' Desabilita
        If .ListIndex >= 0 Then .Selected(.ListIndex) = False ' Tira a seleção
    End With
    
    
End Sub
Private Sub btnSubInclui_Click()
    sDecisaoSub = "Inclusão"
    txbSubcategoria.Enabled = True: lblSubcategoria.Enabled = True
    txbSubcategoria.SetFocus
    txbSubcategoria.Text = ""
    Call BotoesDecisaoSubcategoriaEsconde
    lstSubcategorias.ListIndex = -1
End Sub
Private Sub btnSubAltera_Click()
    
    sDecisaoSub = "Alteração"
    
    With txbSubcategoria
        .Enabled = True: lblSubcategoria.Enabled = True
        .SetFocus
    End With
    
    Call BotoesDecisaoSubcategoriaEsconde
    
End Sub
Private Sub btnSubExclui_Click()
    
    sDecisaoSub = "Exclusão"
    
    txbSubcategoria.Enabled = True: lblSubcategoria.Enabled = True
    
    Call BotoesDecisaoSubcategoriaEsconde
    
    btnSubConfirma.SetFocus
    
End Sub
Private Sub btnSubConfirma_Click()

    If txbSubcategoria.Text <> Empty Then
    
        oSubcategoria.Subcategoria = txbSubcategoria.Text
        oSubcategoria.CategoriaID = CLng(lstCategorias.List(lstCategorias.ListIndex, 1))
        
        If sDecisaoSub = "Inclusão" Then
            
            oSubcategoria.Inclui
            
        ElseIf sDecisaoSub = "Alteração" Then
            
            oSubcategoria.Altera CLng(lstSubcategorias.List(lstSubcategorias.ListIndex, 1))
            
        ElseIf sDecisaoSub = "Exclusão" Then
            
            oSubcategoria.Exclui CLng(lstSubcategorias.List(lstSubcategorias.ListIndex, 1))
            
        End If
        
        Call lstSubcategoriasCarregar
        
        ' Caso tenha algum item selecionado na ListBox, tira a seleção
        lstSubcategorias.ListIndex = -1
        Call BotoesDecisaoSubcategoriaExibe
        txbSubcategoria.Text = ""
        
        btnSubAltera.Enabled = False
        btnSubExclui.Enabled = False
        txbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False
        
    End If
End Sub
Private Sub btnSubCancela_Click()
    Call BotoesDecisaoSubcategoriaExibe
    
    txbSubcategoria.Text = ""
    txbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False
    btnSubAltera.Enabled = False
    btnSubExclui.Enabled = False
    btnSubInclui.SetFocus
    
    With lstSubcategorias
        .Enabled = True
        .ListIndex = -1
    End With
End Sub
Private Sub cbbGrupo_Change() ' Subrotina executada ao atualizar ComboBox de tipo de categoria
    Call lstCategoriasCarregar
    btnCatInclui.Enabled = True
    btnSubInclui.Enabled = False
    txbCategoria.Text = ""
    txbSubcategoria.Text = ""
End Sub
Private Sub lstCategoriasCarregar()
    
    Dim col As New Collection

    Set col = oCategoria.PreencheListBox("categoria", cbbGrupo.List(cbbGrupo.ListIndex, 1))
    
    With lstCategorias
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 2                    ' Determina número de colunas
        .ColumnWidths = "120 pt; 0pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oCategoria.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oCategoria.Categoria
            .List(.ListCount - 1, 1) = oCategoria.ID
        Next n
        
    End With
    
End Sub
Private Sub BotoesDecisaoCategoriaEsconde()
    btnCatInclui.Visible = False
    btnCatAltera.Visible = False
    btnCatExclui.Visible = False
    btnSubInclui.Enabled = False
    btnCatConfirma.Visible = True
    btnCatCancela.Visible = True
End Sub
Private Sub BotoesDecisaoCategoriaExibe()
    btnCatInclui.Visible = True
    btnCatAltera.Visible = True
    btnCatExclui.Visible = True
    btnCatConfirma.Visible = False
    btnCatCancela.Visible = False
    btnCatAltera.Enabled = False
    btnCatExclui.Enabled = False
End Sub
Private Sub BotoesDecisaoSubcategoriaEsconde()
    btnSubInclui.Visible = False
    btnSubAltera.Visible = False
    btnSubExclui.Visible = False
    btnSubConfirma.Visible = True
    btnSubCancela.Visible = True
    btnCatInclui.Enabled = False
    btnCatAltera.Enabled = False
    btnCatExclui.Enabled = False
End Sub
Private Sub BotoesDecisaoSubcategoriaExibe()
    btnSubInclui.Visible = True
    btnSubAltera.Visible = True
    btnSubExclui.Visible = True
    btnSubConfirma.Visible = False
    btnSubCancela.Visible = False
    btnSubAltera.Enabled = False
    btnSubExclui.Enabled = False
    btnCatInclui.Enabled = True
End Sub

Private Sub lstSubcategorias_Change()

    If lstSubcategorias.ListIndex > -1 Then
        txbSubcategoria.Text = oSubcategoria.Subcategoria
    End If
    
    btnSubInclui.Enabled = True
    btnSubAltera.Enabled = True
    btnSubExclui.Enabled = True
    
End Sub
Private Sub UserForm_Terminate()
    Set oCategoria = Nothing
    Set oSubcategoria = Nothing
    Call Desconecta
End Sub
