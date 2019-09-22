VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ValidaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oValidaData    As MSForms.TextBox
Attribute m_oValidaData.VB_VarHelpID = -1
Private WithEvents m_oValidaDecimal As MSForms.TextBox
Attribute m_oValidaDecimal.VB_VarHelpID = -1


' VALIDAÇÃO DE CAMPO DE DATA
Public Property Set ValidaData(ByVal oValidaData As MSForms.TextBox)

    Set m_oValidaData = oValidaData

End Property

Private Sub m_oValidaData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 9 Or KeyCode = 13 Then
        If IsDate(m_oValidaData.Text) Then
            m_oValidaData.Text = Format(m_oValidaData.Text, "dd/mm/yyyy")
            Exit Sub
        Else
            m_oValidaData.Text = Date
        End If
    End If
    
End Sub
Private Sub m_oValidaData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    With m_oValidaData
        Select Case KeyAscii
            Case 8                      ' Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}"   ' Emula o TAB
            Case 48 To 57
                If .SelStart = 2 Then .SelText = "/"
                If .SelStart = 5 Then .SelText = "/"
            Case Else: KeyAscii = 0     ' Ignora os outros caracteres
        End Select
    End With
End Sub

' VALIDAÇÃO DE CAMPO DECIMAL
Public Property Set ValidaDecimal(ByVal oValidaDecimal As MSForms.TextBox)

    Set m_oValidaDecimal = oValidaDecimal

End Property


















Private txb As TextBox



















Public Function SomenteNumerosEVirgula(Caracter As MSForms.ReturnInteger, Texto As String) As Integer

    Select Case Caracter
        Case 8          ' Backspace (seta de apagar)
        Case 48 To 57   ' Números de 0 a 9
        Case 44         ' Vírgula
        If InStr(Texto, ",") Then 'Se o campo já tiver vírgula então ele não adiciona
            Caracter = 0 'Não adiciona a vírgula caso ja tenha
        Else
            Caracter = 44 'Adiciona uma vírgula
        End If
        Case Else
            Caracter = 0 'Não deixa nenhuma outra caractere ser escrito
            MsgBox "Só é permitido número.", vbInformation
    End Select
    
    SomenteNumerosEVirgula = Caracter

End Function
Public Function CampoObrigatorio(NomeCampo As String, Conteudo As String) As Variant
    
    If Conteudo = Empty Then
        MsgBox "O campo " & NomeCampo & " é obrigatório", vbCritical
        CampoObrigatorio = &HC0C0FF
    Else
        CampoObrigatorio = &H80000005
    End If
        
End Function
Public Function CampoValor(Texto As String) As String
    
    If IsNumeric(Texto) Then
        CampoValor = Format(Texto, "#,##0.00")
    Else
        CampoValor = Empty
    End If

End Function
Public Function CampoData(Texto As String) As String
    
    If IsDate(Texto) Then
        CampoData = Format(Texto, "dd/mm/yyyy")
    Else
        CampoData = Empty
    End If

End Function
Public Sub SelecionaConteudo(Campo As TextBox, Comprimento As Integer)

    Campo.SelStart = 0
    Campo.SelLength = Comprimento

End Sub




