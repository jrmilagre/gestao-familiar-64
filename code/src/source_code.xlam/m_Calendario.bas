Attribute VB_Name = "m_Calendario"
Option Private Module
Option Explicit

'---Para obter uma data do calend�rio na Plan1 e na c�lula A1, basta escrever o seguinte c�digo:
'--- Plan1.Range("A1") = GetCalendario

Public Const sMascaraData   As String = "DD/MM/YYYY"   '---formata��o de datas
Public dtDate               As Date
Dim Rotulos()               As New c_Calendario  '---vetor que armazena todos os Label de dia do Calend�rio

Function GetCalendario() As Date
    ' Fun��o GetCalendario
    
    ' Declara vari�veis
    Dim lTotalRotulos As Long   ' Total de r�tulos
    Dim Ctrl As control
    Dim frm As f_Calendario      ' Formul�rio
    
    Set frm = New f_Calendario ' Cria novo objeto setando formul�rio nele
    
    ' Atribui cada um dos Label num elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "l?c?" Then
            lTotalRotulos = lTotalRotulos + 1
            ReDim Preserve Rotulos(1 To lTotalRotulos)
            Set Rotulos(lTotalRotulos).lblGrupo = Ctrl
        End If
    Next Ctrl
    frm.Show
    
    ' Se a data escolhida for nula ou inv�lida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag
    Else
        GetCalendario = dtDate
    End If
    Unload frm
End Function
    

