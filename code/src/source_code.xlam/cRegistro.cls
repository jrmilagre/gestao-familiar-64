VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TMovimentacao
    ID              As Long
    NomeFantasia    As String
    RazaoSocial     As String
    Endereco        As String
    Numero          As String
    Bairro          As String
    Cidade          As String
    Estado          As String
    Pais            As String
    DataCadastro    As Date
End Type

Private This As TMovimentacao
Private Const tbl As String = "tbl_movimentacoes"
