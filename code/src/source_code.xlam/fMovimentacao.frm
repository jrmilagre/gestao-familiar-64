VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMovimentacao 
   Caption         =   ":: Registrar ::"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "fMovimentacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta          As New cConta

Private Sub UserForm_Initialize()

    

End Sub

Private Sub UserForm_Terminate()
    Set oConta = Nothing
    Call Desconecta
End Sub
