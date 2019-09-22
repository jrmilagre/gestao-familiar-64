VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lblGrupo As MSForms.CommandButton
Attribute lblGrupo.VB_VarHelpID = -1

Private Sub lblGrupo_Click()
    lblGrupo.Parent.Tag = lblGrupo.Tag
    lblGrupo.Parent.Hide
End Sub
