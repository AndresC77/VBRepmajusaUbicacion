Attribute VB_Name = "modCentrarChild"
Public Sub Centrar(nomForm As String)
    'Centra esta forma dentro de la forma MDI
    With ActiveForm
        .Left = (mdiPrincipal.Width - .Width) / 2
        .Top = ((mdiPrincipal.Height - .Height) / 2) - (.Height / 8)
    End With
End Sub
