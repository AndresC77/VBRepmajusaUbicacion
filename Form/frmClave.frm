VERSION 5.00
Begin VB.Form frmClave 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de Supervisor"
   ClientHeight    =   1530
   ClientLeft      =   6435
   ClientTop       =   5295
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtClave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la Clave:"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ret As Boolean
Public dblPrecio As String
Public strClaveMAESTRA As String
Private Sub cmdAceptar_Click()
    If txtClave.Text = strClaveMAESTRA Then
        If IsNumeric(dblPrecio) = True Then
            MsgBox "Permiso Consedido" & vbNewLine & vbNewLine & "Se facturará a: " & dblPrecio, vbInformation, "Seguridad"
        ElseIf dblPrecio = "Bodega" Then
            MsgBox "Permiso Consedido" & vbNewLine & vbNewLine & "Para el Cambio de Bodega", vbInformation, "Seguridad"
        ElseIf dblPrecio = "Fecha" Then
            MsgBox "Permiso Consedido" & vbNewLine & vbNewLine & "Para el Cambio de Fecha", vbInformation, "Seguridad"
        ElseIf dblPrecio = "Precio" Then
            MsgBox "Permiso Consedido" & vbNewLine & vbNewLine & "Para el Cambio de Precio", vbInformation, "Seguridad"
        ElseIf dblPrecio = "Anulacion" Then
            MsgBox "Permiso Consedido" & vbNewLine & vbNewLine & "Para la Anulacion", vbInformation, "Seguridad"
        End If
        Ret = True
        Unload Me
    Else
        MsgBox "Clave mal Ingresada", vbInformation, "Seguridad"
    End If
End Sub

Private Sub cmdcancelar_Click()
    Ret = False
    Unload Me
End Sub

Private Sub Form_Load()
'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    Ret = False
End Sub
