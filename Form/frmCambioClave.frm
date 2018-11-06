VERSION 5.00
Begin VB.Form frmCambioClave 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Clave"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambioClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   128
      TabIndex        =   4
      Top             =   120
      Width           =   3720
      Begin VB.TextBox txtClave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txtClave1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Clave:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reconfirme Clave:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2032
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   487
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Public INICIO As Boolean


Private Sub cmbAceptar_Click()
    Dim strSQL As String
    If txtClave1.Text = txtClave2.Text Then
        If txtClave1.Text <> strClave Then
        
            strSQL = "EXEC sp_password '" & strClave & "','" & txtClave1.Text & "', '" & strUsuario & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            
            strSQL = " UPDATE usuario SET " & _
                     " usu_ultimamod=CURRENT_TIMESTAMP " & _
                     " WHERE usu_codigo='" & strUsuario & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            
            MsgBox "La clave fue cambiada." & vbNewLine & "Ingrese al sistema con la nueva clave", vbInformation, "Cambio de Clave"
            ''Unload mdiPrincipal
            Dim frmX As Form
            
            For Each frmX In Forms
               Unload frmX
               Set frmX = Nothing
            Next
            End

        Else
            MsgBox "Ingrese una clave diferente a la anterior", vbCritical, "Cambio de Clave"
        End If
    Else
        MsgBox "Tiene algun error al ingresar la clave, vueva a intentarlo", vbCritical, "Cambio de Clave"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    INICIO = False
End Sub

Private Sub cmdcancelar_Click()
    If INICIO = False Then
        Unload Me
    Else
       Dim frmX As Form
        For Each frmX In Forms
           Unload frmX
           Set frmX = Nothing
        Next
        End
        'Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
