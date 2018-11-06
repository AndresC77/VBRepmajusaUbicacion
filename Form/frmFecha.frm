VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFecha 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFecha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   1605
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21561345
      CurrentDate     =   38589
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Objeto As Object

Private Sub cmbAceptar_Click()
    Objeto.Tag = Fecha
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - mdiPrincipal.Height / 40
End Sub
