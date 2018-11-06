VERSION 5.00
Begin VB.Form frmFormula 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmula"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3990
   Begin VB.CommandButton cmdBorrar 
      Cancel          =   -1  'True
      Caption         =   "&Borrar"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1448
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Fórmula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdOperador 
         Caption         =   "/"
         Height          =   480
         Index           =   3
         Left            =   3000
         TabIndex        =   15
         Top             =   1920
         Width           =   480
      End
      Begin VB.CommandButton cmdOperador 
         Caption         =   "X"
         Height          =   480
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   1440
         Width           =   480
      End
      Begin VB.CommandButton cmdOperador 
         Caption         =   "-"
         Height          =   480
         Index           =   1
         Left            =   3000
         TabIndex        =   13
         Top             =   960
         Width           =   480
      End
      Begin VB.CommandButton cmdOperador 
         Caption         =   "+"
         Height          =   480
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   480
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "&Añadir"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   1290
         Width           =   735
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "&Añadir"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   9
         Top             =   578
         Width           =   735
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "1"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtFormula 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "1"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.ComboBox cmbConstante 
         Height          =   330
         ItemData        =   "frmFormula.frx":030A
         Left            =   240
         List            =   "frmFormula.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblCantidad 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fórmula:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblConstantes 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Constantes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2648
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   248
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Objeto As Object

Private Sub cmbAceptar_Click()
    Objeto.Text = Me.txtFormula
    Unload Me
End Sub

Private Sub cmdAñadir_Click(Index As Integer)
    Dim Constante1 As String
    If Index = 0 Then
        Select Case Me.cmbConstante.ListIndex
        Case 0
            Constante1 = "SueldoBas"
        Case 1
            Constante1 = "SueldoMes"
        Case 2
            Constante1 = "SueldoIESS"
        Case 3
            Constante1 = "ImpRentaMes"
        End Select
        txtFormula = txtFormula & Constante1
    Else
        If Trim(txtCantidad) = "" Then Exit Sub
        txtFormula = txtFormula & txtCantidad
    End If
    
    Alternar (False)
End Sub

Private Sub cmdBorrar_Click()
    txtFormula = ""
    Alternar (True)
    Me.cmbAceptar.Enabled = True
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdOperador_Click(Index As Integer)
    Select Case Index
    Case 0
        txtFormula = txtFormula & "+"
    Case 1
        txtFormula = txtFormula & "-"
    Case 2
        txtFormula = txtFormula & "*"
    Case 3
        txtFormula = txtFormula & "/"
    End Select
    Alternar (True)
    
End Sub

Public Sub Alternar(Prendido As Boolean)
    If Prendido = True Then
        Me.lblCantidad.Enabled = True
        Me.lblConstantes.Enabled = True
        Me.txtCantidad.Enabled = True
        Me.cmbConstante.Enabled = True
        Me.cmdAñadir(0).Enabled = True
        Me.cmdAñadir(1).Enabled = True
        
        Me.cmdOperador(0).Enabled = False
        Me.cmdOperador(1).Enabled = False
        Me.cmdOperador(2).Enabled = False
        Me.cmdOperador(3).Enabled = False
        
        Me.cmbAceptar.Enabled = False
    Else
        Me.lblCantidad.Enabled = False
        Me.lblConstantes.Enabled = False
        Me.txtCantidad.Enabled = False
        Me.cmbConstante.Enabled = False
        Me.cmdAñadir(0).Enabled = False
        Me.cmdAñadir(1).Enabled = False
        
        Me.cmdOperador(0).Enabled = True
        Me.cmdOperador(1).Enabled = True
        Me.cmdOperador(2).Enabled = True
        Me.cmdOperador(3).Enabled = True
        
        Me.cmbAceptar.Enabled = True
    End If
End Sub
Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)
    Me.cmbConstante.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtCantidad_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtCantidad_Validate(Cancel As Boolean)
    txtCantidad = Val(txtCantidad)
End Sub

Private Sub txtFormula_Change()
    If Trim(txtFormula) = "" Then
        Me.cmdBorrar.Enabled = False
    Else
        Me.cmdBorrar.Enabled = True
    End If
End Sub
