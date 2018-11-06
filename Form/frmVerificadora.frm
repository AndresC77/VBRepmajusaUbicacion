VERSION 5.00
Begin VB.Form frmVerificadora 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de datos de Verificadora"
   ClientHeight    =   2310
   ClientLeft      =   5970
   ClientTop       =   4650
   ClientWidth     =   6495
   Icon            =   "frmVerificadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Verificadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   1920
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   4170
         TabIndex        =   3
         Top             =   360
         Width           =   1920
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   4170
         TabIndex        =   4
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   4170
         TabIndex        =   5
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lbldireccion 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblTelefono 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3240
         TabIndex        =   10
         Top             =   750
         Width           =   495
      End
      Begin VB.Label lblFax 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   1110
         Width           =   495
      End
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1740
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmVerificadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#####################################################################################'
'#  Forma para el ingreso y modificación de Verificadoras                            #
'#  frmVerificadora V1.0                                                             #
'#  Copyright (C) 2002                                                               #
'#                                                                                   #
'#  Ventana para el ingreso y modificación de Verificadoras.                         #
'#  Permitirá almacenar en la base de datos nuevas verificadoras y modificar sus     #
'#  nombres, dependiendo de la propiedad Tag, la cual se cambiará en la              #
'#  ventana frmSelVerificadora y desde esta se llamará a esta ventana.               #
'#                                                                                   #
'#  Tablas que se maneja:                                                            #
'#    verificadora: En esta tabla se almacenan las nuevos verificadoras y se         #
'#                  modifican los datos de estos.                                    #
'#                                                                                   #
'#  Procedimientos INTERNOS:                                                         #
'#  Procedimientos EXTERNOS:                                                         #
'#                                                                                   #
'#  Objetos de la forma:                                                             #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos               #
'#                                                                                   #
'#                                                                                   #
'#####################################################################################
'/**********************************************************************************/'

Private clsCon_Def As clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmbAceptar_Click()
  Dim strSql As String
    ' Si la variable Tag es N ingresa una nueva verificadora
    If Me.Tag = "N" Then
    'Consulta para ingresar los datos a la base de datos
        strSql = " INSERT INTO verificadora(ver_codigo,ver_nombre,ver_direccion,ver_telf,ver_fax,ver_email,ver_fechamod,ver_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & UCase(txtNombre.Text) & "','" & UCase(txtDireccion.Text) & "','" & txtTelefono.Text & "','" & txtFax.Text & "','" & txtEmail.Text & "'," & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
        
    ' Si la variable Tag es M se modifican los datos de la verificadora
    ElseIf Me.Tag = "M" Then
    'Consulta para modificar los datos de la verificadora seleccionada
        strSql = " UPDATE verificadora" & _
                 " SET ver_nombre='" & UCase(txtNombre.Text) & "',ver_direccion='" & UCase(txtDireccion.Text) & "',ver_telf='" & txtTelefono.Text & "',ver_email='" & txtEmail.Text & "',ver_fax='" & txtFax.Text & "',ver_fechamod=CURRENT_TIMESTAMP,ver_usumod='" & strUsuario & "' " & _
                 " WHERE ver_codigo='" & txtCodigo.Text & "'"
    End If
    On Error GoTo errhandler
        clsCon_Def.Ejecutar (strSql), "M"
        Unload Me
        Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar Datos de Verificadora"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de nueva Verificadora"
    End If
    
    Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub



Private Sub txtCodigo_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub TxtDireccion_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtEmail_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtNombre_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtFax_GotFocus()
    Seleccionar_Contenido
End Sub
