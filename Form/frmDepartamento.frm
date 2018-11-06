VERSION 5.00
Begin VB.Form frmDepartamento 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Departamento"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDepartamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4470
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Departamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   128
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   2085
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1005
         Width           =   1920
      End
      Begin VB.TextBox txtNombreP 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2085
         TabIndex        =   1
         Top             =   585
         Width           =   1920
      End
      Begin VB.TextBox txtCodigoP 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2085
         TabIndex        =   0
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2085
         TabIndex        =   3
         Top             =   1365
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   855
         Left            =   2085
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1725
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Departamento:"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Area:"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   540
      End
      Begin VB.Label LabDescripcion 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1725
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2288
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   728
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
End
Attribute VB_Name = "frmDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Departamentos                            #
'#  frmDepartamento V1.0                                                        #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de Departamento.                    #
'#  Permitirá almacenar en la base de datos nuevos Departamentos y modificar    #
'#  sus nombres, dependiendo de la propiedad Tag, la cual se cambiará en la     #
'#  ventana frmSelDepartamento y desde esta se llamará a esta ventana.          #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#       Departamento: En esta tabla se almacenan las nuevos Departamentos y se #
'#               modifican los datos de estos.                                  #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

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
    ' Si se esta ingresando una nuevo Departamento
    If Me.Tag = "N" Then
    ' Almacenamiento de los datos de la nuevo area
        strSql = " INSERT INTO departamento(dto_codigo,emp_codigo,are_codigo,dto_nombre,dto_descripcion,dto_fechamod,dto_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & strEmpresa & "','" & txtCodigoP.Text & "','" & UCase(txtNombre.Text) & "','" & UCase(txtDescripcion.Text) & "', " & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
    ' Si se esta modificando el Departamento
    ElseIf Me.Tag = "M" Then
    'Almacenamiento de los cambios realizados al Area
        strSql = " UPDATE departamento " & _
                 " SET dto_nombre='" & UCase(txtNombre.Text) & "', dto_descripcion='" & UCase(txtDescripcion.Text) & "',dto_fechamod=CURRENT_TIMESTAMP,dto_usumod='" & strUsuario & "' " & _
                 " WHERE dto_codigo='" & txtCodigo.Text & "' AND are_codigo='" & txtCodigoP.Text & "'AND emp_codigo='" & strEmpresa & "'"
    End If
    'On Error GoTo errhandler
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

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar Departamento"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nuevo Departamento"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
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
