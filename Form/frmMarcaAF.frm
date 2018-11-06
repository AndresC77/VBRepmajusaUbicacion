VERSION 5.00
Begin VB.Form frmMarcaAF 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Marcas"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMarcaAF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3840
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1973
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   413
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Marcas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   53
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   1920
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   1920
      End
      Begin VB.Label LblNombre 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Marca:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label LblCodigo 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   390
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMarcaAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Marcas para los activos fijospara ubicar #
'#  frmMarcaAF V1.0                                                               #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de las marcas.                      #
'#  Permitirá almacenar en la base de datos nuevas marcas y modificar           #
'#  sus nombres y descripciones, dependiendo de la propiedad Tag,               #
'#  la cual se cambiará en la ventana frmSelMarca y desde esta se llamará       #
'#  a esta ventana.                                                             #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#        marca_activo_fijo: En esta tabla se almacenan las nuevas marcas y se  #
'#               modifican los datos de estas.                                  #
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
    ' Si se esta ingresando una nueva marca
    If Me.Tag = "N" Then
    ' Almacenamiento de los datos del nueva linea
        strSql = " SELECT mar_act_fij_codigo FROM marca_activo_fijo where mar_act_fij_codigo='" & UCase(txtCodigo.Text) & "' "
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount <= 0 Then
            strSql = " INSERT INTO marca_activo_fijo(mar_act_fij_codigo,mar_act_fij_nombre,mar_act_fij_fechamod,mar_act_fij_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & UCase(txtNombre.Text) & "', " & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
        Else
            MsgBox "La marca que ingresó, ya existe." & vbCrLf & "Por favor cambie el código", vbExclamation, "Error Marca"
            txtCodigo.SetFocus
            txtCodigo.SelStart = 0
            txtCodigo.SelLength = Len(txtCodigo)
            Exit Sub
        End If
    ' Si se esta modificando la marca
    ElseIf Me.Tag = "M" Then
    'Almacenamiento de los cambios realizados a la marca
        strSql = " UPDATE marca_activo_fijo " & _
                 " SET mar_act_fij_nombre='" & UCase(txtNombre.Text) & "'," & _
                 " mar_act_fij_fechamod=CURRENT_TIMESTAMP,mar_act_fij_usumod='" & strUsuario & "' " & _
                 " WHERE mar_act_fij_codigo='" & txtCodigo.Text & "'  "
        frmSelMarcaAF.dcmbCodigo = txtCodigo.Text
        frmSelMarcaAF.dcmbNombre.Text = txtNombre.Text
        
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

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar datos de la Marca"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nueva Marca"
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
        SendKeys "{TAB}"
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

Private Sub txtDescripcion_GotFocus()
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
