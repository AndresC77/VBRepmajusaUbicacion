VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTipoSum 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Tipos para Suministros"
   ClientHeight    =   2280
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
   Icon            =   "frmTipoSum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3975
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo Suministro"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   1920
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCuentas 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Contable:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label LblNombre 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Tipo:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label LblCodigo 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmTipoSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Tipos de Sumnistros para ubicar a los    #
'#  clientes y proveedores                                                      #
'#  frmTipoSum V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de los tipos de sumnistros          #
'#  Permitirá almacenar en la base de datos nuevas unidades y modificar         #
'#  sus nombres, dependiendo de la propiedad Tag, la cual se cambiará en la     #
'#  ventana frmSeltiposum y desde esta se llamará a esta ventana.               #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#       tipo_suministro: En esta tabla se almacenan los nuevos tipos de        #
'#               sumnistros modifican los datos de estas.                       #
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
        strSql = " SELECT tip_sum_codigo " & _
                 " FROM tipo_suministro " & _
                 " WHERE tip_sum_codigo='" & UCase(txtCodigo.Text) & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount <= 0 Then
            strSql = " INSERT INTO tipo_suministro(tip_sum_codigo,emp_codigo,tip_sum_nombre,tip_sum_ctaconta,tip_sum_fechamod,tip_sum_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & strEmpresa & "','" & UCase(txtNombre.Text) & "', " & _
                 " '" & dcmbCuentas.Text & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
        Else
            MsgBox "La tipo que ingresó, ya existe." & vbCrLf & "Por favor cambie el código", vbExclamation, "Error Tipo"
            txtCodigo.SetFocus
            txtCodigo.SelStart = 0
            txtCodigo.SelLength = Len(txtCodigo)
            Exit Sub
        End If
    ' Si se esta modificando la marca
    ElseIf Me.Tag = "M" Then
    'Almacenamiento de los cambios realizados a la marca
        strSql = " UPDATE tipo_suministro " & _
                 " SET tip_sum_nombre='" & UCase(txtNombre.Text) & "' " & _
                 ",tip_sum_ctaconta='" & dcmbCuentas.Text & "',tip_sum_fechamod=CURRENT_TIMESTAMP,tip_sum_usumod='" & strUsuario & "' " & _
                 " WHERE tip_sum_codigo='" & txtCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        frmSelTipoSum.dcmbCodigo = txtCodigo.Text
        frmSelTipoSum.dcmbNombre.Text = UCase(txtNombre.Text)
        
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

Private Sub dcmbCuentas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "UN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCuentas
    End If
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strSql = " Select cta_codigo From ctaconta " & _
             " Where emp_codigo = '" & strEmpresa & "' And cta_subcta = 0 " & _
             " Order By cta_codigo "
    'Ejecuta la consulta anterior
    clsCon_Def.Ejecutar (strSql)
    'Muestra los datos de los códigos de las cuentas en un datacombo
    Set dcmbCuentas.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCuentas.ListField = "cta_codigo"
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar datos del Tipo"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de un Nuevo Tipo"
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
