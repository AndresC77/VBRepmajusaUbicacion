VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpresa 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Nuevas Empresas"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6480
   Begin VB.Frame fraSel 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Copiar datos segun empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   6255
      Begin VB.CheckBox chkFormaPago 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Formas de Pago"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkDefMov 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Definiciones de Movimientos Inventario"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CheckBox chkParametro 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Parametros Generales"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkRetencion 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Retenciones"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkPlanCuenta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Plan de Cuenta"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcmbEmpresa 
         Height          =   330
         Left            =   1080
         TabIndex        =   20
         Top             =   360
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   582
         _Version        =   393216
         IntegralHeight  =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblEmpresas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Nueva Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   113
      TabIndex        =   10
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   2820
      End
      Begin VB.TextBox txtRuc 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   750
         Width           =   1785
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1470
         Width           =   2085
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   1470
         Width           =   2085
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1110
         Width           =   4935
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1830
         Width           =   4995
      End
      Begin VB.TextBox txtUrl 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   2190
         Width           =   4995
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2400
         TabIndex        =   17
         Top             =   405
         Width           =   600
      End
      Begin VB.Label lblRuc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC/CI:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   795
         Width           =   540
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1515
         Width           =   765
      End
      Begin VB.Label lblFax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3360
         TabIndex        =   13
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1875
         Width           =   405
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   2235
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3293
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1733
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de las Empresas                             #
'#  frmEmpresa V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de las empresas del sistema.        #
'#  Permitirá almacenar en la base de datos nuevas empresas y modificar         #
'#  los datos de estas, dependiendo de la propiedad Tag, la cual se cambiará    #
'# en la ventana frmSelEmpresa y desde esta se llamará a esta ventana.          #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    empresa: En esta tabla se almacenan las nuevas empresas y se              #
'#             modifican los datos de estas.                                    #
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
    ' Si se esta ingresando una nueva empresa
    On Error GoTo errhandler
    If Me.Tag = "N" Then
    ' Almacenamiento de los datos de la nueva empresa
        strSql = " INSERT INTO empresa(emp_codigo,emp_nombre,emp_ruc,emp_direccion,emp_telf,emp_fax,emp_email,emp_url,emp_fechamod,emp_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & UCase(txtNombre.Text) & "','" & _
                 txtRuc.Text & "','" & UCase(txtDireccion.Text) & "','" & txtTelefono.Text & "','" & _
                 txtFax.Text & "','" & txtEmail.Text & "','" & txtUrl.Text & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsCon_Def.Ejecutar (strSql), "M"
        strSql = " INSERT INTO sucursal(suc_codigo,emp_codigo,suc_nombre,dep_codigo,suc_ctaconta_ventas,suc_ctaconta_ventas_sp,suc_ctaconta_servicios,suc_ctaconta_servicios_sp, " & _
                 " suc_ctaconta_costoventa,suc_direccion,suc_telefono,suc_ciudad,suc_fechamod,suc_usumod) " & _
                 " VALUES ('001','" & UCase(txtCodigo.Text) & "','MATRIZ','PRI','','','',''," & _
                 "'','" & UCase(txtDireccion.Text) & "','" & txtTelefono.Text & "','',CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsCon_Def.Ejecutar (strSql), "M"
        strSql = " INSERT INTO deposito(dep_codigo,emp_codigo,dep_nombre,dep_direccion,dep_telf,dep_fax,dep_email,dep_ctaconta,dep_fechamod,dep_usumod) " & _
                 " VALUES ('PRI','" & UCase(txtCodigo.Text) & "','PRINCIPAL','PRI','" & UCase(txtDireccion.Text) & "','" & txtTelefono.Text & "','" & txtFax.Text & "','" & txtEmail.Text & "'," & _
                 "'',CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsCon_Def.Ejecutar (strSql), "M"
    ' Creación de parámetros para las empresas
        If chkPlanCuenta.value = 1 Then
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM ctaconta WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO ctaconta" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        End If
        If chkRetencion.value = 1 Then
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM retencion WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO retencion" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        End If
        If chkParametro.value = 1 Then
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM parametro WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO parametro" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        Else
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM parametro WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "',par_numero=0,par_texto=''"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO parametro" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        
        End If
        If chkDefMov.value = 1 Then
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM tipo_ingreso WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO tipo_ingreso" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM tipo_egreso WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO tipo_egreso" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        Else
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM tipo_ingreso WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "',tip_ing_ctaconta='',tip_ing_ctaconta2=''"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO tipo_ingreso" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM tipo_egreso WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "',tip_ing_ctaconta='',tip_ing_ctaconta2=''"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO tipo_egreso" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        End If
        If chkFormaPago.value = 1 Then
            strSql = " CREATE TEMPORARY TABLE tmp" & _
                     " SELECT * FROM forma_pago WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " UPDATE tmp" & _
                     " SET emp_codigo='" & txtCodigo.Text & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " INSERT INTO forma_pago" & _
                     " SELECT * FROM tmp WHERE emp_codigo='" & dcmbEmpresa.BoundText & "'"
            clsCon_Def.Ejecutar (strSql), "M"
            
            strSql = " DROP TABLE tmp"
            clsCon_Def.Ejecutar (strSql), "M"
        End If
    ' Si se esta modificando
    Else
    ' Almacenamiento de los cambios realizados a la empresa
        strSql = " UPDATE empresa " & _
                 " SET emp_nombre='" & UCase(txtNombre.Text) & "',emp_ruc='" & txtRuc.Text & _
                 "',emp_direccion='" & UCase(txtDireccion.Text) & "',emp_telf='" & txtTelefono.Text & _
                 "',emp_fax='" & txtFax.Text & "',emp_email='" & txtEmail.Text & _
                 "',emp_url='" & txtUrl.Text & "',emp_fechamod=CURRENT_TIMESTAMP,emp_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & txtCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql), "M"
    End If
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
        Me.Caption = "Modificar datos de la Empresa"
        txtCodigo.Enabled = False
    ' Consulta para conocer los datos actuales de la empresa a modificar
        strSql = "SELECT * FROM empresa WHERE emp_codigo='" & txtCodigo.Text & "'"
        On Error GoTo errhandler
            clsCon_Def.Ejecutar (strSql)
            txtNombre.Text = clsCon_Def.adorec_Def("emp_nombre")
            txtRuc.Text = clsCon_Def.adorec_Def("emp_ruc")
            txtDireccion.Text = clsCon_Def.adorec_Def("emp_direccion")
            txtTelefono.Text = clsCon_Def.adorec_Def("emp_telf")
            txtFax.Text = clsCon_Def.adorec_Def("emp_fax")
            txtEmail.Text = clsCon_Def.adorec_Def("emp_email")
            txtUrl.Text = clsCon_Def.adorec_Def("emp_url")
            fraSel.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nueva Empresa"
        ' Consulta para conocer las empresas a las que el usuario tiene acceso
        strSql = " SELECT DISTINCT CONCAT(empresa_usu.emp_codigo,'-',sucursal.suc_codigo) as codigo,CONCAT(empresa.emp_nombre,' - ', suc_nombre) as nombre,dep_codigo,empresa_usu.emp_codigo,sucursal.suc_codigo " & _
                 " FROM empresa INNER JOIN empresa_usu " & _
                 " ON empresa.emp_codigo=empresa_usu.emp_codigo " & _
                 " INNER JOIN sucursal ON empresa.emp_codigo=sucursal.emp_codigo " & _
                 " WHERE empresa_usu.usu_codigo='" & strUsuario & "' " & _
                 " ORDER BY empresa.emp_nombre "
        clsCon_Def.Ejecutar strSql
        
        Set dcmbEmpresa.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbEmpresa.ListField = "nombre"
        dcmbEmpresa.BoundColumn = "codigo"
        fraSel.Enabled = True
        
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

Private Sub txtFax_GotFocus()
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

Private Sub txtRuc_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtUrl_GotFocus()
    Seleccionar_Contenido
End Sub
