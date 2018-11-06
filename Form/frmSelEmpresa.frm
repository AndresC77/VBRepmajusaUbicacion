VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelEmpresa 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Empresa"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   5670
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Selección de Empresa"
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
      Left            =   143
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtPtoFactura 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4200
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   1080
      End
      Begin MSDataListLib.DataCombo dcmbEmpresa 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   582
         _Version        =   393216
         IntegralHeight  =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Punto de Facturación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2520
         TabIndex        =   10
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label lblEmpresas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   765
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar Empresa"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2670
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   2670
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNueva 
      Caption         =   "&Nueva Empresa"
      Height          =   405
      Left            =   990
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de seleccion de empresa                                               #
'#  frmSelEmpresa V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la seleccion y el cambio de empresa en la que se trabajará.    #
'#  Permitirá ademas tener acceso al ingreso de nuevas empresas y modificación  #
'#  de los datos de estas.                                                      #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    empresa: Utilizada para obtener la lista de empresas disponibles,         #
'#             presentando estos nombres en el combo y en el textbox se         #
'#             presentará los códigos de la empresa seleccionada.               #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#    mdiPrincipal.Crear_Menu: Proceso para la creación del menú controlando    #
'#                             los permisos que tiene cada usuario.             #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As clsConsulta

Private Sub Form_Activate()
    HoyDia = Hoy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmdAceptar_Click()
' Selección de una empresa y almacenamiento de su código en la variable strEmpresa y
' Cambia el Caption de mdiPrincipal para conocer la empresa seleccionada
    Dim strNombreEmpresa As String
    If txtCodigo.Text <> "" Then
        strNombreEmpresa = dcmbEmpresa.Text
        If Trim(strPtoFactura) = "" Then
            frmSelNegocio.Show vbModal
        End If
        
        clsCon_Def.Filtrar "codigo='" & txtCodigo.Text & "'"
        strEmpresa = clsCon_Def.adorec_Def("emp_codigo")
        strSucursal = clsCon_Def.adorec_Def("suc_codigo")
        strBodega = clsCon_Def.adorec_Def("dep_codigo")
        PorIVA = clsCon_Def.adorec_Def("poriva")
        CodigoIVA = clsCon_Def.adorec_Def("cod_iva_codigo")
        GeneraDocElec = clsCon_Def.adorec_Def("elec")
        PtoEmiDocEle = clsCon_Def.adorec_Def("PtoEmiDocEle")
        ReservaNotaCredito = clsCon_Def.adorec_Def("PNC")
        
        mdiPrincipal.Crear_Menu
        mdiPrincipal.Caption = strSucursal & "  NEED - Enlace Digital  "
        mdiPrincipal.StatusBar.Panels(1).Text = "   " & strNombreEmpresa & ":   " & Right(txtCodigo.Text, 3) & "   "
        mdiPrincipal.StatusBar.Panels(2).Text = "   PUNTO DE FACTURACION:   " & txtPtoFactura.Text & "   "
        mdiPrincipal.StatusBar.Panels(3).Text = "   SUC :   " & strSucursal & "   "
        mdiPrincipal.StatusBar.Panels(4).Text = "   USUARIO:   " & strUsuario & "   "
        
        mdiPrincipal.menNuevoBrowser_Click
        Unload Me
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de una empresa, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código de la empresa que se modificará
    frmEmpresa.Show
    frmEmpresa.txtCodigo.Text = Me.txtCodigo.Text
    frmEmpresa.Tag = "M"
End Sub

Private Sub cmdNueva_Click()
' Crea una nueva empresa, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará una nueva empresa
    frmEmpresa.Show
    frmEmpresa.Tag = "N"
End Sub

Private Sub dcmbEmpresa_Change()
' Chequea la empresa seleccionada y escribe su códio en el textbox
    Dim strComparar As String
    On Error GoTo errhandler
        If dcmbEmpresa.MatchedWithList = True Then
            txtCodigo.Text = dcmbEmpresa.BoundText
            cmdModificar.Enabled = True
            cmdAceptar.Enabled = True
        Else
            txtCodigo.Text = ""
            cmdModificar.Enabled = False
            cmdAceptar.Enabled = False
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
    'Carga la lista de empresas al combo
    Dim strSql As String
    
    '****Inicialmente deshabilito las opciones de Nueva y Modificar empresa
    
    cmdNueva.Visible = False
    cmdModificar.Visible = False
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    txtPtoFactura.Text = strPtoFactura
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        HoyDia = Hoy
        strSql = "select count(emp_codigo) from empresa"
        clsCon_Def.Ejecutar strSql
        
    ' En caso de no haber empresas se visualizan las opciones
    ' de nueva (habilitada) y modificar (desabilitada)
        If (clsCon_Def.adorec_Def.Fields(0).Value = 0) Then
            cmdNueva.Visible = True
            cmdModificar.Visible = True
            'Centra esta forma dentro de la forma MDI
            Me.Left = (mdiPrincipal.Width - Me.Width) / 2
            Me.Top = 0
        End If
    ' Consulta para conocer las empresas a las que el usuario tiene acceso
        strSql = " SELECT DISTINCT CONCAT(empresa_usu.emp_codigo,'-',sucursal.suc_codigo) as codigo,CONCAT(empresa.emp_nombre,' - ', suc_nombre) as nombre,dep_codigo," & _
                 " empresa_usu.emp_codigo,sucursal.suc_codigo,parametro.par_numero as poriva,p2.par_numero as elec,p2.par_texto as PtoEmiDocEle,p3.par_numero as PNC,cod_iva_codigo " & _
                 " FROM empresa INNER JOIN empresa_usu " & _
                 " ON empresa.emp_codigo=empresa_usu.emp_codigo " & _
                 " INNER JOIN sucursal ON empresa.emp_codigo=sucursal.emp_codigo " & _
                 " INNER JOIN parametro ON empresa.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV'" & _
                 " INNER JOIN codigo_iva ON parametro.par_numero=codigo_iva.cod_iva_porcentaje" & _
                 " INNER JOIN parametro p2 ON empresa.emp_codigo=p2.emp_codigo AND p2.par_codigo='DEL'" & _
                 " INNER JOIN parametro p3 ON empresa.emp_codigo=p3.emp_codigo AND p3.par_codigo='PNC'" & _
                 " WHERE empresa_usu.usu_codigo='" & strUsuario & "' " & _
                 " ORDER BY CONCAT(empresa.emp_nombre,' - ', suc_nombre) "
        clsCon_Def.Ejecutar strSql
        
        Set dcmbEmpresa.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbEmpresa.ListField = "nombre"
        dcmbEmpresa.BoundColumn = "codigo"
        dcmbEmpresa.BoundText = clsCon_Def.adorec_Def("codigo")
        Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
