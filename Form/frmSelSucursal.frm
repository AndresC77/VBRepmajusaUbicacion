VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelSucursal 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sucursales"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmSelSucursal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   7935
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   113
      TabIndex        =   7
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtCiudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   19
         Top             =   1320
         Width           =   2400
      End
      Begin VB.TextBox txtCtaVenSer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   18
         Top             =   1320
         Width           =   2400
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   6600
      End
      Begin VB.TextBox txtCtaVenPrd 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   2
         Top             =   960
         Width           =   2400
      End
      Begin VB.TextBox txtCtaCosVen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   1680
         Width           =   2400
      End
      Begin VB.TextBox txtBodega 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   1680
         Width           =   2400
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   2400
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1372
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Venta Servicios:"
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
         Left            =   3510
         TabIndex        =   20
         Top             =   1372
         Width           =   1530
      End
      Begin VB.Label lbldireccion 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   652
         Width           =   720
      End
      Begin VB.Label lblCodigo 
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
         Left            =   120
         TabIndex        =   15
         Top             =   292
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4440
         TabIndex        =   14
         Top             =   292
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Venta Productos:"
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
         Left            =   3450
         TabIndex        =   11
         Top             =   1012
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Costo Venta:"
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
         Left            =   3765
         TabIndex        =   10
         Top             =   1732
         Width           =   1275
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1732
         Width           =   600
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1012
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Bancos, y poder modificar, ingresar o eliminar   #
'#  Bancos                                                                      #
'#  frmSelBanco V1.0                                                            #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los bancos que hasta el momento estan ingresados en  #
'#  en el sistema. Desde esta ventana se puede añadir un nuevo banco, modificar #
'#  o eliminar los bancos ya ingresados.                                        #
'#  Esta ventana se llama a la ventana frmBanco en la que se añade y modifica   #
'#  los bancos                                                                  #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    Banaco: En esta tabla se almacenan los nuevos bancos, se modfican los     #
'#            datos de los bancos y se eliminan los bancos ya ingresados.       #
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

Private clsCon_Def As New clsConsulta
Private clsSql As New clsConsulta
Private strSql As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmdModificar_Click()

' Modifica los datos del banco seleccionado, se manda a la variable Tag del formulario una bandera para
' que indique se se va a modificar el banco, además se envia como datos a la forma fmrbanco el código y el nombre del banco
    Dim intPos As Integer
    'Verifica si se ha seleccionado un banco para ser modificado
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione una sucursal", vbInformation, "Sucursales"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmSucursal.Tag = "M"
    frmSucursal.txtCodigo.Text = Me.dcmbCodigo.Text
    frmSucursal.txtNombre.Text = Me.dcmbNombre.Text
    frmSucursal.TxtDireccion.Text = Me.TxtDireccion
    frmSucursal.TxtTelefono.Text = Me.TxtTelefono
    frmSucursal.TxtCiudad.Text = Me.TxtCiudad
    frmSucursal.dcmbBodega.BoundText = Me.txtBodega.Tag
    frmSucursal.dcmbCtaVenPrd.BoundText = Me.txtCtaVenPrd
    frmSucursal.dcmbCtaVenSer.BoundText = Me.txtCtaVenSer
    frmSucursal.dcmbCtaCosVen.BoundText = Me.txtCtaCosVen
    frmSucursal.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa un nuevo banco, se manda a la variable Tag del formulario una bandera para
' que indique se se va a ingresar un nuevo banco
    frmSucursal.Tag = "N"
    frmSucursal.Show
End Sub
Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
'Muestra el nombre relacionado con el código del depósito en el momento de seleccionar uno del combobox
    If clsCon_Def.adorec_Def.RecordCount = 0 Then Exit Sub
    clsCon_Def.adorec_Def.MoveFirst
    clsCon_Def.adorec_Def.Find "suc_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsCon_Def.adorec_Def.EOF = True Then
        dcmbNombre = ""
        dcmbNombre.BoundText = ""
        TxtDireccion = ""
        TxtTelefono = ""
        TxtCiudad = ""
        txtBodega = ""
        txtCtaVenPrd = ""
        txtCtaVenSer = ""
        txtCtaCosVen = ""
        cmdModificar.Enabled = False
    Else
        dcmbNombre = clsCon_Def.adorec_Def("suc_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        TxtDireccion = clsCon_Def.adorec_Def("suc_direccion")
        TxtTelefono.Text = clsCon_Def.adorec_Def("suc_telefono")
        TxtCiudad = clsCon_Def.adorec_Def("suc_ciudad")
        txtBodega.Text = clsCon_Def.adorec_Def("dep_nombre")
        txtBodega.Tag = clsCon_Def.adorec_Def("dep_codigo")
        txtCtaVenPrd = clsCon_Def.adorec_Def("suc_ctaconta_ventas")
        txtCtaVenSer = clsCon_Def.adorec_Def("suc_ctaconta_servicios")
        txtCtaCosVen = clsCon_Def.adorec_Def("suc_ctaconta_costoventa")
        cmdModificar.Enabled = True
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbNombre_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub


Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub Form_Activate()
 
    'Muestra la lista de depósitos actualizada
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "suc_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "suc_nombre"
    dcmbNombre.BoundColumn = "suc_codigo"
    If Me.Tag <> "" Then
        dcmbCodigo = ""
        dcmbCodigo = Me.Tag
    ElseIf Not clsCon_Def.adorec_Def.EOF Then
        dcmbCodigo_Change
    End If
End Sub

Private Sub Form_Load()
 Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta los documentos que estan disponibles
        strSql = " SELECT suc_codigo,suc_nombre,dep_nombre,sucursal.dep_codigo,suc_ctaconta_ventas,suc_ctaconta_servicios,suc_ctaconta_costoventa,suc_direccion,suc_telefono,suc_ciudad " & _
                 " FROM sucursal INNER JOIN deposito ON sucursal.emp_codigo=deposito.emp_codigo AND sucursal.dep_codigo=deposito.dep_codigo" & _
                 " ORDER BY suc_codigo"
        
        clsCon_Def.Ejecutar (strSql)
        
        'Muestra los datos de los códigos del depósito
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "suc_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "suc_nombre"
        dcmbNombre.BoundColumn = "suc_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("suc_codigo")
        End If
        Exit Sub
'        ,ban_nombre,ban_direccion,ban_telefono,ban_email,ban_url
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

