VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersona 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de Clientes  y  Proveedores"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersona.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7350
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Personales"
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
      Left            =   128
      TabIndex        =   14
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   2160
         Width           =   2265
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   1800
         Width           =   2265
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   2265
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   2265
      End
      Begin VB.TextBox txtRuc 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1455
         Width           =   2265
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1095
         Width           =   2265
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   735
         Width           =   2265
      End
      Begin MSDataListLib.DataCombo dcmbZona 
         Height          =   330
         Left            =   4680
         TabIndex        =   9
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbCiudad 
         Height          =   330
         Left            =   4680
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPais 
         Height          =   330
         Left            =   4680
         TabIndex        =   7
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbCategoria 
         Height          =   330
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   25
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   24
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   23
         Top             =   2205
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   22
         Top             =   1845
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   2205
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   1845
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pais:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   18
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   17
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3728
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   2168
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma que permite la modificación de datos del cliente y proveedor          #
'#  frmPersona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana  que permite visualizar y modificar los datos de los clientes y     #
'#  proveedores, dependiendo del tag de la forma.  Esta ventana viene de la     #
'#  forma frmComprobanteEgresoComun                                             #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#                                                                              #
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
Private clsPer As New clsConsulta
Private clsCat As New clsConsulta
Private clsPai As New clsConsulta
Private clsZon As New clsConsulta
Private clsCiu As New clsConsulta
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
    Set clsPer = Nothing
    Set clsCat = Nothing
    Set clsPai = Nothing
    Set clsZon = Nothing
    Set clsCiu = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmbAceptar_Click()
    
    strSql = " UPDATE persona " & _
             " SET zon_codigo = '" & UCase(dcmbZona.BoundText) & "', cat_p_codigo = '" & UCase(dcmbCategoria.BoundText) & "', ciu_codigo = '" & UCase(dcmbCiudad.BoundText) & "', per_ruc = '" & txtRuc.Text & "', per_nombre = '" & UCase(txtNombre.Text) & "', per_apellido='" & UCase(txtApellido.Text) & "', per_direccion = '" & UCase(txtDireccion.Text) & "', per_telf = '" & txtTelefono.Text & "', per_fax= '" & txtFax.Text & "', per_email = '" & txtEmail.Text & "',per_fechamod = CURRENT_TIMESTAMP, per_usumod = '" & strUsuario & "' " & _
             " WHERE per_codigo = '" & txtCodigo & " ' "
    clsSql.Ejecutar (strSql)
    MsgBox " Los datos han sido cambiados", vbInformation, "Datos personales"
       
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub dcmbPais_Change()
 'Consulta las ciudades existentes
    strSql = " SELECT ciu_codigo, ciu_nombre" & _
             " FROM ciudad " & _
             " WHERE pai_codigo = '" & dcmbPais.BoundText & "'"
    clsCiu.Ejecutar (strSql)
    
    If clsCiu.adorec_Def.EOF = False Then
        Set dcmbCiudad.RowSource = clsCiu.adorec_Def.DataSource
        dcmbCiudad.ListField = "ciu_nombre"
        dcmbCiudad.BoundColumn = "ciu_codigo"
        dcmbCiudad.Text = clsCiu.adorec_Def("ciu_nombre")
    Else
        dcmbCiudad = ""
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)

'    On Error GoTo errhandler
 
    clsPer.Inicializar AdoConn
    clsCat.Inicializar AdoConn
    clsPai.Inicializar AdoConn
    clsZon.Inicializar AdoConn
    clsCiu.Inicializar AdoConn
    clsSql.Inicializar AdoConn
    
'    Consulta los paises existentes
    strSql = " SELECT pai_codigo, pai_nombre" & _
             " FROM pais "
    clsPai.Ejecutar (strSql)
    If clsPai.adorec_Def.EOF = False Then
        Set dcmbPais.RowSource = clsPai.adorec_Def.DataSource
        dcmbPais.ListField = "pai_nombre"
        dcmbPais.BoundColumn = "pai_codigo"
    Else
        dcmbPais = ""
        dcmbCiudad.Enabled = False
    End If
    
    'Consulta las zonas existentes
    strSql = " SELECT zon_codigo, zon_nombre" & _
             " FROM zona "
    clsZon.Ejecutar (strSql)
    If clsZon.adorec_Def.EOF = False Then
        Set dcmbZona.RowSource = clsZon.adorec_Def.DataSource
        dcmbZona.ListField = "zon_nombre"
        dcmbZona.BoundColumn = "zon_codigo"
    Else
        dcmbZona = ""
    End If
    
End Sub



Private Sub txtApellido_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtCodigo_Change()
    
    strSql = " SELECT distinct persona.zon_codigo, persona.cat_p_codigo, persona.ciu_codigo, per_ruc, per_nombre, per_apellido, per_direccion, per_telf, per_fax, per_email, cat_p_nombre, zon_nombre, ciu_nombre, ciudad.pai_codigo, pai_nombre " & _
             " FROM ((((persona INNER JOIN categoria_p ON persona.cat_p_codigo = categoria_p.cat_p_codigo)INNER JOIN zona ON persona.zon_codigo = zona.zon_codigo)INNER JOIN ciudad ON persona.ciu_codigo = ciudad.ciu_codigo)INNER JOIN pais ON ciudad.pai_codigo = pais.pai_codigo)" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND per_codigo = '" & txtCodigo.Text & "' "
    clsPer.Ejecutar (strSql)
    
    dcmbCategoria.Text = clsPer.adorec_Def("cat_p_nombre")
    dcmbCiudad.Text = clsPer.adorec_Def("ciu_nombre")
    dcmbPais.Text = clsPer.adorec_Def("pai_nombre")
    dcmbZona.Text = clsPer.adorec_Def("zon_nombre")
    txtNombre.Text = clsPer.adorec_Def("per_nombre")
    txtApellido.Text = clsPer.adorec_Def("per_apellido")
    txtTelefono.Text = clsPer.adorec_Def("per_telf")
    txtFax.Text = clsPer.adorec_Def("per_fax")

    If Me.Tag = "P" Then
    
        strSql = " SELECT distinct cat_p_codigo, cat_p_nombre" & _
                 " FROM categoria_p " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo = 'P'"
        clsCat.Ejecutar (strSql)
    
        Set dcmbCategoria.RowSource = clsCat.adorec_Def.DataSource
        dcmbCategoria.ListField = "cat_p_nombre"
        dcmbCategoria.BoundColumn = "cat_p_codigo"
       
    ElseIf Me.Tag = "C" Then
    
        strSql = " SELECT distinct cat_p_codigo, cat_p_nombre" & _
                 " FROM categoria_p " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo = 'C'"
        clsCat.Ejecutar (strSql)
    
        Set dcmbCategoria.RowSource = clsCat.adorec_Def.DataSource
        dcmbCategoria.ListField = "cat_p_nombre"
        dcmbCategoria.BoundColumn = "cat_p_codigo"
       
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
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

Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtRuc_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Contenido
End Sub

