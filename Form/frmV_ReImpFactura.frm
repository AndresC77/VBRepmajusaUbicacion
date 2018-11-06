VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmV_ReImpFactura 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re Impresión de Facturas"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_ReImpFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   5655
   Begin VB.CheckBox chkFacturaTicket 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Fac.Ticket"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkConGuia 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Con Guia"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   2100
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impresíon de Facturas"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5415
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   720
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCotizacion 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   1080
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3660
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmV_ReImpFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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
Private strSql As String
Private clsSql As New clsConsulta

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.MatchedWithList = True Then
        strSql = " SELECT egr_codigo " & _
                 " FROM egreso INNER JOIN persona ON (egreso.emp_codigo = persona.emp_codigo) AND (egreso.per_codigo = persona.per_codigo) " & _
                 " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' AND cat_p_tipo='C' " & _
                 " ORDER BY egreso.egr_codigo "
        clsSql.Ejecutar (strSql)
        cmbCotizacion = ""
        Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
        cmbCotizacion.ListField = "egr_codigo"
    End If
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        LimpiarTodo
        strSql = " SELECT tip_ped_ptofac,tip_ped_facturaticket " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
            chkFacturaTicket.Value = clsSql.adorec_Def("tip_ped_facturaticket")
        End If
    Else
        Exit Sub
    End If
    
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
    cmbCotizacion.BoundText = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    Dim GuiaAutomatica As Boolean
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        Dim RepIDCaja As New frmReporte
        'Dim RepRC As New frmReporte
        RepIDCaja.strNumero = Me.cmbCliente.BoundText
        RepIDCaja.strReporte = "rptIDCaja"
        RepIDCaja.Show
        'RepRC.strNumero = cmbCotizacion.BoundText
        'RepRC.strReporte = "rptRC"
        'RepRC.Show
        If Me.chkFacturaTicket.Value = 0 Then
            frmReporte.strNumero = cmbCotizacion.BoundText
            'listo
            GuiaAutomatica = IIf(chkConGuia.Value = 1, True, False)
            frmReporte.strReporte = IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola")
            frmReporte.Show
        Else
            frmImpresionDirecta.strNumero = cmbCotizacion.BoundText
            frmImpresionDirecta.strReporte = "rptFacturaSola"
            frmImpresionDirecta.Show
            frmImpresionDirecta.optImpresora.Value = True
            frmImpresionDirecta.cmdImprimir_Click
        End If
        
    Else
        MsgBox "No ha seleccionado una factura", vbInformation, "Factura"
    End If
End Sub
Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    cargarTipoPedido
    
End Sub
