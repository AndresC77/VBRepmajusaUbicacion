VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCambioProducto 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Productos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmCambioProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   12600
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingreso"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   12375
      Begin VB.TextBox txtCantIng 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalleIng 
         Height          =   1770
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12135
         _cx             =   61100957
         _cy             =   61082674
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCambioProducto.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblCantINg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   7320
         TabIndex        =   15
         Top             =   2070
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo de Negocio:"
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
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5535
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   375
         Width           =   3975
         _ExtentX        =   7011
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
         TabIndex        =   12
         Top             =   420
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4950
      TabIndex        =   7
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos del Cambio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   9975
      Begin VB.TextBox txtRuc 
         Height          =   285
         Left            =   7320
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   630
         Width           =   1695
         _extentx        =   2990
         _extenty        =   450
         value           =   41988.4286226852
      End
      Begin VB.TextBox TxtObserv 
         Height          =   555
         Left            =   1080
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   5535
      End
      Begin MSDataListLib.DataCombo cmbFactura 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CI/RUC:"
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
         Left            =   6720
         TabIndex        =   20
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         TabIndex        =   18
         Top             =   292
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
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
         Left            =   3000
         TabIndex        =   17
         Top             =   652
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observ:"
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
         TabIndex        =   10
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         TabIndex        =   9
         Top             =   652
         Width           =   495
      End
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   600
      Picture         =   "frmCambioProducto.frx":0471
      Top             =   5520
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   360
      Picture         =   "frmCambioProducto.frx":059D
      Top             =   5520
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmCambioProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso de mercadería a los depòsitos por concepto de         #
'#  importaciones se permite crear estos ingresos                               #
'#  frmIngImportacion  V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite ingresar los productos a los diferentes depòsitos       #
'#  de la compañía por concepto de importaciones , solo se permite el ingreso   #
'#  de tales datos para posteriormente actualizar las existencias.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se almacenan los nuevos ingresos de mercadería #
'#    det_ingreso: En estatabla se almacena los detalles de cada ingreso        #
'#    persona    : Se consulta los proveedores de la empresa                    #
'#    deposito   : Se consulta los depositos o bodegas de la empresa            #
'#    producto   : Se consulta los productos de la empresa                      #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#               limpiarFxGD()   Permite borrar los datos que se encuentran     #
'#                               en el flexGrid para realizar un nuevo ingreso  #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Public Neg As String
Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private clsCon_Prd2 As New clsConsulta
Private strSQL As String
Private ValDias As Long

Private Sub cmbCliente_Change()
    ValDias = 30
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    cmbFactura = ""
    If ValDias = 0 Then ValDias = 30
        
    strSQL = " SELECT lis_pre_codigo " & _
             " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
             " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
             " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
             " Where persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND per_codigo='" & cmbCliente.BoundText & "' "
    clsCon_Def.Ejecutar (strSQL)
    cmbCliente.Tag = clsCon_Def.adorec_Def("lis_pre_codigo")
    strSQL = " SELECT egr_codigo as fac " & _
             " FROM egreso " & _
             " Where emp_codigo='" & strEmpresa & "' And tip_egr_codigo='FAC' " & _
             " AND egr_anulado=0 " & _
             " AND egr_fecha>='" & DateAdd("d", -1 * ValDias, HoyDia) & "' AND per_codigo='" & cmbCliente.BoundText & "'" & _
             " UNION " & _
             " SELECT concat('R',egr_codigo) as fac " & _
             " FROM factura_ryb " & _
             " Where emp_codigo='" & strEmpresa & "' And tip_egr_codigo='FAC' " & _
             " AND egr_anulado=0 " & _
             " AND egr_fecha>='" & DateAdd("d", -1 * ValDias, HoyDia) & "' AND per_codigo='" & cmbCliente.BoundText & "'" & _
             " ORDER BY fac "
    clsCon_Def.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbFactura.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbFactura.ListField = "fac"
        cmbFactura.BoundColumn = "fac"
    Else
        cmbFactura = "No hay facturas del cliente "
    End If
End Sub


Private Sub cmbFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If MsgBox("Desea revisar todas las facturas del cliente?", vbQuestion + vbYesNo + vbDefaultButton2, "Ajustes") = vbYes Then
            frmClave.strClaveMAESTRA = "posa"
            frmClave.Show vbModal
            If frmClave.Ret = False Then
                ValDias = 30
            Else
                ValDias = 1000
            End If
        Else
            ValDias = 30
        End If
        cmbCliente_Validate False
    End If
End Sub

Private Sub cmbFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF4 Then
        If MsgBox("Desea revisar todas las facturas del cliente?", vbQuestion + vbYesNo + vbDefaultButton2, "Ajustes") = vbYes Then
            frmClave.strClaveMAESTRA = "wsed"
            frmClave.Show vbModal
            frmClave.Show vbModal
            If frmClave.Ret = False Then
                ValDias = 30
            Else
                ValDias = 1000
            End If
        Else
            ValDias = 30
        End If
        
    End If
End Sub

Private Sub cmbFactura_Validate(Cancel As Boolean)
    CargaProductos
End Sub

Private Sub cmbNegocio_Change()
    ValDias = 30
    If cmbNegocio.BoundText <> "" Then
        strSQL = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSQL
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If strPtoFactura <> clsCon_Def.adorec_Def(0) Then
                LimpiarTodo
            End If
            strPtoFactura = clsCon_Def.adorec_Def(0)
        End If
    Else
        Exit Sub
    End If
    strSQL = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
End Sub

Private Sub cmdAceptar_Click()
    Dim clsCambio As New clsCambio
    Dim booGuardar As Boolean
    Dim i As Long, j As Long, cIng As Long, no As Boolean, cEgr As Long
    no = False
    
    If cmbCliente.BoundText = "" Then
        MsgBox "Seleccione el cliente al cual realizará el cambio de producto", vbInformation, "Cambio de Productos"
        cmbCliente.SetFocus
        Exit Sub
    End If
    
    
    clsCambio.Inicializar AdoConn, AdoConnMaster
    Dim DocCambio As String
    
    booGuardar = clsCambio.NuevoCambio(True, strSucursal, strPtoFactura, cmbCliente.BoundText, dtpFecha.Value, cmbFactura.BoundText, UCase(TxtObserv.Text))
    If booGuardar = True Then
        With vsfgDetalleIng
            For i = 1 To .Rows - 1
                clsCambio.NuevoDet .TextMatrix(i, 1), .TextMatrix(i, 2), FormatoD4(.TextMatrix(i, 4)), .TextMatrix(i, 6)
            Next i
        End With
        MsgBox " Los datos han sido ingresados" & vbNewLine & clsCambio.strDoc, vbInformation, "Cambio de Productos"
    End If
    
    If booGuardar = True Then
        Dim rpMov1 As New frmReporte
        rpMov1.strNumero = FormatoD0(clsCambio.strDoc)
        rpMov1.strReporte = "rptTckAjuste"
        rpMov1.Show
        Dim rpMov2 As New frmReporte
        rpMov2.strNumero = FormatoD0(clsCambio.strDoc)
        rpMov2.strReporte = "rptAjuste"
        rpMov2.Show
    End If
    
    
    Set clsCambio = Nothing
        
    Unload Me
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim rpMov1 As New frmReporte
    rpMov1.strNumero = "10020000352"
    rpMov1.strReporte = "rptTckAjuste"
    rpMov1.Show
    Dim rpMov2 As New frmReporte
    rpMov2.strNumero = "10020000352"
    rpMov2.strReporte = "rptAjuste"
    rpMov2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSQL = ""
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_Prd2 = Nothing
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    ValDias = 30
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd2.Inicializar AdoConn, AdoConnMaster
    dtpFecha.Value = HoyDia
    dtpFecha.Enabled = False
    
    cargarTipoPedido
        
    strSQL = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
    
    PonerBotones
End Sub
Private Sub txtRuc_Validate(Cancel As Boolean)
    
    If Trim(txtRuc.Text) <> "" Then
        strSQL = " SELECT per_codigo " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND per_ruc='" & txtRuc.Text & "'"
        clsCon_Def.Ejecutar strSQL
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If cmbCliente.BoundText <> clsCon_Def.adorec_Def(0) Then
                cmbCliente.BoundText = clsCon_Def.adorec_Def(0)
                cmbCliente_Validate False
            End If
        Else
            MsgBox "No se encontró un cliente con CI/RUC " & txtRuc.Text, vbInformation, "CI/RUC"
            If cmbCliente.Text <> "" Then
                LimpiarTodo
            Else
                txtRuc.Text = ""
            End If
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 3 Or Col = 5 Or Col = 8 Then Cancel = True
End Sub

Private Sub VsfgDetalleIng_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleIng.MouseRow
    c = vsfgDetalleIng.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = (vsfgDetalleIng.Rows - 1)) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleIng.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleIng.Cell(flexcpLeft, r, c) + vsfgDetalleIng.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         vsfgDetalleIng.RemoveItem (r)
         PonerBotones
         CalculaTotal
    Else
        vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (vsfgDetalleIng.Rows - 1)
        vsfgDetalleIng.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            vsfgDetalleIng.Cell(flexcpPicture, i, 0) = imgBtnUp
            vsfgDetalleIng.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub

Private Sub cargarTipoPedido()
    strSQL = " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo LIKE '" & Neg & "' " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSQL
    Set cmbNegocio.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSQL = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo LIKE '" & Neg & "' " & _
             " AND tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
End Sub

Private Sub ConsultaProducto(PrdCodigo As String)


        If Left(cmbFactura.Text, 1) = "R" Then
            strSQL = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio,det_egr_cantidad,producto.prd_sku " & _
                     " FROM det_factura_ryb INNER JOIN producto ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
                     " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                     " WHERE det_factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_factura_ryb.tip_egr_codigo = 'FAC' " & _
                     " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
                     " AND prd_baja=0 AND det_factura_ryb.prd_codigo='" & PrdCodigo & "'" & _
                     " ORDER BY prd_nombre "
        Else
            strSQL = " SELECT producto.prd_codigo, prd_nombre,det_egr_precio as prd_precio,det_egr_cantidad,producto.prd_sku " & _
                     " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
                     " AND det_egreso.prd_codigo=producto.prd_codigo " & _
                     "  " & _
                     "  " & _
                     " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_egreso.tip_egr_codigo = 'FAC' " & _
                     " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                     " AND det_egreso.prd_codigo='" & PrdCodigo & "' " & _
                     " UNION " & _
                     " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio,det_cam_cantidad,producto.prd_sku " & _
                     " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo " & _
                     " AND cambio.cam_codigo=det_cambio.cam_codigo " & _
                     " INNER JOIN producto ON det_cambio.emp_codigo=producto.emp_codigo " & _
                     " AND det_cambio.prd_codigo_ped=producto.prd_codigo " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                     " WHERE cambio.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_cambio.tip_ing_codigo = 'ICA' " & _
                     " AND cambio.cam_factura = '" & cmbFactura.BoundText & "' " & _
                     " AND prd_baja=0 AND det_cambio.prd_codigo_ped='" & PrdCodigo & "' " & _
                     " ORDER BY prd_nombre "
        End If
        clsCon_Prd.Ejecutar strSQL

End Sub

Private Sub CargaProductos()

    'Carga los motivos de ajuste
    strSQL = " SELECT mot_aju_codigo,mot_aju_nombre " & _
             " FROM motivo_ajuste " & _
             " Where emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY mot_aju_codigo "
    clsCon_Def.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(1) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "*mot_aju_codigo, mot_aju_nombre", "mot_aju_codigo")
    
    If cmbFactura.MatchedWithList = True Then
'        'Consulto los productos de la factura
'        If Left(cmbFactura.Text, 1) = "R" Then
'            strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio,det_egr_cantidad " & _
'                     " FROM det_factura_ryb INNER JOIN producto ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
'                     " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
'                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
'                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
'                     " WHERE det_factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
'                     " AND det_factura_ryb.tip_egr_codigo = 'FAC' " & _
'                     " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
'                     " AND prd_baja=0 " & _
'                     " ORDER BY prd_nombre "
'        Else
'            strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio,det_egr_cantidad " & _
'                     " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
'                     " AND det_egreso.prd_codigo=producto.prd_codigo " & _
'                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
'                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
'                     " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
'                     " AND det_egreso.tip_egr_codigo = 'FAC' " & _
'                     " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
'                     " AND prd_baja=0 " & _
'                     " UNION " & _
'                     " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio,det_cam_cantidad " & _
'                     " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo " & _
'                     " AND cambio.cam_codigo=det_cambio.cam_codigo " & _
'                     " INNER JOIN producto ON det_cambio.emp_codigo=producto.emp_codigo " & _
'                     " AND det_cambio.prd_codigo_ped=producto.prd_codigo " & _
'                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
'                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
'                     " WHERE cambio.emp_codigo = '" & strEmpresa & "' " & _
'                     " AND det_cambio.tip_ing_codigo = 'ICA' " & _
'                     " AND cambio.cam_factura = '" & cmbFactura.BoundText & "' " & _
'                     " AND prd_baja=0 " & _
'                     " ORDER BY prd_nombre "
'        End If
'        clsCon_Prd.Ejecutar strSql
        'vsfgDetalleIng.ColComboList(3) = vsfgDetalleIng.BuildComboList(clsCon_Prd.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
        
        'Consulto los productos para el cambio
        
'        strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
'                 " FROM producto INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
'                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
'                 " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
'                 " AND lis_pre_p_precio!=0 " & _
'                 " AND prd_baja=0 ORDER BY prd_codigo "
'        clsCon_Prd2.Ejecutar strSql
'        vsfgDetalleIng.ColComboList(6) = vsfgDetalleIng.BuildComboList(clsCon_Prd2.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
        
        
        If Left(cmbFactura.Text, 1) = "R" Then
            strSQL = " SELECT DISTINCT pp.prd_codigo, pp.prd_nombre,lis_pre_p_precio as prd_precio,pp.prd_sku " & _
                     " FROM det_factura_ryb INNER JOIN producto ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
                     " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
                     " INNER JOIN producto pp ON producto.emp_codigo=pp.emp_codigo " & _
                     " AND LEFT(producto.prd_sku,8)=LEFT(pp.prd_sku,8) " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                     " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_egreso.tip_egr_codigo = 'FAC' AND pp.prd_baja=0 " & _
                     " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
                     " ORDER BY prd_nombre"
        Else
            strSQL = " SELECT DISTINCT pp.prd_codigo, pp.prd_nombre,lis_pre_p_precio as prd_precio,pp.prd_sku " & _
                     " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
                     " AND det_egreso.prd_codigo=producto.prd_codigo " & _
                     " INNER JOIN producto pp ON producto.emp_codigo=pp.emp_codigo " & _
                     " AND LEFT(producto.prd_sku,8)=LEFT(pp.prd_sku,8) " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                     " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                     " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_egreso.tip_egr_codigo = 'FAC' AND pp.prd_baja=0 " & _
                     " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                     " ORDER BY prd_nombre"
        End If
        clsCon_Prd2.Ejecutar strSQL
        vsfgDetalleIng.ColComboList(7) = vsfgDetalleIng.BuildComboList(clsCon_Prd2.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    End If

End Sub

Private Sub vsfgDetalleIng_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If vsfgDetalleIng.TextMatrix(Row, Col) <> "" And Not IsNumeric(vsfgDetalleIng.TextMatrix(Row, Col)) Then
            MsgBox "Ingrese valores numéricos en Cantidad", vbInformation, "Detalle"
            vsfgDetalleIng.TextMatrix(Row, Col) = 0
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldCol = 4 Then
        If vsfgDetalleIng.TextMatrix(OldRow, 4) > vsfgDetalleIng.TextMatrix(OldRow, 9) Then
            MsgBox "No puede pedir el cambio de mas de " & vsfgDetalleIng.TextMatrix(OldRow, 9) & " prendas", vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(OldRow, 4) = 0
            Cancel = True
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        If vsfgDetalleIng.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero un motivo", vbInformation, "Motivos"
            vsfgDetalleIng.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
        ConsultaProducto vsfgDetalleIng.TextMatrix(Row, 2)
        'clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        If clsCon_Prd.adorec_Def.RecordCount > 0 Then
            vsfgDetalleIng.TextMatrix(Row, 3) = clsCon_Prd.adorec_Def("prd_nombre")
            vsfgDetalleIng.TextMatrix(Row, 4) = 1
            vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_sku")
            vsfgDetalleIng.TextMatrix(Row, 9) = ConsutaCantidadMaxima(vsfgDetalleIng.TextMatrix(Row, 2))
        Else
            MsgBox "Producto no encontrado en la factura", vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(Row, 2) = ""
            vsfgDetalleIng.TextMatrix(Row, 3) = ""
            vsfgDetalleIng.TextMatrix(Row, 4) = ""
            vsfgDetalleIng.TextMatrix(Row, 5) = ""
        End If
    ElseIf Col = 6 Then
        vsfgDetalleIng.TextMatrix(Row, 7) = vsfgDetalleIng.TextMatrix(Row, 6)
        clsCon_Prd2.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 6) & "'"
        
        If clsCon_Prd2.adorec_Def.RecordCount > 0 Then
            vsfgDetalleIng.TextMatrix(Row, 8) = clsCon_Prd2.adorec_Def("prd_sku")
        Else
            MsgBox "Producto no encontrado el producto", vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(Row, 6) = ""
            vsfgDetalleIng.TextMatrix(Row, 7) = ""
        End If
        
    ElseIf Col = 7 Then
        vsfgDetalleIng.TextMatrix(Row, 6) = vsfgDetalleIng.TextMatrix(Row, 7)
        clsCon_Prd2.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 7) & "'"
        vsfgDetalleIng.TextMatrix(Row, 8) = clsCon_Prd2.adorec_Def("prd_sku")
    End If
    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 2) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 3) <> "" And Val(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 4)) <> 0 Then
        vsfgDetalleIng.AddItem ""
        vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 0) = vsfgDetalleIng.Rows - 1
        vsfgDetalleIng.Cell(flexcpPicture, vsfgDetalleIng.Rows - 1, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, vsfgDetalleIng.Rows - 1, 0) = flexAlignRightCenter
    End If
    CalculaTotal
End Sub

Private Function ConsutaCantidadMaxima(strProducto As String) As Long
    
    If Left(cmbFactura.Text, 1) = "R" Then
        strSQL = " SELECT MAX(det_egr_cantidad) as cantidad " & _
                 " FROM det_factura_ryb INNER JOIN producto ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
                 " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE det_factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_factura_ryb.tip_egr_codigo = 'FAC' " & _
                 " AND det_factura_ryb.prd_codigo='" & strProducto & "' " & _
                 " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_factura_ryb.emp_codigo,det_factura_ryb.egr_codigo,det_factura_ryb.prd_codigo " & _
                 " ORDER BY cantidad DESC "
    Else
        strSQL = " SELECT MAX(det_egr_cantidad) as cantidad " & _
                 " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
                 " AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_egreso.tip_egr_codigo = 'FAC' " & _
                 " AND det_egreso.prd_codigo='" & strProducto & "' " & _
                 " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_egreso.emp_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo "
        strSQL = strSQL & " UNION " & _
                 " SELECT MAX(det_cam_cantidad) " & _
                 " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo " & _
                 " AND cambio.cam_codigo=det_cambio.cam_codigo " & _
                 " INNER JOIN producto ON det_cambio.emp_codigo=producto.emp_codigo " & _
                 " AND det_cambio.prd_codigo_ped=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE cambio.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_cambio.tip_ing_codigo = 'ICA' " & _
                 " AND det_cambio.prd_codigo_ped='" & strProducto & "' " & _
                 " AND cambio.cam_factura = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_cambio.emp_codigo,det_cambio.cam_codigo,det_cambio.prd_codigo_ped " & _
                 " ORDER BY cantidad DESC "
    End If
    clsCon_Def.Ejecutar strSQL
    ConsutaCantidadMaxima = FormatoD0(clsCon_Def.adorec_Def("cantidad"))
End Function

Private Sub CalculaTotal()
    Dim i As Long
    Dim totalIng As Double
    Dim CantIng As Double
    totalIng = 0
    CantIng = 0
    
    For i = 1 To vsfgDetalleIng.Rows - 1
        If Left(vsfgDetalleIng.TextMatrix(i, 5), 8) <> Left(vsfgDetalleIng.TextMatrix(i, 8), 8) And vsfgDetalleIng.TextMatrix(i, 6) <> "" Then
            MsgBox "no coinciden los precios del producto de la fila " & i, vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(i, 6) = ""
            vsfgDetalleIng.TextMatrix(i, 7) = ""
            vsfgDetalleIng.TextMatrix(i, 8) = ""
        End If
        CantIng = CantIng + FormatoD4(vsfgDetalleIng.TextMatrix(i, 4))
    Next i
    
    txtCantIng.Text = FormatoD4(CantIng)
    
End Sub
