VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmV_ReImprimirFacPed 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReImprimir Facturas y Pedidos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmV_ReImprimirFacPed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8790
   Begin VB.TextBox txtPag 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6480
      TabIndex        =   22
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CheckBox chkFacturaTicket 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Fac.Ticket"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Todos"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.OptionButton optNotaEntregaSuministro 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Nota Entrega Sum."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   480
      Width           =   2175
   End
   Begin VB.OptionButton optFactura 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Facturas"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnvioCorreo 
      Caption         =   "Envio Correo"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdImpFacturas 
      Caption         =   "Facturas"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdImpPedidos 
      Caption         =   "Pedidos"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CheckBox chkFechas 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "ACTUALIZAR"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Top             =   120
      Width           =   5280
      _ExtentX        =   9313
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   8595
      _cx             =   15161
      _cy             =   8493
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmV_ReImprimirFacPed.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
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
   Begin MSComCtl2.DTPicker dtpFechaFactura 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   66256899
      CurrentDate     =   37463
   End
   Begin MSComCtl2.DTPicker Fecha1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   6
      Top             =   855
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   66256899
      CurrentDate     =   37463
   End
   Begin MSComCtl2.DTPicker Fecha2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   8
      Top             =   855
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   66256899
      CurrentDate     =   37463
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   1320
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
   Begin VB.CheckBox chkConGuia 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Con Guia"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   6660
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pag:"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   1357
      Width           =   315
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pedidos:"
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
      TabIndex        =   12
      Top             =   6682
      Width           =   2205
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
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
      Left            =   4560
      TabIndex        =   10
      Top             =   855
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   855
      Width           =   510
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
      Left            =   360
      TabIndex        =   2
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmV_ReImprimirFacPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################
'#  Forma para ver un pedido ya confirmado de bodega que está listo para ser
'#  facturado.
'#  frmV_VerPedConfirm V1.0
'#  Copyright (C) 2002
'#
'#  Opciones que permite:
'#  *   En una lista se despliegan los pedidos confirmados con sus detalles de
'#      cabecera como el cliente y el vendedor que lo atiende y el estado del
'#      mismo.
'#  *   De igual manera es necesario seleccionar el tipo de facturación que se
'#      va a aplicar al pedido.
'#  *   Es necesario también seleccionar la forma de pago.
'#  *   El usuario puede seleccionar los posibles recargos que puede generar
'#      la facturación de un pedido.
'#
'#  Procesos internos que maneja:
'#  *   La lista que muestra los distintos pedidos, se refresca automáticamente
'#      cada 20 segundos para buscar un nuevo pedido confirmado.
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los
'#      detalles del mismo en un segundo grid.
'#  *   Una vez que el pedido ha sido facturado su estado pasa a vendido.
'#  *   Se pueden ver solo los pedidos que están confirmados y los que ya
'#      se han vendido el día de hoy.
'#  *   Una vez que se va a facturar el pedido se generan automáticamente las
'#      respectivas retenciones que puede tener un cliente.
'#
'#  Tablas que maneja:
'#
'#  persona:
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el
'#      pedido que se está confirmando.
'#  *   También se extrae el nombre del vendedor asignado al pedido.
'#  pedido:
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.
'#  det_pedido:
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.
'#  persona_ret:
'#  *   De esta tala se extraen las diferentes retenciones que puede tener un
'#      cliente determinado para luego aplicarlas a esta factura.
'#  retencion:
'#  *   De aquí se extraen los valores y descripciones de las retenciones, que
'#      se aplicarán posteriormente.
'#  existencia:
'#  *   En esta tabla se actualizan las cantidades existentes de los productos
'#      vendidos.
'#  det_egreso:
'#  *   En esta tabla se guardan los detalles del nuevo documento de egreso de
'#      productos.
'#  ocargo:
'#  *   De esta tabla se extraen los diferentes recargos que se puede manejar
'#      al realizar un nuevo egreso de productos de bodega, como pueden ser:
'#      transporte, fletes, etc.
'#  det_egreso_c:
'#  *   En esta tabla se guardan los diferentes recargos que puede tener esta
'#      nueva compra o egreso de productos.
'#  det_egreso_ret:
'#  *   En esta tabla se guardan los valores de las retenciones aplicadas a este
'#      ingreso de productos a bodega.
'#
'################################################################################

Private clsPedidos As New clsConsulta
Private strSql As String
Private strListaFactura As String
Private strListaPedido As String
Private lngNFacNPed As Long
Private strTipoDoc As String

Private Sub chkCIRUC_Click()
    cmbNegocio_Change
End Sub

Private Sub cmdLimpiar_Click()
    'Limpia el contenido del grid de detalles
    VSFG.Clear 1
    VSFG.Rows = 2
End Sub


Private Sub chkFechas_Click()
    If chkFechas.Value = 1 Then
        Fecha1.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha1.Enabled = False
        Fecha2.Enabled = False
    End If
End Sub

Private Sub chkTodos_Click()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = chkTodos.Value
    Next i
End Sub

Private Sub cmbNegocio_Change()
    Dim strFiltro As String
    Dim clsTipoPed As New clsConsulta
    clsTipoPed.Inicializar AdoConn, AdoConnMaster
    If optFactura.Value = True Then
        strTipoDoc = "FAC"
    Else
        strTipoDoc = "NET"
    End If
    cmdLimpiar_Click
    
    clsTipoPed.Ejecutar " SELECT tip_ped_facturaticket FROM tipo_pedido WHERE emp_codigo='" & strEmpresa & "' AND tip_ped_codigo='" & Me.cmbNegocio.BoundText & "'"
    
    chkFacturaTicket.Value = clsTipoPed.adorec_Def("tip_ped_facturaticket")
    
    strFiltro = ""
    If chkFechas.Value = 1 Then
        strFiltro = " AND egr_fechamod BETWEEN '" & Format(Fecha1.Value, "yyyy-mm-dd hh:mm:ss") & "' AND '" & Format(Fecha2.Value, "yyyy-mm-dd hh:mm:ss") & "' "
    End If
    If cmbNegocio.BoundText <> "" Then
            'Consulta todos los pedidos que pasan a bodega para ser revisados
            strSql = " SELECT '" & chkTodos.Value & "' as sel, ped_codigo, egr_codigo,LEFT(egr_fecha,10) as fecha,egr_fechamod,egr_usumod, " & _
                     " CONCAT(COALESCE(persona.per_apellido,''),' ',COALESCE(persona.per_nombre,'')) as cli,persona.per_email, " & _
                     " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,N9.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,N8.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,N7.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,N6.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,N5.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,N4.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,N3.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,N2.per_email," & _
                     " IIF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,N1.per_email,''))))))))) as emailpapa, " & _
                     " IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),persona.per_celular,''),egr_total,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1" & _
                     " FROM egreso INNER JOIN pedido ON egreso.emp_codigo=pedido.emp_codigo " & _
                     " AND egreso.tip_egr_codigo=pedido.ped_tip_egr_codigo " & _
                     " AND egreso.egr_codigo=pedido.ped_egr_codigo " & _
                     " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
                     " AND egreso.per_codigo=persona.per_codigo " & _
                     " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'"
            strSql = strSql & _
                     " LEFT JOIN persona as N1 ON N1.emp_codigo=persona.emp_codigo " & _
                     " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
                     " LEFT JOIN persona as N2 ON N2.emp_codigo=persona.emp_codigo " & _
                     " AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1 " & _
                     " LEFT JOIN persona as N3 ON persona.emp_codigo = N3.emp_codigo " & _
                     " AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1 " & _
                     " LEFT JOIN persona as N4 ON persona.emp_codigo = N4.emp_codigo " & _
                     " AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1 " & _
                     " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                     " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                     " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                     " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                     " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                     " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                     " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                     " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                     " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                     " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                     " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egr_anulado=0 " & _
                     " AND egreso.tip_egr_codigo='" & strTipoDoc & "' " & _
                     " AND egreso.egr_fecha='" & Format(dtpFechaFactura.Value, "yyyy-mm-dd") & "' " & _
                     strFiltro & _
                     " ORDER BY egr_codigo "
            clsPedidos.Ejecutar (strSql)
            
    Else
        Exit Sub
    End If
'    clsPedidos.Actualizar
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFG.DataSource = clsPedidos.adorec_Def.DataSource
    lblTotal.Caption = "Total pedidos: " & (VSFG.Rows - 1)
End Sub

Public Sub cmdEnvioCorreo_Click()
    Dim i As Long
    Dim egr As String
    Dim emailFactura As String
    Dim emailPapaFactura As String
    Dim ClienteFactura As String
    Dim RepFactura As New frmReporte
    Dim LiderListaCliente As String
    Dim ListaCliente As String
    Dim egrTot As Double

    For i = 1 To VSFG.Rows - 1
        Me.Caption = "ReImprimir Facturas y Pedidos - " & i & "/" & (VSFG.Rows - 1)
        VSFG.Select i, 0
        VSFG.ShowCell i, 0
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            egr = VSFG.TextMatrix(i, 2)
            ClienteFactura = VSFG.TextMatrix(i, 6)
            emailFactura = VSFG.TextMatrix(i, 7)
            emailPapaFactura = VSFG.TextMatrix(i, 8)
            egrTot = VSFG.TextMatrix(i, 10)
            If Trim(emailFactura) & Trim(emailPapaFactura) <> "" Then
                RepFactura.strNumero = egr
                'listo
                GuiaAutomatica = False
                RepFactura.strReporte = IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola")
                RepFactura.Show
                RepFactura.Form_Activate
                RepFactura.VSRpt.RenderToFile "Factura" & egr & ".pdf", vsrPDF
                'Unload RepFactura
                EnviarMail NombreComercial & " Facturacion", CorreoServicioAlCliente, ClienteFactura, Trim(emailFactura) & "; " & Trim(emailPapaFactura), "", "Factura " & egr, _
                        "Estimad@" & vbNewLine & _
                        ClienteFactura & vbNewLine & _
                        "Adjunto encontrarás tu factura emitida el " & Format(VSFG.TextMatrix(i, 3), "yyyy-mm-dd") & "." & vbNewLine & _
                        "Recuerda que es una factura electrónica y la puedes descargar en nuestro sitio web " & _
                        "www.rbimportadores.com . " & vbNewLine & vbNewLine & _
                        "Saludos Cordiales" & vbNewLine & _
                        "Facturación" & vbNewLine & _
                        NombreComercial, "Factura" & egr & ".pdf"
                Kill "Factura" & egr & ".pdf"
                If Trim(emailFactura) = "" Then
                    ListaCliente = ListaCliente & ClienteFactura & vbNewLine
'                    EnviarMail NombreComercial & " Facturacion", CorreoServicioAlCliente, ClienteFactura, CorreoAsistenteCos, "", "Factura " & egr & " Cliente sin Email", _
'                            "Estimad@. El cliente " & vbNewLine & _
'                            ClienteFactura & vbNewLine & _
'                            "No esta registrado un correo electronico" & vbNewLine & _
'                            "Saludos Cordiales" & vbNewLine & _
'                            "Facturación" & vbNewLine & _
'                            NombreComercial
                ElseIf Trim(emailPapaFactura) = "" Then
                    LiderListaCliente = LiderListaCliente & ClienteFactura & vbNewLine
'                    EnviarMail NombreComercial & " Facturacion", CorreoServicioAlCliente, ClienteFactura, CorreoAsistenteCos, "", "Factura " & egr & " Lider sin Email", _
'                            "Estimad@. El lider inmediato del cliente " & vbNewLine & _
'                            ClienteFactura & vbNewLine & _
'                            "No esta registrado un correo electronico" & vbNewLine & _
'                            "Saludos Cordiales" & vbNewLine & _
'                            "Facturación" & vbNewLine & _
'                            NombreComercial
                End If

            Else
                ListaCliente = ListaCliente & ClienteFactura & vbNewLine
                LiderListaCliente = LiderListaCliente & ClienteFactura & vbNewLine
'                EnviarMail NombreComercial & " Facturacion", CorreoServicioAlCliente, ClienteFactura, CorreoAsistenteCos, "", "Factura " & egr & " Cliente y Lider sin Email", _
'                        "Estimad@. El cliente y el lider inmediato del cliente " & vbNewLine & _
'                        ClienteFactura & vbNewLine & _
'                        "No tienen registrado correo electronico" & vbNewLine & _
'                        "Saludos Cordiales" & vbNewLine & _
'                        "Facturación" & vbNewLine & _
'                        NombreComercial

            End If
        End If
    Next i
    EnviarMail NombreComercial & " Facturacion", CorreoServicioAlCliente, "Asistente COS", CorreoAsistenteCos & ";" & CorreoServicioAlCliente, "", "Clientes y Lideres sin Email", _
            "Estimad@. El siguiente listado son los clientes que no tienen email registrado: " & vbNewLine & _
            ListaCliente & vbNewLine & vbNewLine & _
            "El siguiente listado son los clientes que su lider inmediato no tienen email registrado:" & vbNewLine & _
            LiderListaCliente & vbNewLine & vbNewLine & _
            "Saludos Cordiales" & vbNewLine & _
            "Facturación" & vbNewLine & _
            NombreComercial

    Unload RepFactura
    MsgBox "Envio de correos Finalizado"
End Sub

Private Sub cmdImpFacturas_Click()
    Dim i As Long
    Dim RepFactura As New frmReporte
    
    Dim GuiaAutomatica As Boolean
    
    strListaFactura = ""
    lngNFacNPed = 0
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            strListaFactura = strListaFactura & VSFG.TextMatrix(i, 2) & ","
            lngNFacNPed = lngNFacNPed + 1
        End If
    Next i
    strListaFactura = Left(strListaFactura, Len(strListaFactura) - 1)
    
    If Me.chkFacturaTicket.Value = 0 Then
    
        RepFactura.strNumero = strListaFactura
        
        'listo
        GuiaAutomatica = IIf(chkConGuia.Value = 1, True, False)
        RepFactura.strReporte = IIf(strTipoDoc = "FAC", IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola"), "rptNotaEntregaSuministro")
        RepFactura.VSPrint.Copies = 1
        RepFactura.VSPrint.Collate = colFalse
        RepFactura.Show
    ElseIf Me.chkFacturaTicket.Value = 1 Then
        frmImpresionDirecta.strNumero = strListaFactura
        frmImpresionDirecta.lngPag = FormatoD0(txtPag.Text)
        frmImpresionDirecta.strReporte = "rptPedido"
        frmImpresionDirecta.Show
        frmImpresionDirecta.optImpresora.Value = True
        'frmImpresionDirecta.Form_Activate
        'frmImpresionDirecta.cmdImprimir_Click
    End If

End Sub

Private Sub cmdImpPedidos_Click()
    Dim i As Long
    Dim clsBloq As New clsConsulta
    Dim RepStk As New frmReporte
    strListaPedido = ""
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            strListaPedido = strListaPedido & VSFG.TextMatrix(i, 1) & ","
        End If
    Next i
    strListaPedido = Left(strListaPedido, Len(strListaPedido) - 1)
    
    If ImpresoraEtiqueta = "" Then
        RepStk.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = RepStk.VSPrint.Device
        GuardarImpresoras
    End If
    RepStk.VSPrint.Device = ImpresoraEtiqueta
    
    RepStk.VSPrint.PaperWidth = 7669.292
    RepStk.VSPrint.PaperHeight = 3885.039
    RepStk.strNumero = strListaPedido
    RepStk.strReporte = "rptSTKDespacho"
    RepStk.strTipo = 5
    RepStk.VSPrint.Copies = 1
    RepStk.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidos = Nothing
    Set clsPed = Nothing
    Set clsSql = Nothing
    Set clsTFac = Nothing
    Set clsRecargos = Nothing
    Set clsFPago = Nothing
    Set clsFormaPago = Nothing
    Set clsTC = Nothing
    Set clsRet = Nothing
    Set clsExis = Nothing
End Sub

Public Sub cmdActualizar_Click()
    cmbNegocio_Change
    
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtPedido" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsPedidos.Ejecutar strSql
    If clsPedidos.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsPedidos.adorec_Def(0)
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa los objetos de conexión con la base de datos
    clsPedidos.Inicializar AdoConn, AdoConnMaster
        
    cargarTipoPedido
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False, True, False, True, False, False
    'Coloca la fecha actual
    dtpFechaFactura.Value = HoyDia
    Fecha1.Value = HoyDia
    Fecha2.Value = HoyDia
    'cmbNegocio_Change
        
End Sub

Private Sub optFactura_Click()
    If optFactura.Value = True Then
        cmdImpFacturas.Caption = "Facturas"
        lblFecha.Caption = "Fecha Factura"
        VSFG.TextMatrix(0, 2) = "Facturas"
        VSFG.TextMatrix(0, 3) = "Fecha Factura"
        VSFG.TextMatrix(0, 4) = "Fecha Facturacion"
        cmbNegocio_Change
    End If
End Sub

Private Sub optNotaEntregaSuministro_Click()
    If optNotaEntregaSuministro.Value = True Then
        cmdImpFacturas.Caption = "NotaEntrega"
        lblFecha.Caption = "Fecha Nota"
        VSFG.TextMatrix(0, 2) = "Notas"
        VSFG.TextMatrix(0, 3) = "Fecha Nota"
        VSFG.TextMatrix(0, 4) = "Fecha Emision"
        cmbNegocio_Change
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

