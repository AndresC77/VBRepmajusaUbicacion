VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCarteraPedidos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Pendientes de pago"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmCarteraPedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdEnvioCorreo 
      Caption         =   "Envio SMS"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
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
      Top             =   120
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
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
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8595
      _cx             =   15161
      _cy             =   10186
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCarteraPedidos.frx":030A
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
      TabIndex        =   5
      Top             =   6600
      Width           =   2325
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
Attribute VB_Name = "frmCarteraPedidos"
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

Private Sub cmbNegocio_Change()
    Dim strFiltro As String
    
    cmdLimpiar_Click
    
    strFiltro = ""
    If cmbNegocio.BoundText <> "" Then
            'Consulta todos los pedidos que pasan a bodega para ser revisados
        strSql = " SELECT '1'as sel, pedido.ped_codigo,ped_fechamod,CONCAT(per_apellido, ' ',per_nombre) as cli," & _
                 " per_celular, " & _
                 "  " & _
                 "" & _
                 " ROUND(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00),2) " & _
                 "+ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00) * (par_numero)/100.00,2) " & _
                 "-COALESCE(doc_pag_valor,0.00),2) as d, " & _
                 " 1 as tipoM "
        strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                 " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                 " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                 " FROM doc_pago_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                 " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                 " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                 " AND pedido.per_codigo=pag.per_codigo " & _
                 " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
                 " AND (ped_celular='' or ped_celular is null) " & _
                 " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
                 " GROUP BY pedido.ped_codigo,ped_fechamod,per_apellido,per_nombre, per_celular,par_numero,doc_pag_valor,pedido.ped_dctoadicional "
        If Right(Ahora, 8) > "10:00:00" Then
            strSql = strSql & " UNION " & _
                     " SELECT '1'as sel, pedido.ped_codigo,ped_fechamod,CONCAT(per_apellido, ' ',per_nombre) as cli," & _
                     " per_celular, " & _
                     " " & _
                     " " & _
                     " ROUND(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00),2) " & _
                     "+ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00) * (par_numero)/100.00,2) " & _
                     "-COALESCE(doc_pag_valor,0.00),2) as d, " & _
                     " 2 as tipoM "
            strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                     " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                     " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                     " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                     " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                     " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                     " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                     " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                     " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                     " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                     " FROM doc_pago_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                     " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                     " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                     " AND pedido.per_codigo=pag.per_codigo " & _
                     " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
                     " AND (ped_celular!='' and ped_celular is not null) " & _
                     " AND (ped_celular2='' or ped_celular2 is null) " & _
                     " AND ped_fechamod<= DATEADD(d,1,CURRENT_TIMESTAMP)" & _
                     " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
                     " GROUP BY pedido.ped_codigo,ped_fechamod,per_apellido,per_nombre, per_celular,par_numero,doc_pag_valor,pedido.ped_dctoadicional "
        End If
        If Right(Ahora, 8) < "10:00:00" Then
            strSql = strSql & " UNION " & _
                     " SELECT '1' as sel, pedido.ped_codigo,ped_fechamod,CONCAT(per_apellido, ' ',per_nombre) as cli," & _
                     " per_celular, " & _
                     "  " & _
                     "" & _
                     " ROUND(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00),2) " & _
                     "+ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00) * (par_numero)/100.00,2) " & _
                     "-COALESCE(doc_pag_valor,0.00),2) as d, " & _
                     " 3 as tipoM "
            strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                     " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                     " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                     " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                     " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                     " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                     " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                     " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                     " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                     " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                     " FROM doc_pago_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                     " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                     " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                     " AND pedido.per_codigo=pag.per_codigo " & _
                     " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
                     " AND (ped_celular!='' and ped_celular is not null) " & _
                     " AND (ped_celular2!='' and ped_celular2 is not null) " & _
                     " AND (ped_celular3='' or ped_celular3 is null) " & _
                     " AND ped_fechamod<= DATEADD(d,2,CURRENT_TIMESTAMP)" & _
                     " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
                     " GROUP BY pedido.ped_codigo,ped_fechamod,per_apellido,per_nombre, per_celular,par_numero,doc_pag_valor,pedido.ped_dctoadicional "
        End If
        clsPedidos.Ejecutar strSql
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
    Dim egrTot As Double
    Dim clsProveeSMS As New clsConsulta
    Dim SMS As New clsEnvioSMS
    clsProveeSMS.Inicializar AdoConn, AdoConnMaster
    clsProveeSMS.Ejecutar " SELECT par_texto FROM parametro WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='SMS'", "L"
    SMS.Inicializar (clsProveeSMS.adorec_Def("par_texto"))
    'SMS.Enviar "prueba", "0998023203"
    For i = 1 To VSFG.Rows - 1
        VSFG.Select i, 0
        VSFG.ShowCell i, 0
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            egr = VSFG.TextMatrix(i, 1)
            'ClienteFactura = VSFG.TextMatrix(i, 3)
            egrTot = VSFG.TextMatrix(i, 5)
            
            If Trim(VSFG.TextMatrix(i, 4)) <> "" Then
'                SMS.Enviar "Estimado ejecutivo JSN-VPC mil disculpas por el error que tuvimos en el valor " & _
'                               "total del pedido enviado en el anterior SMS." & _
'                               "Enseguida te llegara otro con el valor correcto. MUCHAS GRACIAS", Trim(VSFG.TextMatrix(i, 4))
                If VSFG.TextMatrix(i, 6) = 1 Then
                    SMS.Enviar "Estimado ejecutivo JSN-VPC el pedido " & egr & _
                               " se cargo hoy al banco por el valor de " & FormatoD2(egrTot) & _
                               ".Tu codigo para el pago en los puntos habilitados es 9" & Right(egr, 7), Trim(VSFG.TextMatrix(i, 4))
                    strSql = " UPDATE pedido " & _
                             " SET ped_celular='" & VSFG.TextMatrix(i, 4) & "'" & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ped_codigo='" & egr & "'"
                ElseIf VSFG.TextMatrix(i, 6) = 2 Then
                    SMS.Enviar "Estimado ejecutivo JSN-VPC te recordamos que tienes el pedido " & egr & _
                               " cargado en el banco por el valor de " & FormatoD2(egrTot) & _
                               ".Tu codigo para el pago es 9" & Right(egr, 7), Trim(VSFG.TextMatrix(i, 4))
                    strSql = " UPDATE pedido " & _
                             " SET ped_celular2='" & VSFG.TextMatrix(i, 4) & "'" & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ped_codigo='" & egr & "'"
                ElseIf VSFG.TextMatrix(i, 6) = 3 Then
                    SMS.Enviar "Estimado ejecutivo JSN-VPC tienes un pedido cargado en el banco por " & FormatoD2(egrTot) & _
                               " que esta proximo a ser anulado por no pago." & _
                               "El codigo para el pago es 9" & Right(egr, 7), Trim(VSFG.TextMatrix(i, 4))
                    strSql = " UPDATE pedido " & _
                             " SET ped_celular3='" & VSFG.TextMatrix(i, 4) & "'" & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ped_codigo='" & egr & "'"
                End If
                clsPedidos.Ejecutar strSql, "M"
            End If
        'Else
            'MsgBox "AAA"
        End If
    Next i
    Set SMS = Nothing
    MsgBox "Envio de SMS Finalizado"
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

Public Sub cmdcancelar_Click()
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
    
    'Coloca la fecha actual
    cmbNegocio_Change
        
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
