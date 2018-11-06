VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCuadreAsientos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revision de Asientos Descuadrados"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "frmCuadreAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7485
   Begin VB.CommandButton cmdCobrosDescuadrados 
      Caption         =   "Cobros Descuadrados"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdNotaCreditoProDescuadrada 
      Caption         =   "Notas de Credito Prov Descuadradas"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton cmdCompraDescuadrada 
      Caption         =   "Compras Descuadradas"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdNotaCreditoCliDescuadrada 
      Caption         =   "Notas de Credito Cli Descuadradas"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdNotaCredCliNoGenerada 
      Caption         =   "Notas de Credito Cli No Generadas"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton cmdFacturaDescuadrada 
      Caption         =   "Facturas Descuadradas"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton cmdFacturaNoGenerada 
      Caption         =   "Facturas No generadas"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   2892
      TabIndex        =   0
      Top             =   4080
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3240
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   3540
      _cx             =   28252260
      _cy             =   28251731
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCuadreAsientos.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Value           =   42039.720462963
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha desde:"
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
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmCuadreAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String

Private Sub cmdCobrosDescuadrados_Click()
    Dim clsSql As New clsConsulta
    Dim strFecha As String
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    
    strSql = " UPDATE doc_pago, det_asiento" & _
             " SET det_asiento.cta_codigo='1.1.01.01.001' " & _
             " WHERE doc_pago.emp_codigo = det_asiento.emp_codigo " & _
             " and doc_pago.asi_numasiento=det_asiento.asi_numasiento " & _
             " and (doc_pago.tip_doc_pag_codigo='ch' or doc_pago.tip_doc_pag_codigo='chc' or doc_pago.tip_doc_pag_codigo='efc' or doc_pago.tip_doc_pag_codigo='' or doc_pago.tip_doc_pag_codigo is null) " & _
             " and cta_codigo!='1.1.01.01.001' " & _
             " and det_asi_debe!=0 " & _
             " and doc_pag_fecha_recepcion>='" & strFecha & "' "
    clsSql.Ejecutar strSql, "LOCAL"
    
    strSql = " UPDATE doc_pago, det_asiento " & _
             " SET det_asiento.cta_codigo='1.1.01.01.003' " & _
             " WHERE doc_pago.emp_codigo = det_asiento.emp_codigo " & _
             " and doc_pago.asi_numasiento=det_asiento.asi_numasiento " & _
             " and (doc_pago.tip_doc_pag_codigo='vou' or doc_pago.tip_doc_pag_codigo='voc') " & _
             " and cta_codigo!='1.1.01.01.003' " & _
             " and det_asi_debe!=0 " & _
             "and doc_pag_fecha_recepcion>='" & strFecha & "' "
    clsSql.Ejecutar strSql, "LOCAL"

    strSql = " UPDATE doc_pago, det_asiento " & _
             " SET det_asiento.cta_codigo='1.1.01.01.002'" & _
             " WHERE doc_pago.emp_codigo = det_asiento.emp_codigo " & _
             " and doc_pago.asi_numasiento2=det_asiento.asi_numasiento " & _
             " and (doc_pago.tip_doc_pag_codigo='chp' or doc_pago.tip_doc_pag_codigo='pfc') " & _
             " and cta_codigo!='1.1.01.01.002' " & _
             " and det_asi_debe!=0 " & _
             "and doc_pag_fecha_recepcion>='" & strFecha & "' "
    clsSql.Ejecutar strSql, "LOCAL"
    
    MsgBox "Revision terminada"
    
End Sub

Private Sub cmdCompraDescuadrada_Click()
    chequeoAsiCOM
End Sub

Private Sub chequeoAsiCOM()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT ing_codigo,ingreso.emp_codigo,IIF(ing_anulado=1,0,ing_total) as ing_total,asiento.asi_numasiento,sum(COALESCE(det_asi_debe,0)) as d,sum(COALESCE(det_asi_haber,0)) as h, abs(round(sum(COALESCE(det_asi_debe,0)),2)-round(sum(COALESCE(det_asi_haber,0)),2)) as dif " & _
             " FROM ingreso inner join asiento " & _
             " ON ingreso.emp_codigo=asiento.emp_codigo" & _
             " AND ingreso.ing_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%COMPRA LOCAL%',ing_codigo,'%') " & _
             " LEFT JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo " & _
             " AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
             " WHERE tip_ing_codigo='COM' " & _
             " AND ing_fecha>='" & strFecha & "' AND ingreso.emp_codigo='" & strEmpresa & "' and ing_anulado=0 " & _
             " GROUP BY ingreso.ing_codigo,ingreso.emp_codigo,asiento.asi_numasiento,ing_total,ing_anulado " & _
             " HAVING ROUND(sum(COALESCE(det_asi_debe,0)), 2) <> ROUND(sum(COALESCE(det_asi_haber,0)), 2) OR ROUND(sum(COALESCE(det_asi_haber,0)), 2)<> ing_total" & _
             " ORDER BY dif DESC "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("ing_codigo")
        CargaAsientoCOM clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("ing_codigo"), clsSql.adorec_Def("asi_numasiento")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsientoCOM(Emp As String, COM As String, Asi As String)
        Dim PerCodigo As String
        Dim COMTotal As Double
        Dim COMIVA As Double
        Dim COMSubTotal As Double
        Dim COMDcto As Double
        Dim COMSubTotalP As Double
        Dim COMSubTotalS As Double
        Dim PerSinIVA As Boolean
        Dim PerSecPub As Boolean
        Dim clsAuxAsi As New clsConsulta
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXP
        strSql = " SELECT ingreso.per_codigo,IIF(ing_anulado=1,0,ing_total) as ing_total," & _
                 " IIF(ing_anulado=1,0,ing_subtotal) as ing_subtotal,IIF(ing_anulado=1,0,ing_dcto) as ing_dcto," & _
                 " IIF(ing_anulado=1,0,ing_impuesto) as ing_impuesto,per_siniva,per_sec_publico," & _
                 " IIF(ing_anulado=1,0,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0))) as totprod," & _
                 " IIF(ing_anulado=1,0,SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0))) as totserv " & _
                 " FROM ingreso INNER JOIN persona " & _
                 " ON ingreso.emp_codigo=persona.emp_codigo " & _
                 " AND ingreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN det_ingreso " & _
                 " ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                 " AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                 " WHERE ingreso.emp_codigo = '" & Emp & "' " & _
                 " AND ingreso.tip_ing_codigo='COM'" & _
                 " AND ingreso.ing_codigo='" & COM & "' " & _
                 " GROUP BY ingreso.ing_codigo,ingreso.emp_codigo,ingreso.tip_ing_codigo,ingreso.per_codigo,ing_anulado,ing_total,ing_subtotal,ing_dcto,ing_impuesto,per_siniva,per_sec_publico"
        clsAuxAsi.Ejecutar (strSql)
        PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
        COMTotal = clsAuxAsi.adorec_Def("ing_total")
        COMIVA = clsAuxAsi.adorec_Def("ing_impuesto")
        COMSubTotal = clsAuxAsi.adorec_Def("ing_subtotal")
        COMDcto = clsAuxAsi.adorec_Def("ing_dcto")
        COMSubTotalP = clsAuxAsi.adorec_Def("totprod")
        COMSubTotalS = clsAuxAsi.adorec_Def("totserv")
        
        If COMSubTotalP <> 0 And COMSubTotalS = 0 Then
            COMSubTotalP = COMSubTotal - COMDcto
        ElseIf COMSubTotalS <> 0 And COMSubTotalP = 0 Then
            COMSubTotalS = COMSubTotal - COMDcto
        ElseIf COMSubTotalS = 0 And COMSubTotalP = 0 Then
            'FacSubTotalS = FacSubTotal
        Else
            COMSubTotalS = COMSubTotal - COMDcto - COMSubTotalP
        End If
        PerSinIVA = False
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        clsAsi.NumAsiento = Asi
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        'cuenta contable IVA COMPRAS
        If FormatoD2(COMIVA) <> 0 Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='IVAC' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", FormatoD2(COMIVA), 0
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(COMSubTotalP) <> 0 Then
            strSql = " SELECT tip_ing_ctaconta " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_ing_codigo='COM' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta"), "", FormatoD2(COMSubTotalP), 0
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(COMSubTotalS) <> 0 Then
            strSql = " SELECT tip_ing_ctaconta2 " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_ing_codigo='COM' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta2"), "", FormatoD2(COMSubTotalS), 0
        End If
        'cuentas contables de RECARGOS
        strSql = " SELECT oca_ctaconta, det_ing_c_cantidad*det_ing_c_precio as Tot " & _
                 " FROM det_ingreso_c INNER JOIN ocargos ON det_ingreso_c.emp_codigo=ocargos.emp_codigo " & _
                 " AND det_ingreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_ingreso_c.emp_codigo='" & Emp & "' " & _
                 " AND det_ingreso_c.tip_ing_codigo='COM' " & _
                 " AND det_ingreso_c.ing_codigo='" & COM & "' "
        clsAuxAsi.Ejecutar strSql, "M"
        While Not clsAuxAsi.adorec_Def.EOF
            If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), "", FormatoD2(clsAuxAsi.adorec_Def("Tot")), 0
            End If
            clsAuxAsi.adorec_Def.MoveNext
        Wend
        'cuentas contables de RETENCIONES
        Dim COMRet As Double
        strSql = " SELECT ret_ctaconta,det_ing_ret_valor " & _
                 " FROM det_ingreso_ret INNER JOIN retencion ON retencion.emp_codigo=det_ingreso_ret.emp_codigo " & _
                 " AND retencion.ret_codigo=det_ingreso_ret.ret_codigo " & _
                 " WHERE det_ingreso_ret.emp_codigo='" & Emp & "' " & _
                 " AND det_ingreso_ret.tip_ing_codigo='COM' " & _
                 " AND det_ingreso_ret.ing_codigo='" & COM & "' "
        clsAuxAsi.Ejecutar strSql, "M"
        While Not clsAuxAsi.adorec_Def.EOF
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("ret_ctaconta"), "", 0, clsAuxAsi.adorec_Def("det_ing_ret_valor")
            COMRet = COMRet + FormatoD2(clsAuxAsi.adorec_Def("det_ing_ret_valor"))
            clsAuxAsi.adorec_Def.MoveNext
        Wend
        'cuenta contable CXP
        strSql = " SELECT cat_p_ctaconta " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
        clsAuxAsi.Ejecutar strSql
        clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", 0, FormatoD2(COMTotal) - FormatoD2(COMRet)
End Sub



Private Sub cmdFacturaDescuadrada_Click()
    chequeoAsiFact
End Sub

Private Sub chequeoAsiFact()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT egr_codigo,egreso.emp_codigo,IIF(egr_anulado=1,0,egr_total) as et,asiento.asi_numasiento," & _
             " sum(COALESCE(det_asi_debe,0)) as d,sum(COALESCE(det_asi_haber,0)) as h, " & _
             " abs(round(sum(COALESCE(det_asi_debe,0)),2)-round(sum(COALESCE(det_asi_haber,0)),2)) as dif " & _
             " FROM egreso inner join asiento " & _
             " ON egreso.emp_codigo=asiento.emp_codigo" & _
             " AND egreso.egr_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%FACTURA%',egr_codigo,'%') " & _
             " LEFT JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo " & _
             " AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
             " WHERE tip_egr_codigo='FAC' " & _
             " AND egr_fecha>='" & strFecha & "' AND egreso.emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY egr_codigo,egreso.emp_codigo,egr_anulado,egr_total,asiento.asi_numasiento " & _
             " HAVING ROUND(sum(COALESCE(det_asi_debe,0)), 2) <> ROUND(sum(COALESCE(det_asi_haber,0)), 2) OR ROUND(sum(COALESCE(det_asi_haber,0)), 2)<>IIF(egr_anulado=1,0,egr_total)" & _
             " ORDER BY dif DESC "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("egr_codigo")
        CargaAsiento clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("egr_codigo"), clsSql.adorec_Def("asi_numasiento")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsiento(Emp As String, Fac As String, Asi As String)
    Dim PerCodigo As String
    Dim FacTotal As Double
    Dim FacIVA As Double
    Dim FacSubTotal As Double
    Dim FacSubTotalCERO As Double
    Dim FacDcto As Double
    Dim FacSubTotalP As Double
    Dim FacSubTotalS As Double
    Dim PerSinIVA As Boolean
    Dim PerSecPub As Boolean
    Dim clsAuxAsi As New clsConsulta
    Dim strCentroCosto As String
On Error GoTo errhandler
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXC
        strSql = " SELECT egreso.per_codigo,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva," & _
                 " per_sec_publico,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totprod," & _
                 " SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totserv " & _
                 " FROM egreso INNER JOIN persona " & _
                 " ON egreso.emp_codigo=persona.emp_codigo " & _
                 " AND egreso.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN det_egreso " & _
                 " ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                 " WHERE egreso.emp_codigo = '" & Emp & "' " & _
                 " AND egreso.tip_egr_codigo='FAC'" & _
                 " AND egreso.egr_codigo='" & Fac & "' " & _
                 " GROUP BY egreso.per_codigo,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva,per_sec_publico"
        clsAuxAsi.Ejecutar (strSql)
        PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
        FacTotal = clsAuxAsi.adorec_Def("egr_total")
        FacIVA = clsAuxAsi.adorec_Def("egr_impuesto")
        FacSubTotal = clsAuxAsi.adorec_Def("egr_subtotal")
        FacDcto = clsAuxAsi.adorec_Def("egr_dcto")
        FacSubTotalP = FormatoD2(clsAuxAsi.adorec_Def("totprod"))
        FacSubTotalS = FormatoD2(clsAuxAsi.adorec_Def("totserv"))
        If FacSubTotalP <> 0 And FacSubTotalS = 0 Then
            FacSubTotalP = FacSubTotal - FacDcto
        ElseIf FacSubTotalS <> 0 And FacSubTotalP = 0 Then
            FacSubTotalS = FacSubTotal - FacDcto
        ElseIf FacSubTotalS = 0 And FacSubTotalP = 0 Then
           'FacSubTotalS = FacSubTotal
        Else
            'FacSubTotalS = FacSubTotal - FacDcto - FacSubTotalP
        End If
        PerSinIVA = False
        If clsAuxAsi.adorec_Def("per_siniva") = 1 Then
            PerSinIVA = True
        End If
        PerSecPub = False
        If clsAuxAsi.adorec_Def("per_sec_publico") = 1 Then
            PerSecPub = True
        End If
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        clsAsi.NumAsiento = Asi
        strSql = " UPDATE cuenta_p_c " & _
                 " SET cue_p_c_st_prod = '" & FacSubTotalP & "', " & _
                 " cue_p_c_st_serv = '" & FacSubTotalS & "' " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'" & _
                 " AND cue_p_c_egr_codigo='" & Fac & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        'cuenta contable CXC
        strCentroCosto = ""
        strSql = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
        clsAuxAsi.Ejecutar (strSql)
        clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", FormatoD2(FacTotal), 0
        strCentroCosto = clsAuxAsi.adorec_Def("cen_cos_codigo")
        'cuenta contable IVA VENTAS
        If FormatoD2(FacIVA) <> 0 And PerSinIVA = False Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='IVAV' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", 0, FormatoD2(FacIVA)
        End If
        If PerSecPub = False Then
            secPub = ""
        Else
            secPub = "_sp"
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(FacSubTotalP) <> 0 Then
            strSql = " SELECT suc.suc_ctaconta_ventas" & secPub & " as tip_egr_ctaconta" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & Emp & "' " & _
                     " AND te.tip_egr_codigo='FAC' " & _
                     " AND suc.suc_codigo in ('001','" & PtoEmiDocEle & "') "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta"), strCentroCosto, 0, FormatoD2(FacSubTotalP)
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(FacSubTotalS) <> 0 Then
            strSql = " SELECT suc.suc_ctaconta_servicios" & secPub & " as tip_egr_ctaconta2" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & Emp & "' " & _
                     " AND te.tip_egr_codigo='FAC' " & _
                     " AND suc.suc_codigo in ('001','" & PtoEmiDocEle & "') "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta2"), strCentroCosto, 0, FormatoD2(FacSubTotalS)
        End If
        'cuentas contables de recargos
        strSql = " SELECT oca_ctaconta, det_egr_c_cantidad*det_egr_c_precio as Tot " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo " & _
                 " AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_egreso_c.emp_codigo='" & Emp & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                 " AND det_egreso_c.egr_codigo='" & Fac & "' "
        clsAuxAsi.Ejecutar strSql
        While Not clsAuxAsi.adorec_Def.EOF
            If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), strCentroCosto, 0, FormatoD2(clsAuxAsi.adorec_Def("Tot"))
            End If
            clsAuxAsi.adorec_Def.MoveNext
        Wend
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description & vbNewLine & strSql
    End Select
End Sub



Private Sub cmdFacturaNoGenerada_Click()
    chequeoNoAsiFact
End Sub

Private Sub chequeoNoAsiFact()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT egr_codigo,egreso.emp_codigo " & _
             " FROM egreso left join asiento " & _
             " ON egreso.emp_codigo=asiento.emp_codigo" & _
             " AND egreso.egr_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%FACTURA%',egr_codigo,'%') " & _
             " WHERE tip_egr_codigo='FAC' " & _
             " AND egr_fecha>='" & strFecha & "' " & _
             " AND egr_codigo<999999999999999" & _
             " AND egreso.emp_codigo='" & strEmpresa & "'" & _
             " AND asiento.asi_numasiento IS NULL "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("egr_codigo")
        CargaAsiento2 clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("egr_codigo")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsiento2(Emp As String, Fac As String)
        Dim PerCodigo As String
        Dim Asi As String
        Dim FechaEgr As String
        Dim FacTotal As Double
        Dim FacIVA As Double
        Dim FacSubTotal As Double
        Dim FacDcto As Double
        Dim FacSubTotalP As Double
        Dim FacSubTotalS As Double
        Dim PerSinIVA As Boolean
        Dim PerSecPub As Boolean
        Dim clsAuxAsi As New clsConsulta
        Dim strCentroCosto As String
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXC
        strSql = " SELECT egreso.per_codigo,egr_fecha,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva," & _
                 " per_sec_publico,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totprod," & _
                 " SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totserv " & _
                 " FROM egreso INNER JOIN persona " & _
                 " ON egreso.emp_codigo=persona.emp_codigo " & _
                 " AND egreso.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN det_egreso " & _
                 " ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                 " WHERE egreso.emp_codigo = '" & Emp & "' " & _
                 " AND egreso.tip_egr_codigo='FAC'" & _
                 " AND egreso.egr_codigo='" & Fac & "' " & _
                 " GROUP BY egreso.per_codigo,egr_fecha,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva,per_sec_publico"
        clsAuxAsi.Ejecutar (strSql)
        PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
        FechaEgr = clsAuxAsi.adorec_Def("egr_fecha")
        FacTotal = clsAuxAsi.adorec_Def("egr_total")
        FacIVA = clsAuxAsi.adorec_Def("egr_impuesto")
        FacSubTotal = clsAuxAsi.adorec_Def("egr_subtotal")
        FacDcto = clsAuxAsi.adorec_Def("egr_dcto")
        FacSubTotalP = clsAuxAsi.adorec_Def("totprod")
        FacSubTotalS = clsAuxAsi.adorec_Def("totserv")
        PerSinIVA = False
        If clsAuxAsi.adorec_Def("per_siniva") = 1 Then
            PerSinIVA = True
        End If
        PerSecPub = False
        If clsAuxAsi.adorec_Def("per_sec_publico") = 1 Then
            PerSecPub = True
        End If
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        clsAsi.NuevoAsiento "F", FechaEgr, 0, 0, FacTotal, "FACTURA " & Fac, True
        Asi = clsAsi.NumAsiento
        strSql = " UPDATE egreso " & _
                 " SET egr_numasiento='" & Asi & "' " & _
                 " WHERE egreso.emp_codigo = '" & Emp & "' " & _
                 " AND egreso.tip_egr_codigo='FAC'" & _
                 " AND egreso.egr_codigo='" & Fac & "'"
        clsAuxAsi.Ejecutar (strSql)
        strSql = " UPDATE cuenta_p_c " & _
                 " SET asi_numasiento='" & Asi & "' " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & Emp & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='C'" & _
                 " AND cuenta_p_c.cue_p_c_egr_codigo='" & Fac & "' "
        clsAuxAsi.Ejecutar (strSql)
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        'cuenta contable CXC
        strCentroCosto = ""
        strSql = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
        clsAuxAsi.Ejecutar (strSql)
        clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", FormatoD2(FacTotal), 0
        strCentroCosto = clsAuxAsi.adorec_Def("cen_cos_codigo")
        'cuenta contable IVA VENTAS
        If FormatoD2(FacIVA) <> 0 And PerSinIVA = False Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='IVAV' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", 0, FormatoD2(FacIVA)
        End If
        If PerSecPub = False Then
            secPub = ""
        Else
            secPub = "_sp"
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(FacSubTotalP) <> 0 Then
            strSql = " SELECT suc.suc_ctaconta_ventas" & secPub & " as tip_egr_ctaconta" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & Emp & "' " & _
                     " AND te.tip_egr_codigo='FAC' " & _
                     " AND suc.suc_codigo in ('001','" & PtoEmiDocEle & "') "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta"), strCentroCosto, 0, FormatoD2(FacSubTotalP)
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(FacSubTotalS) <> 0 Then
            strSql = " SELECT suc.suc_ctaconta_servicios" & secPub & " as tip_egr_ctaconta2" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & Emp & "' " & _
                     " AND te.tip_egr_codigo='FAC' " & _
                     " AND suc.suc_codigo in ('001','" & PtoEmiDocEle & "') "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta2"), strCentroCosto, 0, FormatoD2(FacSubTotalS)
        End If
        'cuentas contables de recargos
        strSql = " SELECT oca_ctaconta, det_egr_c_cantidad*det_egr_c_precio as Tot " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo " & _
                 " AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_egreso_c.emp_codigo='" & Emp & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                 " AND det_egreso_c.egr_codigo='" & Fac & "' "
        clsAuxAsi.Ejecutar strSql
        While Not clsAuxAsi.adorec_Def.EOF
            If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), strCentroCosto, 0, FormatoD2(clsAuxAsi.adorec_Def("Tot"))
            End If
            clsAuxAsi.adorec_Def.MoveNext
        Wend
End Sub



Private Sub cmdNotaCredCliNoGenerada_Click()
    chequeoNoAsiNC
End Sub

Private Sub chequeoNoAsiNC()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT ing_codigo,ingreso.emp_codigo " & _
             " FROM ingreso left join asiento " & _
             " ON ingreso.emp_codigo=asiento.emp_codigo" & _
             " AND ingreso.ing_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%NOTA DE CREDITO%',ing_codigo,'%') " & _
             " WHERE tip_ing_codigo='DCL' " & _
             " AND ing_fecha>='" & strFecha & "' " & _
             " AND asiento.asi_numasiento IS NULL "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("ing_codigo")
        CargaAsientoNC2 clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("ing_codigo")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsientoNC2(Emp As String, NC As String)
        Dim PerCodigo As String
        Dim Asi As String
        Dim FechaIng As String
        Dim NCTotal As Double
        Dim NCIVA As Double
        Dim NCSubTotal As Double
        Dim NCDcto As Double
        Dim NCSubTotalP As Double
        Dim NCSubTotalS As Double
        Dim PerSinIVA As Boolean
        Dim PerSecPub As Boolean
        Dim clsAuxAsi As New clsConsulta
        Dim strCentroCosto As String
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXC
        strSql = " SELECT ingreso.per_codigo,ing_fecha,0 AS ing_total,0 AS ing_subtotal,0 AS ing_dcto," & _
                 " 0 AS ing_impuesto,per_siniva,per_sec_publico,0 as totprod,0 AS totserv " & _
                 " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo " & _
                 " AND ingreso.per_codigo=persona.per_codigo " & _
                 " WHERE ingreso.emp_codigo='" & Emp & "' AND ingreso.tip_ing_codigo='DCL' " & _
                 " AND ingreso.ing_codigo='" & NC & "' and ing_anulado=1 "
        clsAuxAsi.Ejecutar (strSql)
        If clsAuxAsi.adorec_Def.RecordCount = 0 Then
            strSql = " SELECT ingreso.per_codigo,ing_fecha,ing_total,ing_subtotal,ing_dcto,ing_impuesto,per_siniva," & _
                     " per_sec_publico,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0)) as totprod," & _
                     " SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0)) as totserv " & _
                     " FROM ingreso INNER JOIN persona " & _
                     " ON ingreso.emp_codigo=persona.emp_codigo " & _
                     " AND ingreso.per_codigo=persona.per_codigo " & _
                     " INNER JOIN det_ingreso " & _
                     " ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                     " AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                     " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & Emp & "' " & _
                     " AND ingreso.tip_ing_codigo='DCL'" & _
                     " AND ingreso.ing_codigo='" & NC & "' " & _
                     " GROUP BY ingreso.per_codigo,ing_fecha,ing_total,ing_subtotal,ing_dcto,ing_impuesto,per_siniva,per_sec_publico"
            clsAuxAsi.Ejecutar (strSql)
        End If
        If clsAuxAsi.adorec_Def.RecordCount = 0 Then
            strSql = " SELECT ingreso.per_codigo,ing_fecha,ing_total,ing_subtotal,ing_dcto,ing_impuesto,per_siniva," & _
                     " per_sec_publico,0 as totprod,-1*ing_dcto as totserv " & _
                     " FROM ingreso INNER JOIN persona " & _
                     " ON ingreso.emp_codigo=persona.emp_codigo " & _
                     " AND ingreso.per_codigo=persona.per_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & Emp & "' " & _
                     " AND ingreso.tip_ing_codigo='DCL'" & _
                     " AND ingreso.ing_codigo='" & NC & "' "
            clsAuxAsi.Ejecutar (strSql)
        
        End If
        If clsAuxAsi.adorec_Def.RecordCount > 0 Then
            PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
            FechaIng = clsAuxAsi.adorec_Def("ing_fecha")
            NCTotal = clsAuxAsi.adorec_Def("ing_total")
            NCIVA = clsAuxAsi.adorec_Def("ing_impuesto")
            NCSubTotal = clsAuxAsi.adorec_Def("ing_subtotal")
            NCDcto = clsAuxAsi.adorec_Def("ing_dcto")
            NCSubTotalP = clsAuxAsi.adorec_Def("totprod")
            NCSubTotalS = clsAuxAsi.adorec_Def("totserv")
            PerSinIVA = False
            If clsAuxAsi.adorec_Def("per_siniva") = 1 Then
                PerSinIVA = True
            End If
            PerSecPub = False
            If clsAuxAsi.adorec_Def("per_sec_publico") = 1 Then
                PerSecPub = True
            End If
            Dim clsAsi As New clsContable
            clsAsi.Inicializar AdoConn, AdoConnMaster
            clsAsi.NuevoAsiento "A", FechaIng, 0, 0, NCTotal, "NOTA DE CREDITO " & NC, True
            Asi = clsAsi.NumAsiento
            strSql = " UPDATE ingreso " & _
                     " SET ing_numasiento='" & Asi & "' " & _
                     " WHERE ingreso.emp_codigo = '" & Emp & "' " & _
                     " AND ingreso.tip_ing_codigo='DCL'" & _
                     " AND ingreso.ing_codigo='" & NC & "'"
            clsAuxAsi.Ejecutar (strSql)
            strSql = " DELETE FROM det_asiento " & _
                     " WHERE emp_codigo = '" & Emp & "' " & _
                     " AND asi_numasiento='" & Asi & "'"
            clsAuxAsi.Ejecutar strSql, "MASTER"
            'cuenta contable CXC
            strCentroCosto = ""
            strSql = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                    " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                    " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                    " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                    " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                    " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                    " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                    " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                    " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", 0, FormatoD2(NCTotal)
            strCentroCosto = clsAuxAsi.adorec_Def("cen_cos_codigo")
            'cuenta contable IVA VENTAS
            If FormatoD2(NCIVA) <> 0 And PerSinIVA = False Then
                strSql = " SELECT par_texto " & _
                         " FROM parametro " & _
                         " WHERE emp_codigo='" & Emp & "' " & _
                         " AND par_codigo='IVAV' "
                clsAuxAsi.Ejecutar (strSql)
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", FormatoD2(NCIVA), 0
            End If
            'cuenta contable FACTURAS PRODUCTOS
            If FormatoD2(NCSubTotalP) <> 0 Then
                strSql = " SELECT tip_ing_ctaconta " & _
                         " FROM tipo_ingreso " & _
                         " WHERE emp_codigo='" & Emp & "' " & _
                         " AND tip_ing_codigo='DCL' "
                clsAuxAsi.Ejecutar (strSql)
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta"), strCentroCosto, FormatoD2(NCSubTotalP), 0
            End If
            'cuenta contable FACTURAS SERVICIOS
            If FormatoD2(NCSubTotalS) <> 0 Then
                strSql = " SELECT tip_ing_ctaconta2 " & _
                         " FROM tipo_ingreso " & _
                         " WHERE emp_codigo='" & Emp & "' " & _
                         " AND tip_ing_codigo='DCL' "
                clsAuxAsi.Ejecutar (strSql)
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta2"), strCentroCosto, FormatoD2(NCSubTotalS), 0
            End If
            'cuentas contables de recargos
            strSql = " SELECT oca_ctaconta, det_ing_c_cantidad*det_ing_c_precio as Tot " & _
                     " FROM det_ingreso_c INNER JOIN ocargos ON det_ingreso_c.emp_codigo=ocargos.emp_codigo " & _
                     " AND det_ingreso_c.oca_codigo=ocargos.oca_codigo " & _
                     " WHERE det_ingreso_c.emp_codigo='" & Emp & "' " & _
                     " AND det_ingreso_c.tip_ing_codigo='DCL' " & _
                     " AND det_ingreso_c.ing_codigo='" & NC & "' "
            clsAuxAsi.Ejecutar strSql, "M"
            While Not clsAuxAsi.adorec_Def.EOF
                If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                    clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), strCentroCosto, FormatoD2(clsAuxAsi.adorec_Def("Tot")), 0
                End If
                clsAuxAsi.adorec_Def.MoveNext
            Wend
            'cuenta contable DESCUENTO EN VENTAS
            If FormatoD2(NCSubTotalS) = 0 And FormatoD2(NCSubTotalP) = 0 And FormatoD2(NCDcto) <> 0 Then
                strSql = " SELECT par_texto " & _
                         " FROM parametro " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND par_codigo='DCV' "
                clsAuxAsi.Ejecutar strSql
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), strCentroCosto, FormatoD2(NCDcto) * (-1), 0
            End If
        End If
End Sub


Private Sub cmdNotaCreditoCliDescuadrada_Click()
    chequeoAsiNC
End Sub

Private Sub chequeoAsiNC()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT ing_codigo,ingreso.emp_codigo,IIF(ing_anulado=1,0,ing_total) as ing_total," & _
             " asiento.asi_numasiento,sum(COALESCE(det_asi_debe,0)) as d,sum(COALESCE(det_asi_haber,0)) as h, " & _
             " abs(round(sum(COALESCE(det_asi_debe,0)),2)-round(sum(COALESCE(det_asi_haber,0)),2)) as dif " & _
             " FROM ingreso inner join asiento " & _
             " ON ingreso.emp_codigo=asiento.emp_codigo" & _
             " AND ingreso.ing_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%NOTA DE CREDITO%',ing_codigo,'%') " & _
             " LEFT JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo " & _
             " AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
             " WHERE tip_ing_codigo='DCL' " & _
             " AND ing_fecha>='" & strFecha & "' AND ingreso.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY ing_codigo,ingreso.emp_codigo,ing_anulado,ing_total,asiento.asi_numasiento " & _
             " HAVING ROUND(sum(COALESCE(det_asi_debe,0)), 2) <> ROUND(sum(COALESCE(det_asi_haber,0)), 2) OR ROUND(sum(COALESCE(det_asi_haber,0)), 2)<> ing_total" & _
             " ORDER BY dif DESC "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("ing_codigo")
        CargaAsientoNC clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("ing_codigo"), clsSql.adorec_Def("asi_numasiento")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsientoNC(Emp As String, NC As String, Asi As String)
        Dim PerCodigo As String
        Dim NCTotal As Double
        Dim NCIVA As Double
        Dim NCSubTotal As Double
        Dim NCDcto As Double
        Dim NCSubTotalP As Double
        Dim NCSubTotalS As Double
        Dim PerSinIVA As Boolean
        Dim PerSecPub As Boolean
        Dim clsAuxAsi As New clsConsulta
        Dim strCentroCosto As String
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXC
        strSql = " SELECT ingreso.per_codigo,IIF(ing_anulado=1,0,ing_total) as ing_total," & _
                 " IIF(ing_anulado=1,0,ing_subtotal) as ing_subtotal,IIF(ing_anulado=1,0,ing_dcto) as ing_dcto," & _
                 " IIF(ing_anulado=1,0,ing_impuesto) as ing_impuesto,per_siniva,per_sec_publico," & _
                 " IIF(ing_anulado=1,0,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0))) as totprod," & _
                 " IIF(ing_anulado=1,0,SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_ing_cantidad*det_ing_precio,2)-det_ing_dcto,0))) as totserv " & _
                 " FROM ingreso INNER JOIN persona " & _
                 " ON ingreso.emp_codigo=persona.emp_codigo " & _
                 " AND ingreso.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN det_ingreso " & _
                 " ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                 " AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                 " WHERE ingreso.emp_codigo = '" & Emp & "' " & _
                 " AND ingreso.tip_ing_codigo='DCL'" & _
                 " AND ingreso.ing_codigo='" & NC & "' " & _
                 " GROUP BY ingreso.per_codigo,ing_anulado,ing_total,ing_subtotal,ing_dcto,ing_impuesto,per_siniva,per_sec_publico"
        clsAuxAsi.Ejecutar (strSql)
        PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
        NCTotal = clsAuxAsi.adorec_Def("ing_total")
        NCIVA = clsAuxAsi.adorec_Def("ing_impuesto")
        NCSubTotal = clsAuxAsi.adorec_Def("ing_subtotal")
        NCDcto = clsAuxAsi.adorec_Def("ing_dcto")
        NCSubTotalP = clsAuxAsi.adorec_Def("totprod")
        NCSubTotalS = clsAuxAsi.adorec_Def("totserv")
        If NCSubTotalP <> 0 And NCSubTotalS = 0 Then
            NCSubTotalP = NCTotal - NCIVA
        ElseIf NCSubTotalS <> 0 And NCSubTotalP = 0 Then
            NCSubTotalS = NCSubTotal - NCDcto
        'ElseIf NCSubTotalS = 0 And NCSubTotalP = 0 Then
            'NCSubTotalS = NCSubTotal
        Else
            NCSubTotalS = NCSubTotal - NCDcto - NCSubTotalP
        End If
        PerSinIVA = False
        If clsAuxAsi.adorec_Def("per_siniva") = 1 Then
            PerSinIVA = True
        End If
        PerSecPub = False
        If clsAuxAsi.adorec_Def("per_sec_publico") = 1 Then
            PerSecPub = True
        End If
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        clsAsi.NumAsiento = Asi
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        'cuenta contable CXC
        strCentroCosto = ""
        strSql = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
        clsAuxAsi.Ejecutar (strSql)
        clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", 0, FormatoD2(NCTotal)
        strCentroCosto = clsAuxAsi.adorec_Def("cen_cos_codigo")
        'cuenta contable IVA VENTAS
        If FormatoD2(NCIVA) <> 0 And PerSinIVA = False Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='IVAV' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", FormatoD2(NCIVA), 0
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(NCSubTotalP) <> 0 Then
            strSql = " SELECT tip_ing_ctaconta " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_ing_codigo='DCL' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta"), strCentroCosto, FormatoD2(NCSubTotalP), 0
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(NCSubTotalS) <> 0 Then
            strSql = " SELECT tip_ing_ctaconta2 " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_ing_codigo='DCL' "
            clsAuxAsi.Ejecutar (strSql)
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_ing_ctaconta2"), strCentroCosto, FormatoD2(NCSubTotalS), 0
        End If
        'cuentas contables de recargos
        strSql = " SELECT oca_ctaconta, det_ing_c_cantidad*det_ing_c_precio as Tot " & _
                 " FROM det_ingreso_c INNER JOIN ocargos ON det_ingreso_c.emp_codigo=ocargos.emp_codigo " & _
                 " AND det_ingreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_ingreso_c.emp_codigo='" & Emp & "' " & _
                 " AND det_ingreso_c.tip_ing_codigo='DCL' " & _
                 " AND det_ingreso_c.ing_codigo='" & NC & "' "
        clsAuxAsi.Ejecutar strSql, "M"
        While Not clsAuxAsi.adorec_Def.EOF
            If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), strCentroCosto, FormatoD2(clsAuxAsi.adorec_Def("Tot")), 0
            End If
            clsAuxAsi.adorec_Def.MoveNext
        Wend
        'cuenta contable DESCUENTO EN VENTAS
        If FormatoD2(NCSubTotalS) = 0 And FormatoD2(NCSubTotalP) = 0 And FormatoD2(NCDcto) <> 0 Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='DCV' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), strCentroCosto, FormatoD2(NCDcto) * (-1), 0
        End If
End Sub

Private Sub cmdNotaCreditoProDescuadrada_Click()
    chequeoAsiDPV
End Sub

Private Sub chequeoAsiDPV()
    Dim clsSql As New clsConsulta
    Dim i As Long
    clsSql.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(dtpFecha.Value, "YYYY-mm-dd")
    strSql = " SELECT egr_codigo,egreso.emp_codigo,IIF(egr_anulado=1,0,egr_total) as egr_total," & _
             " asiento.asi_numasiento,sum(COALESCE(det_asi_debe,0)) as d,sum(COALESCE(det_asi_haber,0)) as h, " & _
             " abs(round(sum(COALESCE(det_asi_debe,0)),2)-round(sum(COALESCE(det_asi_haber,0)),2)) as dif " & _
             " FROM egreso inner join asiento " & _
             " ON egreso.emp_codigo=asiento.emp_codigo" & _
             " AND egreso.egr_numasiento=asiento.asi_numasiento " & _
             " AND asiento.asi_descripcion like CONCAT('%DEVOLUCION A PROVEEDOR%',egr_codigo,'%') " & _
             " LEFT JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo " & _
             " AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
             " WHERE tip_egr_codigo='DPV' " & _
             " AND egr_fecha>='" & strFecha & "' AND egreso.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY egr_codigo,egreso.emp_codigo,IIF(egr_anulado=1,0,egr_total),asiento.asi_numasiento " & _
             " HAVING ROUND(sum(COALESCE(det_asi_debe,0)), 2) <> ROUND(sum(COALESCE(det_asi_haber,0)), 2) OR ROUND(sum(COALESCE(det_asi_haber,0)), 2)<>IIF(egr_anulado=1,0,egr_total) " & _
             " ORDER BY abs(round(sum(COALESCE(det_asi_debe,0)),2)-round(sum(COALESCE(det_asi_haber,0)),2)) DESC "
    clsSql.Ejecutar strSql, "LOCAL"
    VSFG.Rows = 1
    i = 1
    While Not clsSql.adorec_Def.EOF
        VSFG.AddItem i & vbTab & clsSql.adorec_Def("egr_codigo")
        CargaAsientoDPV clsSql.adorec_Def("emp_codigo"), clsSql.adorec_Def("egr_codigo"), clsSql.adorec_Def("asi_numasiento")
        clsSql.adorec_Def.MoveNext
        i = i + 1
    Wend
    MsgBox "Revision terminada"
End Sub

Private Sub CargaAsientoDPV(Emp As String, Fac As String, Asi As String)
        Dim PerCodigo As String
        Dim FacTotal As Double
        Dim FacIVA As Double
        Dim FacSubTotal As Double
        Dim FacDcto As Double
        Dim FacSubTotalP As Double
        Dim FacSubTotalS As Double
        Dim PerSinIVA As Boolean
        Dim PerSecPub As Boolean
        Dim clsAuxAsi As New clsConsulta
        clsAuxAsi.Inicializar AdoConn, AdoConnMaster
        'cuenta contable CXC
        strSql = " SELECT egreso.per_codigo,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva," & _
                 " per_sec_publico,SUM(IIF(LEFT(prd_codigo,3)!='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totprod," & _
                 " SUM(IIF(LEFT(prd_codigo,3)='PR-',ROUND(det_egr_cantidad*det_egr_precio,2)-det_egr_dcto,0)) as totserv,egr_anulado " & _
                 " FROM egreso INNER JOIN persona " & _
                 " ON egreso.emp_codigo=persona.emp_codigo " & _
                 " AND egreso.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN det_egreso " & _
                 " ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                 " WHERE egreso.emp_codigo = '" & Emp & "' " & _
                 " AND egreso.tip_egr_codigo='DPV'" & _
                 " AND egreso.egr_codigo='" & Fac & "' " & _
                 " GROUP BY egreso.per_codigo,egr_total,egr_subtotal,egr_dcto,egr_impuesto,per_siniva,per_sec_publico,egr_anulado"
        clsAuxAsi.Ejecutar (strSql)
        PerCodigo = clsAuxAsi.adorec_Def("per_codigo")
        If clsAuxAsi.adorec_Def("egr_anulado") = 0 Then
            FacTotal = clsAuxAsi.adorec_Def("egr_total")
            FacIVA = clsAuxAsi.adorec_Def("egr_impuesto")
            FacSubTotal = FacTotal - FacIVA
            FacDcto = clsAuxAsi.adorec_Def("egr_dcto")
            FacSubTotalP = clsAuxAsi.adorec_Def("totprod")
            FacSubTotalS = clsAuxAsi.adorec_Def("totserv")
        Else
            FacTotal = 0
            FacIVA = 0
            FacSubTotal = 0
            FacDcto = 0
            FacSubTotalP = 0
            FacSubTotalS = 0
        
        End If
        If FacSubTotalP <> 0 And FacSubTotalS = 0 Then
            FacSubTotalP = FacSubTotal
        ElseIf FacSubTotalS <> 0 And FacSubTotalP = 0 Then
            FacSubTotalS = FacSubTotal
        ElseIf FacSubTotalS = 0 And FacSubTotalP = 0 Then
            'FacSubTotalS = FacSubTotal
        Else
            FacSubTotalS = FacSubTotal - FacSubTotalP
        End If
        
        PerSinIVA = False
        If clsAuxAsi.adorec_Def("per_siniva") = 1 Then
            PerSinIVA = True
        End If
        PerSecPub = False
        If clsAuxAsi.adorec_Def("per_sec_publico") = 1 Then
            PerSecPub = True
        End If
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        clsAsi.NumAsiento = Asi
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE emp_codigo = '" & Emp & "' " & _
                 " AND asi_numasiento='" & Asi & "'"
        clsAuxAsi.Ejecutar strSql, "MASTER"
        'cuenta contable CXP
        strSql = " SELECT cat_p_ctaconta " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " WHERE persona.per_codigo= '" & PerCodigo & "' AND persona.emp_codigo = '" & Emp & "' "
        clsAuxAsi.Ejecutar strSql
        clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("cat_p_ctaconta"), "", FormatoD2(FacTotal), 0
        'cuenta contable IVA COMPRAS
        If FormatoD2(FacIVA) <> 0 Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='IVAC' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", 0, FormatoD2(FacIVA)
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(FacSubTotalP) <> 0 Then
            strSql = " SELECT tip_egr_ctaconta " & _
                     " FROM tipo_egreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_egr_codigo='DPV' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta"), "", 0, FormatoD2(FacSubTotalP)
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(FacSubTotalS) <> 0 Then
            strSql = " SELECT tip_egr_ctaconta2 " & _
                     " FROM tipo_egreso " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND tip_egr_codigo='DPV' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("tip_egr_ctaconta2"), "", 0, FormatoD2(FacSubTotalS)
        End If
        'cuentas contables de RECARGOS
        strSql = " SELECT oca_ctaconta, det_egr_c_cantidad*det_egr_c_precio as Tot " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo " & _
                 " AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_egreso_c.emp_codigo='" & Emp & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='DPV' " & _
                 " AND det_egreso_c.egr_codigo='" & Fac & "' "
        clsAuxAsi.Ejecutar strSql, "M"
        While Not clsAuxAsi.adorec_Def.EOF
            If FormatoD2(clsAuxAsi.adorec_Def("Tot")) <> 0 Then
                clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("oca_ctaconta"), "", 0, FormatoD2(clsAuxAsi.adorec_Def("Tot"))
            End If
            clsAuxAsi.adorec_Def.MoveNext
        Wend
        'cuenta contable DESCUENTO EN COMPRAS
        If FormatoD2(FacSubTotalS) = 0 And FormatoD2(FacSubTotalP) = 0 And FormatoD2(FacDcto) <> 0 Then
            strSql = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & Emp & "' " & _
                     " AND par_codigo='DCC' "
            clsAuxAsi.Ejecutar strSql
            clsAsi.NuevoDetAsiento clsAuxAsi.adorec_Def("par_texto"), "", 0, FormatoD2(FacDcto) * (-1)
        End If
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

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False, True, True, True, False, False, False
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
