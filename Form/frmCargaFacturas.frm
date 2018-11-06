VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCargaFacturas 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Facturas"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaFacturas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   14850
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5858
      TabIndex        =   2
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   14640
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   9120
         TabIndex        =   10
         Top             =   300
         Width           =   4275
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   13440
         TabIndex        =   6
         Top             =   300
         Width           =   375
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   14415
         _cx             =   1983275858
         _cy             =   1983256385
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaFacturas.frx":030A
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
         TabBehavior     =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   6615
         _cx             =   101395860
         _cy             =   101390357
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   27
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaFacturas.frx":0394
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
         TabBehavior     =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG2_det 
         Height          =   3855
         Left            =   6840
         TabIndex        =   5
         Top             =   4200
         Width           =   7695
         _cx             =   1983264005
         _cy             =   1983257232
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaFacturas.frx":06A0
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
         TabBehavior     =   0
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
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   13320
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   4575
         _ExtentX        =   8070
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
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   6690
         TabIndex        =   11
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Value           =   42439.3923958333
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6120
         TabIndex        =   12
         Top             =   352
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8400
         TabIndex        =   9
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   345
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7538
      TabIndex        =   0
      Top             =   8280
      Width           =   1455
   End
End
Attribute VB_Name = "frmCargaFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
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
Private strSql As String
Private FacYaReg As Long
Private NoHayCli As Long
Private BloqueadoCli As Long
Private ProductoProblema As Long

Private Sub cmdAplicar_Click()
    Dim i As Long
    Dim j As Long
    Dim clsAsiento As New clsContable
    Dim clsEgreso As New clsInventario
    Dim clsCta As New clsCtaXx
    Dim numPed As String
    
'    If ProductoProblema > 0 Then
'        MsgBox "No puede cargar hasta resolver problemas con productos", vbCritical, "Migración"
'        Exit Sub
'    End If
    
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsCta.Inicializar AdoConn, AdoConnMaster
    
    Me.MousePointer = 11
    
    VSFG2.Select 1, VSFG2.Cols - 1
    VSFG2.Sort = flexSortGenericDescending
    
    For i = 1 To VSFG2.Rows - 1
        VSFG2.ShowCell i, 0
        VSFG2.Refresh
        If Val(VSFG2.TextMatrix(i, VSFG2.Cols - 1)) = 2 Then
            'FP, Vendedor, Persona
            clsEgreso.NuevoEgr True, VSFG2.TextMatrix(i, 0), False, VSFG2.TextMatrix(i, 2), VSFG2.TextMatrix(i, 3), VSFG2.TextMatrix(i, 4), VSFG2.TextMatrix(i, 5), VSFG2.TextMatrix(i, 6), Format(VSFG2.TextMatrix(i, 7), "yyyy-MM-dd"), , VSFG2.TextMatrix(i, 9), VSFG2.TextMatrix(i, 10), , VSFG2.TextMatrix(i, 12), VSFG2.TextMatrix(i, 13), FormatoD2(VSFG2.TextMatrix(i, 14)), FormatoD2(VSFG2.TextMatrix(i, 15)), FormatoD2(VSFG2.TextMatrix(i, 16)), FormatoD2(VSFG2.TextMatrix(i, 17)), FormatoD2(VSFG2.TextMatrix(i, 18)), 0, VSFG2.TextMatrix(i, 20), VSFG2.TextMatrix(i, 21)
            numPed = RegistroPedido(VSFG2.TextMatrix(i, 2), VSFG2.TextMatrix(i, 3), VSFG2.TextMatrix(i, 4), VSFG2.TextMatrix(i, 6), VSFG2.TextMatrix(i, 9), Format(VSFG2.TextMatrix(i, 25), "yyyy-MM-dd"), FormatoD2(VSFG2.TextMatrix(i, 14)))
            clsAsiento.NuevoAsiento "F", Format(VSFG2.TextMatrix(i, 7), "yyyy-MM-dd"), 0, 0, FormatoD2(VSFG2.TextMatrix(i, 18)), "FACTURA " & clsEgreso.strDoc
            'Inserta la cabecera del egreso
            clsEgreso.ModificaEgr , , , , , , clsAsiento.NumAsiento
            For j = VSFG2.TextMatrix(i, 22) To VSFG2.TextMatrix(i, 23)
                If FormatoD2(VSFG2_det.TextMatrix(j, 5)) <> 0 Then
                    clsEgreso.NuevoDetEgr VSFG2_det.TextMatrix(j, 4), VSFG2_det.TextMatrix(j, 3), VSFG2_det.TextMatrix(j, 5), VSFG2_det.TextMatrix(j, 6), 0, VSFG2_det.TextMatrix(j, 7), Abs(FormatoD0(VSFG2_det.TextMatrix(j, 9))), VSFG2_det.TextMatrix(j, 8)
                End If
                RegistroDetPedido numPed, VSFG2_det.TextMatrix(j, 4), VSFG2_det.TextMatrix(j, 3), VSFG2_det.TextMatrix(j, 10), VSFG2_det.TextMatrix(j, 5), VSFG2_det.TextMatrix(j, 6), VSFG2_det.TextMatrix(j, 7)
            Next j
            'clsEgreso.NuevoDetEgrRecargo VSFGReca.TextMatrix(i, 1), FormatoD2(VSFGReca.TextMatrix(i, 3))
            
            clsCta.NuevaCta "C", 1, "00", Format(VSFG2.TextMatrix(i, 7), "yyyy-MM-dd"), Format(DateAdd("d", VSFG2.TextMatrix(i, 24), VSFG2.TextMatrix(i, 7)), "yyyy-MM-dd"), VSFG2.TextMatrix(i, 6), "Factura # " & clsEgreso.strDoc & " - " & VSFG2.TextMatrix(i, 10), VSFG2.TextMatrix(i, 2) & VSFG2.TextMatrix(i, 3), Format(Right(VSFG2.TextMatrix(i, 4), 7), "0000000"), VSFG2.TextMatrix(i, 12), VSFG2.TextMatrix(i, 13), clsEgreso.dblTotalProd, clsEgreso.dblTotalServ, clsEgreso.dblTotalProdIVA, clsEgreso.dblTotalServIVA, 2, clsEgreso.dblIVA, clsEgreso.dblSubTotal0, 0, 0, 0, clsEgreso.dblTotal, clsAsiento.NumAsiento
    
            clsCta.IngAsientoEgr clsAsiento, clsEgreso
        Else
            Exit For
        End If
    Next i
    Me.MousePointer = 0
    MsgBox "Facturas Cargadas", vbInformation, "Facturas"
    Unload Me
End Sub
Private Sub RegistroDetPedido(strNum As String, strProducto As String, strDeposito As String, dblCantPedida As Double, dblCantEntregada As Double, dblPrecio As Double, dblDcto As Double)
    
    strSql = " SELECT det_ped_cant_pedida,det_ped_precio,det_ped_dcto " & _
             " FROM det_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_codigo=" & strNum & " " & _
             " AND prd_codigo='" & strProducto & "' " & _
             " AND dep_codigo='" & strDeposito & "' "
    clsCon_Def.Ejecutar (strSql), "M"
    If clsCon_Def.adorec_Def.RecordCount = 0 Then
        strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                 " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
                 " VALUES ('" & strEmpresa & "'," & strNum & ",'" & strProducto & "','" & strDeposito & "','" & dblCantPedida & "'," & _
                 "'" & dblCantEntregada & "' ," & dblPrecio & "," & dblDcto & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
        clsCon_Def.Ejecutar (strSql), "M"
    Else
        strSql = " UPDATE det_pedido " & _
                 " SET det_ped_cant_pedida=det_ped_cant_pedida+" & dblCantPedida & "," & _
                 " det_ped_cant_entregada=det_ped_cant_entregada+" & dblCantEntregada & "," & _
                 " det_ped_precio=" & (clsCon_Def.adorec_Def("det_ped_precio") * clsCon_Def.adorec_Def("det_ped_cant_pedida") + dblCantPedida * dblPrecio) / (clsCon_Def.adorec_Def("det_ped_cant_pedida") + dblCantPedida) & "," & _
                 " det_ped_dcto=det_ped_dcto+" & dblDcto & " " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo=" & strNum & " " & _
                 " AND prd_codigo='" & strProducto & "' " & _
                 " AND dep_codigo='" & strDeposito & "' "
        clsCon_Def.Ejecutar (strSql), "M"
    End If
End Sub

Private Function RegistroPedido(strSuc As String, strPto As String, strNumero As String, strCliente As String, strVendedor As String, strFecha As String, dblSubTotal As Double) As String
Dim clsSqlNum As New clsConsulta
    clsSqlNum.Inicializar AdoConn, AdoConnMaster
    strSql = " LOCK TABLES pedido WRITE "
    clsSqlNum.Ejecutar strSql, "M"
    strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(strSucursal & Fact & "0000001") & "') as num " & _
             " From pedido " & _
             " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(strSuc & strPto) & "%'" & _
             " GROUP BY emp_codigo"
    clsSqlNum.Ejecutar (strSql), "M"
    RegistroPedido = clsSqlNum.adorec_Def("num")
    strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
         " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega,ped_tip_egr_codigo,ped_egr_codigo, ped_fechamod, ped_usumod) " & _
         " VALUES ('" & strEmpresa & "'," & clsSqlNum.adorec_Def("num") & ",'" & strCliente & "','" & strVendedor & "', " & _
         " 'SINTC','0'," & _
         " '" & strFecha & "','2'," & FormatoD2(dblSubTotal) & ",'MIGRACION - FAC: " & Format(strSuc, "000") & Format(strPto, "000") & Format(strNumero, "0000000") & "', " & _
         " '',1,'0','FAC','" & Format(strSuc, "000") & Format(strPto, "000") & Format(strNumero, "0000000") & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
    clsSqlNum.Ejecutar (strSql), "M"
    strSql = " UNLOCK TABLES"
    clsSqlNum.Ejecutar (strSql), "M"

End Function
Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    FacYaReg = 0
    NoHayCli = 0
    BloqueadoCli = 0
    ProductoProblema = 0
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    If (txtArchivo.Text <> "") Then
        Me.MousePointer = 11
        VSFG.ClipSeparators = ";" & vbCr
        VSFG.FixedRows = 0
        VSFG.Rows = 0
        VSFG.LoadGrid txtArchivo.Text, flexFileCustomText
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.LoadGrid txtArchivo.Text, flexFileTabText
'        If Not LeerExcel(txtArchivo.Text) Then
'            Me.MousePointer = 0
'            Exit Sub
'        End If
        VSFG.FixedRows = 1
        j = 1
        k = 1
        FacYaReg = 0
        For i = 1 To VSFG.Rows - 1
'            strSql = " SELECT COALESCE(count(*),0) as n " & _
'                     " FROM persona " & _
'                     " WHERE emp_codigo='" & strEmpresa & "'" & _
'                     " AND per_ruc='" & VSFG.TextMatrix(i, 1) & "' "
'
'            clsCon_Def.Ejecutar strSql
            If i >= VSFG.Rows - 1 Then Exit For
            VSFG.ShowCell i, 0
            VSFG.Refresh
            If UCase(VSFG.TextMatrix(i, 14)) <> "NINGUNA" And UCase(VSFG.TextMatrix(i, 14)) <> "" Then
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 2
                VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
                If VSFG2.TextMatrix(j - 1, 3) = Left(VSFG.TextMatrix(i, 14), 3) And VSFG2.TextMatrix(j - 1, 2) = Right(Left(VSFG.TextMatrix(i, 14), 7), 3) And VSFG2.TextMatrix(j - 1, 4) = Right(VSFG.TextMatrix(i, 14), Len(VSFG.TextMatrix(i, 14)) - 8) Then
                
                    VSFG2.TextMatrix(j - 1, 14) = FormatoD2(VSFG2.TextMatrix(j - 1, 14)) + FormatoD2(IIf(FormatoD2(VSFG.TextMatrix(i, 26)) <> 0, FormatoD4(VSFG.TextMatrix(i, 23)) * FormatoD2(VSFG.TextMatrix(i, 28)), 0)) 'subtotal
                    VSFG2.TextMatrix(j - 1, 15) = FormatoD2(VSFG2.TextMatrix(j - 1, 15)) + FormatoD2(IIf(FormatoD2(VSFG.TextMatrix(i, 26)) = 0, FormatoD4(VSFG.TextMatrix(i, 23)) * FormatoD2(VSFG.TextMatrix(i, 28)), 0)) 'subtotal_0
                    VSFG2.TextMatrix(j - 1, 16) = FormatoD2(VSFG2.TextMatrix(j - 1, 16)) + FormatoD2(VSFG.TextMatrix(i, 27)) 'Dcto
                    VSFG2.TextMatrix(j - 1, 17) = FormatoD2((FormatoD2(VSFG2.TextMatrix(j - 1, 14)) - FormatoD2(VSFG2.TextMatrix(j - 1, 16))) * PorIVA / 100#)   'Impuesto
                    VSFG2.TextMatrix(j - 1, 18) = FormatoD2(FormatoD2(VSFG2.TextMatrix(j - 1, 14)) + FormatoD2(VSFG2.TextMatrix(j - 1, 15)) - FormatoD2(VSFG2.TextMatrix(j - 1, 16)) + FormatoD2(VSFG2.TextMatrix(j - 1, 17))) 'Total
                
                Else
                    VSFG2.AddItem ""
                    VSFG2.TextMatrix(j, 0) = "FAC" 'tipo
                    VSFG2.ShowCell j, 0
                    VSFG2.Refresh
                    VSFG2.TextMatrix(j, 1) = False 'confirma numero
                    VSFG2.TextMatrix(j, 3) = Left(VSFG.TextMatrix(i, 14), 3) 'sucursal
                    VSFG2.TextMatrix(j, 2) = Right(Left(VSFG.TextMatrix(i, 14), 7), 3) 'puntofactura
                    VSFG2.TextMatrix(j, 4) = Right(VSFG.TextMatrix(i, 14), Len(VSFG.TextMatrix(i, 14)) - 8) 'numero
                    strSql = " SELECT per_codigo,persona.for_pag_codigo,for_pag_tiempo,ven_codigo,IF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                             " FROM persona INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo" & _
                             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo" & _
                             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
                             " AND cat_p_tipo='C' " & _
                             " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                             " AND per_ruc='" & VSFG.TextMatrix(i, 4) & "' "
                    clsCon_Def.Ejecutar strSql
                    VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 3
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                        If FormatoD0(clsCon_Def.adorec_Def("per_bloqueado")) = 1 Then
                            VSFG2.Cell(flexcpBackColor, j, 1, j, VSFG2.Cols - 1) = vbMagenta
                            BloqueadoCli = BloqueadoCli + 1
                            If VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 3 Then
                                VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 0
                            End If
                        End If
                        
                            VSFG2.TextMatrix(j, 5) = clsCon_Def.adorec_Def("for_pag_codigo") 'forma de pago
                            VSFG2.TextMatrix(j, 6) = clsCon_Def.adorec_Def("per_codigo") 'persona
                            VSFG2.TextMatrix(j, 24) = clsCon_Def.adorec_Def("for_pag_tiempo") 'tiempo de credito
                            VSFG2.TextMatrix(j, 9) = clsCon_Def.adorec_Def("ven_codigo") 'vendedor
                    Else
                        VSFG2.Cell(flexcpBackColor, j, 2, j, VSFG2.Cols - 1) = vbCyan
                        NoHayCli = NoHayCli + 1
                        If VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 3 Then
                            VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 0
                        End If
                    End If
                    VSFG2.TextMatrix(j, 7) = Format(dtpFecha.Value, "yyyy-mm-dd") 'fecha
                    VSFG2.TextMatrix(j, 25) = Format(Right(VSFG.TextMatrix(i, 3), 4) & "-" & Left(Right(VSFG.TextMatrix(i, 3), 7), 2) & "-" & Left(VSFG.TextMatrix(i, 3), 2), "yyyy-mm-dd") 'fechapedido
                    VSFG2.TextMatrix(j, 8) = VSFG.TextMatrix(i, 2) 'doc2
                    VSFG2.TextMatrix(j, 10) = "MIGRACION" 'Observacion
                    VSFG2.TextMatrix(j, 11) = "" 'ASIENTO
                    VSFG2.TextMatrix(j, 12) = strAutorFactura 'strautor
                    VSFG2.TextMatrix(j, 13) = strCaducaFactura 'strcaduca
                    VSFG2.TextMatrix(j, 14) = FormatoD2(IIf(VSFG.TextMatrix(i, 26) <> 0, FormatoD4(VSFG.TextMatrix(i, 23)) * FormatoD2(VSFG.TextMatrix(i, 28)), 0)) 'subtotal
                    VSFG2.TextMatrix(j, 15) = FormatoD2(IIf(VSFG.TextMatrix(i, 26) = 0, FormatoD4(VSFG.TextMatrix(i, 23)) * FormatoD2(VSFG.TextMatrix(i, 28)), 0)) 'subtotal_0
                    VSFG2.TextMatrix(j, 16) = FormatoD2(VSFG.TextMatrix(i, 27)) 'Dcto
                    VSFG2.TextMatrix(j, 17) = FormatoD2((FormatoD2(VSFG2.TextMatrix(j, 14)) - FormatoD2(VSFG2.TextMatrix(j, 16))) * PorIVA / 100#)   'Impuesto
                    VSFG2.TextMatrix(j, 18) = FormatoD2(FormatoD2(VSFG2.TextMatrix(j, 14)) + FormatoD2(VSFG2.TextMatrix(j, 15)) - FormatoD2(VSFG2.TextMatrix(j, 16)) + FormatoD2(VSFG2.TextMatrix(j, 17))) 'Total
                    VSFG2.TextMatrix(j, 19) = 0 'saldo
                    VSFG2.TextMatrix(j, 20) = False 'Sector Publico
                    VSFG2.TextMatrix(j, 21) = False 'SinIVA
                    VSFG2.TextMatrix(j, 22) = k 'desde
                    If j > 1 Then
                        VSFG2.TextMatrix(j - 1, 23) = k - 1 'hasta
                    End If
                    
                            
                    strSql = " SELECT COALESCE(count(*),0) as n " & _
                             " FROM egreso " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND tip_egr_codigo='FAC' " & _
                             " AND egr_codigo='" & FormatoD0(VSFG2.TextMatrix(j, 2) & VSFG2.TextMatrix(j, 3) & Format(VSFG2.TextMatrix(j, 4), "0000000")) & "' "
                    clsCon_Def.Ejecutar strSql
                    
                    If clsCon_Def.adorec_Def("n") <> 0 Then
                        VSFG2.Cell(flexcpBackColor, j, 3, j, VSFG2.Cols - 1) = vbRed
                        FacYaReg = FacYaReg + 1
                        VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 0
                    Else
                        If VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 3 Then
                            VSFG2.TextMatrix(j, VSFG2.Cols - 1) = 2
                        End If
                    End If
                    
                    j = j + 1
                    If FormatoD2(VSFG.TextMatrix(i, 12)) > 0 Then
                        VSFG2_det.AddItem ""
                        VSFG2_det.TextMatrix(k, 1) = Left(VSFG.TextMatrix(i, 14), 3) 'sucursal
                        VSFG2_det.TextMatrix(k, 0) = Right(Left(VSFG.TextMatrix(i, 14), 7), 3) 'puntofactura
                        VSFG2_det.TextMatrix(k, 2) = Right(VSFG.TextMatrix(i, 14), Len(VSFG.TextMatrix(i, 14)) - 8) 'numero
                        VSFG2_det.ShowCell k, 0
                        VSFG2_det.Refresh
                        VSFG2_det.TextMatrix(k, 3) = "PRI" 'bodega
                        VSFG2_det.TextMatrix(k, 4) = "PR-CARGOO100330TU" 'producto
                        VSFG2_det.TextMatrix(k, 5) = 1 'cantidad
                        VSFG2_det.TextMatrix(k, 10) = 1 'cantidadpedida
                        VSFG2_det.TextMatrix(k, 6) = FormatoD2(FormatoD2(VSFG.TextMatrix(i, 12)) / (1 + PorIVA / 100#))   'precio
                        VSFG2_det.TextMatrix(k, 7) = 0 'descuento
                        VSFG2_det.TextMatrix(k, 8) = 0 'pdcto
                        VSFG2_det.TextMatrix(k, 9) = FormatoD2(IIf(VSFG.TextMatrix(i, 12) <> 0, 1, 0)) 'con iva
                        VSFG2_det.Cell(flexcpBackColor, k, 0, k, VSFG2_det.Cols - 1) = VSFG2.Cell(flexcpBackColor, j - 1, 0, j - 1, 0)
                        k = k + 1
                        
                        
                    VSFG2.TextMatrix(j - 1, 14) = FormatoD2(VSFG2.TextMatrix(j - 1, 14)) + FormatoD2(VSFG2_det.TextMatrix(k - 1, 6)) 'subtotal
                    'VSFG2.TextMatrix(j - 1, 15) = FormatoD2(VSFG2.TextMatrix(j - 1, 15)) + FormatoD2(IIf(VSFG.TextMatrix(i, 26) = 0, VSFG.TextMatrix(i, 23) * VSFG.TextMatrix(i, 28), 0)) 'subtotal_0
                    'VSFG2.TextMatrix(j - 1, 16) = FormatoD2(VSFG2.TextMatrix(j - 1, 16)) + FormatoD2(VSFG.TextMatrix(i, 27)) 'Dcto
                    VSFG2.TextMatrix(j - 1, 17) = FormatoD2((FormatoD2(VSFG2.TextMatrix(j - 1, 14)) - FormatoD2(VSFG2.TextMatrix(j - 1, 16))) * PorIVA / 100#)   'Impuesto
                    VSFG2.TextMatrix(j - 1, 18) = FormatoD2(FormatoD2(VSFG2.TextMatrix(j - 1, 14)) + FormatoD2(VSFG2.TextMatrix(j - 1, 15)) - FormatoD2(VSFG2.TextMatrix(j - 1, 16)) + FormatoD2(VSFG2.TextMatrix(j - 1, 17))) 'Total
                        
                    End If
                    
                End If
                VSFG2_det.AddItem ""
                VSFG2_det.TextMatrix(k, 1) = Left(VSFG.TextMatrix(i, 14), 3) 'sucursal
                VSFG2_det.TextMatrix(k, 0) = Right(Left(VSFG.TextMatrix(i, 14), 7), 3) 'puntofactura
                VSFG2_det.TextMatrix(k, 2) = Right(VSFG.TextMatrix(i, 14), Len(VSFG.TextMatrix(i, 14)) - 8) 'numero
                VSFG2_det.ShowCell k, 0
                VSFG2_det.Refresh
                VSFG2_det.TextMatrix(k, 3) = "PRI" 'bodega
                VSFG2_det.TextMatrix(k, 4) = VSFG.TextMatrix(i, 30) 'producto
                VSFG2_det.TextMatrix(k, 5) = VSFG.TextMatrix(i, 28) 'cantidad
                VSFG2_det.TextMatrix(k, 10) = VSFG.TextMatrix(i, 24) 'cantidadpedida
                VSFG2_det.TextMatrix(k, 6) = FormatoD2(VSFG.TextMatrix(i, 23)) 'precio
                VSFG2_det.TextMatrix(k, 7) = FormatoD2(VSFG.TextMatrix(i, 27)) 'descuento
                VSFG2_det.TextMatrix(k, 8) = 0 'pdcto
                VSFG2_det.TextMatrix(k, 9) = FormatoD2(IIf(VSFG.TextMatrix(i, 26) <> 0, 1, 0)) 'con iva
                VSFG2_det.Cell(flexcpBackColor, k, 0, k, VSFG2_det.Cols - 1) = VSFG2.Cell(flexcpBackColor, j - 1, 0, j - 1, 0)
                k = k + 1
            Else
                VSFG.RemoveItem i
                i = i - 1
            End If
            If i > 0 Then
            VSFG2.TextMatrix(j - 1, 23) = k - 1 'hasta
            If VSFG.TextMatrix(i, 28) > 0 Then
                strSql = "SELECT * FROM producto WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG2_det.TextMatrix(k - 1, 4) & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    VSFG2_det.Cell(flexcpBackColor, VSFG2.TextMatrix(j - 1, 22), 0, VSFG2.TextMatrix(j - 1, 23), VSFG2_det.Cols - 1) = vbGreen
                    VSFG2.Cell(flexcpBackColor, j - 1, 0, j - 1, VSFG2.Cols - 1) = vbGreen
                    VSFG2.TextMatrix(j - 1, VSFG2.Cols - 1) = 0
                    ProductoProblema = ProductoProblema + 1
                End If
            End If
            End If
        Next i
        Me.MousePointer = 0
        If FacYaReg > 0 Or NoHayCli > 0 Or ProductoProblema > 1 Then
            MsgBox "Existen " & FacYaReg & " (rojo) ya ingresadas y " & vbNewLine & NoHayCli & " (celeste) que no estan registradas el cliente" & vbNewLine & ProductoProblema & " (verde) productos con problema" & vbNewLine & BloqueadoCli & " (magenta) cliente bloqueados"
        End If
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpFecha.Value = HoyDia
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
        Set ucrtVSFG.VSFGControl = VSFG2
        Call ucrtVSFG.Inicializar(False, False, False, True, False, True, False, False, False)
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
        End If
        cmbNegocio.Enabled = False
        
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Function LeerExcel(Archivo As String) As Boolean
On Error GoTo SalirExcel
'dimensiones
LeerExcel = False
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim lngUltimaFila As Long, Fil As Long, Col As Long, FilXl As Long, mc As Boolean

Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

'abrir programa Excel
Set xlApp = New Excel.Application
xlApp.Visible = False

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(Archivo, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)

'2. Si no conoces el rango
lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

'lngUltimaFila = 17

If MsgBox("Serán Migrados #" & CStr(lngUltimaFila - 1) & " Desea Continuar?", vbYesNo) = vbNo Then Exit Function
VSFG.Rows = 1
VSFG.Cols = 36
For FilXl = 1 To lngUltimaFila
    VSFG.Rows = Fil + 1
    mc = False
    For Col = 1 To VSFG.Cols - 1
        If Col = 1 And Len(xlHoja.Range(xlHoja.Cells(FilXl, 2), xlHoja.Cells(FilXl, 2))) > 0 Then
            mc = True
            Fil = Fil + 1
        End If
        If mc = True Then
            VSFG.TextMatrix(Fil - 1, Col - 1) = xlHoja.Range(xlHoja.Cells(FilXl, Col), xlHoja.Cells(FilXl, Col))
        End If
    Next Col
Next FilXl
VSFG.Rows = VSFG.Rows + 1
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing
LeerExcel = True
Exit Function
SalirExcel:
    LeerExcel = False
    MsgBox "El Formato del Archivo Excel no es el Correcto."
End Function

