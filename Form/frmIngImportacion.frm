VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmIngImportacion 
   BackColor       =   &H00DDDDDD&
   Caption         =   "Ingreso de Importación"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmIngImportacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   7485
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Importación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   98
      TabIndex        =   13
      Top             =   120
      Width           =   7095
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   4560
         TabIndex        =   27
         Top             =   300
         Width           =   1815
         _extentx        =   3201
         _extenty        =   556
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   600
         TabIndex        =   23
         Top             =   2385
         Width           =   5895
         Begin VB.CheckBox chk_importacion 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Check1"
            Height          =   255
            Left            =   3840
            TabIndex        =   30
            Top             =   2670
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dcbo_importacion 
            Height          =   315
            Left            =   1680
            TabIndex        =   29
            Top             =   2640
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.CommandButton btn_load 
            Caption         =   "Cargar Imp. ""NS"""
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdCargarPed 
            Caption         =   "Cargar de Pedido"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   8
            Top             =   2160
            Width           =   1095
         End
         Begin VSFlex8LCtl.VSFlexGrid vsfgDetalleImp 
            Height          =   1890
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   5655
            _cx             =   136324855
            _cy             =   136318214
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   275
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmIngImportacion.frx":030A
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
         Begin MSDataListLib.DataCombo dcmbPedido 
            Height          =   330
            Left            =   1680
            TabIndex        =   26
            Top             =   2160
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image imgBtnDn 
            Height          =   210
            Left            =   2280
            Picture         =   "frmIngImportacion.frx":03D6
            Top             =   2160
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Image imgBtnUp 
            Height          =   210
            Left            =   2040
            Picture         =   "frmIngImportacion.frx":0502
            Top             =   2160
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label lblTotal 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   3960
            TabIndex        =   24
            Top             =   2220
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos Provedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   825
         Width           =   6855
         Begin VB.TextBox txtDirProveedor 
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   1005
            Width           =   2565
         End
         Begin VB.TextBox txtTelProveedor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            TabIndex        =   5
            Top             =   630
            Width           =   2130
         End
         Begin VB.TextBox txtFaxProveedor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            TabIndex        =   6
            Top             =   1005
            Width           =   2130
         End
         Begin VB.TextBox txtRucProveedor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            TabIndex        =   4
            Top             =   255
            Width           =   2130
         End
         Begin MSDataListLib.DataCombo dcmbCodP 
            Height          =   330
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtNomP 
            Height          =   315
            Left            =   960
            TabIndex        =   2
            Top             =   270
            Width           =   2565
         End
         Begin VB.Label lblCodProveedor 
            BackColor       =   &H00BAA892&
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
            Height          =   225
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   750
         End
         Begin VB.Label lblDirProveedor 
            BackColor       =   &H00BAA892&
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
            Height          =   300
            Left            =   120
            TabIndex        =   21
            Top             =   1012
            Width           =   885
         End
         Begin VB.Label lblTelProveedor 
            BackColor       =   &H00BAA892&
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
            Height          =   345
            Left            =   3645
            TabIndex        =   20
            Top             =   615
            Width           =   915
         End
         Begin VB.Label lblRucProveedor 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "RUC:"
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
            Height          =   255
            Left            =   3645
            TabIndex        =   19
            Top             =   285
            Width           =   480
         End
         Begin VB.Label lblFaxProveedor 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Height          =   240
            Left            =   3645
            TabIndex        =   18
            Top             =   1042
            Width           =   450
         End
      End
      Begin VB.TextBox txtObs 
         Height          =   930
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   6000
         Width           =   6855
      End
      Begin VB.TextBox txtNumIngreso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label lblObserv 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   5745
         Width           =   1410
      End
      Begin VB.Label lblFecha 
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
         Left            =   3885
         TabIndex        =   15
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lblNumIngreso 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Ingreso:"
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
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1245
      TabIndex        =   10
      Top             =   7350
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   2805
      TabIndex        =   11
      Top             =   7365
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4365
      TabIndex        =   12
      Top             =   7350
      Width           =   1455
   End
End
Attribute VB_Name = "frmIngImportacion"
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

Private clsConsu As New clsConsulta
Private clsCon_Def As New clsConsulta
Private clsCon_Aux As New clsConsulta
Private clsCon_Pro As New clsConsulta
Private clsExis As New clsConsulta
Private strSql As String

Private cargadoIXC As Boolean
Private strNumeroIXC As String
Private strTipoIXC As String

'Variables Globales
Public i_importacion As Integer

Private Sub btn_load_Click()
  
    vsfgDetalleImp.Clear 1
    vsfgDetalleImp.Rows = 2

    Dim i As Long
    
    strSql = " SELECT ing_codigo,for_pag_codigo,ing_fecha FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IXC' " & _
             " AND ing_factura='" & dcbo_importacion.BoundText & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cargadoIXC = True
        strNumeroIXC = clsCon_Def.adorec_Def("ing_codigo")
        dtpFecha.value = clsCon_Def.adorec_Def("ing_fecha")
        strTipoIXC = "IXC"
        'Me.CmbFpago.BoundText = clsAux.adorec_Def("for_pag_codigo")
        strSql = " SELECT dep_codigo,prd_codigo,det_ing_precio as precio,det_ing_cantidad as cant " & _
                 " FROM det_ingreso " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND ing_codigo=" & strNumeroIXC & " AND tip_ing_codigo='" & strTipoIXC & "' "
    Else
        cargadoIXC = False
        strNumeroIXC = "0"
        strTipoIXC = ""
        strSql = " SELECT contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,0 as precio,SUM(det_con_mer_cantidad) as cant " & _
                 " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
                 " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                 " INNER JOIN producto " & _
                 " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
                 " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
                 " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_contenedor_mercaderia.mov_codigo=" & strNumeroIXC & " AND det_contenedor_mercaderia.tip_mov_codigo='" & strTipoIXC & "' " & _
                 " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_importacion.BoundText & "'" & _
                 " GROUP BY contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,prd_nombre " & _
                 " HAVING SUM(det_con_mer_cantidad)!=0 " & _
                 " ORDER BY prd_nombre "
'
'        strSql = " SELECT contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,SUM(det_con_mer_cantidad) as cant " & _
'                 " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
'                 " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
'                 " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
'                 " INNER JOIN det_contenedor_mercaderia " & _
'                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
'                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
'                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
'                 " INNER JOIN producto " & _
'                 " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
'                 " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
'                 " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
'                 " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_importacion.BoundText & "%'" & _
'                 " GROUP BY contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,prd_nombre having SUM(det_con_mer_cantidad)!=0 " & _
'                 " ORDER BY prd_nombre "
    End If
    clsCon_Def.Ejecutar strSql
    i = 1
    While Not clsCon_Def.adorec_Def.EOF
        vsfgDetalleImp.TextMatrix(i, 1) = clsCon_Def.adorec_Def("dep_codigo")
        vsfgDetalleImp.TextMatrix(i, 2) = clsCon_Def.adorec_Def("prd_codigo")
        vsfgDetalleImp.TextMatrix(i, 4) = clsCon_Def.adorec_Def("cant")
        i = i + 1
        clsCon_Def.adorec_Def.MoveNext
    Wend
End Sub

Private Sub chk_importacion_Click()
If chk_importacion.value = 1 Then
  cmdCargarPed.Enabled = False
  dcmbPedido.Enabled = False
  btn_load.Enabled = True
  dcbo_importacion.Enabled = True
Else
  cmdCargarPed.Enabled = True
  dcmbPedido.Enabled = True
  btn_load.Enabled = False
  dcbo_importacion.Enabled = False
  vsfgDetalleImp.Clear
End If
End Sub

Private Sub cmdCargarPed_Click()
    Dim i As Long
    strSql = " SELECT 'TRA' as bod, prd_codigo,det_ped_imp_cantidad " & _
             " FROM det_pedido_imp " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_imp_codigo='" & dcmbPedido.BoundText & "' "
    clsCon_Def.Ejecutar strSql
    i = 1
    While Not clsCon_Def.adorec_Def.EOF
        vsfgDetalleImp.TextMatrix(i, 1) = clsCon_Def.adorec_Def("bod")
        vsfgDetalleImp.TextMatrix(i, 2) = clsCon_Def.adorec_Def("prd_codigo")
        vsfgDetalleImp.TextMatrix(i, 4) = clsCon_Def.adorec_Def("det_ped_imp_cantidad")
        i = i + 1
        clsCon_Def.adorec_Def.MoveNext
    Wend

End Sub

Private Sub dcbo_importacion_Change()
If (dcbo_importacion.BoundText <> "") Then
      i_importacion = CInt(dcbo_importacion.BoundText)
  End If
End Sub

Private Sub dcmbCodP_Validate(Cancel As Boolean)
        
    strSql = " SELECT rec_mer_codigo,CONCAT(rec_mer_codigo,' ()') as n " & _
             " FROM recepcion_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo like '" & dcmbCodP.BoundText & "' " & _
             " AND rec_mer_codigo NOT IN (" & _
                " SELECT DISTINCT ing_factura " & _
                " FROM ingreso " & _
                " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IXC' AND ing_anulado=0" & _
             " ) " & _
             " UNION" & _
             " SELECT rec_mer_codigo,CONCAT(rec_mer_codigo,' (',ing_codigo,')') as n " & _
             " FROM recepcion_mercaderia INNER JOIN ingreso " & _
             " ON recepcion_mercaderia.emp_codigo=ingreso.emp_codigo " & _
             " AND recepcion_mercaderia.per_codigo=ingreso.per_codigo " & _
             " AND recepcion_mercaderia.rec_mer_codigo=ingreso.ing_factura " & _
             " AND ingreso.tip_ing_codigo='IXC' and ing_anulado=0 " & _
             " WHERE recepcion_mercaderia.emp_codigo='" & strEmpresa & "' " & _
             " AND recepcion_mercaderia.per_codigo like '" & dcmbCodP.BoundText & "' " & _
             " ORDER BY rec_mer_codigo"
    
    clsCon_Def.Ejecutar strSql
    
    Set dcbo_importacion.RowSource = clsCon_Def.adorec_Def.DataSource
    dcbo_importacion.ListField = "n"
    dcbo_importacion.BoundColumn = "rec_mer_codigo"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_Def = Nothing
    Set clsCon_Aux = Nothing
    Set clsCon_Pro = Nothing
    Set clsExis = Nothing
End Sub



Private Sub cmdAceptar_Click()
    Dim x As Date
    Dim i As Long, j As Long
        
    Dim f As Date
    Dim d As String
    Dim m As String
    Dim Y As String
    Dim ff As String
    Dim ant As Long
    Dim cargado As Boolean
    
    Dim clsIngresoAux As New clsInventario

    If dcbo_importacion.BoundText <> "" And dcbo_importacion.Enabled = True Then
        cargado = True
    Else
        cargado = False
    End If

    clsExis.Inicializar AdoConn, AdoConnMaster

    'Valido los datos de # de Ingreso, Proveedor, fecha de Ingreso, etc.
    
    If (txtNumIngreso.Text = "") Then
        MsgBox "Número de Ingreso de Importación incorrecto", vbExclamation, "SisAdmi - Ingreso de Importación"
        cmdNuevo.SetFocus
        Exit Sub
    End If
    If (dcmbCodP.BoundText = "" Or txtNomP.Text = "" Or txtRucProveedor.Text = "") Then
        MsgBox "Datos del Proveedor incorrectos, verifíquelos", vbExclamation, "SisAdmi - Ingreso de Importación"
        dcmbCodP.SetFocus
        Exit Sub
    End If
           
    ff = Format(dtpFecha.value, "yyyy-mm-dd")
    'Verifica si la fecha ingresada es correcta
    If (IsDate(ff)) = False Then
        MsgBox "La fecha de Ingreso no es correcta", vbExclamation, "SisAdmi - Ingreso de Importación"
        Exit Sub
    End If
    'valido que no haga filas vacias
    band = 0
    
    For i = 1 To vsfgDetalleImp.Rows - 1
        band = 0
        For j = 1 To vsfgDetalleImp.Cols - 1
            If vsfgDetalleImp.TextMatrix(i, j) = "" Then band = band + 1
        Next j
        
        If band = vsfgDetalleImp.Cols - 2 Then
            ant = vsfgDetalleImp.Rows
            vsfgDetalleImp.RemoveItem i
            If vsfgDetalleImp.Rows = ant Then
                vsfgDetalleImp.RemoveItem i
            End If
        End If
        
    Next i
    'Verifica que existan datos en el FlexGrid
    If vsfgDetalleImp.Rows = 1 Then
    
        MsgBox "El ingreso no tiene detalle", vbExclamation, "SisAdmi - Ingreso de Importación"
        vsfgDetalleImp.AddItem ""
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
        vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
        If vsfgDetalleImp.Rows > 2 Then
            vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) = vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 2, 1)
        End If
   Else
       
       
       'Verifica que existan datos en el FlexGrid
        For i = 1 To vsfgDetalleImp.Rows - 1
             If (vsfgDetalleImp.TextMatrix(i, 1) = "" Or vsfgDetalleImp.TextMatrix(i, 2) = "") And i < vsfgDetalleImp.Rows - 1 Then
                If i = 1 Then
                    MsgBox "El ingreso no tiene detalle", vbExclamation, "SisAdmi - Ingreso de Importación"
                    Exit Sub
                End If
    
                For j = 1 To vsfgDetalleImp.Cols - 1
                    If (vsfgDetalleImp.TextMatrix(i, j) = "" Or vsfgDetalleImp.TextMatrix(2, j) = "") Then
                        MsgBox "Dato incorrecto en: " & vsfgDetalleImp.TextMatrix(0, j) & " ,fila: " & i, vbExclamation, "SisAdmi - Ingreso de Importación"
                        Exit Sub
                    End If
                Next j
    
             End If
        Next i
    
        
           
        i = vsfgDetalleImp.Rows
        If (TxtTotal.Text = "") Then
            Exit Sub
        End If
        If (i - 1 <> 0) Then ' Si existen detalles, almaceno.
        
            Mensaje = "Existen " & i - 1 & " detalle(s) en el ingreso, desea guardar?" ' Define el mensaje.
            Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
            Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
            respuesta = MsgBox(Mensaje, Estilo, Título)
            
            'Recorro el FlexGrid para almacenar los detalles del ingreso
            If respuesta = vbYes Then
                Dim aux As Double
                
                strSql = "select COALESCE(max(ing_codigo),0) as t from ingreso where emp_codigo = '" & strEmpresa & "' and tip_ing_codigo = 'IIM'" & _
                         " GROUP BY emp_codigo"
                
                clsConsu.Ejecutar (strSql)
        
                If (IsNull(clsConsu.adorec_Def.Fields(0).value)) Then
                
                    aux = 1
                    
                Else
                    aux = clsConsu.adorec_Def.Fields(0).value + 1
            
                End If
                
                txtNumIngreso.Text = aux
    
                Dim clsIngreso As New clsInventario
                clsIngreso.Inicializar AdoConn, AdoConnMaster
                clsIngreso.NuevoIng IIf(cargado = True, False, True), "IIM", False, , , , , dcmbCodP.BoundText, ff, , , txtObs.Text, , , , , , , , , , , , , IIf(cargadoIXC = True, strNumeroIXC, "")
                
                For aux = 1 To i - 1
                    clsIngreso.NuevoDetIng vsfgDetalleImp.TextMatrix(aux, 2), vsfgDetalleImp.TextMatrix(aux, 1), vsfgDetalleImp.TextMatrix(aux, 4), vsfgDetalleImp.TextMatrix(aux, 5), vsfgDetalleImp.TextMatrix(aux, 5)
                Next aux
                
                
                If cargadoIXC = True Then
                    clsIngresoAux.Inicializar AdoConn, AdoConnMaster
                    clsIngresoAux.AnularIng strNumeroIXC, "IXC", , "CONTABILIZADO EN " & clsIngreso.strTipo & " " & clsIngreso.strDoc
                End If
                If cargado = True And cargadoIXC = False Then
                
                    strSql = " UPDATE det_contenedor_mercaderia " & _
                             " SET det_contenedor_mercaderia.tip_mov_codigo='" & clsIngreso.strTipo & "'," & _
                             " det_contenedor_mercaderia.mov_codigo='" & clsIngreso.strDoc & "'" & _
                             " FROM det_recepcion_mercaderia, contenedor_mercaderia,det_contenedor_mercaderia WHERE det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                             " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                             " AND contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                             " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                             " AND det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                             " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_importacion.BoundText & "'"
                    clsCon_Def.Ejecutar strSql
                End If
                
                InicializarContenedorRecurrente
                MsgBox "Ingreso almacenado", vbInformation, "SisAdmi - Ingreso de Importación"
                
                Call actualizar_importacion(i_importacion)
                
                Dim rpMov As New frmReporte
                rpMov.strNumero = clsIngreso.strDoc
                rpMov.strTipo = clsIngreso.strTipo
                rpMov.strReporte = "rptIngresoMercaderia"
                rpMov.Show
                Call CmdSalir_Click
            
                    
            End If
            
        End If
    End If
    
End Sub

Private Sub actualizar_importacion(p_actualizar As Integer)


' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        'Se conecta a la base de SQL Server 2005
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "UPDATE importacion set imp_estado=" & 2 & " where imp_id = " & p_actualizar, cnn, adOpenDynamic, adLockOptimistic
    End If

End Sub


Private Sub cmdNuevo_Click()
'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    clsCon_Pro.Inicializar AdoConn, AdoConnMaster
    
    
    Dim var As Long
    
    dtpFecha.value = HoyDia
    
    'Consulta del nùmero de ingreso último, se agrega uno para el nuevo ingreso
    strSql = "select COALESCE(max(ing_codigo),0) as t from ingreso where emp_codigo = '" & strEmpresa & "' and tip_ing_codigo = 'IIM'" & _
             " GROUP BY emp_codigo"
    clsConsu.Ejecutar (strSql)
    txtNumIngreso.Text = ""
    If (IsNull(clsConsu.adorec_Def.Fields(0).value)) Then
    
        txtNumIngreso.Text = "1"
        
    Else
        txtNumIngreso.Text = clsConsu.adorec_Def.Fields(0).value + 1
        
    End If
        
    txtRucProveedor.Text = ""
    txtDirProveedor.Text = ""
    txtTelProveedor.Text = ""
    txtNomP.Text = ""
    txtFaxProveedor.Text = ""
    txtObs.Text = ""
    TxtTotal.Text = ""
    dcmbCodP.Text = ""
    
    'limpia el FlexGrid

    Call limpiarFxGD
         
    'Ejecuta un SQL contra la base de datos para consultar los proveedores.
    strSql = " select per_codigo, CONCAT(per_apellido,' ',per_nombre) as per_nombre, " & _
             " per_apellido, per_ruc, per_direccion, " & _
             " per_telf, per_fax from persona " & _
             " where emp_codigo= '" & strEmpresa & "' and cat_p_tipo = 'P' " & _
             " order by per_apellido,per_nombre"
    clsConsu.Ejecutar (strSql), "M"
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Proveedores ingresados en el Sistema", vbInformation, "SisAdmi - Ingreso de Importación"
        Exit Sub
    Else
        Set dcmbCodP.RowSource = clsConsu.adorec_Def.DataSource
        dcmbCodP.ListField = "per_nombre"
        dcmbCodP.BoundColumn = "per_codigo"
    End If
    strSql = "select dep_codigo, dep_nombre from deposito where emp_codigo = '" & strEmpresa & "' "
    clsCon_Aux.Ejecutar (strSql)
    
    If (clsCon_Aux.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Depósitos creados", vbInformation, "SisAdmi - Ingreso de Importaciones"
        Exit Sub
    Else
        'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgDetalleImp
    
        vsfgDetalleImp.ColComboList(1) = vsfgDetalleImp.BuildComboList(clsCon_Aux.adorec_Def, "dep_nombre, *dep_codigo", "dep_codigo")
       
    End If
    
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 order by prd_nombre"
    clsCon_Pro.Ejecutar (strSql)
    vsfgDetalleImp.ColComboList(3) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    
    
    'Consulto los productos de la empresa
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 order by prd_codigo"
    clsCon_Pro.Ejecutar (strSql)
    
    If (clsCon_Pro.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Productos creados", vbInformation, "SisAdmi - Ingreso de Importaciones"
        Exit Sub
    Else
   
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(2) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    End If
    
    
End Sub

Private Sub CmdSalir_Click()
   Unload Me
   frmVerIngImp.Show
         
End Sub



Private Sub dcmbCodP_Change()
On Error GoTo errhandler
 'Muestra el nombre del proveedor relacionado con el código seleccionado
' o ingresado en el combo Proveedores al momento de hacer un cambio en el combo
If dcmbCodP.Text = "" Then
Exit Sub
End If
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "per_codigo = '" & dcmbCodP.BoundText & "'", , adSearchForward
   
    If clsConsu.adorec_Def.EOF = False Then
        'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
        txtNomP.Text = clsConsu.adorec_Def("per_apellido") & " " & clsConsu.adorec_Def("per_nombre")
        
        txtRucProveedor.Text = clsConsu.adorec_Def("per_ruc")
        txtDirProveedor.Text = clsConsu.adorec_Def("per_direccion")
        txtTelProveedor.Text = clsConsu.adorec_Def("per_telf")
        txtFaxProveedor.Text = clsConsu.adorec_Def("per_fax")
        strSql = " SELECT ped_imp_codigo,CONCAT(ped_imp_codigo,' - ',LEFT(ped_imp_fecha_ped,10)) as ped_imp_nom " & _
                 " FROM pedido_importacion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & dcmbCodP.BoundText & "' " & _
                 " AND ped_imp_estado IN ('AD','TL') "
        clsCon_Def.Ejecutar strSql
        dcmbPedido.ListField = "ped_imp_nom"
        dcmbPedido.BoundColumn = "ped_imp_codigo"
        Set dcmbPedido.RowSource = clsCon_Def.adorec_Def.DataSource
    Else
        'MsgBox "No existe el Proveedor ingresado", vbInformation, "SisAdmi - PROVEEDOR"
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        TxtTotal.Text = ""
    End If
    
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

Private Sub Form_Activate()
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 ORDER BY prd_nombre"
    clsCon_Pro.Ejecutar (strSql)
    
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(3) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    
    
    'Consulto los productos de la empresa
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 ORDER BY prd_codigo"
    clsCon_Pro.Ejecutar (strSql)
    
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(2) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
End Sub

Private Sub Form_Load()
Dim var As Long
' Objetos de conexion para SQL Server

    Me.Width = 8125
    Me.Height = 8400
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    clsCon_Pro.Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    'txtFechaIng.Text = format(HoyDia, "dd/mm/yyyy")
    
    'Descompone la fecha actual  en día, mes y año
    dtpFecha.value = HoyDia
    
    'Consulta del nùmero de ingreso último, se agrega uno para el nuevo ingreso
    strSql = "select COALESCE(max(ing_codigo),0) as num from ingreso where emp_codigo = '" & strEmpresa & "' and tip_ing_codigo = 'IIM'" & _
             " GROUP BY emp_codigo"
    clsConsu.Ejecutar (strSql)
    If clsConsu.adorec_Def.EOF Then
        txtNumIngreso.Text = "1"
    Else
        txtNumIngreso.Text = clsConsu.adorec_Def("num") + 1
        
    End If
    txtNumIngreso.Enabled = False
    txtRucProveedor.Enabled = False
    txtDirProveedor.Enabled = False
    txtTelProveedor.Enabled = False
    txtNomP.Enabled = False
    txtFaxProveedor.Enabled = False
       
    'Ejecuta un SQL contra la base de datos
    strSql = " select per_codigo, CONCAT(per_apellido, ' ',per_nombre) as per_nombre, " & _
             " per_apellido, per_ruc, per_direccion, " & _
             " per_telf, per_fax from persona " & _
             " where emp_codigo= '" & strEmpresa & "' and cat_p_tipo = 'P' " & _
             " order by per_apellido,per_nombre"
    clsConsu.Ejecutar (strSql)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Proveedores ingresados en el Sistema", vbInformation, "SisAdmi"
        Exit Sub
    Else
        Set dcmbCodP.RowSource = clsConsu.adorec_Def.DataSource
        dcmbCodP.ListField = "per_nombre"
        dcmbCodP.BoundColumn = "per_codigo"
    End If
    strSql = "select dep_codigo, dep_nombre from deposito where emp_codigo = '" & strEmpresa & "' "
    clsCon_Aux.Ejecutar (strSql)
    
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgDetalleImp
    'vsfgGrupo.BuildComboList(clsCon_Def.adorec_Def, "*gru_nombre, gru_codigo", "gru_nombre")
    vsfgDetalleImp.ColComboList(1) = vsfgDetalleImp.BuildComboList(clsCon_Aux.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 ORDER BY prd_nombre"
    clsCon_Pro.Ejecutar (strSql)
    
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(3) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    
    
    'Consulto los productos de la empresa
    strSql = "select prd_codigo, prd_nombre,prd_costo from producto where emp_codigo = '" & strEmpresa & "' and prd_baja=0 ORDER BY prd_codigo"
    clsCon_Pro.Ejecutar (strSql)
    
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(2) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    'Insertamos el botón de eliminar en cada una de las filas
    
    ' initializa el flexgrid
    vsfgDetalleImp.Editable = flexEDKbdMouse
    vsfgDetalleImp.AllowUserResizing = flexResizeBoth
    
    ' Agrega un botón en el grid
    
    vsfgDetalleImp.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalleImp.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
       
        End Select
End Sub

Private Sub txtFechaIng_KeyPress(KeyAscii As Integer)
    'Validación de caracteres ingresados para que solo ingrese números y el caracter "/"
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub vsfgDetalleImp_EnterCell()
    Dim Col As Long
    Dim Row As Long
    Col = vsfgDetalleImp.Col
    Row = vsfgDetalleImp.Row
    ' Aumenta una fila en caso de ser necesario.
    
    'MsgBox CStr(Col) & vsfgDetalleImp.TextMatrix(Row, 4)
  
    
    Dim SumaT As Long, c As Long
        SumaT = 0
        For c = 1 To vsfgDetalleImp.Rows - 1
        SumaT = SumaT + Val(vsfgDetalleImp.TextMatrix(c, 4))
        'MsgBox Val(vsfgDetalleImp.TextMatrix(1, Col))
        '= Format(CStr(Val(TxtTotal.Text) + Val(vsfgDetalleImp.TextMatrix(Row, Col))), "####0.00")
        Next c
        TxtTotal.Text = FormatoD2(SumaT)
        
    If (vsfgDetalleImp.Row = (vsfgDetalleImp.Rows - 1) And Trim(vsfgDetalleImp.TextMatrix(Row, 4)) <> "" And vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 3) <> "") Then
        vsfgDetalleImp.AddItem ""
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
        vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
        If vsfgDetalleImp.Rows > 2 Then
             vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) = vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 2, 1)
        End If
   End If
   
   
End Sub

Private Sub vsfgDetalleImp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleImp.MouseRow
    c = vsfgDetalleImp.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = (vsfgDetalleImp.Rows - 1)) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleImp.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
   
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleImp.Cell(flexcpLeft, r, c) + vsfgDetalleImp.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleImp.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "AHORA DEBE ELIMINAR ESTA FILA!"
    
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        
        'Recorro el FlexGrid para almacenar los detalles del ingreso
        
        If respuesta = vbYes Then
            Dim i As Long
        
            TxtTotal.Text = FormatoD2(FormatoD2(TxtTotal.Text) - FormatoD2(vsfgDetalleImp.TextMatrix(r, 4)))
            vsfgDetalleImp.RemoveItem (r)
            For i = 1 To (vsfgDetalleImp.Rows - 1)
                vsfgDetalleImp.TextMatrix(i, 0) = i
                vsfgDetalleImp.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalleImp.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            Next i
        Else
            vsfgDetalleImp.Cell(flexcpPicture, r, c) = imgBtnUp
        
        End If
    
        
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub vsfgDetalleImp_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'para que no se pueda escribir en las columnas que se indica
'''  If NewCol = 3 Then
'''        If Abs(NewCol - OldCol) = 1 Then
'''            If NewCol > OldCol Then
'''                SendKeys vbKeyTab
'''            Else
'''                SendKeys vbkeyleft
'''            End If
'''        Else
'''            Cancel = True
'''        End If
'''    End If
End Sub

Private Sub vsfgDetalleImp_CellChanged(ByVal Row As Long, ByVal Col As Long)

    If (Col = 2 And vsfgDetalleImp.TextMatrix(Row, Col) <> "") Then
        clsCon_Pro.adorec_Def.MoveFirst
        clsCon_Pro.adorec_Def.Find "prd_codigo = '" & vsfgDetalleImp.TextMatrix(Row, Col) & "' ", , adSearchForward
        If (clsCon_Pro.adorec_Def.EOF = False) Then
            vsfgDetalleImp.TextMatrix(Row, Col + 1) = vsfgDetalleImp.TextMatrix(Row, Col)
            vsfgDetalleImp.TextMatrix(Row, 5) = clsCon_Pro.adorec_Def("prd_costo")
        End If
    ElseIf (Col = 3 And vsfgDetalleImp.TextMatrix(Row, Col) <> "") Then
        clsCon_Pro.adorec_Def.MoveFirst
        clsCon_Pro.adorec_Def.Find "prd_codigo = '" & vsfgDetalleImp.TextMatrix(Row, Col) & "' ", , adSearchForward
        If (clsCon_Pro.adorec_Def.EOF = False) Then
            vsfgDetalleImp.TextMatrix(Row, Col - 1) = vsfgDetalleImp.TextMatrix(Row, Col)
            vsfgDetalleImp.TextMatrix(Row, 5) = clsCon_Pro.adorec_Def("prd_costo")
        End If
    End If
    If Col = 4 And vsfgDetalleImp.TextMatrix(Row, Col) <> "" Then
        'vsfgDetalleImp.TextMatrix(Row, Col + 2) = CStr(Val(clsCon_Pro.adorec_Def("prd_costo")) * Val(vsfgDetalleImp.TextMatrix(Row, Col)))
        Dim SumaT As Long
        SumaT = 0
        For c = 1 To vsfgDetalleImp.Rows - 1
        SumaT = SumaT + Val(vsfgDetalleImp.TextMatrix(c, 4))
       '  MsgBox Val(vsfgDetalleImp.TextMatrix(c, Col))
        '= Format(CStr(Val(TxtTotal.Text) + Val(vsfgDetalleImp.TextMatrix(Row, Col))), "####0.00")
        Next c
        TxtTotal.Text = FormatoD2(SumaT)
        'aumenta una fila mas si hace falta
        If vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, Col) <> "" And vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, Col - 1) <> "" Then
            vsfgDetalleImp.AddItem ""
            vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
            vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
            vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
            If vsfgDetalleImp.Rows > 2 Then
                 vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) = vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 2, 1)
            End If
        End If
    End If

End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Long
    vsfgDetalleImp.Tag = "N"
    
    vsfgDetalleImp.Clear 1
    vsfgDetalleImp.Rows = 2
    vsfgDetalleImp.Tag = "T"
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Se da un tab al presionar enter para que al ingresar un dato pase al siguiente campo
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If

End Sub

Private Sub vsfgDetalleImp_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'Valido que solo se pueda dar enter en el campo Desc. Producto
'''    If (Col = 3) Then
'''        If KeyAscii <> 13 Then
'''            KeyAscii = 0
'''        End If
'''    End If
    
    
    
    'Valido que solo se pueda ingresar números  en el campo cantidad
    
    If Col = 4 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfgDetalleImp_RowColChange()
    'Se envía un espacio en blanco al recorrer el flexGrid para desplegar los combos que existan
    'SendKeys " "
End Sub



