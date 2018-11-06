VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecepcionMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Recepción de Mercaderia"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15960
   Icon            =   "frmRecepcionMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15960
   Begin VB.CheckBox chkCrearIngresoMercaderia 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Crear Ingreso de Mercaderia"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Orden de Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   7425
      Begin VB.TextBox txtFormaPago 
         Height          =   285
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   960
         Width           =   2400
      End
      Begin VB.TextBox txtObservacionOrdenCompra 
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1320
         Width           =   6000
      End
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   2400
      End
      Begin VB.TextBox txtFechaEntrega 
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   2400
      End
      Begin VB.TextBox txtFechaEnvio 
         Height          =   285
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   2400
      End
      Begin VB.TextBox txtFechaOrden 
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   2400
      End
      Begin MSDataListLib.DataCombo cmbOrdenCompra 
         Height          =   330
         Left            =   1305
         TabIndex        =   16
         Top             =   240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. de Pago:"
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
         Left            =   4035
         TabIndex        =   32
         Top             =   990
         Width           =   810
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         TabIndex        =   27
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Left            =   4260
         TabIndex        =   25
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. de Entrega:"
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
         Left            =   195
         TabIndex        =   23
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. de Envio:"
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
         Left            =   3960
         TabIndex        =   21
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. de Orden:"
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
         Left            =   300
         TabIndex        =   19
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ord. Compra:"
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
         TabIndex        =   17
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCrearRecepcion 
      Caption         =   "&Crear Recepción"
      Height          =   360
      Left            =   4080
      TabIndex        =   2
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5880
      TabIndex        =   1
      Top             =   8400
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Recepción de Mercadería"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
      Begin VB.TextBox txtFac 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   2760
      End
      Begin VB.TextBox TxtObser 
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   6000
      End
      Begin MSComCtl2.DTPicker dtpFecha 
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
         Left            =   4755
         TabIndex        =   5
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   16842755
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbProveedor 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Fact / Aux:"
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
         Left            =   135
         TabIndex        =   12
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   420
         TabIndex        =   10
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
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
         Left            =   4215
         TabIndex        =   6
         Top             =   660
         Width           =   495
      End
      Begin VB.Label LblObser 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3360
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   7380
      _cx             =   13017
      _cy             =   5927
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRecepcionMercaderia.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG1 
      Height          =   8640
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   8220
      _cx             =   14499
      _cy             =   15240
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
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRecepcionMercaderia.frx":042C
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
   Begin MSDataListLib.DataCombo cmbEstadoOrdenCompra 
      Height          =   330
      Left            =   1425
      TabIndex        =   28
      Top             =   7920
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
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
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo Estado:"
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
      Left            =   270
      TabIndex        =   29
      Top             =   7980
      Width           =   1050
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Left            =   4320
      TabIndex        =   8
      Top             =   4155
      Width           =   555
   End
End
Attribute VB_Name = "frmRecepcionMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String

Private Sub cmbBodega_Validate(Cancel As Boolean)
    CargaUbica
End Sub

Private Sub CargaUbica()
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & cmbBodega.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsCon_Def.Ejecutar strSql
    Set cmbUbicacion.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbUbicacion.ListField = "ubi_bod_codigo"
    cmbUbicacion.BoundColumn = "ubi_bod_codigo"
End Sub

Private Sub cmbOrdenCompra_Validate(Cancel As Boolean)
    CargaOrdenCompra
    CargaContenedores
End Sub

Private Sub CargaOrdenCompra()
    Dim clsAux As New clsConsulta
    Dim codigoEstado As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    If cmbOrdenCompra.BoundText <> "" Then
        strSql = " SELECT ord_com_codigo,est_ord_com_descripcion,orden_compra.est_ord_com_codigo,ord_com_fecha,ord_com_fecha_envio,ord_com_fecha_entrega,ord_com_observacion, " & _
                 " forma_pago.for_pag_codigo,forma_pago.for_pag_nombre " & _
                 " FROM orden_compra INNER JOIN est_orden_compra ON orden_compra.est_ord_com_codigo=est_orden_compra.est_ord_com_codigo " & _
                 " INNER JOIN forma_pago ON orden_compra.emp_codigo=forma_pago.emp_codigo " & _
                 " AND orden_compra.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " WHERE orden_compra.emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & cmbOrdenCompra.BoundText & "'"
        clsAux.Ejecutar strSql
        If clsAux.adorec_Def.RecordCount > 0 Then
            txtEstado.Text = clsAux.adorec_Def("est_ord_com_descripcion")
            txtFechaOrden.Text = clsAux.adorec_Def("ord_com_fecha")
            txtFechaEnvio.Text = clsAux.adorec_Def("ord_com_fecha_envio")
            txtFechaEntrega.Text = clsAux.adorec_Def("ord_com_fecha_entrega")
            txtFormaPago.Text = clsAux.adorec_Def("for_pag_nombre")
            txtFormaPago.Tag = clsAux.adorec_Def("for_pag_codigo")
            txtObservacionOrdenCompra.Text = clsAux.adorec_Def("ord_com_observacion")
            codigoEstado = clsAux.adorec_Def("est_ord_com_codigo")
            strSql = " SELECT est_ord_com_codigo,est_ord_com_descripcion " & _
                     " FROM est_orden_compra WHERE est_ord_com_codigo>=10 or est_ord_com_codigo='" & codigoEstado & "'"
            clsAux.Ejecutar strSql
            cmbEstadoOrdenCompra.ListField = "est_ord_com_descripcion"
            cmbEstadoOrdenCompra.BoundColumn = "est_ord_com_codigo"
            Set cmbEstadoOrdenCompra.RowSource = clsAux.adorec_Def.DataSource
            'cmbEstadoOrdenCompra.Text = txtEstado.Text
            cmbEstadoOrdenCompra.BoundText = codigoEstado
        Else
            txtEstado.Text = ""
            txtFechaOrden.Text = ""
            txtFechaEnvio.Text = ""
            txtFechaEntrega.Text = ""
            txtFormaPago.Text = ""
            txtFormaPago.Tag = ""
            txtObservacionOrdenCompra.Text = ""
        End If
    End If
End Sub

Private Sub cmbProveedor_Validate(Cancel As Boolean)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ord_com_codigo " & _
             " FROM orden_compra " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo='" & cmbProveedor.BoundText & "'" & _
             " AND est_ord_com_codigo in (0,1,2,10,11,12)"
    clsAux.Ejecutar strSql
    cmbOrdenCompra.ListField = "ord_com_codigo"
    cmbOrdenCompra.BoundColumn = "ord_com_codigo"
    
    Set cmbOrdenCompra.RowSource = clsAux.adorec_Def.DataSource
    CargaContenedores
End Sub

Private Sub cmdCrearRecepcion_Click()
    Dim num As String
    Dim i As Long
    Dim clsAux As New clsConsulta
    Dim clsIngreso As New clsInventario
    clsAux.Inicializar AdoConn, AdoConnMaster
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    strSql = " BEGIN TRAN "
    clsAux.Ejecutar strSql, "M"
    strSql = " SELECT COALESCE(MAX(rec_mer_codigo),0)+1 as n " & _
             " FROM recepcion_mercaderia WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "'"
    clsAux.Ejecutar strSql, "M"
    num = 1
    If clsAux.adorec_Def.RecordCount > 0 Then
        num = clsAux.adorec_Def("n")
    End If

    strSql = " INSERT INTO recepcion_mercaderia(emp_codigo,rec_mer_codigo,per_codigo,est_rec_mer_codigo," & _
             " ord_com_codigo,rec_mer_factura,rec_mer_fecha,rec_mer_observacion,rec_mer_fechamod,rec_mer_usumod)" & _
             " VALUES('" & strEmpresa & "','" & num & "','" & cmbProveedor.BoundText & "','1'," & _
             " '" & FormatoD0(cmbOrdenCompra.BoundText) & "','" & UCase(txtFac.Text) & "','" & dtpFecha.value & "','" & UCase(TxtObser.Text) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar strSql, "M"
    strSql = " COMMIT TRAN "
    clsAux.Ejecutar strSql, "M"
    With VSFG
        For i = 1 To .Rows - 1
            If Abs(.TextMatrix(i, 0)) = 1 Then
                strSql = " INSERT INTO det_recepcion_mercaderia (emp_codigo,rec_mer_codigo,con_mer_codigo," & _
                         " det_rec_mer_fechamod,det_rec_mer_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & num & "','" & .TextMatrix(i, 1) & "', " & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsAux.Ejecutar strSql, "M"
                strSql = "UPDATE contenedor_mercaderia SET est_con_mer_codigo=1 WHERE emp_codigo='" & strEmpresa & "' AND con_mer_codigo='" & .TextMatrix(i, 1) & "' AND est_con_mer_codigo=0"
                clsAux.Ejecutar strSql, "M"
            End If
        Next i
    End With
    
    If cmbOrdenCompra.BoundText <> "" Then
        strSql = " UPDATE orden_compra SET est_ord_com_codigo='" & cmbEstadoOrdenCompra.BoundText & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_codigo='" & cmbOrdenCompra.BoundText & "'"
        clsAux.Ejecutar strSql, "M"
    End If
    
    If chkCrearIngresoMercaderia.value = 1 Then
        If MsgBox("Desea crear Ingreso de Mercaderia por Contabilizar?", vbQuestion + vbYesNo, "Inventario") = vbYes Then
            clsIngreso.NuevoIng False, "IXC", False, PtoEmiDocEle, strPtoFactura, , txtFormaPago.Tag, cmbProveedor.BoundText, dtpFecha.value, num, , UCase(TxtObser.Text) & vbNewLine & UCase(txtObservacionOrdenCompra), , , , VSFG1.TextMatrix(VSFG1.Rows - 3, VSFG1.Cols - 1), , , , VSFG1.TextMatrix(VSFG1.Rows - 3, VSFG1.Cols - 1)
            With VSFG1
                For i = 1 To .Rows - 4
                    If Abs(.TextMatrix(i, 3)) > 0 Then
                        clsIngreso.NuevoDetIng .TextMatrix(i, 1), .TextMatrix(i, 0), .TextMatrix(i, 3), .TextMatrix(i, 7)
                    End If
                Next i
            End With
            
            strSql = " UPDATE det_contenedor_mercaderia " & _
                     " SET det_contenedor_mercaderia.tip_mov_codigo='" & clsIngreso.strTipo & "'," & _
                     " det_contenedor_mercaderia.mov_codigo='" & clsIngreso.strDoc & "'" & _
                     " FROM det_recepcion_mercaderia ,contenedor_mercaderia,det_contenedor_mercaderia " & _
                     " WHERE det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " AND contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " AND det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & num & "'"
            clsCon_Def.Ejecutar strSql
            clsIngreso.RevisarProductosEnPedidosReprogramados
        End If
    End If
    
    Unload Me
    
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
    
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ', per_nombre) as nomb " & _
             " FROM persona WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='P' " & _
             " ORDER BY nomb "
    clsCon_Def.Ejecutar strSql
    Set cmbProveedor.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbProveedor.ListField = "nomb"
    cmbProveedor.BoundColumn = "per_codigo"
    dtpFecha.value = Ahora
    VSFG1.MergeCells = flexMergeRestrictRows
    VSFG1.MergeCol(0) = True
    VSFG1.Subtotal flexSTSum, -1, 2, , vbYellow, , True, "TOTAL"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarContenedor UCase(txtLector.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub AgregarContenedor(codigo As String)
    Dim i As Long
    Dim pas As Boolean
    pas = False
    With VSFG
        For i = 1 To .Rows - 1
            If codigo = .TextMatrix(i, 1) Then
                .ShowCell i, 0
                .Select i, 0
                If Abs(.TextMatrix(i, 0)) = 1 Then
                    .TextMatrix(i, 0) = 0
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
                Else
                    .TextMatrix(i, 0) = 1
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbYellow
                End If
                pas = True
                Exit For
            End If
        Next i
        If pas = False Then
            MsgBox "El contenedor " & codigo & " no esta disponible para ser asignado a esta Recepción"
        Else
            CargarDetalle
        End If
        '.ShowCell 1, 2
    End With
End Sub

Private Sub CargarDetalle()
    Dim strContenedores As String
    Dim i As Long
    With VSFG
        For i = 1 To .Rows - 1
            If Abs(.TextMatrix(i, 0)) = 1 Then
                strContenedores = strContenedores & .TextMatrix(i, 1) & ","
            End If
        Next i
        If Len(strContenedores) > 0 Then
            strContenedores = Left(strContenedores, Len(strContenedores) - 1)
        Else
            strContenedores = "''"
        End If
    End With
    VSFG1.Clear 1
    VSFG1.Rows = 3
    If Len(strContenedores) > 2 Then
        If cmbOrdenCompra.BoundText <> "" Then
            strSql = " SELECT dep_codigo,producto.prd_codigo,prd_nombre,COALESCE(cantContenedores,0),det_ord_com_cantidad,COALESCE(cantRecibida,0),det_ord_com_cantidad-(COALESCE(cantContenedores,0)+COALESCE(cantRecibida,0)),det_ord_com_precio,cantContenedores*det_ord_com_precio " & _
                     " FROM det_orden_compra INNER JOIN preproducto_producto ON det_orden_compra.emp_codigo=preproducto_producto.emp_codigo " & _
                     " AND det_orden_compra.pre_codigo=preproducto_producto.pre_codigo " & _
                     " AND det_orden_compra.col_codigo=preproducto_producto.col_codigo " & _
                     " AND det_orden_compra.tal_codigo=preproducto_producto.tal_codigo " & _
                     " INNER JOIN producto " & _
                     " ON preproducto_producto.emp_codigo=producto.emp_codigo " & _
                     " AND preproducto_producto.prd_codigo=producto.prd_codigo "
            strSql = strSql & " LEFT JOIN (" & _
                     " SELECT contenedor_mercaderia.emp_codigo,dep_codigo,prd_codigo,SUM(det_con_mer_cantidad) as cantContenedores " & _
                     " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                     " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND contenedor_mercaderia.con_mer_codigo in (" & strContenedores & ")" & _
                     " GROUP BY contenedor_mercaderia.emp_codigo,dep_codigo,prd_codigo " & _
                     " ) cm ON producto.emp_codigo=cm.emp_codigo AND producto.prd_codigo=cm.prd_codigo"
            strSql = strSql & " LEFT JOIN (" & _
                     " SELECT recepcion_mercaderia.emp_codigo,prd_codigo,SUM(det_con_mer_cantidad) as cantRecibida " & _
                     " FROM recepcion_mercaderia INNER JOIN det_recepcion_mercaderia " & _
                     " ON recepcion_mercaderia.emp_codigo=det_recepcion_mercaderia.emp_codigo " & _
                     " AND recepcion_mercaderia.rec_mer_codigo=det_recepcion_mercaderia.rec_mer_codigo " & _
                     " INNER JOIN contenedor_mercaderia ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " INNER JOIN det_contenedor_mercaderia " & _
                     " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " WHERE recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND recepcion_mercaderia.ord_com_codigo='" & cmbOrdenCompra.BoundText & "' " & _
                     " GROUP BY recepcion_mercaderia.emp_codigo,prd_codigo " & _
                     " ) rm ON producto.emp_codigo=rm.emp_codigo AND producto.prd_codigo=rm.prd_codigo"
            strSql = strSql & " WHERE det_orden_compra.emp_codigo='" & strEmpresa & "' " & _
                     " AND det_orden_compra.ord_com_codigo='" & cmbOrdenCompra.BoundText & "' "
        Else
            strSql = " SELECT dep_codigo,producto.prd_codigo,prd_nombre,SUM(det_con_mer_cantidad),0,0,0,0,0 " & _
                     " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                     " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " INNER JOIN producto " & _
                     " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
                     " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
                     " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND contenedor_mercaderia.con_mer_codigo in (" & strContenedores & ")" & _
                     " GROUP BY dep_codigo,producto.prd_codigo,prd_nombre "
        End If
        clsCon_Def.Ejecutar strSql
        Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
    End If
    VSFG1.MergeCells = flexMergeRestrictRows
    VSFG1.MergeCol(0) = True
    VSFG1.Subtotal flexSTSum, -1, 3, , vbYellow, , True, "TOTAL"
    VSFG1.Subtotal flexSTSum, -1, 4, , vbYellow, , True, "TOTAL"
    VSFG1.Subtotal flexSTSum, -1, 5, , vbYellow, , True, "TOTAL"
    VSFG1.Subtotal flexSTSum, -1, 8, , vbYellow, , True, "TOTAL"
    VSFG1.TextMatrix(VSFG1.Rows - 1, 7) = "SubTotal"
    VSFG1.AddItem ""
    VSFG1.TextMatrix(VSFG1.Rows - 1, 8) = FormatoD2(VSFG1.TextMatrix(VSFG1.Rows - 2, 8) * PorIVA / 100)
    VSFG1.TextMatrix(VSFG1.Rows - 1, 7) = "IVA"
    VSFG1.AddItem ""
    VSFG1.TextMatrix(VSFG1.Rows - 1, 8) = FormatoD2(VSFG1.TextMatrix(VSFG1.Rows - 3, 8) * PorIVA / 100) + VSFG1.TextMatrix(VSFG1.Rows - 3, 8)
    VSFG1.TextMatrix(VSFG1.Rows - 1, 7) = "TOTAL"
    VSFG1.Cell(flexcpBackColor, VSFG1.Rows - 3, 0, VSFG1.Rows - 1, VSFG1.Cols - 1) = vbYellow
    VSFG1.Cell(flexcpFontBold, VSFG1.Rows - 3, 0, VSFG1.Rows - 1, VSFG1.Cols - 1) = True
End Sub

Private Sub CargaDetalle()
    Dim i As Long
    Dim strContenedores As String
    With VSFG
        For i = 1 To .Rows - 1
            If Abs(.TextMatrix(i, 0)) = 1 Then
                strContenedores = strContenedores & .TextMatrix(i, 1) & ","
            End If
        Next i
    End With
    strContenedores = Left(strContenedores, Len(strContenedores) - 1)
    strSql = " SELECT det_contenedor_mercaderia.prd_codigo,prd_nombre,SUM(det_con_mer_cantidad) " & _
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
             " AND det_recepcion_mercaderia.rec_mer_codigo in (" & strContenedores & ")" & _
             " GROUP BY det_contenedor_mercaderia.prd_codigo,prd_nombre " & _
             " ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    VSFG2.Subtotal flexSTSum, -1, 2, , vbBlue, vbWhite, True, "TOTAL"
End Sub

Private Sub CargaContenedores()
    If cmbOrdenCompra.BoundText <> "" Then
        strSql = " SELECT DISTINCT '0' as sel, contenedor_mercaderia.con_mer_codigo,con_mer_fecha,est_con_mer_descripcion,dep_nombre,ubi_bod_codigo,con_mer_observacion,con_mer_fechamod,con_mer_usumod " & _
                 " FROM contenedor_mercaderia " & _
                 " INNER JOIN est_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.est_con_mer_codigo=est_contenedor_mercaderia.est_con_mer_codigo" & _
                 " INNER JOIN deposito " & _
                 " ON contenedor_mercaderia.emp_codigo=deposito.emp_codigo " & _
                 " AND contenedor_mercaderia.dep_codigo=deposito.dep_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo" & _
                 " INNER JOIN (SELECT preproducto_producto.emp_codigo,preproducto_producto.prd_codigo " & _
                 " FROM det_orden_compra INNER JOIN preproducto_producto " & _
                 " ON det_orden_compra.emp_codigo=preproducto_producto.emp_codigo " & _
                 " AND det_orden_compra.pre_codigo=preproducto_producto.pre_codigo " & _
                 " AND det_orden_compra.col_codigo=preproducto_producto.col_codigo " & _
                 " AND det_orden_compra.tal_codigo=preproducto_producto.tal_codigo " & _
                 " WHERE det_orden_compra.emp_codigo='" & strEmpresa & "'" & _
                 " AND det_orden_compra.ord_com_codigo='" & cmbOrdenCompra.BoundText & "') oc" & _
                 " ON det_contenedor_mercaderia.emp_codigo=oc.emp_codigo " & _
                 " AND det_contenedor_mercaderia.prd_codigo=oc.prd_codigo " & _
                 " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND contenedor_mercaderia.est_con_mer_codigo=0" & _
                 " AND contenedor_mercaderia.ord_com_codigo='" & cmbOrdenCompra.BoundText & "'" & _
                 " ORDER BY contenedor_mercaderia.con_mer_codigo "
    Else
        strSql = " SELECT DISTINCT '0' as sel, con_mer_codigo,con_mer_fecha,est_con_mer_descripcion,dep_nombre,ubi_bod_codigo,con_mer_observacion,con_mer_fechamod,con_mer_usumod " & _
                 " FROM contenedor_mercaderia " & _
                 " INNER JOIN est_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.est_con_mer_codigo=est_contenedor_mercaderia.est_con_mer_codigo" & _
                 " INNER JOIN deposito " & _
                 " ON contenedor_mercaderia.emp_codigo=deposito.emp_codigo " & _
                 " AND contenedor_mercaderia.dep_codigo=deposito.dep_codigo " & _
                 " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND contenedor_mercaderia.est_con_mer_codigo=0" & _
                 " ORDER BY con_mer_codigo "
    End If
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With VSFG
        .ShowCell Row, 0
        .Select Row, 0
        If Abs(.TextMatrix(Row, 0)) = 1 Then
            .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = vbYellow
        Else
            .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = vbWhite
        End If
        CargarDetalle
    End With
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
