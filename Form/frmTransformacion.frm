VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTransformacion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transformaciones"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "frmTransformacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5190
      TabIndex        =   11
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Producto Terminado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   9975
      Begin VB.TextBox TxtCantIng 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cmbProductoIng 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.CommandButton btn_cargar 
         Caption         =   "Cargar Ingresos de ""NS"""
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox TxtSubTotalIng 
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
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalleIng 
         Height          =   1650
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9735
         _cx             =   58082067
         _cy             =   58067806
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransformacion.frx":030A
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
      Begin MSDataListLib.DataCombo dcbo_ingreso 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label16 
         BackColor       =   &H00DDDDDD&
         Caption         =   "No. Recepción"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   7560
         TabIndex        =   21
         Top             =   1950
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Materia Prima"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   9975
      Begin VB.TextBox TxtCantEgr 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtSubTotalEgr 
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
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalleEgr 
         Height          =   2610
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9735
         _cx             =   58082067
         _cy             =   58069500
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransformacion.frx":03F2
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
         Begin VSFlex8Ctl.VSFlexGrid VSFGAbrir 
            Height          =   1260
            Left            =   0
            TabIndex        =   22
            Top             =   1320
            Visible         =   0   'False
            Width           =   5625
            _cx             =   116074178
            _cy             =   116066478
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTransformacion.frx":04DA
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin MSDataListLib.DataCombo cmbProductoEgr 
            Height          =   315
            Left            =   3000
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
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
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   7560
         TabIndex        =   18
         Top             =   2910
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos de la Transformación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9975
      Begin VB.TextBox TxtObserv 
         Height          =   645
         Left            =   5040
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtNumAux 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   615
         Width           =   1815
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _extentx        =   2566
         _extenty        =   661
         value           =   42816.4426157407
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   3780
         TabIndex        =   15
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Auxiliar"
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
         TabIndex        =   14
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Doc"
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
         TabIndex        =   13
         Top             =   300
         Width           =   1035
      End
   End
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   360
      Picture         =   "frmTransformacion.frx":057A
      Top             =   6960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   120
      Picture         =   "frmTransformacion.frx":06A6
      Top             =   6960
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmTransformacion"
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

Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private strSql As String
Private cargado As Boolean
Private cargadoIXC As Boolean
Private strNumeroIXC As String
Private strTipoIXC As String

Private Sub btn_cargar_Click()
    Dim clsAux As New clsConsulta
    Dim i As Long
    clsAux.Inicializar AdoConn, AdoConnMaster
    vsfgDetalleIng.Clear 1
    vsfgDetalleIng.Rows = 2
    cargado = True
    strSql = " SELECT ing_codigo,for_pag_codigo,ing_fecha FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IXC' " & _
             " AND ing_factura='" & dcbo_ingreso.BoundText & "'"
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        cargadoIXC = True
        strNumeroIXC = clsAux.adorec_Def("ing_codigo")
        dtpFecha.value = clsAux.adorec_Def("ing_fecha")
        strTipoIXC = "IXC"
        
        strSql = " SELECT dep_codigo,prd_codigo,det_ing_precio as precio,det_ing_cantidad as cant " & _
                 " FROM det_ingreso " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND ing_codigo=" & strNumeroIXC & " AND tip_ing_codigo='" & strTipoIXC & "' "
    Else
        cargadoIXC = False
        strNumeroIXC = "0"
        strTipoIXC = ""
        
        strSql = " SELECT contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,SUM(det_con_mer_cantidad) as cant " & _
                 " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
                 " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 AND mov_codigo=0 and tip_mov_codigo='' " & _
                 " INNER JOIN producto " & _
                 " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
                 " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
                 " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_ingreso.BoundText & "%'" & _
                 " GROUP BY contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,prd_nombre " & _
                 " ORDER BY prd_nombre "
    End If
    clsAux.Ejecutar strSql
    i = 1
    While Not clsAux.adorec_Def.EOF
        
        vsfgDetalleIng.Cell(flexcpPicture, i, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        
        vsfgDetalleIng.TextMatrix(i, 1) = clsAux.adorec_Def("dep_codigo")
        vsfgDetalleIng.TextMatrix(i, 2) = clsAux.adorec_Def("prd_codigo")
        vsfgDetalleIng.TextMatrix(i, 4) = clsAux.adorec_Def("cant")
        i = i + 1
        clsAux.adorec_Def.MoveNext
    Wend
End Sub

Private Sub cmdAbrir_Click()
    Dim num As Integer
    
    Dim strPath As String
    Dim strLinea As String
    Dim Arch As String
    'Arch = cmbTDoc.Text & ".xls"
    VSFGAbrir.Clear 1
    VSFGAbrir.Rows = 1
        
    If vsfgDetalleEgr.Rows > 1 Then
        strPath = Trim(App.Path)
        cmdArchivo.DialogTitle = "Abrir"
        'cmdArchivo.DefaultExt = strPath
        cmdArchivo.InitDir = strPath
        'cmdArchivo.FileName = Arch
        cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Documento de Excel 2007|*xlsx|Todos los Archivos|*.*"
        cmdArchivo.ShowOpen
        num = FreeFile
        Archivo = cmdArchivo.FileName
        If Archivo <> "" Then
            VSFGAbrir.LoadGrid Archivo, flexFileExcel
            If MsgBox("El archivo es Materia Prima??", vbYesNo + vbQuestion, "Transformacion") = vbYes Then
                vsfgDetalleEgr.Rows = 1
                With VSFGAbrir
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 0) <> "" Then
                            vsfgDetalleEgr.AddItem "", i
                            vsfgDetalleEgr.TextMatrix(i, 1) = .TextMatrix(i, 0)
                            vsfgDetalleEgr.TextMatrix(i, 2) = .TextMatrix(i, 1)
                            vsfgDetalleEgr.TextMatrix(i, 4) = .TextMatrix(i, 2)
                            '.ShowCell i + 1, 1
                            
                            vsfgDetalleEgr.Cell(flexcpPicture, i, 0) = imgBtnUp
                            vsfgDetalleEgr.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                        End If
                    Next i
                    If vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 1) = "" Then
                        vsfgDetalleEgr.RemoveItem vsfgDetalleEgr.Rows - 1
                    End If
                End With
            ElseIf MsgBox("El archivo es Producto Terminado??", vbYesNo + vbQuestion, "Transformacion") = vbYes Then
                vsfgDetalleIng.Rows = 1
                With VSFGAbrir
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 0) <> "" Then
                            vsfgDetalleIng.AddItem "", i
                            vsfgDetalleIng.TextMatrix(i, 1) = .TextMatrix(i, 0)
                            vsfgDetalleIng.TextMatrix(i, 2) = .TextMatrix(i, 1)
                            vsfgDetalleIng.TextMatrix(i, 4) = .TextMatrix(i, 2)
                            '.ShowCell i + 1, 1
                            
                            vsfgDetalleIng.Cell(flexcpPicture, i, 0) = imgBtnUp
                            vsfgDetalleIng.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                        End If
                    Next i
                    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) = "" Then
                        vsfgDetalleIng.RemoveItem vsfgDetalleIng.Rows - 1
                    End If
                End With
            
            End If
        End If
    Else
        MsgBox "No se tiene información para guardar", vbInformation, "Guardar"
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim clsEgreso As New clsInventario
    Dim clsIngresoAux As New clsInventario
    Dim booGuardar As Boolean
    Dim i As Long
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    booGuardar = clsEgreso.NuevoEgr(True, "ETN", False, strSucursal, strPtoFactura, , , , dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotalEgr.Text), , , 0, FormatoD2(TxtSubTotalEgr.Text))
    If booGuardar = True Then
        With vsfgDetalleEgr
            For i = 1 To .Rows - 1
                clsEgreso.NuevoDetEgr .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD4(.TextMatrix(i, 4)), FormatoD4(.TextMatrix(i, 5)), FormatoD4(.TextMatrix(i, 5))
            Next i
        End With
        If cargado = True Then
            clsEgreso.NuevoIng False, "ITN", False, strSucursal, strPtoFactura, , , , dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotalEgr.Text), , , , FormatoD2(TxtSubTotalEgr.Text), , , , , IIf(cargadoIXC = True, strNumeroIXC, "")
        Else
            clsEgreso.NuevoIng True, "ITN", False, strSucursal, strPtoFactura, , , , dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotalEgr.Text), , , , FormatoD2(TxtSubTotalEgr.Text), , , , , IIf(cargadoIXC = True, strNumeroIXC, "")
        End If
        With vsfgDetalleIng
            For i = 1 To .Rows - 1
                clsEgreso.NuevoDetIng .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD4(.TextMatrix(i, 4)), FormatoD4(.TextMatrix(i, 5)), FormatoD4(.TextMatrix(i, 5))
            Next i
            InicializarContenedorRecurrente
        End With
        
        If cargadoIXC = True Then
            clsIngresoAux.Inicializar AdoConn, AdoConnMaster
            clsIngresoAux.AnularIng strNumeroIXC, "IXC", , "CONTABILIZADO EN " & clsEgreso.strTipo & " " & clsEgreso.strDoc
        End If
        If cargado = True And cargadoIXC = False Then
            strSql = " UPDATE det_contenedor_mercaderia " & _
                     " SET det_contenedor_mercaderia.tip_mov_codigo='" & clsEgreso.strTipo & "'," & _
                     " det_contenedor_mercaderia.mov_codigo='" & clsEgreso.strDoc & "'" & _
                     " FROM det_recepcion_mercaderia ,contenedor_mercaderia,det_contenedor_mercaderia WHERE det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " AND contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " AND det_contenedor_mercaderia.tip_mov_codigo='' " & _
                     " AND det_contenedor_mercaderia.mov_codigo=0 " & _
                     " AND det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_ingreso.BoundText & "'"
            clsCon_Def.Ejecutar strSql
        End If
        
        MsgBox " Los datos han sido ingresado", vbInformation, "Egresos"
        
        Dim rpTra As New frmReporte
        rpTra.strNumero = clsEgreso.strDoc
        rpTra.strReporte = "rptTransformacionMercaderia"
        rpTra.Show
        
        Unload Me
    End If
End Sub

Private Sub CmdSalir_Click()
Dim rpTra As New frmReporte
        rpTra.strNumero = "10010000001"
        rpTra.strReporte = "rptTransformacionMercaderia"
        rpTra.Show
    
    Unload Me
End Sub

Private Sub Form_Activate()
    CargaProductos
    Dim clsConAUX As New clsConsulta
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT rec_mer_codigo " & _
             " FROM recepcion_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    
    strSql = " SELECT rec_mer_codigo,CONCAT(rec_mer_codigo,' ()') as n " & _
             " FROM recepcion_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
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
             " ORDER BY rec_mer_codigo"
    
    clsConAUX.Ejecutar strSql
    
    Set dcbo_ingreso.RowSource = clsConAUX.adorec_Def.DataSource
    dcbo_ingreso.ListField = "n"
    dcbo_ingreso.BoundColumn = "rec_mer_codigo"
    
End Sub

'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    dtpFecha.value = HoyDia
    'CargaProductos
    cargado = False
End Sub

Private Sub CargaProductos()
    'Carga los depositos
    strSql = "SELECT dep_codigo, dep_nombre FROM deposito WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql
    vsfgDetalleIng.ColComboList(1) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    vsfgDetalleEgr.ColComboList(1) = vsfgDetalleEgr.BuildComboList(clsCon_Def.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    'Carga los productos
'    strSql = " SELECT producto.prd_codigo, prd_nombre,prd_costo,prd_costo as prd_precio " & _
'             " FROM producto " & _
'             " WHERE producto.emp_codigo = '" & strEmpresa & "' ORDER BY prd_codigo "
'    clsCon_Prd.Ejecutar strSql
'    vsfgDetalleIng.ColComboList(2) = vsfgDetalleIng.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
'    vsfgDetalleEgr.ColComboList(2) = vsfgDetalleEgr.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    'Consulto los productos de la empresa
    strSql = " SELECT producto.prd_codigo, prd_nombre,prd_costo,prd_costo as prd_precio " & _
             " FROM producto " & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' ORDER BY prd_nombre "
    clsCon_Prd.Ejecutar strSql
'    vsfgDetalleIng.ColComboList(3) = vsfgDetalleIng.BuildComboList(clsCon_Prd.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
'    vsfgDetalleEgr.ColComboList(3) = vsfgDetalleIng.ColComboList(3) 'vsfgDetalleEgr.BuildComboList(clsCon_Prd.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
End Sub

Private Sub vsfgDetalleEgr_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleEgr.TextMatrix(Row, 2) & "'"
        vsfgDetalleEgr.TextMatrix(Row, 3) = clsCon_Prd.adorec_Def("prd_nombre")
        vsfgDetalleEgr.TextMatrix(Row, 4) = 0
        vsfgDetalleEgr.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalleEgr.TextMatrix(Row, 6) = 0
'    ElseIf Col = 3 Then
'        vsfgDetalleEgr.TextMatrix(Row, 2) = vsfgDetalleEgr.TextMatrix(Row, 3)
'        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleEgr.TextMatrix(Row, 2) & "'"
'        vsfgDetalleEgr.TextMatrix(Row, 4) = 0
'        vsfgDetalleEgr.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
'        vsfgDetalleEgr.TextMatrix(Row, 6) = 0
    ElseIf Col = 4 Then
        vsfgDetalleEgr.TextMatrix(Row, 6) = FormatoD4(FormatoD4(vsfgDetalleEgr.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalleEgr.TextMatrix(Row, 5)))
        CalculaTotal
    End If
    If vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 1) <> "" And vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 2) <> "" And vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 3) <> "" And Val(vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 4)) <> 0 Then
        vsfgDetalleEgr.AddItem ""
        vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Rows - 1, 0) = vsfgDetalleEgr.Rows - 1
        vsfgDetalleEgr.Cell(flexcpPicture, vsfgDetalleEgr.Rows - 1, 0) = imgBtnUp
        vsfgDetalleEgr.Cell(flexcpPictureAlignment, vsfgDetalleEgr.Rows - 1, 0) = flexAlignRightCenter
    End If
End Sub

Private Sub vsfgDetalleEgr_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If vsfgDetalleEgr.Col = 3 And KeyCode = vbKeyF4 And Trim(vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Row, vsfgDetalleEgr.Col)) <> "" And Len(Trim(vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Row, vsfgDetalleEgr.Col))) >= 2 Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " " & _
                 " " & _
                 " " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " " & _
                 " " & _
                 " AND prd_nombre LIKE '" & Trim(vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Row, vsfgDetalleEgr.Col)) & "%' " & _
                 " ORDER BY producto.prd_nombre "
        clsAux.Ejecutar strSql
        
        Set cmbProductoEgr.RowSource = clsAux.adorec_Def.DataSource
        cmbProductoEgr.ListField = "prd_nombre"
        cmbProductoEgr.BoundColumn = "prd_codigo"
        cmbProductoEgr.Visible = True
        cmbProductoEgr.SetFocus
        cmbProductoEgr = ""
    End If
    Set clsAux = Nothing
End Sub

Private Sub cmbProductoEgr_Validate(Cancel As Boolean)
    vsfgDetalleEgr.TextMatrix(vsfgDetalleEgr.Row, 2) = cmbProductoEgr.BoundText
    cmbProductoEgr.Visible = False
    vsfgDetalleEgr.SetFocus
    vsfgDetalleEgr.Col = 2
    vsfgDetalleEgr.EditCell
End Sub

Private Sub vsfgDetalleIng_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If vsfgDetalleIng.Col = 3 And KeyCode = vbKeyF4 And Trim(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Row, vsfgDetalleIng.Col)) <> "" And Len(Trim(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Row, vsfgDetalleIng.Col))) >= 2 Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " " & _
                 " " & _
                 " " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " " & _
                 " " & _
                 " AND prd_nombre LIKE '" & Trim(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Row, vsfgDetalleIng.Col)) & "%' " & _
                 " ORDER BY producto.prd_nombre "
        clsAux.Ejecutar strSql
        
        Set cmbProductoIng.RowSource = clsAux.adorec_Def.DataSource
        cmbProductoIng.ListField = "prd_nombre"
        cmbProductoIng.BoundColumn = "prd_codigo"
        cmbProductoIng.Visible = True
        cmbProductoIng.SetFocus
        cmbProductoIng = ""
    End If
    Set clsAux = Nothing
End Sub

Private Sub cmbProductoIng_Validate(Cancel As Boolean)
    vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Row, 2) = cmbProductoIng.BoundText
    cmbProductoIng.Visible = False
    vsfgDetalleIng.SetFocus
    vsfgDetalleIng.Col = 2
    vsfgDetalleIng.EditCell
End Sub

Private Sub vsfgDetalleIng_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col >= 5 Then Cancel = True
End Sub

Private Sub vsfgDetalleEgr_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col >= 5 Then Cancel = True
End Sub

Private Sub vsfgDetalleIng_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 3) = clsCon_Prd.adorec_Def("prd_nombre")
        vsfgDetalleIng.TextMatrix(Row, 4) = 1
        vsfgDetalleIng.TextMatrix(Row, 5) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(Row, 6)) / FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)))
        If FormatoD4(vsfgDetalleIng.TextMatrix(Row, 6)) = 0 Then
            vsfgDetalleIng.TextMatrix(Row, 6) = 0
        End If
'    ElseIf Col = 3 Then
'        vsfgDetalleIng.TextMatrix(Row, 2) = vsfgDetalleIng.TextMatrix(Row, 3)
'        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
'        vsfgDetalleIng.TextMatrix(Row, 4) = 1
'        vsfgDetalleIng.TextMatrix(Row, 5) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(Row, 6)) / FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)))
'        If FormatoD4(vsfgDetalleIng.TextMatrix(Row, 6)) = 0 Then
'            vsfgDetalleIng.TextMatrix(Row, 6) = 0
'        End If
    ElseIf Col = 4 Then
        If FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)) <> 0 Then
            vsfgDetalleIng.TextMatrix(Row, 5) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(Row, 6)) / FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)))
        Else
            vsfgDetalleIng.TextMatrix(Row, 5) = 0
        End If
        CalculaTotal
    End If
    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 2) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 3) <> "" And Val(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 4)) <> 0 Then
        vsfgDetalleIng.AddItem ""
        vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 0) = vsfgDetalleIng.Rows - 1
        vsfgDetalleIng.Cell(flexcpPicture, vsfgDetalleIng.Rows - 1, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, vsfgDetalleIng.Rows - 1, 0) = flexAlignRightCenter
    End If
End Sub

Private Sub CalculaTotal()
    Dim i As Long
    Dim CantIng As Double
    TxtSubTotalEgr.Text = 0
    TxtCantEgr = 0
    For i = 1 To vsfgDetalleEgr.Rows - 1
        TxtSubTotalEgr.Text = FormatoD4(FormatoD4(TxtSubTotalEgr.Text) + FormatoD4(vsfgDetalleEgr.TextMatrix(i, 6)))
        TxtCantEgr.Text = FormatoD4(FormatoD4(TxtCantEgr.Text) + FormatoD4(vsfgDetalleEgr.TextMatrix(i, 4)))
    Next i
    TxtSubTotalIng.Text = TxtSubTotalEgr.Text
    CantIng = 0
    txtCantIng = 0
    For i = 1 To vsfgDetalleIng.Rows - 1
        CantIng = CantIng + FormatoD4(vsfgDetalleIng.TextMatrix(i, 4))
    Next i
    txtCantIng.Text = FormatoD4(CantIng)
    For i = 1 To vsfgDetalleIng.Rows - 1
        If CantIng <> 0 And vsfgDetalleIng.TextMatrix(i, 2) <> "" Then
        vsfgDetalleIng.TextMatrix(i, 6) = FormatoD2(TxtSubTotalEgr.Text) / CantIng * FormatoD4(vsfgDetalleIng.TextMatrix(i, 4))
        vsfgDetalleIng.TextMatrix(i, 5) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(i, 6)) / FormatoD4(vsfgDetalleIng.TextMatrix(i, 4)))
        End If
        
    Next i
End Sub
