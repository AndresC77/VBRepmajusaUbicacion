VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVerMovimiento 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Mercaderia"
   ClientHeight    =   11400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerMovimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11400
   ScaleWidth      =   12510
   Begin VB.CommandButton cmdReasignarContenedores 
      Caption         =   "Reasig.Conten."
      Height          =   375
      Left            =   8228
      TabIndex        =   84
      Top             =   10920
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   4876
      TabIndex        =   83
      Top             =   10920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6556
      TabIndex        =   82
      Top             =   10920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   3196
      TabIndex        =   81
      Top             =   10920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "&Copiar"
      Height          =   375
      Left            =   1508
      TabIndex        =   80
      Top             =   10920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9908
      TabIndex        =   79
      Top             =   10920
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   240
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraRet 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Retención"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   2520
      Width           =   6735
      Begin VB.TextBox txtFechaR 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDocumentoR 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtSerieR 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtAutorizacionR 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1110
         Width           =   1335
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGRet 
         Height          =   855
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   4305
         _cx             =   58334634
         _cy             =   58328548
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerMovimiento.frx":030A
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Autorizacion"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2250
         TabIndex        =   10
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label15 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   585
      End
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
      Height          =   5895
      Left            =   1268
      TabIndex        =   17
      Top             =   4920
      Width           =   9975
      Begin VB.TextBox TxtObserv 
         Height          =   525
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   5160
         Width           =   8295
      End
      Begin VB.TextBox TxtSubTotal 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox TxtDesc 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TxtIva 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox TxtRecargo 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4320
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGDetalle 
         Height          =   3300
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   9705
         _cx             =   58344159
         _cy             =   58332861
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerMovimiento.frx":036D
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Recargos:"
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
         Height          =   1455
         Left            =   1440
         TabIndex        =   18
         Top             =   3600
         Width           =   4575
         Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4305
            _cx             =   58334634
            _cy             =   58328971
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmVerMovimiento.frx":0477
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
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   180
         TabIndex        =   45
         Top             =   5160
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   7680
         TabIndex        =   44
         Top             =   4710
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recargos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7680
         TabIndex        =   43
         Top             =   4350
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7680
         TabIndex        =   42
         Top             =   3870
         Width           =   825
      End
      Begin VB.Label LblIva 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA X%"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7680
         TabIndex        =   41
         Top             =   4110
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7680
         TabIndex        =   40
         Top             =   3630
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Extras del Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   368
      TabIndex        =   12
      Top             =   3960
      Width           =   11775
      Begin VB.TextBox txtDocumento 
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtSerie 
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtAutorizacion 
         Height          =   315
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtAsiento 
         Height          =   315
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   1800
      End
      Begin VB.TextBox txtCaduca 
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   1800
      End
      Begin VB.TextBox txtFPago 
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   1800
      End
      Begin VB.TextBox txtNumAux 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5760
         TabIndex        =   56
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizacion"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8880
         TabIndex        =   55
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3000
         TabIndex        =   54
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   53
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8880
         TabIndex        =   47
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "No.Aux."
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Caduc Doc"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5760
         TabIndex        =   15
         Top             =   630
         Width           =   795
      End
      Begin VB.Label lblFpago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Pago"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3000
         TabIndex        =   14
         Top             =   645
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Left            =   5640
         TabIndex        =   77
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Left            =   8760
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   8760
         TabIndex        =   63
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   330
            ItemData        =   "frmVerMovimiento.frx":04F7
            Left            =   1320
            List            =   "frmVerMovimiento.frx":0522
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   210
            Value           =   -1  'True
            Width           =   255
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
            Left            =   480
            TabIndex        =   68
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   16842755
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
            Left            =   1920
            TabIndex        =   69
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   16842755
            CurrentDate     =   37463
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   72
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   71
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   70
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.TextBox txtNum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   735
         Width           =   3450
      End
      Begin VB.CheckBox chkNum 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por No. de Documento"
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
         Left            =   5160
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   2685
      End
      Begin VB.CheckBox chkFiltroPersona 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Tipo de Persona"
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
         Left            =   90
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton optEgr 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Egresos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optIng 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Ingresos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   90
         TabIndex        =   3
         Top             =   735
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPersona 
         Height          =   330
         Left            =   90
         TabIndex        =   48
         Top             =   1560
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Doc"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   495
         Width           =   3450
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Ingreso"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   59
         Top             =   495
         Width           =   4905
      End
      Begin VB.Label lblPersona 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Personas"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   58
         Top             =   1335
         Width           =   4905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos de la Persona"
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
      TabIndex        =   24
      Top             =   2520
      Width           =   5415
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   615
         Width           =   4515
      End
      Begin VB.TextBox txtTelf 
         Height          =   315
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtRUC 
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtPersona 
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   4515
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Telf."
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2850
         TabIndex        =   32
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "CI/RUC"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   31
         Top             =   990
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   30
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   675
      End
   End
   Begin MSDataListLib.DataCombo cmbCotizacion 
      Height          =   330
      Left            =   4005
      TabIndex        =   74
      Top             =   2160
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGGuardar 
      Height          =   1260
      Left            =   360
      TabIndex        =   78
      Top             =   8760
      Visible         =   0   'False
      Width           =   7785
      _cx             =   58340772
      _cy             =   58329262
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerMovimiento.frx":058B
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000050&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número de Doc"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4005
      TabIndex        =   76
      Top             =   1920
      Width           =   4500
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO ANULADO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8820
      TabIndex        =   75
      Top             =   2040
      Width           =   3435
   End
End
Attribute VB_Name = "frmVerMovimiento"
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
Private strSQL As String
Private clsSql As New clsConsulta
Private FechaI As Variant
Private FechaF As Variant


Private Sub cmdFormato_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        drptDetIngreso.Tag = cmbCotizacion.BoundText
        drptDetIngreso.TipIng = cmbCliente.BoundText
        drptDetIngreso.Show
    Else
        MsgBox "No ha seleccionado ningun Ingreso", vbInformation, "Impresión Ingresos"
    End If
End Sub

Private Sub chkFiltroFecha_Click()
    If chkFiltroFecha.Value = 1 Then
        fraFecha.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        If Option1.Value = True Then
            lblMes.Enabled = True
            cmbMesI.Enabled = True
        ElseIf Option2.Value = True Then
            Fecha1.Enabled = True
            Label1.Enabled = True
            Fecha1.Enabled = True
            chkFechas.Enabled = True
            If chkFechas.Value = 1 Then
                Label2.Enabled = True
                Fecha2.Enabled = True
            End If
        End If
    Else
        fraFecha.Enabled = False
        
        Fecha2.Enabled = False
        Label1.Enabled = False
        Fecha1.Enabled = False
        Label2.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        
        Option1.Enabled = False
        Option2.Enabled = False
        lblMes.Enabled = False
        cmbMesI.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub chkFiltroPersona_Click()
    If chkFiltroPersona.Value = 1 Then
        lblPersona.Enabled = True
        dcmbPersona.Enabled = True
    Else
        lblPersona.Enabled = False
        dcmbPersona.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub chkNum_Click()
    If chkNum.Value = 1 Then
        lblNum.Enabled = True
        txtNum.Enabled = True
    Else
        lblNum.Enabled = False
        txtNum.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub cmbCotizacion_Change()
    If cmbCotizacion.MatchedWithList = True Then
        CargaDocumento
        If Me.Tag = True Then
            If MesCerrado(txtFecha.Text) = True Then
                cmdAnular.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub cmdAnular_Click()
    Dim clsInv As New clsInventario
    clsInv.Inicializar AdoConn, AdoConnMaster
    If optEgr.Value = True Then
        clsInv.strIE = "E"
        clsInv.AnularEgr cmbCotizacion.BoundText, cmbCliente.BoundText
    ElseIf optIng.Value = True Then
        clsInv.strIE = "I"
        clsInv.AnularIng cmbCotizacion.BoundText, cmbCliente.BoundText
    End If
    cmbCotizacion_Change
End Sub

Private Sub cmdBuscar_Click()
    ActualizaNumero
End Sub

Private Sub cmdCopiar_Click()
    Dim i As Long
    Dim j As Long
    If Me.Tag = "I" Then
        frmNuevoIngreso.vsfgDetalle.Clear 1
        frmNuevoIngreso.vsfgDetalle.Rows = vsfgDetalle.Rows
        frmNuevoIngreso.txtNumAux.Text = cmbCotizacion.BoundText
        For i = 1 To vsfgDetalle.Rows - 1
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 1) = vsfgDetalle.TextMatrix(i, 1)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 2) = vsfgDetalle.TextMatrix(i, 2)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 4) = vsfgDetalle.TextMatrix(i, 4)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 5) = vsfgDetalle.TextMatrix(i, 5)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 6) = vsfgDetalle.TextMatrix(i, 6)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 7) = vsfgDetalle.TextMatrix(i, 7)
            frmNuevoIngreso.vsfgDetalle.TextMatrix(i, 8) = vsfgDetalle.TextMatrix(i, 8)
        Next i
    ElseIf Me.Tag = "E" Then
        frmNuevoEgreso.vsfgDetalle.Clear 1
        frmNuevoEgreso.vsfgDetalle.Rows = vsfgDetalle.Rows
        For i = 1 To vsfgDetalle.Rows - 1
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 1) = vsfgDetalle.TextMatrix(i, 1)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 2) = vsfgDetalle.TextMatrix(i, 2)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 4) = vsfgDetalle.TextMatrix(i, 4)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 5) = vsfgDetalle.TextMatrix(i, 5)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 6) = vsfgDetalle.TextMatrix(i, 6)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 7) = vsfgDetalle.TextMatrix(i, 7)
            frmNuevoEgreso.vsfgDetalle.TextMatrix(i, 8) = vsfgDetalle.TextMatrix(i, 8)
        Next i
    End If
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim num As Integer
    
    Dim strPath As String
    Dim strLinea As String
    Dim Arch As String
    Arch = IIf(optIng.Value = True, "Ingreso", "Egreso") & ".xls"
    VSFGGuardar.Clear 1
    VSFGGuardar.Rows = 1
    
    If vsfgDetalle.Rows > 1 Then
        strPath = Trim(App.Path)
        cmdArchivo.DialogTitle = "Guardar"
        'cmdArchivo.DefaultExt = strPath
        cmdArchivo.InitDir = strPath
        cmdArchivo.FileName = Arch
        cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Documento de Excel 2007|*xlsx|Todos los Archivos|*.*"
        cmdArchivo.ShowSave
        num = FreeFile
        Archivo = cmdArchivo.FileName
        If Archivo <> "" Then
            With vsfgDetalle
                For i = 1 To .Rows - 1
                    strLinea = .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 2) & vbTab & .TextMatrix(i, 4) & vbTab & .TextMatrix(i, 7) & vbTab & "0.00"
                    VSFGGuardar.AddItem strLinea
                Next i
            End With
            VSFGGuardar.SaveGrid Archivo, flexFileExcel
        End If
    Else
        MsgBox "No se tiene información para guardar", vbInformation, "Guardar"
    End If
End Sub

Private Sub cmdReasignarContenedores_Click()
    Dim i As Long
    Dim Reasignar As Boolean
    Dim tip As String
    Dim Tabla As String
    Dim producto As String
    Dim clsAux As New clsConsulta
    Dim clsAux2 As New clsConsulta
    Dim clsMov As New clsContenedor
    
    clsAux.Inicializar AdoConn, AdoConnMaster
    clsMov.Inicializar AdoConn, AdoConnMaster
    If optEgr.Value = True Then
        tip = "egr"
        Tabla = "egreso"
    Else
        tip = "ing"
        Tabla = "ingreso"
    End If
    strSQL = " SELECT prd_codigo,sum(det_" & tip & "_cantidad) as cant" & _
             " FROM det_" & Tabla & " " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_" & tip & "_codigo='" & cmbCliente.BoundText & "'" & _
             " AND " & tip & "_codigo='" & cmbCotizacion.BoundText & "'" & _
             " GROUP BY prd_codigo"
    clsSql.Ejecutar strSQL, "L"
    While Not clsSql.adorec_Def.EOF
        Reasignar = False
        producto = clsSql.adorec_Def("prd_codigo")
        strSQL = " SELECT prd_codigo,sum(det_con_mer_cantidad) as cant" & _
                 " FROM det_contenedor_mercaderia " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND tip_mov_codigo='" & cmbCliente.BoundText & "'" & _
                 " AND mov_codigo='" & cmbCotizacion.BoundText & "'" & _
                 " AND prd_codigo='" & producto & "' " & _
                 " GROUP BY prd_codigo"
        clsAux.Ejecutar strSQL, "L"
        If clsAux.adorec_Def.RecordCount > 0 Then
            If clsSql.adorec_Def("cant") <> clsAux.adorec_Def("cant") Then
                Reasignar = True
            End If
        Else
            Reasignar = True
        End If
        If Reasignar = True Then
            strSQL = " DELETE " & _
                     " FROM det_contenedor_mercaderia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND tip_mov_codigo='" & cmbCliente.BoundText & "'" & _
                     " AND mov_codigo='" & cmbCotizacion.BoundText & "'" & _
                     " AND prd_codigo='" & producto & "' "
            clsAux.Ejecutar strSQL, "M"
            strSQL = " SELECT dep_codigo,prd_codigo,det_" & tip & "_cantidad as cant" & _
                     " FROM det_" & Tabla & " " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND tip_" & tip & "_codigo='" & cmbCliente.BoundText & "'" & _
                     " AND " & tip & "_codigo='" & cmbCotizacion.BoundText & "'" & _
                     " AND prd_codigo='" & producto & "' "
            clsAux.Ejecutar strSQL, "L"
            
            If optEgr.Value = True Then
                'clsMov.EgresarPrendas producto, clsAux.adorec_Def("cant"), "PRI','DES", cmbCliente.BoundText, cmbCotizacion.BoundText
                clsMov.EgresarPrendas producto, clsAux.adorec_Def("cant"), clsAux.adorec_Def("dep_codigo"), cmbCliente.BoundText, cmbCotizacion.BoundText
            Else
                clsMov.IngresarPrendas producto, clsAux.adorec_Def("cant"), clsAux.adorec_Def("dep_codigo"), cmbCliente.BoundText, cmbCotizacion.BoundText
            End If
            
        End If
        clsSql.adorec_Def.MoveNext
    Wend
    MsgBox "Reasignacion terminada"
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

Private Sub cmbCliente_Change()
    If cmbCliente.MatchedWithList = True Then
        ActualizaNumero
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
'        drptDetMovimiento.Tag = cmbCotizacion.BoundText
'        If optIng.Value = True Then
'            drptDetMovimiento.strTipo = "ingreso"
'        Else
'            drptDetMovimiento.strTipo = "egreso"
'        End If
'        drptDetMovimiento.TipDoc = cmbCliente.BoundText
'        drptDetMovimiento.Show
        Dim tip As String
        If optIng.Value = True Then
            If cmbCliente.BoundText = "ITN" Then
                tip = "rptTransformacionMercaderia"
            ElseIf cmbCliente.BoundText = "ITR" Then
                tip = "rptTransferencia"
            ElseIf cmbCliente.BoundText = "ICA" Then
                tip = "rptCambioProducto"
            Else
                tip = "rptIngresoMercaderia"
            End If
            
            
        Else
            If cmbCliente.BoundText = "ETN" Then
                tip = "rptTransformacionMercaderia"
            ElseIf cmbCliente.BoundText = "ETR" Then
                tip = "rptTransferencia"
            ElseIf cmbCliente.BoundText = "ECA" Then
                tip = "rptCambioProducto"
            ElseIf cmbCliente.BoundText = "NET" Then
                tip = "rptNotaEntregaSuministro"
            Else
                tip = "rptEgresoMercaderia"
            End If
            
        End If
        Dim rpMov As New frmReporte
        rpMov.strNumero = cmbCotizacion.BoundText
        rpMov.strTipo = cmbCliente.BoundText
        rpMov.strReporte = tip
        rpMov.Show
    Else
        MsgBox "No ha seleccionado un Movimiento", vbInformation, "Imprimir"
    End If
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    clsSql.Inicializar AdoConn, AdoConnMaster
    'Carga los personas
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,IIF(cat_p_tipo='C',CONCAT(' (',tip_ped_nombre,')'),''),' (',cat_p_tipo,')') per " & _
             " FROM persona " & _
             " LEFT JOIN tipo_pedido " & _
             " ON tipo_pedido.emp_codigo=persona.emp_codigo " & _
             " AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
             " WHERE persona.emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')')"
    clsSql.Ejecutar strSQL
    dcmbPersona.ListField = "per"
    dcmbPersona.BoundColumn = "per_codigo"
    Set dcmbPersona.RowSource = clsSql.adorec_Def.DataSource
        'Selecciona el mes actual
    Fecha1 = Format(HoyDia, "yyyy-mm-dd")
    Fecha2 = Format(HoyDia, "yyyy-mm-dd")
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    optIng_Click
    lblEstado.Caption = ""
End Sub



Private Sub optEgr_Click()
    If optEgr.Value = True Then
        lblTipo.Caption = "Tipo de Egreso"
        ActualizarTipo
    End If
End Sub

Private Sub optIng_Click()
    If optIng.Value = True Then
        lblTipo.Caption = "Tipo de Ingreso"
        ActualizarTipo
    End If
End Sub

Private Sub ActualizarTipo()
    Dim strTipo As String
    If optIng.Value = True Then
        strTipo = "ingreso"
    Else
        strTipo = "egreso"
    End If
    strSQL = " SELECT tip_" & Left(strTipo, 3) & "_codigo,tip_" & Left(strTipo, 3) & "_nombre " & _
             " FROM tipo_" & strTipo & " " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY tip_" & Left(strTipo, 3) & "_nombre "
    clsSql.Ejecutar strSQL
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "tip_" & Left(strTipo, 3) & "_nombre"
    cmbCliente.BoundColumn = "tip_" & Left(strTipo, 3) & "_codigo"
End Sub

Private Sub ActualizaNumero()
    Dim strTipo As String
    Dim strWhere As String
    If optIng.Value = True Then
        strTipo = "ingreso"
    Else
        strTipo = "egreso"
    End If
    If chkFiltroPersona.Value = 1 Then
        strWhere = strWhere & " AND per_codigo ='" & dcmbPersona.BoundText & "' "
    End If
    If chkNum.Value = 1 Then
        strWhere = strWhere & " AND " & Left(strTipo, 3) & "_numero ='" & txtNum.Text & "' "
    End If
    If chkFiltroFecha.Value = 1 Then
        If chkFechas.Value = 0 Then
            Fecha2 = Fecha1
        End If
        If Option1.Value = True Then
            strWhere = strWhere & " AND LEFT(" & Left(strTipo, 3) & "_fecha,10) BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
        ElseIf Option2.Value = True Then
            strWhere = strWhere & " AND LEFT(" & Left(strTipo, 3) & "_fecha,10) BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
        End If
    End If
    strSQL = " SELECT " & Left(strTipo, 3) & "_codigo as doc, CONCAT(" & Left(strTipo, 3) & "_codigo,IIF(" & Left(strTipo, 3) & "_anulado=1,' ANULADO ',''),' - ',COALESCE(" & Left(strTipo, 3) & "_factura,''),' - ',COALESCE(" & Left(strTipo, 3) & "_observacion,'')) as nombre " & _
             " FROM " & strTipo & " " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_" & Left(strTipo, 3) & "_codigo='" & cmbCliente.BoundText & "'" & _
             strWhere & _
             " ORDER BY " & Left(strTipo, 3) & "_codigo DESC"
    clsSql.Ejecutar strSQL
    cmbCotizacion = ""
    cmbCotizacion.ListField = "nombre"
    cmbCotizacion.BoundColumn = "doc"
    Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
End Sub

Private Sub CargaDocumento()
    Dim strTipo As String
    If optIng.Value = True Then
        strTipo = "ingreso"
    Else
        strTipo = "egreso"
    End If
    strSQL = " SELECT " & Left(strTipo, 3) & "_fecha as fecha, CONCAT(COALESCE(per_apellido,''),' ',COALESCE(per_nombre,'')) as nombre,COALESCE(per_ruc,'') as ruc,COALESCE(per_telf,'') as telf,COALESCE(per_direccion,'') as dir," & Left(strTipo, 3) & "_observacion as obs, " & _
             " " & Left(strTipo, 3) & "_serie as serie," & Left(strTipo, 3) & "_numero as numero," & Left(strTipo, 3) & "_autorizacion as autorizacion," & Left(strTipo, 3) & "_caduca as caduca," & Left(strTipo, 3) & "_anulado as anulado,COALESCE(for_pag_nombre,'') as fp, " & _
             " COALESCE(com_ret_serie,'') as com_ret_serie,COALESCE(com_ret_numero,0) as com_ret_numero,COALESCE(com_ret_autorizacion,'') as com_ret_autorizacion,COALESCE(com_ret_fecha,'') as com_ret_fecha," & _
             " " & Left(strTipo, 3) & "_subtotal as st," & Left(strTipo, 3) & "_subtotal_o as st0," & Left(strTipo, 3) & "_impuesto as iva," & Left(strTipo, 3) & "_dcto as dcto," & Left(strTipo, 3) & "_total as total,COALESCE(" & Left(strTipo, 3) & "_numasiento,'') as asi" & _
             " FROM " & strTipo & " LEFT JOIN persona ON " & strTipo & ".emp_codigo=persona.emp_codigo AND " & strTipo & ".per_codigo=persona.per_codigo " & _
             " LEFT JOIN forma_pago ON " & strTipo & ".emp_codigo=forma_pago.emp_codigo AND " & strTipo & ".for_pag_codigo=forma_pago.for_pag_codigo " & _
             " LEFT JOIN cuenta_p_c ON " & strTipo & ".emp_codigo=cuenta_p_c.emp_codigo AND " & strTipo & "." & Left(strTipo, 3) & "_serie=cuenta_p_c.cue_p_c_serie AND " & strTipo & "." & Left(strTipo, 3) & "_numero=cuenta_p_c.cue_p_c_numero AND " & strTipo & "." & Left(strTipo, 3) & "_autorizacion=cuenta_p_c.cue_p_c_autorizacion AND " & strTipo & "." & Left(strTipo, 3) & "_caduca=cuenta_p_c.cue_p_c_caduca " & _
             " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
             " WHERE " & strTipo & ".emp_codigo='" & strEmpresa & "' " & _
             " AND tip_" & Left(strTipo, 3) & "_codigo='" & cmbCliente.BoundText & "'" & _
             " AND " & Left(strTipo, 3) & "_codigo='" & Me.cmbCotizacion.BoundText & "' " & _
             " ORDER BY " & Left(strTipo, 3) & "_codigo DESC"
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def("anulado") = 0 Then
        lblEstado.Caption = ""
    Else
        lblEstado.Caption = "ANULADO"
    End If
    TxtSubTotal.Text = clsSql.adorec_Def("st")
    TxtRecargo.Text = clsSql.adorec_Def("st0")
    TxtDesc.Text = clsSql.adorec_Def("dcto")
    TxtIva.Text = clsSql.adorec_Def("iva")
    TxtTotal.Text = clsSql.adorec_Def("total")
    txtPersona.Text = clsSql.adorec_Def("nombre")
    txtAsiento.Text = clsSql.adorec_Def("asi")
    If Me.Tag = True Then
        If txtAsiento.Text = "" Then
            If clsSql.adorec_Def("anulado") = 0 Then
                cmdAnular.Visible = True
            Else
                cmdAnular.Visible = False
            End If
        Else
            cmdAnular.Visible = False
        End If
    End If
    txtRuc.Text = clsSql.adorec_Def("ruc")
    txtTelf.Text = clsSql.adorec_Def("telf")
    txtDireccion.Text = clsSql.adorec_Def("dir")
    txtFecha.Text = Left(clsSql.adorec_Def("fecha"), 10)
    txtFPago.Text = clsSql.adorec_Def("fp")
    txtSerie.Text = clsSql.adorec_Def("serie")
    txtDocumento.Text = clsSql.adorec_Def("numero")
    txtAutorizacion.Text = clsSql.adorec_Def("autorizacion")
    txtCaduca.Text = clsSql.adorec_Def("caduca")
    TxtObserv.Text = clsSql.adorec_Def("obs")
    txtFechaR.Text = Left(clsSql.adorec_Def("com_ret_fecha"), 10)
    txtSerieR.Text = clsSql.adorec_Def("com_ret_serie")
    txtDocumentoR.Text = clsSql.adorec_Def("com_ret_numero")
    txtAutorizacionR.Text = clsSql.adorec_Def("com_ret_autorizacion")
    
    strSQL = " SELECT dep_codigo,producto.prd_codigo, prd_nombre,det_" & Left(strTipo, 3) & "_cantidad,det_" & Left(strTipo, 3) & "_precio,det_" & Left(strTipo, 3) & "_dcto,ROUND(det_" & Left(strTipo, 3) & "_cantidad * det_" & Left(strTipo, 3) & "_precio - det_" & Left(strTipo, 3) & "_dcto,2),det_" & Left(strTipo, 3) & "_costo " & _
             " FROM det_" & strTipo & " INNER JOIN producto ON det_" & strTipo & ".emp_codigo=producto.emp_codigo AND det_" & strTipo & ".prd_codigo=producto.prd_codigo " & _
             " WHERE det_" & strTipo & ".emp_codigo='" & strEmpresa & "' " & _
             " AND det_" & strTipo & ".tip_" & Left(strTipo, 3) & "_codigo='" & cmbCliente.BoundText & "'" & _
             " AND det_" & strTipo & "." & Left(strTipo, 3) & "_codigo='" & Me.cmbCotizacion.BoundText & "' " & _
             " ORDER BY prd_codigo DESC"
    clsSql.Ejecutar strSQL
    vsfgDetalle.Clear 1
    Set vsfgDetalle.DataSource = clsSql.adorec_Def.DataSource
    strSQL = " SELECT ocargos.oca_codigo, oca_nombre,det_" & Left(strTipo, 3) & "_c_precio " & _
             " FROM det_" & strTipo & "_c INNER JOIN ocargos ON det_" & strTipo & "_c.emp_codigo=ocargos.emp_codigo AND det_" & strTipo & "_c.oca_codigo=ocargos.oca_codigo " & _
             " WHERE det_" & strTipo & "_c.emp_codigo='" & strEmpresa & "' " & _
             " AND det_" & strTipo & "_c.tip_" & Left(strTipo, 3) & "_codigo='" & cmbCliente.BoundText & "'" & _
             " AND det_" & strTipo & "_c." & Left(strTipo, 3) & "_codigo='" & Me.cmbCotizacion.BoundText & "' " & _
             " ORDER BY oca_codigo DESC"
    clsSql.Ejecutar strSQL
    VSFGReca.Clear 1
    Set VSFGReca.DataSource = clsSql.adorec_Def.DataSource
    strSQL = " SELECT ret_nombre,ret_porcentaje,det_" & Left(strTipo, 3) & "_ret_valor " & _
             " FROM det_" & strTipo & "_ret INNER JOIN retencion ON det_" & strTipo & "_ret.emp_codigo=retencion.emp_codigo AND det_" & strTipo & "_ret.ret_codigo=retencion.ret_codigo " & _
             " WHERE det_" & strTipo & "_ret.emp_codigo='" & strEmpresa & "' " & _
             " AND det_" & strTipo & "_ret.tip_" & Left(strTipo, 3) & "_codigo='" & cmbCliente.BoundText & "'" & _
             " AND det_" & strTipo & "_ret." & Left(strTipo, 3) & "_codigo='" & Me.cmbCotizacion.BoundText & "' " & _
             " AND '" & txtFechaR.Text & "' BETWEEN ret_fechaini AND ret_fechafin " & _
             " ORDER BY ret_nombre DESC"
    clsSql.Ejecutar strSQL
    VSFGRet.Clear 1
    Set VSFGRet.DataSource = clsSql.adorec_Def.DataSource
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Label1.Enabled = False
        Fecha1.Enabled = False
        Label2.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        cmdBuscar.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        Label1.Enabled = True
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.Value = 1 Then
            Label2.Enabled = True
            Fecha2.Enabled = True
        End If
        cmdBuscar.Enabled = True
    End If
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1", "yyyy-mm-dd")
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal, "yyyy-mm-dd")
        DiaFinal = DiaFinal - 1
    Wend
    cmdBuscar.Enabled = True
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub chkFechas_Click()
    If chkFechas.Value = 1 Then
        Label1.Caption = "Fecha Inicial"
        Label2.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Label1.Caption = "Fecha"
        Label2.Enabled = False
        Fecha2.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub
