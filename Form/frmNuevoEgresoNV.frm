VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNuevoEgresoNV 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egreso de Mercaderia"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "frmNuevoEgresoNV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdAplicaDcto 
      Caption         =   "Aplicar &Dcto."
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   8550
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   1815
      Left            =   6840
      TabIndex        =   44
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtAutorizacionR 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1350
         Width           =   1335
      End
      Begin VB.TextBox txtSerieR 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtDocumentoR 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFechaR 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         Format          =   64028675
         CurrentDate     =   37463
      End
      Begin VB.Label Label15 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   210
         TabIndex        =   48
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
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
         Left            =   210
         TabIndex        =   47
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Autorizacion"
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
         Left            =   210
         TabIndex        =   46
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
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
         Left            =   210
         TabIndex        =   45
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   9975
      Begin VB.TextBox txtCaduca 
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1060
      End
      Begin VB.TextBox txtNumAux 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   615
         Width           =   1815
      End
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtAutorizacion 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   990
         Width           =   1815
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
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
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
         Format          =   64028675
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSComCtl2.DTPicker dtpCaduca 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   8400
         TabIndex        =   9
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         CustomFormat    =   "MM/yyy"
         Format          =   64028675
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   7920
         TabIndex        =   51
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   6960
         TabIndex        =   52
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lblFpago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         TabIndex        =   49
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Caducidad del Doc"
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
         Left            =   6960
         TabIndex        =   43
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
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
         Left            =   3840
         TabIndex        =   35
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizacion"
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
         Left            =   3840
         TabIndex        =   34
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
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
         Left            =   3840
         TabIndex        =   33
         Top             =   270
         Width           =   375
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
         TabIndex        =   32
         Top             =   300
         Width           =   1035
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
         TabIndex        =   31
         Top             =   630
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Egreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   6615
      Begin MSDataListLib.DataCombo dcmbCodP 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Top             =   660
         Width           =   5625
         _ExtentX        =   9922
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
      Begin MSDataListLib.DataCombo cmbTDoc 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   300
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblCodProveedor 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Persona"
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
         TabIndex        =   27
         Top             =   720
         Width           =   750
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
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   9975
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "Copiar Desde"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   975
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
         Height          =   2175
         Left            =   1155
         TabIndex        =   36
         Top             =   3600
         Width           =   7935
         Begin VB.TextBox TxtObserv 
            Height          =   285
            Left            =   360
            MaxLength       =   250
            TabIndex        =   22
            Top             =   1800
            Width           =   7335
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   360
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1440
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
            Left            =   6480
            TabIndex        =   18
            Top             =   600
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   840
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1080
            Width           =   1215
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
            Height          =   1095
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   4305
            _cx             =   7594
            _cy             =   1931
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
            FormatString    =   $"frmNuevoEgresoNV.frx":030A
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
            Left            =   420
            TabIndex        =   42
            Top             =   1560
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
            Left            =   5580
            TabIndex        =   41
            Top             =   1470
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recargos:"
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
            Left            =   5580
            TabIndex        =   40
            Top             =   1110
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Var SubT:"
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
            Left            =   5580
            TabIndex        =   39
            Top             =   630
            Width           =   735
         End
         Begin VB.Label LblIva 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IVA X%"
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
            Left            =   5580
            TabIndex        =   38
            Top             =   870
            Width           =   570
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
            Left            =   5580
            TabIndex        =   37
            Top             =   390
            Width           =   630
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalle 
         Height          =   3330
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   9735
         _cx             =   17171
         _cy             =   5874
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmNuevoEgresoNV.frx":038A
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
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   120
         Picture         =   "frmNuevoEgresoNV.frx":04DB
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   360
         Picture         =   "frmNuevoEgresoNV.frx":0611
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   8550
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5197
      TabIndex        =   24
      Top             =   8550
      Width           =   1455
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
      TabIndex        =   28
      Top             =   6120
      Width           =   1410
   End
End
Attribute VB_Name = "frmNuevoEgresoNV"
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
Private clsCon_TipDoc As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsFPago As New clsConsulta
Private strSQL As String
Private ModPreCos As Boolean
Private PreCos As String
Private IngAsi As Boolean
Private DctoTotal As Boolean

Private Sub CmbFpago_Change()
    cmdAplicaDcto_Click
End Sub

Private Sub cmbTDoc_Change()
    clsCon_TipDoc.Filtrar "tip_egr_codigo='" & cmbTDoc.BoundText & "' "
    If dcmbCodP.Tag <> "S" Then
        CargarPersonas clsCon_TipDoc.adorec_Def("tip_egr_persona")
    End If
    If Val(clsCon_TipDoc.adorec_Def("tip_egr_retencion")) = 1 Then
        fraRet.Visible = True
        If Left(clsCon_TipDoc.adorec_Def("tip_egr_cx_p_c"), 1) = "P" Then
            strSQL = " SELECT COALESCE(com_ret_serie,'') as com_ret_serie,COALESCE(com_ret_numero,'0')+1 as com_ret_numero,COALESCE(com_ret_autorizacion,'') as com_ret_autorizacion " & _
                     " FROM comprobante_retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='" & Left(clsCon_TipDoc.adorec_Def("tip_egr_cx_p_c"), 1) & "' " & _
                     " ORDER BY com_ret_numero DESC LIMIT 0,1 "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                txtSerieR.Text = clsCon_Def.adorec_Def("com_ret_serie")
                txtDocumentoR.Text = clsCon_Def.adorec_Def("com_ret_numero")
                txtAutorizacionR.Text = clsCon_Def.adorec_Def("com_ret_autorizacion")
                dtpFechaR.Value = Format(HoyDia, "yyyy-mm-dd")
            End If
        End If
    Else
        fraRet.Visible = False
    End If
    If Right(clsCon_TipDoc.adorec_Def("tip_egr_cx_p_c"), 1) = "S" Then
        IngAsi = True
    Else
        IngAsi = False
    End If
    If Left(clsCon_TipDoc.adorec_Def("tip_egr_cx_p_c"), 1) <> "N" Then
        lblFpago.Visible = True
        CmbFpago.Visible = True
    Else
        lblFpago.Visible = False
        CmbFpago.Visible = False
    End If
    CmbFpago.Tag = Left(clsCon_TipDoc.adorec_Def("tip_egr_cx_p_c"), 1)
    If Right(clsCon_TipDoc.adorec_Def("tip_egr_cos_pre"), 1) = "S" Then
        ModPreCos = True
    Else
        ModPreCos = False
    End If
    PreCos = Left(clsCon_TipDoc.adorec_Def("tip_egr_cos_pre"), 1)
    'CargaProductos
    If clsCon_TipDoc.adorec_Def("tip_egr_persona") = "N" Then
        dcmbCodP.BoundText = "%"
    End If
    If Val(clsCon_TipDoc.adorec_Def("tip_egr_impuesto")) = 1 Then
        TxtIva.Enabled = True
    Else
        TxtIva.Enabled = False
        TxtIva.Text = 0
    End If
    If Val(clsCon_TipDoc.adorec_Def("tip_egr_recargo")) = 1 Then
        VSFGReca.Enabled = True
    Else
        VSFGReca.Enabled = False
        TxtRecargo = 0
    End If
    If clsCon_TipDoc.adorec_Def("tip_egr_numsri") = "F" Then
        txtSerie.Text = strSucursal & strPtoFactura
        txtDocumento.Text = ""
        txtAutorizacion.Text = strAutorFactura
        txtCaduca.Text = ""
        txtSerie.Locked = True
        txtDocumento.Locked = True
        txtAutorizacion.Locked = True
        dtpCaduca.Enabled = False
    ElseIf clsCon_TipDoc.adorec_Def("tip_egr_numsri") = "P" Then
        strSQL = " SELECT COALESCE(egr_serie,'') as egr_serie,COALESCE(egr_numero,'0')+1 as egr_numero,COALESCE(egr_autorizacion,'') as egr_autorizacion,COALESCE(egr_caduca,'00/0000') as egr_caduca " & _
                 " FROM egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_egr_codigo='" & cmbTDoc.BoundText & "' " & _
                 " AND per_codigo LIKE '" & dcmbCodP.BoundText & "' " & _
                 " ORDER BY egr_fecha DESC,egr_numero DESC,egr_codigo DESC LIMIT 0,1 "
        clsCon_Def.Ejecutar strSQL
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtSerie.Text = clsCon_Def.adorec_Def("egr_serie")
            txtDocumento.Text = clsCon_Def.adorec_Def("egr_numero")
            txtAutorizacion.Text = clsCon_Def.adorec_Def("egr_autorizacion")
            If clsCon_Def.adorec_Def("egr_caduca") <> "00/0000" And clsCon_Def.adorec_Def("egr_caduca") <> "" Then
                dtpCaduca.Value = clsCon_Def.adorec_Def("egr_caduca")
                txtCaduca.Text = clsCon_Def.adorec_Def("egr_caduca")
            Else
                dtpCaduca.Value = Format(HoyDia, "mm\/yyyy")
                txtCaduca.Text = ""
            End If
        Else
            txtSerie.Text = ""
            txtCaduca.Text = ""
            txtDocumento.Text = ""
            txtAutorizacion.Text = ""
            dtpCaduca.Value = Format(HoyDia, "mm\/yyyy")
        End If
        txtSerie.Locked = False
        txtDocumento.Locked = False
        txtAutorizacion.Locked = False
        dtpCaduca.Enabled = True
    End If
    If cmbTDoc.BoundText = "NOT" Then
        cmdAplicaDcto.Visible = False
    Else
        cmdAplicaDcto.Visible = False
    End If
    
    cargarVendedor (cmbTDoc.BoundText)
    
End Sub

Private Sub cmdAceptar_Click()
    Dim clsEgreso As New clsInventario
    Dim clsAsiento As New clsContable
    Dim clsCta As New clsCtaXx
    Dim i As Long
    Dim strObserv As String
    Dim booGuardar As Boolean
    Dim TotalRet As Double
    Dim cue_p_c_codigo As Double
    Dim strTipCompAsiento As String
    Dim ConfNum As Boolean
    Dim DblSub As Double
    If fraRet.Visible = True Then
        If Trim(txtSerieR.Text) = "" Or Trim(txtDocumentoR.Text) = "" Or Trim(txtAutorizacionR.Text) = "" Then
            MsgBox "Llene los campos de la Retencion", vbInformation, "Retenciones"
            Exit Sub
        End If
    End If
    If CmbFpago.BoundText = "TAR" Then
        DblSub = FormatoD2(TxtSubTotal.Text) - FormatoD2(TxtDesc.Text)
        vsfgDetalle.AddItem ""
        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = "PRI"
        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 2) = "PR-TAR"
        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 4) = 1
        'vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 5) = FormatoD2(DblSub * 0.1)
    End If
    If cmbTDoc.BoundText = "NOT" Then
        If Me.cmbVendedor.MatchedWithList = False Then
            MsgBox "Llene los campo de Vendedor", vbInformation, "Vendedores"
            Exit Sub
        End If
    End If
    If CmbFpago.Visible = False Then
        If CmbFpago.MatchedWithList = True Then
            MsgBox "Seleccione la Forma de Pago", vbInformation, "Forma de Pago"
            Exit Sub
        End If
    End If
    If txtSerie.Locked = False And txtDocumento.Locked = False And txtAutorizacion.Locked = False Then
        If Trim(txtSerie.Text) = "" Or Trim(txtDocumento.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
            MsgBox "Llene los campos del Documento", vbInformation, "Documento"
            Exit Sub
        End If
    End If
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    If cmbTDoc.BoundText = "NOT" Then
         If cmbVendedor.Text = "" Then
            MsgBox "Seleccione un vendedor", vbInformation, "Vendedor"
            Exit Sub
        Else
            booGuardar = clsEgreso.NuevoEgr(True, cmbTDoc.BoundText, True, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.Value, txtNumAux.Text, cmbVendedor.BoundText, UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text))
        End If
    Else
        booGuardar = clsEgreso.NuevoEgr(True, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.Value, txtNumAux.Text, , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text))
    End If
    If booGuardar = True Then
        If clsEgreso.strTipo = "DPV" Then
            strTipCompAsiento = "N"
        End If
        strObserv = UCase(cmbTDoc.Text & clsEgreso.strDoc & vbNewLine & "PERSONA: " & dcmbCodP.Text & vbNewLine & "DOCUMENTO: " & txtSerie.Text & Format(txtDocumento.Text, "0000000") & vbNewLine & TxtObserv.Text)
        IngAsi = False
        If IngAsi = True Then
            'clsAsiento.Inicializar AdoConn, AdoConnMaster
            'clsAsiento.NuevoAsiento "V", dtpFecha.value, 0, 0, TxtTotal.Text, strObserv
            'clsEgreso.ModificaEgr , , , , , , clsAsiento.NumAsiento
        End If
        With vsfgDetalle
            For i = 1 To .Rows - 1
                clsEgreso.NuevoDetEgr .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD0(.TextMatrix(i, 4)), FormatoD4(.TextMatrix(i, 5)), FormatoD4(.TextMatrix(i, 8)), FormatoD4(.TextMatrix(i, 6)), Abs(FormatoD0(.TextMatrix(i, 9)))
            Next i
        End With
        With VSFGReca
            For i = 1 To .Rows - 1
                clsEgreso.NuevoDetEgrRecargo .TextMatrix(i, 1), FormatoD2(.TextMatrix(i, 3))
            Next i
        End With
        clsEgreso.DetRetenciones

        If IngAsi = True And CmbFpago.Visible = True Then
            clsFPago.adorec_Def.MoveFirst
            strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
            clsFPago.adorec_Def.Find strComparar
            'Inserta un nuevo registro de la cuenta por cobrar*/
            clsCta.Inicializar AdoConn, AdoConnMaster
            If CmbFpago.Tag = "P" Then
                clsCta.NuevaCta CmbFpago.Tag, 1, "02", dtpFecha, Format(IIf(CmbFpago.Visible = True, DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha), dtpFecha), "yyyy-MM-dd"), dcmbCodP.BoundText, strObserv, txtSerie.Text, txtDocumento.Text, txtAutorizacion.Text, txtCaduca.Text, FormatoD2(clsEgreso.dblTotalProd), FormatoD2(clsEgreso.dblTotalServ), FormatoD2(clsEgreso.dblTotalProdIVA), FormatoD2(clsEgreso.dblTotalServIVA), 2, FormatoD2(clsEgreso.dblIVA), FormatoD2(clsEgreso.dblSubTotal0), 0, 0, 0, FormatoD2(clsEgreso.dblTotal), clsAsiento.NumAsiento
                'clsCta.IngRetencionPersonaIng clsEgreso, txtSerieR.Text, txtDocumentoR.Text, txtAutorizacion.Text, Format(dtpFechaR.Value, "yyyy-mm-dd")
                clsCta.IngAsientoEgr clsAsiento, clsEgreso
                MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
                If fraRet.Visible = True Then
                    clsCta.VerRet
                End If
            ElseIf CmbFpago.Tag = "C" Then
                Dim TipDoc As String
                If cmbTDoc.BoundText = "NOT" Then
                    TipDoc = "2"
                Else
                    TipDoc = "1"
                End If
                clsCta.NuevaCta CmbFpago.Tag, 1, "00", dtpFecha, Format(IIf(CmbFpago.Visible = True, DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha), dtpFecha), "yyyy-MM-dd"), dcmbCodP.BoundText, strObserv, txtSerie.Text, txtDocumento.Text, txtAutorizacion.Text, txtCaduca.Text, FormatoD2(clsEgreso.dblTotalProd), FormatoD2(clsEgreso.dblTotalServ), FormatoD2(clsEgreso.dblTotalProdIVA), FormatoD2(clsEgreso.dblTotalServIVA), 2, FormatoD2(clsEgreso.dblIVA), FormatoD2(clsEgreso.dblSubTotal0), 0, 0, 0, FormatoD2(clsEgreso.dblTotal), clsAsiento.NumAsiento
                'clsCta.IngRetencionPersonaegr clsIngreso, txtSerieR.Text, txtDocumentoR.Text, txtAutorizacion.Text, Format(dtpFechaR.Value, "yyyy-mm-dd")
                clsCta.IngAsientoEgr clsAsiento, clsEgreso
                MsgBox " Los datos han sido ingresado", vbInformation, "Egresos"
                If fraRet.Visible = True Then
                    clsCta.VerRet
                End If
            ElseIf CmbFpago.Tag = "S" Then
                clsCta.IngAsientoEgr clsAsiento, clsEgreso
                MsgBox " Los datos han sido ingresado", vbInformation, "Egresos"
            End If
            Set clsCta = Nothing
            Set clsAsiento = Nothing
        End If
        
        Dim rpMov As New frmReporte
        'rpMov.strNumero = clsEgreso.strDoc
        'rpMov.strTipo = clsEgreso.strTipo
        'rpMov.strReporte = "rptEgresoMercaderia"
        'rpMov.Show
        
        If cmbTDoc.BoundText = "NOT" Then
            Dim RepNotaVenta As New frmReporte
            RepNotaVenta.strNumero = clsEgreso.strDoc
            RepNotaVenta.strReporte = "rptNotaVenta"
            RepNotaVenta.Show
            'frmCobrosNotasVenta.txtValor = FormatoD2(Me.TxtTotal)
            'frmCobrosNotasVenta.CodigoPersona = dcmbCodP.BoundText
            'frmCobrosNotasVenta.Descripcion = dcmbCodP & " - NOTA DE VENTA: " & clsEgreso.strDoc
            'frmCobrosNotasVenta.Show 1
        End If
        
        Unload Me
    End If
End Sub

Private Sub cmdAplicaDcto_Click()
    Dim strFiltro As String
    Dim Dcto As Double
    Dim i As Long
    If CmbFpago.MatchedWithList = True Then
        strFiltro = "for_pag_codigo='" & CmbFpago.BoundText & "'"
        clsFPago.Filtrar strFiltro
        Dcto = 0
        For i = 1 To vsfgDetalle.Rows - 1
            vsfgDetalle.TextMatrix(i, 6) = FormatoD4(FormatoD4(vsfgDetalle.TextMatrix(i, 4)) * FormatoD4(vsfgDetalle.TextMatrix(i, 5)) * (Dcto + FormatoD2(vsfgDetalle.TextMatrix(i, 10))) / 100#)
        Next i
    End If
End Sub

Private Sub cmdCopiar_Click()
    frmVerMovimiento.Tag = "E"
    frmVerMovimiento.cmdCopiar.Visible = True
    frmVerMovimiento.cmdAnular.Visible = False
End Sub



Private Sub dcmbCodP_Change()
    dcmbCodP.Tag = "S"
    cmbTDoc_Change
    dcmbCodP.Tag = ""
End Sub

Private Sub dtpCaduca_Change()
    txtCaduca.Text = Format(dtpCaduca.Value, "mm\/yyyy")
End Sub

'Private Sub Form_Activate()
'    CargaProductos
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_TipDoc = Nothing
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_TipDoc.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    dtpFecha = HoyDia
    dtpFechaR = HoyDia
    DctoTotal = False
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'IVA
    TxtIva.Tag = PorIVA
    'Carga los egresos
    strSQL = " SELECT tip_egr_codigo, tip_egr_nombre,tip_egr_impuesto,tip_egr_persona,tip_egr_cx_p_c,tip_egr_recargo,tip_egr_numsri,tip_egr_cos_pre,tip_egr_retencion " & _
             " FROM tipo_egreso WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_TipDoc.Ejecutar strSQL
    cmbTDoc.ListField = "tip_egr_nombre"
    cmbTDoc.BoundColumn = "tip_egr_codigo"
    Set cmbTDoc.RowSource = clsCon_TipDoc.adorec_Def.DataSource
    'Carga los depositos
    strSQL = "SELECT dep_codigo, dep_nombre FROM deposito WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSQL
    vsfgDetalle.ColComboList(1) = vsfgDetalle.BuildComboList(clsCon_Def.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    
    'Consulta los recargos que puede manejar una empresa
    strSQL = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSQL)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Insertamos el botón de eliminar en cada una de las filas
    vsfgDetalle.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalle.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSQL = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar (strSQL)
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
    
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        End Select
End Sub


Private Sub cargarVendedor(Tipo As String)
    If UCase(Tipo) = "NOT" Then
        strSQL = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) nombre " & _
                 " FROM vendedor WHERE emp_codigo = '" & strEmpresa & "' "
        clsCon_Def.Ejecutar strSQL
        Set cmbVendedor.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbVendedor.ListField = "nombre"
        cmbVendedor.BoundColumn = "ven_codigo"
    Else
        cmbVendedor.Visible = False
        lblVendedor.Visible = False
    End If
End Sub


Private Sub TxtDesc_Validate(Cancel As Boolean)
    Dim i As Long
    Dim Dcto As Double
    Dim subST As Double
    DctoTotal = True
    subST = 0
    Dcto = FormatoD2(TxtDesc.Text)
    For i = 1 To vsfgDetalle.Rows - 1
        If FormatoD2(subST - FormatoD2(TxtSubTotal.Text)) = 0 Then Exit For
        vsfgDetalle.TextMatrix(i, 6) = FormatoD4((FormatoD2(vsfgDetalle.TextMatrix(i, 4)) * FormatoD2(vsfgDetalle.TextMatrix(i, 5))) / FormatoD2(FormatoD2(TxtSubTotal.Text) - subST) * FormatoD4(Dcto))
        Dcto = Dcto - vsfgDetalle.TextMatrix(i, 6)
        subST = subST + (FormatoD2(vsfgDetalle.TextMatrix(i, 4)) * FormatoD2(vsfgDetalle.TextMatrix(i, 5)))
    Next i
    DctoTotal = False
End Sub

Private Sub vsfgDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col >= 8 Then
        Cancel = True
    End If
    If ModPreCos = False Then
        If Col = 5 Or Col = 6 Or Col = 7 Then
            Cancel = True
        End If
    End If
    If Col = 2 Or Col = 3 Then
        If Trim(Me.vsfgDetalle.TextMatrix(Row, 1)) <> "" Then
            CargaProductos Row
        End If
    End If
End Sub

Private Sub vsfgDetalle_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalle.MouseRow
    c = vsfgDetalle.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = 1) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalle.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
   
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalle.Cell(flexcpLeft, r, c) + vsfgDetalle.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "AHORA DEBE ELIMINAR ESTA FILA!"
    
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        
        'Recorro el FlexGrid para almacenar los detalles del ingreso
        
        If respuesta = vbYes Then
            Dim i As Long
            vsfgDetalle.RemoveItem (r)
            For i = 1 To (vsfgDetalle.Rows - 1)
                vsfgDetalle.TextMatrix(i, 0) = i
                vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            Next i
            CalculaTotal
        Else
            vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnUp
        
        End If
    Cancel = True
End Sub

Private Sub vsfgDetalle_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        vsfgDetalle.TextMatrix(Row, 3) = vsfgDetalle.TextMatrix(Row, 2)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalle.TextMatrix(Row, 2) & "'"
        If vsfgDetalle.TextMatrix(Row, 2) <> "PR-TAR" Then
            vsfgDetalle.TextMatrix(Row, 4) = 0
            vsfgDetalle.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_precio")
            vsfgDetalle.TextMatrix(Row, 6) = 0
            vsfgDetalle.TextMatrix(Row, 7) = 0
            vsfgDetalle.TextMatrix(Row, 8) = clsCon_Prd.adorec_Def("prd_costo")
            vsfgDetalle.TextMatrix(Row, 9) = clsCon_Prd.adorec_Def("prd_iva")
            vsfgDetalle.TextMatrix(Row, 10) = clsCon_Prd.adorec_Def("lis_pre_p_promo")
        Else
            vsfgDetalle.TextMatrix(Row, 4) = 0
            vsfgDetalle.TextMatrix(Row, 5) = FormatoD2((FormatoD2(TxtSubTotal.Text) - FormatoD2(TxtDesc.Text)) * 0.1)
            vsfgDetalle.TextMatrix(Row, 6) = 0
            vsfgDetalle.TextMatrix(Row, 7) = 0
            vsfgDetalle.TextMatrix(Row, 8) = 0
            vsfgDetalle.TextMatrix(Row, 9) = 1
            vsfgDetalle.TextMatrix(Row, 10) = 0
        
        End If
    ElseIf Col = 3 Then
        vsfgDetalle.TextMatrix(Row, 2) = vsfgDetalle.TextMatrix(Row, 3)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalle.TextMatrix(Row, 2) & "'"
        vsfgDetalle.TextMatrix(Row, 4) = 0
        vsfgDetalle.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_precio")
        vsfgDetalle.TextMatrix(Row, 6) = 0
        vsfgDetalle.TextMatrix(Row, 7) = 0
        vsfgDetalle.TextMatrix(Row, 8) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalle.TextMatrix(Row, 9) = clsCon_Prd.adorec_Def("prd_iva")
        vsfgDetalle.TextMatrix(Row, 10) = clsCon_Prd.adorec_Def("lis_pre_p_promo")
    ElseIf Col = 4 Or Col = 5 Or Col = 6 Then
        If Col = 4 Then
            vsfgDetalle.TextMatrix(Row, 6) = 0
        End If
        vsfgDetalle.TextMatrix(Row, 7) = FormatoD2(FormatoD0(vsfgDetalle.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalle.TextMatrix(Row, 5)) - FormatoD4(vsfgDetalle.TextMatrix(Row, 6)))
        CalculaTotal
    ElseIf Col = 7 Then
        If FormatoD0(vsfgDetalle.TextMatrix(Row, 4)) <> 0 Then
            vsfgDetalle.TextMatrix(Row, 5) = FormatoD4((FormatoD2(vsfgDetalle.TextMatrix(Row, 7)) + FormatoD4(vsfgDetalle.TextMatrix(Row, 6))) / FormatoD0(vsfgDetalle.TextMatrix(Row, 4)))
        Else
            vsfgDetalle.TextMatrix(Row, 5) = 0
        End If
        CalculaTotal
    ElseIf Col = 1 Then
        vsfgDetalle.TextMatrix(Row, 2) = ""
        vsfgDetalle.TextMatrix(Row, 3) = ""
        vsfgDetalle.TextMatrix(Row, 4) = ""
        vsfgDetalle.TextMatrix(Row, 5) = ""
        vsfgDetalle.TextMatrix(Row, 6) = ""
        vsfgDetalle.TextMatrix(Row, 7) = ""
        vsfgDetalle.TextMatrix(Row, 8) = ""
        vsfgDetalle.TextMatrix(Row, 9) = 0
        vsfgDetalle.TextMatrix(Row, 10) = 0
        CalculaTotal
    End If
    If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 2) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 3) <> "" And Val(vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 4)) <> 0 Then
        If vsfgDetalle.Rows >= 28 And Me.cmbTDoc.BoundText = "NOT" Then
            MsgBox "Ya supero los 28 productos, debe hacer una nueva nota de venta", vbInformation, "Nota de Venta"
        Else
            vsfgDetalle.AddItem ""
            vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = strBodega
            vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 0) = vsfgDetalle.Rows - 1
            vsfgDetalle.Cell(flexcpPicture, vsfgDetalle.Rows - 1, 0) = imgBtnUp
            vsfgDetalle.Cell(flexcpPictureAlignment, vsfgDetalle.Rows - 1, 0) = flexAlignRightCenter
        End If
    End If
    
End Sub
Private Sub CalculaTotal()
    Dim i As Long
    Dim subIVA As Double
    Dim Suma As Double
    Dim subIVASDcto As Double
    Dim sumaSDcto As Double
    Suma = 0
    subIVA = 0
    sumaSDcto = 0
    subIVASDcto = 0
    If DctoTotal = False Then
        TxtDesc.Text = 0
    End If
    CalcuReca
    For i = 1 To vsfgDetalle.Rows - 1
        If Abs(FormatoD0(vsfgDetalle.TextMatrix(i, 9))) = 0 Then
            subIVA = FormatoD2(FormatoD2(subIVA) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)))
            subIVASDcto = FormatoD2(FormatoD2(subIVASDcto) + FormatoD4(vsfgDetalle.TextMatrix(i, 4)) * FormatoD4(vsfgDetalle.TextMatrix(i, 5)))
        Else
            Suma = FormatoD2(FormatoD2(Suma) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)))
            sumaSDcto = FormatoD2(FormatoD2(sumaSDcto) + FormatoD4(vsfgDetalle.TextMatrix(i, 4)) * FormatoD4(vsfgDetalle.TextMatrix(i, 5)))
        End If
        If DctoTotal = False Then
            TxtDesc.Text = FormatoD2(FormatoD2(TxtDesc.Text) + FormatoD4(vsfgDetalle.TextMatrix(i, 6)))
        End If
    Next i
    TxtSubTotal.Text = sumaSDcto
    TxtRecargo.Tag = FormatoD2(FormatoD2(TxtRecargo.Text) + subIVA)
    TxtRecargo.Text = FormatoD2(FormatoD2(TxtRecargo.Text) + subIVASDcto)
    If TxtIva.Enabled = True Then
        TxtIva.Text = FormatoD2(FormatoD2(Suma) * Val(TxtIva.Tag) / 100#)
    End If
    TxtTotal.Text = FormatoD2(Suma) + FormatoD2(TxtIva.Text) + FormatoD2(TxtRecargo.Tag)
End Sub

Private Sub vsfgDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
Private Sub CargarPersonas(Tipo As String)
'Carga los personas
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')') per " & _
             " FROM persona " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND cat_p_tipo LIKE '" & Tipo & "'" & _
             " ORDER BY CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')')"
    clsCon_Def.Ejecutar strSQL
    dcmbCodP.ListField = "per"
    dcmbCodP.BoundColumn = "per_codigo"
    Set dcmbCodP.RowSource = clsCon_Def.adorec_Def.DataSource
End Sub
Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 And VSFGReca.TextMatrix(OldRow, 1) <> "" Then
        VSFGReca.AddItem ""
        PonerBotones
    End If
End Sub

Private Sub VSFGReca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 0 del recargo
    If Col = 2 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    With VSFGReca
        ' only interesetd in left button
        If Button <> 1 Then Exit Sub
        
        ' get cell that was clicked
        Dim r&, c&
        r = .MouseRow
        c = .MouseCol
        
        ' make sure the click was on the sheet
        If r < 0 Or c < 0 Then Exit Sub
        
        If (c <> 0 Or r = (.Rows - 1)) Then Exit Sub
         
        ' make sure the click was on a cell with a button
        If .Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        
        ' click was on a button: do the work
         .Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
            
        'Recorro el FlexGrid para poner números a las filas
            
        If respuesta = vbYes Then
             Dim i As Integer
              .RemoveItem (r)
             PonerBotones
             CalculaTotal
        Else
             .Cell(flexcpPicture, r, c) = imgBtnUp
        End If
        
        ' cancel default processing
        ' note: this is not strictly necessary in this case, because
        '       the dialog box already stole the focus etc, but let's be safe.
        Cancel = True
    End With
End Sub

Private Sub VSFGReca_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Busca y coloca el valor del recargo seleccionado
    If Row > 0 And VSFGReca.TextMatrix(Row, 1) <> "" And Col <> 3 Then
        clsRecargos.Filtrar "oca_codigo='" & VSFGReca.TextMatrix(Row, 1) & "'"
        VSFGReca.TextMatrix(Row, 2) = clsRecargos.adorec_Def("oca_nombre")
        VSFGReca.TextMatrix(Row, 3) = clsRecargos.adorec_Def("oca_precio")
        clsRecargos.QuitarFiltro
        'Verifica que no se haya escogido antes el mismo recargo, en ese caso suma sus valores
        For i = 1 To VSFGReca.Rows - 1
            If VSFGReca.TextMatrix(Row, 1) = VSFGReca.TextMatrix(i, 1) And Row <> i Then
                VSFGReca.TextMatrix(i, 3) = Val(VSFGReca.TextMatrix(i, 3)) + (VSFGReca.TextMatrix(Row, 3))
                VSFGReca.RemoveItem Row
                PonerBotones
                Exit For
            End If
        Next i
    End If
    CalculaTotal
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With VSFGReca
        For i = 1 To (.Rows - 1)
            .TextMatrix(i, 0) = i
            If conBot = True Then
                'Coloca los botones de elimniar fila en el grid
                .Cell(flexcpPicture, i, 0) = imgBtnUp
                .Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub

Private Sub CalcuReca()
    'Calcula el total del pedido
    Dim Suma As Double
    For i = 1 To VSFGReca.Rows - 1
        Suma = Suma + FormatoD2(VSFGReca.TextMatrix(i, 3))
    Next i
    TxtRecargo = Format(Suma, "####0.00")
End Sub

Private Sub CargaProductos(Optional Row As Long)
    If vsfgDetalle.TextMatrix(Row, 2) = "" Or vsfgDetalle.TextMatrix(Row, 3) = "" Or vsfgDetalle.ColComboList(2) = "" Then
        If PreCos = "C" Then
            'Carga los productos
            strSQL = " SELECT producto.prd_codigo, prd_nombre,prd_costo,prd_costo as prd_precio,0 as lis_pre_p_promo,prd_iva, SUM(existencia.exi_cantidad) as exi_cantidad " & _
                     " FROM producto INNER JOIN existencia ON existencia.prd_codigo=producto.prd_codigo AND existencia.emp_codigo=producto.emp_codigo AND existencia.dep_codigo='" & Me.vsfgDetalle.TextMatrix(Row, 1) & "'" & _
                     " WHERE producto.emp_codigo = '" & strEmpresa & "' AND prd_baja=0 GROUP BY prd_codigo HAVING exi_cantidad!=0 ORDER BY prd_codigo "
            clsCon_Prd.Ejecutar strSQL
            vsfgDetalle.ColComboList(2) = vsfgDetalle.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre, exi_cantidad", "prd_codigo")
            'Consulto los productos de la empresa
            strSQL = " SELECT producto.prd_codigo, prd_nombre,prd_costo,prd_costo as prd_precio,0 as lis_pre_p_promo,prd_iva, SUM(existencia.exi_cantidad) as exi_cantidad " & _
                     " FROM producto INNER JOIN existencia ON existencia.prd_codigo=producto.prd_codigo AND existencia.emp_codigo=producto.emp_codigo AND existencia.dep_codigo='" & Me.vsfgDetalle.TextMatrix(Row, 1) & "'" & _
                     " WHERE producto.emp_codigo = '" & strEmpresa & "' AND prd_baja=0 GROUP BY prd_codigo HAVING exi_cantidad!=0 ORDER BY prd_nombre "
            clsCon_Def.Ejecutar strSQL
            vsfgDetalle.ColComboList(3) = vsfgDetalle.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre, exi_cantidad", "prd_codigo")
        Else
            'Carga los productos
            strSQL = " SELECT producto.prd_codigo, prd_nombre,prd_costo,lis_pre_p_precio as prd_precio,lis_pre_p_promo,prd_iva, SUM(existencia.exi_cantidad) as exi_cantidad " & _
                     " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo" & _
                     " INNER JOIN producto ON persona.emp_codigo=producto.emp_codigo AND prd_baja=0 " & _
                     " INNER JOIN existencia ON existencia.prd_codigo=producto.prd_codigo AND existencia.emp_codigo=producto.emp_codigo AND existencia.dep_codigo='" & Me.vsfgDetalle.TextMatrix(Row, 1) & "' AND existencia.exi_cantidad!=0 " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo AND producto.prd_codigo=lista_precio_p.prd_codigo AND categoria_p.lis_pre_codigo=lista_precio_p.lis_pre_codigo AND lis_pre_p_precio!=0" & _
                     " WHERE persona.emp_codigo = '" & strEmpresa & "' AND persona.per_codigo='" & dcmbCodP.BoundText & "'" & _
                     " GROUP BY prd_codigo " & _
                     " HAVING exi_cantidad!=0 " & _
                     " ORDER BY prd_codigo "
            clsCon_Prd.Ejecutar strSQL
            vsfgDetalle.ColComboList(2) = vsfgDetalle.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre, exi_cantidad", "prd_codigo")
            'Consulto los productos de la empresa
            strSQL = " SELECT producto.prd_codigo, prd_nombre,prd_costo,lis_pre_p_precio as prd_precio,lis_pre_p_promo,prd_iva, SUM(existencia.exi_cantidad) as exi_cantidad " & _
                     " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo" & _
                     " INNER JOIN producto ON persona.emp_codigo=producto.emp_codigo AND prd_baja=0 " & _
                     " INNER JOIN existencia ON existencia.prd_codigo=producto.prd_codigo AND existencia.emp_codigo=producto.emp_codigo AND existencia.dep_codigo='" & Me.vsfgDetalle.TextMatrix(Row, 1) & "' AND existencia.exi_cantidad!=0 " & _
                     " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo AND producto.prd_codigo=lista_precio_p.prd_codigo AND categoria_p.lis_pre_codigo=lista_precio_p.lis_pre_codigo AND lis_pre_p_precio!=0 " & _
                     " WHERE persona.emp_codigo = '" & strEmpresa & "' AND persona.per_codigo='" & dcmbCodP.BoundText & "'" & _
                     " GROUP BY prd_codigo " & _
                     " HAVING exi_cantidad!=0 " & _
                     " ORDER BY prd_nombre "
            clsCon_Def.Ejecutar strSQL
            vsfgDetalle.ColComboList(3) = vsfgDetalle.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre, exi_cantidad", "prd_codigo")
        End If
    End If
End Sub


