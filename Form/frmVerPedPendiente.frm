VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerPedPendiente 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Pedidos Pendientes de Facturar"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmVerPedPendiente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   9225
   Begin VB.CommandButton cmdLiberar 
      Caption         =   "Liberar NC"
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton CmdPicking 
      Caption         =   "Pasar a Picking"
      Height          =   375
      Left            =   1680
      TabIndex        =   40
      Top             =   9600
      Width           =   1455
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
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9015
      Begin VB.OptionButton optPreVenta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "PreVeta"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7440
         TabIndex        =   43
         Top             =   720
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optNC 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Notas de Credito"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7440
         TabIndex        =   41
         Top             =   240
         Width           =   1500
      End
      Begin VB.OptionButton optPedPendiente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pendientes"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5880
         TabIndex        =   38
         Top             =   960
         Width           =   1500
      End
      Begin VB.OptionButton optPedGuardado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Borrador"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   240
         Width           =   1380
      End
      Begin VB.OptionButton optPedAsignado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Procesado"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5880
         TabIndex        =   36
         Top             =   480
         Width           =   1500
      End
      Begin VB.OptionButton optPedConfirmado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Confirmados"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5880
         TabIndex        =   35
         Top             =   720
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   375
         Width           =   4455
         _ExtentX        =   7858
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
         Left            =   1200
         TabIndex        =   20
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   66322435
         CurrentDate     =   37463
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   21
         Top             =   780
         Width           =   495
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
         TabIndex        =   17
         Top             =   420
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdPreFactura 
      Caption         =   "Ver PreFactura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   9015
      Begin VB.Frame Frame3 
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
         Height          =   1695
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   8535
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
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   240
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
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1320
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
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   480
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
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
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
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   1215
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
            Height          =   1095
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   4305
            _cx             =   1994202538
            _cy             =   1994196875
            Appearance      =   1
            BorderStyle     =   1
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
            FormatString    =   $"frmVerPedPendiente.frx":030A
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total pedido:"
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
            Left            =   5940
            TabIndex        =   34
            Top             =   1350
            Width           =   1065
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
            Left            =   5940
            TabIndex        =   33
            Top             =   990
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento:"
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
            Left            =   5940
            TabIndex        =   32
            Top             =   510
            Width           =   825
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
            Left            =   5940
            TabIndex        =   31
            Top             =   750
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal pedido:"
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
            Left            =   5880
            TabIndex        =   30
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   8715
         _cx             =   1994210316
         _cy             =   1994199204
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerPedPendiente.frx":038A
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
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
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
      Begin MSDataListLib.DataCombo CmbTipoFac 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbTC 
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   690
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista:"
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
         Left            =   6000
         TabIndex        =   19
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento (%):"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Factura:"
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
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   60
      End
      Begin VB.Label LblDetalle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle del Pedido Nº"
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
         Left            =   330
         TabIndex        =   9
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Pedidos:"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   9015
      Begin VB.CheckBox chkSeleccionar 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Seleccionar todos"
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
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "ACT"
         Height          =   1815
         Left            =   8640
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8475
         _cx             =   1994209893
         _cy             =   1994198145
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
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerPedPendiente.frx":04B7
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
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   360
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
      End
   End
   Begin VB.CommandButton CmdDeBaja 
      Caption         =   "Dar de Baja"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton cmdBajaPendiente 
      Caption         =   "Baja x Pendiente"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   0
      Picture         =   "frmVerPedPendiente.frx":072D
      ToolTipText     =   "Elimina una Fila"
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   240
      Picture         =   "frmVerPedPendiente.frx":0863
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmVerPedPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsPedidos As New clsConsulta
Private clsPed As New clsConsulta
Private clsSql As New clsConsulta
Private clsTFac As New clsConsulta
Private clsTC As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsFPago As New clsConsulta
Private clsFormaPago As New clsConsulta
Private clsRet As New clsConsulta
Private clsLstPrds As New clsConsulta
Private clsExis As New clsConsulta
Private IVA As Double
Private CodPer As String
Private codCot As String
Private codVen As String
Private codTC As String
Private FechaUltFac As String
Private strClaveMAESTRA As String
Private SecPublico As Boolean
Private SinIVA As Boolean
Private maxItem As Long


Private Sub chkSeleccionar_Click()
    Dim tip As Integer
    tip = 0
    If CBool(chkSeleccionar.Value) = True Then
        tip = 1
    End If
    For i = 1 To VSFGPeds.Rows - 1
        If VSFGPeds.TextMatrix(i, 8) = "Facturado" Or VSFGPeds.TextMatrix(i, 8) = "De Baja" Then
            VSFGPeds.TextMatrix(i, 0) = 0
        Else
            VSFGPeds.TextMatrix(i, 0) = tip
        End If
    Next i
End Sub

Private Sub cmbNegocio_Change()
    Dim intEstado As Integer
    cmdLiberar.Enabled = False
    If optPreVenta.Value = True Then
        intEstado = -3
    ElseIf optPedPendiente.Value = True Then
        intEstado = -2
    ElseIf Me.optPedGuardado.Value = True Then
        intEstado = -1
    ElseIf Me.optPedAsignado.Value = True Then
        intEstado = 0
    ElseIf Me.optNC.Value = True Then
        intEstado = 0
        cmdLiberar.Enabled = True
    ElseIf Me.optPedConfirmado.Value = True Then
        intEstado = 1
    End If
    cmdLimpiar_Click
    If cmbNegocio.BoundText <> "" Then
        
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strPtoFactura = clsSql.adorec_Def(0)
            If Me.optNC.Value = True Then
                'Consulta todos los pedidos que pasan a bodega para ser revisados
                strSql = " SELECT '' as seleccionar,ped_codigo, ped_fecha, ped_observacion as nombC, " & _
                         " '' as nombV,'' as nombG,'' as nombD, " & _
                         " ped_observacion, est_descripcion, '' as tipo_fac_descripcion, pedido.per_codigo, cot_codigo, " & _
                         " pedido.ven_codigo,'' as per_observacion,pedido.tar_cre_codigo,'' as tar_cre_nombre,''," & _
                         " 0,0,0,0,'' " & _
                         " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
                         " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado='" & intEstado & "' " & _
                         " AND ped_fecha<='" & Format(Me.dtpFecha.Value, "yyyy-mm-dd") & " 23:59:59' AND ped_codigo LIKE CONCAT('100001'+0,'%') " & _
                         " ORDER BY ped_estado,ped_codigo "
                clsPedidos.Ejecutar (strSql)
            Else
                'Consulta todos los pedidos que pasan a bodega para ser revisados
                strSql = " SELECT '' as seleccionar,ped_codigo, ped_fecha, CONCAT(persona.per_apellido,' ',persona.per_nombre) as nombC, " & _
                         " CONCAT(ven_apellido,' ',ven_nombre) as nombV,CONCAT(GZ.per_apellido,' ',GZ.per_nombre) as nombG,CONCAT(DI.per_apellido,' ',DI.per_nombre) as nombD, " & _
                         " ped_observacion, est_descripcion, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
                         " pedido.ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
                         " persona.per_sec_publico,persona.per_siniva,persona.per_fac_flete,persona.per_dcto,pedido.tar_cre_codigo " & _
                         " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
                         " INNER JOIN persona ON pedido.per_codigo = persona.per_codigo " & _
                         " AND pedido.emp_codigo = persona.emp_codigo " & _
                         " INNER JOIN tipo_factura ON pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
                         " AND vendedor.emp_codigo = persona.emp_codigo " & _
                         " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
                         " AND persona.per_codigo_ref = GZ.per_codigo " & _
                         " LEFT JOIN persona as DI ON persona.emp_codigo = DI.emp_codigo " & _
                         " AND persona.per_codigo_ref2 = DI.per_codigo " & _
                         " LEFT JOIN tarjeta_credito ON pedido.emp_codigo = tarjeta_credito.emp_codigo AND pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo " & _
                         " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado='" & intEstado & "' " & _
                         " AND ped_fecha<='" & Format(Me.dtpFecha.Value, "yyyy-mm-dd") & " 23:59:59' AND (ped_codigo LIKE CONCAT('" & strSucursal & strPtoFactura & "'+0,'%') OR ped_codigo LIKE CONCAT('" & "001" & strPtoFactura & "'+0,'%')) " & _
                         " ORDER BY ped_estado,ped_codigo "
                clsPedidos.Ejecutar (strSql)
            End If
        End If
    Else
        Exit Sub
    End If
    clsPedidos.Actualizar
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
End Sub



Private Sub cmbTC_Change()
    Dim i As Long, comis As Double
    Dim ProdAux As String
    Dim CantAux As Long
    Dim Tipo As Integer
    Dim cade As String
    
    If cmbTC.MatchedWithList = True Then
        If cmbTC <> "" Then
        '****** TARJETAS
        strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
                 " FROM tarjeta_credito " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " ORDER BY tar_cre_nombre "
        clsTC.Ejecutar strSql
    
        clsTC.Filtrar "tar_cre_codigo='" & cmbTC.BoundText & "' "
        Tipo = FormatoD0(clsTC.adorec_Def("tip_com_codigo"))
        
        If Tipo = 1 Then 'No comision
            ''dblComision = clsTC.adorec_Def("tar_cre_porcentaje")
            cade = ""
            dblComision = "0"
        ElseIf Tipo = 2 Then
            cade = ""
            dblComision = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
        ElseIf Tipo = 3 Then
            dblComision = "0"
            comis = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
        End If
    
            strSqlPrd = " SELECT dep_codigo, det_pedido.prd_codigo, det_ped_precio,prd_nombre,ROUND((det_ped_cant_entregada * det_ped_precio)-IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & "),2) as total,det_ped_cant_entregada,det_ped_cant_pedida, " & _
                    " (producto.prd_costo/(1 - 0.1)) as prd_costo, det_ped_precio*'" & 1 + dblComision / 100# & "' as lis_precio,prd_cambia_precio, " & _
                    " (IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & ")) as dcto, " & _
                    " ROUND((det_ped_cant_entregada * det_ped_precio)-IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 19)) & ")*(" & 1 + dblComision / 100# & "),2) as totales, " & _
                    " det_ped_precio,det_ped_dcto " & _
                    " FROM pedido INNER JOIN det_pedido ON pedido.ped_codigo = det_pedido.ped_codigo " & _
                    " AND pedido.emp_codigo = det_pedido.emp_codigo " & _
                    " INNER JOIN persona ON persona.per_codigo=pedido.per_codigo AND persona.emp_codigo=pedido.emp_codigo " & _
                    " INNER JOIN categoria_p ON persona.cat_p_codigo=categoria_p.cat_p_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.emp_codigo=categoria_p.emp_codigo " & _
                    " INNER JOIN producto " & _
                    " ON det_pedido.emp_codigo = producto.emp_codigo AND det_pedido.prd_codigo = producto.prd_codigo " & _
                    " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                    " INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                    " AND producto.emp_codigo=lista_precio_p.emp_codigo AND lista_precio_p.lis_pre_codigo=categoria_p.lis_pre_codigo " & _
                    " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                    " AND producto_promo.prd_pro_fechaini<=LEFT(pedido.ped_fechamod,10) AND producto_promo.prd_pro_fechafin>=LEFT(pedido.ped_fechamod,10) AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                    " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                    " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                    " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                    " det_pedido.ped_codigo='" & VSFGPeds.TextMatrix(VSFGPeds.Row, 0) & "' " & _
                    " ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo "
            clsLstPrds.Ejecutar strSqlPrd
    
            VSFG.Tag = "P"
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 1) <> "" Then
                ProdAux = VSFG.TextMatrix(i, 1)
                'CantAux = VSFG.TextMatrix(i, 4)
                VSFG.TextMatrix(i, 1) = ""
                VSFG.TextMatrix(i, 1) = ProdAux
                'VSFG.TextMatrix(i, 4) = CantAux
            End If
        Next i
        VSFG.Tag = "N"
            If Tipo = 3 Then
                strSql = " SELECT prd_codigo,prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='PR-TAR' "
                clsSql.Ejecutar strSql
                If clsSql.adorec_Def.RecordCount > 0 Then
                    VSFG.AddItem VSFG.TextMatrix(1, 0) & vbTab & clsSql.adorec_Def(0) & vbTab & clsSql.adorec_Def(1) & vbTab & "1" & vbTab & "1" & vbTab & (FormatoD4(TxtSubTotal.Text) * (FormatoD4(comis) / 100#)) & vbTab & "0" & vbTab & (FormatoD4(TxtSubTotal.Text) * (FormatoD4(comis) / 100#)) & vbTab & "1"
                End If
            Else
                For i = 1 To VSFG.Rows - 1
                    If VSFG.TextMatrix(i, 1) = "PR-TAR" Then
                        VSFG.RemoveItem i
                        Exit For
                    End If
                Next i
            End If
            
            
        End If
    End If

    CalcuTotal
    'txtComi.Text = FormatoD2(FormatoD4(TxtTotal.Text) * FormatoD4(comis / 100.00))
End Sub

Private Sub cmdBajaPendiente_Click()
    'Verifica si no se completa el pedido
    Dim Resp As Integer, Peds As String
    Peds = ""
    
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            Peds = Peds & VSFGPeds.TextMatrix(i, 1) & ","
        End If
    Next i
    Peds = IIf(Right(Peds, 1) = ",", Left(Peds, Len(Peds) - 1), Peds)
    'Verifica que el usuario esté seguro de dar de baja al pedido
    Resp = MsgBox("Está seguro de dar de baja al pedido Nº. " & Peds, vbInformation + vbYesNo, "De Baja")
    If Resp = vbNo Then
        Exit Sub
    End If
    'Verifica si se quiere realizar un backorder de los productos no completados en el pedido
    
'****** PEDIDO
    'Da de baja al pedido
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            strSql = " UPDATE pedido SET ped_estado=6, " & _
                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
                     " ped_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & VSFGPeds.TextMatrix(i, 1)
            clsSql.Ejecutar (strSql), "M"
            LiberarIncentivo VSFGPeds.TextMatrix(i, 1)
        '****** COTIZACION
            'Actualiza el estado de la cotización relacionada a vigente en caso de que esta exista
            codCot = VSFGPeds.TextMatrix(i, 11)
            If codCot <> "" Then
                strSql = " UPDATE cotizacion SET cot_estado=0, " & _
                         " cot_fechamod=CURRENT_TIMESTAMP, " & _
                         " cot_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & codCot & "' "
                clsSql.Ejecutar (strSql), "M"
            End If
        End If
    Next i
    'Limpia los grids que mostraban datos del pedido
    
    MsgBox "Pedido No. " & Peds & " dado(s) de Baja.", vbInformation, "De Baja"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource

End Sub

Private Sub cmdLiberar_Click()
    'Verifica si no se completa el pedido
    Dim Resp As Integer, Peds As String
    Peds = ""
    
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            Peds = Peds & VSFGPeds.TextMatrix(i, 1) & ","
        End If
    Next i
    Peds = IIf(Right(Peds, 1) = ",", Left(Peds, Len(Peds) - 1), Peds)
    'Verifica que el usuario esté seguro de dar de baja al pedido
    Resp = MsgBox("Está seguro de dar de baja al pedido Nº. " & Peds, vbInformation + vbYesNo, "De Baja")
    If Resp = vbNo Then
        Exit Sub
    End If
    'Verifica si se quiere realizar un backorder de los productos no completados en el pedido
    
'****** PEDIDO
    'Da de baja al pedido
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            strSql = " UPDATE pedido SET ped_estado=7, " & _
                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
                     " ped_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & VSFGPeds.TextMatrix(i, 1)
            clsSql.Ejecutar (strSql), "M"
            
        '****** COTIZACION
            'Actualiza el estado de la cotización relacionada a vigente en caso de que esta exista
            codCot = VSFGPeds.TextMatrix(i, 11)
            If codCot <> "" Then
                strSql = " UPDATE cotizacion SET cot_estado=0, " & _
                         " cot_fechamod=CURRENT_TIMESTAMP, " & _
                         " cot_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & codCot & "' "
                clsSql.Ejecutar (strSql), "M"
            End If
        End If
    Next i
    'Limpia los grids que mostraban datos del pedido
    
    MsgBox "Pedido No. " & Peds & " dado(s) de Baja.", vbInformation, "De Baja"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource

End Sub

Private Sub CmdPicking_Click()
    'Verifica si no se completa el pedido
    Dim Resp As Integer, Peds As String
    Peds = ""
    
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            Peds = Peds & VSFGPeds.TextMatrix(i, 1) & ","
        End If
    Next i
    Peds = IIf(Right(Peds, 1) = ",", Left(Peds, Len(Peds) - 1), Peds)
    'Verifica que el usuario esté seguro de dar de baja al pedido
    Resp = MsgBox("Está seguro de Pasar a Picking " & VSFGPeds.Rows - 1 & " al pedido Nº. " & Peds, vbInformation + vbYesNo, "Picking")
    If Resp = vbNo Then
        Exit Sub
    End If
    'Verifica si se quiere realizar un backorder de los productos no completados en el pedido
    
'****** PEDIDO
    'Da de baja al pedido
    
    strSql = " EXEC Sp_RevisionExistencias"
    clsSql.Ejecutar (strSql), "M"
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
'txtObser.Text = txtObser.Text & vbNewLine & "Estado Ant: " & clsSql.adorec_Def("est_descripcion") & " / Fecha: " & clsSql.adorec_Def("ped_fecha")
'            strSQL = " UPDATE pedido,est_pedido SET ped_estado=0, " & _
'                     " ped_observacion=CONCAT(ped_observacion,'" & vbNewLine & "Estado Ant: ',est_descripcion,' / Fecha: ',ped_fecha), " & _
'                     " ped_fecha=CURRENT_TIMESTAMP, " & _
'                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
'                     " ped_usumod='" & strUsuario & "' " & _
'                     " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_codigo=" & VSFGPeds.TextMatrix(i, 1) & _
'                     " AND pedido.ped_estado=est_pedido.est_codigo"
'            clsSql.Ejecutar (strSQL), "M"
'            If VSFGPeds.TextMatrix(i, 1) = "10020956948" Then
'                MsgBox "AAa"
'            End If
            strSql = " EXEC Sp_PasaPedidoAPicking '" & strEmpresa & "','" & VSFGPeds.TextMatrix(i, 1) & "'"
            clsSql.Ejecutar (strSql), "M"
        '****** COTIZACION
            'Actualiza el estado de la cotización relacionada a vigente en caso de que esta exista
            codCot = VSFGPeds.TextMatrix(i, 11)
            If codCot <> "" Then
                strSql = " UPDATE cotizacion SET cot_estado=0, " & _
                         " cot_fechamod=CURRENT_TIMESTAMP, " & _
                         " cot_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & codCot & "' "
                clsSql.Ejecutar (strSql), "M"
            End If
        End If
    Next i
    'Limpia los grids que mostraban datos del pedido
    
    MsgBox "Pedido No. " & Peds & " pasado(s) a Picking.", vbInformation, "Picking"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource

End Sub

Private Sub cmdPreFactura_Click()
    Dim RepFactura As New frmReporte
    Dim cadena As String, i As Long
    cadena = ""
    strSql = " DROP TABLE IF EXISTS recs "
    clsSql.Ejecutar strSql
    
    strSql = " CREATE TEMPORARY TABLE recs( " & _
             " cod VARCHAR(5)," & _
             " prod VARCHAR(20)," & _
             " prec DECIMAL(14,2)) "
    clsSql.Ejecutar strSql
    For i = 1 To VSFGReca.Rows - 1
        If VSFGReca.TextMatrix(i, 1) <> "" Then
            strSql = " INSERT INTO recs VALUES('" & VSFGReca.TextMatrix(i, 1) & "','" & VSFGReca.TextMatrix(i, 2) & "'," & FormatoD4(VSFGReca.TextMatrix(i, 3)) & ")"
            clsSql.Ejecutar strSql
        End If
    Next i
    RepFactura.strNumero = VSFGPeds.TextMatrix(VSFGPeds.Row, 1)
    RepFactura.strTipo = FormatoD2(TxtRecargo.Text)
    RepFactura.strAsiento = IVA
    RepFactura.strReporte = "rptPreFactura"
    RepFactura.Show
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
    TxtRecargo = FormatoD2(Suma)
End Sub

Private Sub CalcuTotal()
    'Calcula es total del pedido
    Dim Suma As Double
    Dim SumaIVA As Double
    Dim SumaDcto As Double
    Dim sumaSDcto As Double
    Dim SumaIVASDcto As Double
    
    CalcuReca
    Suma = 0
    SumaIVA = 0
    sumaSDcto = 0
    SumaIVASDcto = 0
    SumaDcto = 0
    For i = 1 To VSFG.Rows - 1
        If Val(FormatoD2(VSFG.TextMatrix(i, 4))) <> 0 Then
            If Abs(FormatoD0(VSFG.TextMatrix(i, 8))) = 1 Then
                Suma = Suma + FormatoD2(FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5))) - FormatoD4(VSFG.TextMatrix(i, 6)))
                sumaSDcto = sumaSDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
            Else
                SumaIVA = SumaIVA + FormatoD2(FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5))) - FormatoD4(VSFG.TextMatrix(i, 6)))
                SumaIVASDcto = SumaIVASDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
            End If
        SumaDcto = SumaDcto + FormatoD2(VSFG.TextMatrix(i, 6))
        End If
    Next i
    TxtRecargo.Tag = FormatoD2(TxtRecargo.Text) + FormatoD2(SumaIVA)
    TxtRecargo.Text = FormatoD2(TxtRecargo.Text) + FormatoD2(SumaIVASDcto)
    'Coloca los totales parciales de la factura
    TxtDesc.Text = FormatoD2(SumaDcto)
    TxtSubTotal = FormatoD2(sumaSDcto)
    If SinIVA = False Then
        TxtIva = FormatoD2((Suma) * IVA / 100)
    Else
        TxtIva = 0
    End If
    TxtTotal = FormatoD2(Suma + TxtIva + Val(TxtRecargo.Tag))
End Sub

Private Sub CalcuDesc()
    Dim strSql As String
    TxtDesc = 0
    If VSFGPeds.Row > 1 Then
        strSql = " SELECT COALESCE(SUM(det_ped_cant_entregada*det_ped_precio),0) as suman," & _
                 " SUM(ROUND(det_ped_cant_entregada*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2)) as Descu " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_codigo=" & VSFGPeds.TextMatrix(VSFGPeds.Row, 1) & " GROUP BY pedido.ped_codigo"
        clsPed.Ejecutar (strSql)
        If clsPed.adorec_Def.RecordCount > 0 Then
           ' TxtDesc.Text = FormatoD2(clsPed.adorec_Def("Descu"))
        Else
            TxtDesc.Text = "0.00"
        End If
    End If
End Sub

Private Sub cmdActualizar_Click()
    'TmrAct_Timer
    strSql = " SELECT '' as seleccionar,ped_codigo, ped_fecha, CONCAT(persona.per_apellido,' ',persona.per_nombre) as nombC, " & _
             " CONCAT(ven_apellido,' ',ven_nombre) as nombV,CONCAT(GZ.per_apellido,' ',GZ.per_nombre) as nombG,CONCAT(DI.per_apellido,' ',DI.per_nombre) as nombD, " & _
             " ped_observacion, est_descripcion, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
             " pedido.ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
             " persona.per_sec_publico,persona.per_siniva,persona.per_fac_flete,persona.per_dcto,pedido.tar_cre_codigo " & _
             " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
             " INNER JOIN persona ON pedido.per_codigo = persona.per_codigo " & _
             " AND pedido.emp_codigo = persona.emp_codigo " & _
             " INNER JOIN tipo_factura ON pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
             " AND vendedor.emp_codigo = persona.emp_codigo " & _
             " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
             " AND persona.per_codigo_ref = GZ.per_codigo " & _
             " LEFT JOIN persona as DI ON persona.emp_codigo = DI.emp_codigo " & _
             " AND persona.per_codigo_ref2 = DI.per_codigo " & _
             " LEFT JOIN tarjeta_credito ON pedido.emp_codigo = tarjeta_credito.emp_codigo AND pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado<>0 " & _
             " AND (ped_fecha='" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' OR ped_estado=1) AND ped_codigo LIKE CONCAT('" & strSucursal & strPtoFactura & "'+0,'%') " & _
             " ORDER BY ped_estado,ped_fechamod,ped_codigo "
    clsPedidos.Ejecutar (strSql)
    
    clsPedidos.Actualizar
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Function PrdEntregar() As Long
    Dim i As Long
    Dim num As Long
    num = 0
    For i = 1 To VSFG.Rows - 1
        If FormatoD4(VSFG.TextMatrix(i, 5)) > 0 Then
            num = num + 1
        End If
    Next i
    PrdEntregar = num
End Function


Private Sub CmdDeBaja_Click()
    'Verifica si no se completa el pedido
    Dim Resp As Integer, Peds As String
    Peds = ""
    
        
    
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            Peds = Peds & VSFGPeds.TextMatrix(i, 1) & ","
        End If
    Next i
    Peds = IIf(Right(Peds, 1) = ",", Left(Peds, Len(Peds) - 1), Peds)
    'Verifica que el usuario esté seguro de dar de baja al pedido
    Resp = MsgBox("Está seguro de dar de baja al pedido Nº. " & Peds, vbInformation + vbYesNo, "De Baja")
    If Resp = vbNo Then
        Exit Sub
    End If
    'Verifica si se quiere realizar un backorder de los productos no completados en el pedido
    
'****** PEDIDO
    'Da de baja al pedido
    For i = 1 To VSFGPeds.Rows - 1
        If CBool(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = True Then
            strSql = " UPDATE pedido SET ped_estado=3, " & _
                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
                     " ped_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & VSFGPeds.TextMatrix(i, 1)
            clsSql.Ejecutar (strSql), "M"
            LiberarIncentivo VSFGPeds.TextMatrix(i, 1), VSFGPeds.TextMatrix(i, 10)
        '****** COTIZACION
            'Actualiza el estado de la cotización relacionada a vigente en caso de que esta exista
            codCot = VSFGPeds.TextMatrix(i, 11)
            If codCot <> "" Then
                strSql = " UPDATE cotizacion SET cot_estado=0, " & _
                         " cot_fechamod=CURRENT_TIMESTAMP, " & _
                         " cot_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & codCot & "' "
                clsSql.Ejecutar (strSql), "M"
            End If
        End If
    Next i
    'Limpia los grids que mostraban datos del pedido
    
    MsgBox "Pedido No. " & Peds & " dado(s) de Baja.", vbInformation, "De Baja"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
End Sub


Private Sub cmdLimpiar_Click()
    'Limpia el contenido del grid de detalles
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGReca.Clear 1
    VSFGReca.Rows = 2
    VSFGReca.Enabled = False

    cmdPreFactura.Enabled = False

    TxtSubTotal = ""
    TxtTotal = ""
    TxtRecargo = ""
    TxtIva = ""
    TxtDesc = ""
    txtDescuento = ""
    cmbTC = ""
    LblPedido = "-"
    CmbTipoFac = ""
    CmbFpago = ""
    TxtObserv = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
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
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set ucrtVSFG.VSFGControl = VSFGPeds
    ucrtVSFG.Inicializar False, False, False, True, True, True, False, False, True, "900"
    'Inicializa los objetos de conexión con la base de datos
    clsPedidos.Inicializar AdoConn, AdoConnMaster
    clsPed.Inicializar AdoConn, AdoConnMaster
    clsTC.Inicializar AdoConn, AdoConnMaster
    clsTFac.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsLstPrds.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    clsRet.Inicializar AdoConn, AdoConnMaster
    clsExis.Inicializar AdoConn, AdoConnMaster
    clsFormaPago.Inicializar AdoConn, AdoConnMaster
    '****** CREDITO
    VSFG.Tag = "N"
    'Coloca los datos de los vendedores en un listado
    
    '****** TARJETAS
    strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
             " FROM tarjeta_credito " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY tar_cre_nombre "
    clsTC.Ejecutar (strSql)
    Set cmbTC.RowSource = clsTC.adorec_Def.DataSource
    cmbTC.ListField = "tar_cre_nombre"
    cmbTC.BoundColumn = "tar_cre_codigo"
    'cmbTC.BoundText = "EFE"
    
    dtpFecha.Value = HoyDia
    
    'Consulta todos los pedidos que pasan a bodega para ser revisados
    strSql = " SELECT '' as seleccionar,ped_codigo, ped_fecha, CONCAT(persona.per_apellido,' ',persona.per_nombre) as nombC, " & _
             " CONCAT(ven_apellido,' ',ven_nombre) as nombV,CONCAT(GZ.per_apellido,' ',GZ.per_nombre) as nombG,CONCAT(DI.per_apellido,' ',DI.per_nombre) as nombD, " & _
             " ped_observacion, est_descripcion, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
             " pedido.ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
             " persona.per_sec_publico,persona.per_siniva,persona.per_fac_flete,persona.per_dcto,pedido.tar_cre_codigo " & _
             " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
             " INNER JOIN persona ON pedido.per_codigo = persona.per_codigo " & _
             " AND pedido.emp_codigo = persona.emp_codigo " & _
             " INNER JOIN tipo_factura ON pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
             " AND vendedor.emp_codigo = persona.emp_codigo " & _
             " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
             " AND persona.per_codigo_ref = GZ.per_codigo " & _
             " LEFT JOIN persona as DI ON persona.emp_codigo = DI.emp_codigo " & _
             " AND persona.per_codigo_ref2 = DI.per_codigo " & _
             " LEFT JOIN tarjeta_credito ON pedido.emp_codigo = tarjeta_credito.emp_codigo AND pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado<>0 " & _
             " AND (ped_fecha='" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' OR ped_estado=1) AND ped_codigo LIKE CONCAT('" & strSucursal & strPtoFactura & "'+0,'%') " & _
             " ORDER BY ped_estado,ped_codigo "
    clsPedidos.Ejecutar (strSql)
    'Muestra los datos de los distintos proyectos de trabajo en un listado
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
    
       
    cargarTipoPedido
    
    'Consulta los recargos que puede manejar una empresa
    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Obtiene el IVA vigente para realizar la factura
    IVA = PorIVA
    LblIva = "IVA " & IVA & " %:"

    'Consulta todos los tipos de factura y los muestra en un combo
    strSql = " SELECT * FROM tipo_factura ORDER BY tipo_fac_descripcion"
    clsTFac.Ejecutar (strSql)
    Set CmbTipoFac.RowSource = clsTFac.adorec_Def.DataSource
    CmbTipoFac.ListField = "tipo_fac_descripcion"
    CmbTipoFac.BoundColumn = "tipo_fac_codigo"
    'Coloca los botones de eliminar fila en el grid de recargos
    PonerBotones
    'Coloca la fecha actual
    
End Sub



Private Sub optNC_Click()
    cmbNegocio_Change
    
End Sub

Private Sub optPedAsignado_Click()
    cmbNegocio_Change
End Sub

Private Sub optPedConfirmado_Click()
    cmbNegocio_Change
End Sub

Private Sub optPedGuardado_Click()
    cmbNegocio_Change
End Sub

Private Sub optPedPendiente_Click()
    cmbNegocio_Change
End Sub

Private Sub optPreVenta_Click()
    cmbNegocio_Change
End Sub

Private Sub VSFGPeds_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    Else
        If VSFGPeds.TextMatrix(Row, 8) = "Facturado" Or VSFGPeds.TextMatrix(Row, 8) = "De Baja" Then
            Cancel = True
        End If
    
    End If
End Sub

Private Sub VSFGPeds_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Marca toda la fila con otra tonalidad si el pedido puede ser vendido
    If Col = 8 And Row <> 0 And VSFGPeds.TextMatrix(Row, Col) = "Confirmado" Then
        VSFGPeds.Select Row, 0, Row, VSFGPeds.Cols - 1
        VSFGPeds.FillStyle = flexFillRepeat
        VSFGPeds.CellBackColor = &HC0C0FF
    End If
End Sub


Private Sub VSFGPeds_DblClick()
    FechaUltFac = ""
    Dim strForma As String, codDC As String
    Dim Fp As String
    'Verifica cuando se da un doble click sobre una fila del grid de pedidos
'''    dtpFecha.value = HoyDia
    'Limpia el grid de recargos
    VSFGReca.Rows = 2
    VSFGReca.Clear 1
    VSFG.Tag = "N"
    
    If VSFGPeds.Row > 0 Then
        'Consulta el detalle de un pedido específico
        strSql = " SELECT dep_codigo, det_pedido.prd_codigo, prd_nombre, det_ped_cant_pedida, det_ped_cant_entregada, det_ped_precio,  " & _
                 " IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 20)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 20)) & ") as det_ped_dcto," & _
                 " ROUND((det_ped_cant_entregada * det_ped_precio)-det_ped_dcto,2) as total,prd_iva " & _
                 " FROM ((pedido INNER JOIN det_pedido ON (pedido.ped_codigo = det_pedido.ped_codigo) " & _
                 " AND (pedido.emp_codigo = det_pedido.emp_codigo)) INNER JOIN producto " & _
                 " ON (det_pedido.emp_codigo = producto.emp_codigo) AND (det_pedido.prd_codigo = producto.prd_codigo)) " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND producto_promo.prd_pro_fechaini<=LEFT(pedido.ped_fechamod,10) AND producto_promo.prd_pro_fechafin>=LEFT(pedido.ped_fechamod,10) AND producto_promo.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                 " det_pedido.ped_codigo= " & VSFGPeds.TextMatrix(VSFGPeds.Row, 1) & _
                 " ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo "
        clsPed.Ejecutar (strSql)
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsPed.adorec_Def.DataSource
        'Muestra el número del pedido a modificar
        LblPedido.Caption = VSFGPeds.TextMatrix(VSFGPeds.Row, 1)
        'Muestra el tipo de factura sugerido
        CmbTipoFac = VSFGPeds.TextMatrix(VSFGPeds.Row, 9)
        'Obtiene el código del cliente
        CodPer = VSFGPeds.TextMatrix(VSFGPeds.Row, 10)
        codVen = VSFGPeds.TextMatrix(VSFGPeds.Row, 12)
        codTC = VSFGPeds.TextMatrix(VSFGPeds.Row, 14)
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 17))) = 0 Then
            SecPublico = False
        Else
            SecPublico = True
        End If
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 18))) = 0 Then
            SinIVA = False
        Else
            SinIVA = True
        End If
        
        txtDescuento.Text = FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 20))
        cmbTC.BoundText = ""
        cmbTC.BoundText = VSFGPeds.TextMatrix(VSFGPeds.Row, 21)
        '*** FACTURA FLETE
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 19))) = 1 Then
            VSFGReca.TextMatrix(1, 1) = "FLC"
            VSFGReca.AddItem ""
            PonerBotones
        End If
        
        'Obtiene el código de la cotización relacionada
        codCot = VSFGPeds.TextMatrix(VSFGPeds.Row, 11)
        'Calcula los totales de la factura
        CalcuDesc
        CalcuTotal
        
        'Verifica que no se haya facturado o que haya algún producto en el pedido para no poder volver a facturarlo
        If VSFGPeds.TextMatrix(VSFGPeds.Row, 8) = "Facturado" Or VSFGPeds.TextMatrix(VSFGPeds.Row, 8) = "De Baja" Then
            cmdPreFactura.Enabled = False
            VSFGReca.Enabled = False
        Else
            cmdPreFactura.Enabled = True
            VSFGReca.Enabled = False
        End If
        If VSFGPeds.TextMatrix(VSFGPeds.Row, 13) <> "" Then
            MsgBox VSFGPeds.TextMatrix(VSFGPeds.Row, 13), vbInformation, "Observaciones"
        End If
        'Verifica que no se pueda volver a dar de baja un pedido ya bajado
        
    End If
End Sub

Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
    'And VSFGReca.TextMatrix(OldRow, 1) <> ""
    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 Then
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
             CalcuDesc
             CalcuTotal
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
    CalcuDesc
    CalcuTotal
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

