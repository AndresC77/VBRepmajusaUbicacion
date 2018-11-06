VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPagos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11760
   Begin VB.CheckBox chkNombre2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "El cheque sale con otro nombre"
      Height          =   375
      Left            =   7920
      TabIndex        =   43
      Top             =   9120
      Width           =   2760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   11535
      Begin TabDlg.SSTab SSTab1 
         Height          =   2295
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "CXP"
         TabPicture(0)   =   "frmPagos.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "VSFG1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "CXC"
         TabPicture(1)   =   "frmPagos.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vsfgCXC"
         Tab(1).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid VSFG1 
            Height          =   1815
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   10695
            _cx             =   99699121
            _cy             =   99683457
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   8388608
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPagos.frx":0342
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
            PicturesOver    =   -1  'True
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   1
            OwnerDraw       =   0
            Editable        =   1
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
         Begin VSFlex8Ctl.VSFlexGrid vsfgCXC 
            Height          =   1815
            Left            =   -74880
            TabIndex        =   54
            Top             =   360
            Width           =   10695
            _cx             =   99699121
            _cy             =   99683457
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   8388608
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPagos.frx":0505
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
            PicturesOver    =   -1  'True
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
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   6000
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   8760
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   645
         Left            =   2340
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   6480
         Width           =   7215
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   8760
         Width           =   1275
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   8760
         Width           =   1275
      End
      Begin VB.OptionButton optproveedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optcliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   0
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9585
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Forma de Pago"
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
         Height          =   2535
         Left            =   480
         TabIndex        =   32
         Top             =   3000
         Width           =   2295
         Begin VB.OptionButton optTransferenciaAuto 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Transferencia Auto"
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
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton optTransferencia 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Transferencia"
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
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   800
            Width           =   1455
         End
         Begin VB.OptionButton optOtros 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Otros"
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
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton optNDebito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Nota de Débito"
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
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1640
            Width           =   1455
         End
         Begin VB.OptionButton optcheque 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Cheque"
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
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   520
            Width           =   1095
         End
         Begin VB.OptionButton optNCredito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Nota de Crédito"
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
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   1360
            Width           =   1455
         End
         Begin VB.OptionButton optefectivo 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Efectivo"
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
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtSaldoReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   10245
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   6960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDisponible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   10245
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   7680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPrevisto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   10245
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   8640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10245
         TabIndex        =   28
         Top             =   7920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9525
         TabIndex        =   27
         Top             =   8880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   2655
         Left            =   6840
         TabIndex        =   22
         Top             =   3360
         Width           =   4335
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   285
            Left            =   1560
            TabIndex        =   50
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            Value           =   42719.3532523148
         End
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   1605
            TabIndex        =   9
            Top             =   1080
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   1605
            TabIndex        =   10
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1605
            TabIndex        =   11
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbTipo 
            Height          =   315
            Left            =   1605
            TabIndex        =   12
            Top             =   2160
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin NEED2.dtpFecha dtpFechaDoc 
            Height          =   285
            Left            =   1560
            TabIndex        =   51
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            Value           =   42719.3532523148
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Doc:"
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
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de nota:"
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
            TabIndex        =   40
            Top             =   2175
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria:"
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
            TabIndex        =   26
            Top             =   1845
            Width           =   1245
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
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
            TabIndex        =   25
            Top             =   1485
            Width           =   510
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de documento:"
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
            TabIndex        =   24
            Top             =   1110
            Width           =   1350
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Pago:"
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
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1275
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGRet 
         Height          =   2535
         Left            =   2880
         TabIndex        =   18
         Top             =   3480
         Width           =   3855
         _cx             =   54270608
         _cy             =   54268279
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPagos.frx":0705
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
         AutoSizeMouse   =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   1260
         TabIndex        =   14
         Top             =   7200
         Width           =   8880
         _cx             =   54279471
         _cy             =   54266586
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPagos.frx":0799
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNota 
         Height          =   315
         Left            =   675
         TabIndex        =   45
         Top             =   5640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
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
         TabIndex        =   47
         Top             =   5685
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         TabIndex        =   46
         Top             =   6000
         Width           =   450
      End
      Begin VB.Label Label8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Colocar cuenta contable del proveedor en la fila 2"
         Height          =   855
         Left            =   120
         TabIndex        =   42
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   1440
         TabIndex        =   39
         Top             =   6480
         Width           =   900
      End
      Begin VB.Label lbltotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Left            =   4500
         TabIndex        =   38
         Top             =   8820
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   2820
         Picture         =   "frmPagos.frx":0889
         ToolTipText     =   "Elimina una Fila"
         Top             =   8040
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   3180
         Picture         =   "frmPagos.frx":09BF
         Top             =   8040
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario:"
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
         Left            =   2640
         TabIndex        =   37
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL A PAGAR:"
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
         Left            =   8145
         TabIndex        =   36
         Top             =   3030
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "Real"
         Enabled         =   0   'False
         Height          =   255
         Left            =   10245
         TabIndex        =   35
         Top             =   6600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "disponible"
         Enabled         =   0   'False
         Height          =   255
         Left            =   10245
         TabIndex        =   34
         Top             =   7320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "previsto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   10245
         TabIndex        =   33
         Top             =   8280
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   4320
      TabIndex        =   15
      Top             =   9180
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5940
      TabIndex        =   16
      Top             =   9180
      Width           =   1575
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso del comprobante de egresos comunes                         #
'#  frmComprobanteEgresoComun V1.0                                              #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el comprobante de egresos comunes                     #
'#  Permite ingresar los datos de egresos comunes y sus detalles                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  COMP_EGRESO: Esta tabla almacena los datos del comprobante                  #
'#  PERSONA: donde se guardan los datos de los benficiarios de los comprobantes #
'#  DET_COMP_EGRESO: Guarda los detalles del comprobante de Egreso              #
'#  RET_COMP_EGRESO: Guarda las retenciones que puede tener el comprobante      #
'#  CTA_BANCO: consulta los datos del numero de cuenta y el último cheque       #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsCta1 As New clsConsulta
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private clsPag As New clsConsulta
Private clsPer As New clsConsulta
Private clsSql As New clsConsulta
Private clsEgr As New clsConsulta
Private clsCod As New clsConsulta
Private clsPgd As New clsConsulta
Private clsAsi As New clsConsulta
Private clsRet As New clsConsulta
Private clsNot As New clsConsulta
Private booCambiar As Boolean
Private strSQL As String
Private numComp As String
Private numAsi As String
Private Descripcion As String
Private strPersona As String

Private Sub dcmbNota_Click(Area As Integer)
    If dcmbNota.MatchedWithList = True Then
        clsNot.Filtrar "egr_codigo='" & dcmbNota.BoundText & "'"
        txtSaldo.Text = FormatoD2(clsNot.adorec_Def("sal"))
        txtSaldo.Tag = FormatoD2(clsNot.adorec_Def("egr_saldo"))
        dcmbNota.Tag = clsNot.adorec_Def("egr_numasiento")
        txtDocumento.Text = dcmbNota.Text
    Else
        txtDocumento.Text = ""
        txtSaldo.Text = "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsCta1 = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
    Set clsPag = Nothing
    Set clsPer = Nothing
    Set clsSql = Nothing
    Set clsEgr = Nothing
    Set clsCod = Nothing
    Set clsPgd = Nothing
    Set clsAsi = Nothing
    Set clsRet = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
    
    For i = 1 To (VSFG1.Rows - 1)
        VSFG1.TextMatrix(i, 0) = i
    Next i
End Sub

Private Sub saldodisponible()
    
    'Calcula el saldo disponible de la cuenta bancaria
    
    strSQL = " SELECT  sum(com_egr_ch_valor) as valor" & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND com_egr_ch_estado = 'GIRADO' AND cta_ban_numero = '" & dcmbCuenta.Text & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'  AND com_egr_ch_fecha <= CURRENT_TIMESTAMP " & _
             " GROUP BY cta_ban_numero "
    clsSql.Ejecutar strSQL
    
    If Not IsNull(clsSql.adorec_Def("valor")) And clsSql.adorec_Def.EOF = False Then
        Valor = clsSql.adorec_Def("valor")
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
    Else
        Valor = 0
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
    End If
End Sub

Private Sub CalcuTotal()

   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double

    'Calcula total debe
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")

    'Calcula total haber
    For i = 1 To VSFG.Rows - 1
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")
    TxtTotal.Text = FormatoD2(FormatoD2(txtTotalDebe.Text) - FormatoD2(txtTotalHaber.Text))
End Sub
Private Sub pagos()
    For i = 1 To VSFG1.Rows - 1
        Suma = Suma + Val(VSFG1.TextMatrix(i, 10))
    Next i
    txtValor = Format(Suma, "##0.00")
    LlenarVariableDescripcion
End Sub
Private Sub Limpiar()
    Dim strSQL As String
    VSFG1.Clear 1
    VSFG1.Rows = 2
    VSFG.Clear 1
    VSFG.Rows = 3
    'dcmbBeneficiario.Text = ""
    dcmbBanco.Text = ""
    txtDocumento = ""
    txtDescripcion = ""
    txtTotalHaber = 0
    txtTotalDebe = 0
    TxtTotal = 0
    txtSaldoReal = 0
    txtDisponible = 0
    txtPrevisto = 0
    txtp = 0
    txtD = 0

    txtValor = 0
End Sub


Private Sub cmdAceptar_Click()
    Dim ElAsiento As String
    Dim strTipoD As String
    Dim ffch As String
    Dim ffchdoc As String
    Dim EstadoCash As Integer
    'Comprueba que todos los datos esten ingresados
    ffch = Format(dtpFecha.Value, "yyyy-mm-dd")
    ffchdoc = Format(dtpFechaDoc.Value, "yyyy-mm-dd")
    If (IsDate(ffch) = False) Or (IsDate(ffchdoc) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Pagos"
        Exit Sub
    End If

    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd

    a = VSFG.Rows - 1
'    For i = 1 To a
'        For j = i + 1 To a
'            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) Then
'                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
'                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
'                VSFG.RemoveItem j
'                a = a - 1
'                j = j - 1
'            End If
'            If j >= a Then
'                Exit For
'            End If
'        Next j
'    Next i

    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber And optNCredito.Value = False Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "Pagos"
        Exit Sub
    ElseIf FormatoD2(txtSaldo.Text) < FormatoD2(txtValor.Text) And optNCredito.Value = True Then
        MsgBox "El Saldo de la Nota de Credito debe ser menor al abono de las Facturas", vbInformation, "Pagos"
        Exit Sub
    Else
          'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If optNCredito.Value = False And VSFG.TextMatrix(1, 1) = "" Or txtDescripcion = "" Or dcmbBeneficiario = "" Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "Pagos"
            Exit Sub
        Else
            maxegr = 0
            strTipoD = "EFECTIVO"
            Dim clsAsiento As New clsContable
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            If optNCredito.Value = False And optNDebito.Value = False Then
                If Me.optcheque.Value = True Or Me.optTransferenciaAuto.Value = True Then
                    clsAsiento.NuevoAsiento "E", ffch, 0, 0, FormatoD2(txtTotalDebe), "PAGO INCOMPLETO"
                Else
                    clsAsiento.NuevoAsiento "D", ffch, 0, 0, FormatoD2(txtTotalDebe), "PAGO INCOMPLETO"
                End If
                
            Else
                clsAsiento.NumAsiento = dcmbNota.Tag
                'clsAsiento.ModificarAsiento FormatoD2(txtTotalDebe), FormatoD2(txtTotalDebe), ffch, , , "PAGO INCOMPLETO"
                'clsAsiento.NuevoAsiento "E", ffch, 0, 0, Formatod2(txtTotalDebe), "PAGO INCOMPLETO", False
                'clsAsiento.EliminarAsiento False, True
            End If
            ElAsiento = "NULL"
            If dcmbBanco.Enabled = True And (optcheque.Value = True Or optTransferenciaAuto.Value = True) Then
                'Calcula el máximo código de comprobante de egreso
                If optTransferenciaAuto.Value = True Then
                    strTipoD = "TRANSFERENCIA DIRECTA"
                    EstadoCash = 1
                Else
                    strTipoD = "CHEQUE"
                    EstadoCash = 0
                End If
                Dim Nombre2 As String
                If chkNombre2.Value = 1 And optcheque.Value = True Then
                    Nombre2 = Trim(UCase(InputBox("Ingrese el nombre que saldrá en el cheque: ", "Nombre del cheque")))
                    Me.txtDescripcion = UCase(Me.txtDescripcion) & " GIRADO A: " & Nombre2
                End If
                
                If Trim(Nombre2) = "" Then
                    Nombre2 = "NULL"
                Else
                    Nombre2 = "'" & Nombre2 & "'"
                End If
                If numComp = "" And numAsi = "" Then
                    strSQL = " SELECT COALESCE(max(com_egr_codigo),0) as egr  " & _
                             " FROM comp_egreso " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " GROUP BY emp_codigo"
                    clsEgr.Ejecutar strSQL
                    maxegr = FormatoD0(clsEgr.adorec_Def("egr")) + 1 'valor del código del egreso comun +1
                Else
                    maxegr = numComp
                    strSQL = " DELETE FROM comp_egreso WHERE com_egr_codigo='" & maxegr & "' and emp_codigo='" & strEmpresa & "' "
                    clsSql.Ejecutar strSQL, "M"
                End If
                'Ingreso de datos en comp_egreso
                strSQL = " INSERT INTO comp_egreso (com_egr_codigo, emp_codigo,asi_numasiento, cta_ban_numero, ban_codigo, per_codigo, " & _
                         " com_egr_fecha, com_egr_descripcion, com_egr_ch_fecha,com_egr_ch_num, com_egr_ch_estado,com_egr_ch_valor,com_egr_conciliado, " & _
                         " com_egr_nombre2,com_egr_proceso_cash,com_egr_fechamod, com_egr_usumod) " & _
                         " VALUES ('" & maxegr & "','" & strEmpresa & "','" & clsAsiento.NumAsiento & "','" & dcmbCuenta.BoundText & "','" & dcmbBanco.BoundText & "','" & dcmbBeneficiario.BoundText & "', " & _
                         " '" & ffch & "','BENEFICIARIO: " & dcmbBeneficiario.Text & vbNewLine & "BANCO: " & dcmbBanco.Text & " CTA: " & dcmbCuenta.Text & " CH.No: " & txtDocumento.Text & vbNewLine & UCase(txtDescripcion) & "','" & ffchdoc & "', '" & FormatoD0(txtDocumento) & "','GIRADO','" & txtValor & "'," & _
                         " 0," & Nombre2 & ",'" & EstadoCash & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsSql.Ejecutar strSQL, "M"
                
            ElseIf dcmbBanco.Enabled = True And optTransferencia.Value = True Then
            'Calcula el máximo código de comprobante de egreso
                strTipoD = "TRANSFERENCIA BANCARIA"
                strSQL = " SELECT COALESCE(max(not_d_c_codigo),0) as num " & _
                         " FROM nota_d_c" & _
                         " WHERE emp_codigo = '" & strEmpresa & _
                         "' AND tip_not_d_c='D'" & _
                         " GROUP BY emp_codigo"
                clsEgr.Ejecutar strSQL
                maxegr = clsEgr.adorec_Def("num") + 1 'valor del código del egreso comun +1
                
                
                strSQL = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha,not_d_c_descripcion, not_d_c_monto,asi_numasiento,not_d_c_conciliado, not_d_c_fechamod, not_d_c_usumod) " & _
                         " VALUES ('D','" & maxegr & "', '" & dcmbCuenta.BoundText & "', '" & dcmbBanco.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtDocumento.Text & "','" & ffch & "','" & txtDescripcion.Text & "','" & txtValor.Text & "','" & clsAsiento.NumAsiento & "',0, CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsSql.Ejecutar strSQL, "M"
                
            ElseIf optNCredito.Value = True Then
                strTipoD = "NOTA DE CRÉDITO"
                ElAsiento = "'" & clsAsiento.NumAsiento & "'"
            ElseIf optNDebito.Value = True Then
                strTipoD = "NOTA DE DÉBITO"
                ElAsiento = "'" & clsAsiento.NumAsiento & "'"
            Else
                strTipoD = "EFECTIVO"
            End If
            If optNCredito.Value = False And optNDebito.Value = False Then
                With VSFG
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                            clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                        End If
                    Next i
                End With
            End If
            If optNCredito.Value = True Or optNDebito.Value = True Then
                Dim clsInventa As New clsInventario
                clsInventa.Inicializar AdoConn, AdoConnMaster
                If optNCredito.Value = True Then
                    clsInventa.strTipo = "DPV"
                    clsInventa.strIE = "E"
                    clsInventa.strDoc = dcmbNota.BoundText
                    clsInventa.ModificaEgr , , , , , , , , , , , , , , , FormatoD2(txtSaldo.Tag) + FormatoD2(txtValor.Text)
                Else
                    clsInventa.strTipo = "DCL"
                    clsInventa.strIE = "I"
                    clsInventa.strDoc = dcmbNota.BoundText
                    clsInventa.ModificaIng , , , , , , , , , , , , , , , FormatoD2(txtSaldo.Tag) + FormatoD2(txtValor.Text)
                End If
                Set clsInventa = Nothing
            End If
            'Ingreso de datos en la tabla pago
            n = VSFG1.Rows - 1
            ElAsiento = clsAsiento.NumAsiento
            For i = 1 To n
                k = VSFG1.TextMatrix(i, 10)
                If (VSFG1.TextMatrix(i, 10) <> "" Or VSFG1.TextMatrix(i, 10) <> "0") And VSFG1.TextMatrix(i, 1) <> "0" Then
                    'Calcula el máximo codigo de pago para la cuenta
                     strSQL = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                              " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                              "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                              "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                              " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "'  AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'P'" & _
                              " GROUP BY pago.emp_codigo"
                    clsCod.Ejecutar strSQL
                    If clsCod.adorec_Def.EOF Then
                        maxpag = 1
                    Else
                        maxpag = clsCod.adorec_Def("pag") + 1
                    End If
                    Dim ValorPago As Double
                    If optNDebito.Value = True Then
                        ValorPago = FormatoD2(VSFG1.TextMatrix(i, 10)) * -1
                    Else
                        ValorPago = FormatoD2(VSFG1.TextMatrix(i, 10))
                    End If
                    
                    strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & Val(VSFG1.TextMatrix(i, 2)) & "', 'P', '" & Val(maxpag) & "', '" & ffch & "', '" & ValorPago & "', 'NC:" & txtDocumento & "', '" & txtDescripcion & "', " & _
                             " '" & maxegr & "','" & ElAsiento & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsPag.Ejecutar strSQL, "M"
                End If
                If FormatoD2(VSFG1.TextMatrix(i, 9)) <= FormatoD2(VSFG1.TextMatrix(i, 10)) And optNDebito.Value = False Then
                    strSQL = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_fechapago='" & ffch & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_tipo = 'P' AND emp_codigo = '" & strEmpresa & "' "
                    clsPgd.Ejecutar strSQL, "M"
                End If

            Next i
            If optNCredito.Value = False And optNDebito.Value = False Then
                clsAsiento.ModificarAsiento FormatoD2(txtTotalDebe), FormatoD2(txtTotalHaber), , , , "BENEFICIARIO: " & dcmbBeneficiario.Text & vbNewLine & "BANCO: " & dcmbBanco.Text & " CTA: " & dcmbCuenta.Text & " " & strTipoD & ".No: " & txtDocumento.Text & vbNewLine & txtDescripcion
            End If

            MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"


            'Actualiza los valores de los saldos
            If dcmbBanco.Enabled = True Then
                Dim strChUlt As String
                If booCambiar = True Then
                    strChUlt = " '" & FormatoD0(txtDocumento.Text) & "'"
                Else
                    strChUlt = " cta_ban_ch_ultimo"
                End If
                strSQL = " UPDATE cta_banco " & _
                         " SET cta_ban_ch_ultimo = " & strChUlt & ", cta_ban_saldoreal= '" & txtSaldoReal & "',cta_ban_saldoprevisto= '" & txtPrevisto & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                         " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
                clsSql.Ejecutar strSQL, "M"
                If optcheque.Value = True Then
                    Dim CompEgr As New frmReporte
                    CompEgr.strReporte = "rptComprobanteEgreso"
                    CompEgr.strNumero = maxegr
                    CompEgr.Show
                    Dim Cheque As New frmReporte
                    Cheque.strReporte = "rptCheque"
                    Cheque.strNumero = maxegr
                    Cheque.Show
                Else
                    Dim Asi As New frmReporte
                    Asi.strAsiento = clsAsiento.NumAsiento
                    Asi.strReporte = "rptAsiento"
                    Asi.Show
                End If
            Else
                Dim Asi2 As New frmReporte
                Asi2.strAsiento = clsAsiento.NumAsiento
                Asi2.strReporte = "rptAsiento"
                Asi2.Show
            End If
            Set clsAsiento = Nothing
        End If
    End If
    Unload Me
    'frmPagos.Show
End Sub

Private Sub LlenarVariableDescripcion()
    Dim Coma As String
    Dim TextoInicio As String
    Dim NumItems As Integer
    Dim obs As String
    NumItems = 0
    Descripcion = ""
    obs = ""
    TextoInicio = "FACTURA: "
    For i = 1 To VSFG1.Rows - 1
        If Val(VSFG1.TextMatrix(i, 1)) = -1 Then
            If VSFG1.TextMatrix(i, 2) <> VSFG1.TextMatrix(i - 1, 2) Then
                NumItems = NumItems + 1
                If (NumItems > 1) Then
                    Coma = ", "
                    TextoInicio = ""
                End If
                obs = " (" & VSFG1.TextMatrix(i, 5) & ")"
                Descripcion = TextoInicio & Descripcion & Coma & VSFG1.TextMatrix(i, 4) & obs
            End If
        End If
    Next i
    PonerDescripcion2
End Sub

Private Sub PonerDescripcion2()
    Dim Cadena1 As String
    Dim Cadena2 As String
    Dim Cadena3 As String
    Dim Cadena4 As String
    Dim Cadena5 As String
    
    If optcheque.Value = True Then
        Cadena1 = "CHEQUE "
        If txtDocumento <> "" Then
            Cadena2 = txtDocumento & " - "
        End If
        If dcmbBanco <> "" Then
            Cadena3 = dcmbBanco & " "
        End If
    ElseIf optefectivo.Value = True Then
        Cadena1 = "EFECTIVO "
    ElseIf optNCredito.Value = True Then
        Cadena1 = "NOTA DE CRÉDITO " & dcmbNota.Text & " (" & dcmbNota.BoundText & ") "
    ElseIf optNDebito.Value = True Then
        Cadena1 = "NOTA DE DÉBITO " & dcmbNota.Text & " (" & dcmbNota.BoundText & ") "
    ElseIf optTransferenciaAuto.Value = True Then
        Cadena1 = "TRANSFERENCIA DIRECTA "
        If dcmbBanco <> "" Then
            Cadena3 = dcmbBanco & " "
        End If
    ElseIf optTransferencia.Value = True Then
        Cadena1 = "TRANSFERENCIA BANCARIA "
'        If txtDocumento <> "" Then
'            Cadena2 = txtDocumento
'        End If
        If dcmbBanco <> "" Then
            Cadena3 = dcmbBanco & " "
        End If
    ElseIf optOtros.Value = True Then
        Cadena1 = "OTRA FORMA DE PAGO "
    End If
    If Descripcion <> "" Then
        Cadena4 = Descripcion & " - "
    End If
    If dcmbBeneficiario <> "" Then
        Cadena5 = dcmbBeneficiario & " - "
    End If
    txtDescripcion = Cadena5 & Cadena4 & Cadena1 & Cadena2 & Cadena3
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBanco_Change()
    dcmbCuenta = ""
    VSFG.TextMatrix(1, 4) = txtValor
    dcmbBanco.Tag = dcmbBanco.BoundText

    strSQL = " SELECT cta_ban_numero, cta_ban_ch_ultimo as ban, cta_ban_ctaconta,cta_ban_saldoreal, cta_ban_saldoprevisto" & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSQL
    If clsCtb.adorec_Def.EOF = False Then
        Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
        dcmbCuenta.ListField = ("cta_ban_numero")
    Else
        dcmbCuenta = ""
    End If
    LlenarVariableDescripcion
End Sub

'Private Sub dcmbBeneficiario_Change()
Private Sub dcmbBeneficiario_Validate(Cancel As Boolean)
    Dim CtaBlanco As String
    optefectivo.Value = True
    cmdAceptar.Enabled = True
    VSFG1.Enabled = True
    Limpiar
    VSFGRet.Clear 1
    
    strPersona = "'" & dcmbBeneficiario.BoundText & "'"
    strSQL = " SELECT * " & _
             " FROM persona_cuenta " & _
             " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        optTransferenciaAuto.Enabled = True
    Else
        optTransferenciaAuto.Enabled = False
    End If
    
    
    strSQL = " SELECT per_codigo_rel " & _
             " FROM persona_relacion " & _
             " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        While Not clsSql.adorec_Def.EOF
            strPersona = strPersona & ",'" & clsSql.adorec_Def("per_codigo_rel") & "'"
            clsSql.adorec_Def.MoveNext
        Wend
    End If
    
    strSQL = " SELECT cat_p_ctaconta " & _
                 " FROM persona INNER JOIN categoria_p " & _
                 " ON persona.emp_codigo=categoria_p.emp_codigo " & _
                 " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                 " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                 " WHERE persona.emp_codigo = '" & strEmpresa & "'" & _
                 " AND persona.per_codigo = '" & dcmbBeneficiario.BoundText & "'"
     clsCta.Ejecutar strSQL
    CtaBlanco = clsCta.adorec_Def("cat_p_ctaconta")
    strSQL = " SELECT ret_nombre,' ',ret_ctaconta,ret_codigo" & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY ret_ctaconta"
     clsCta.Ejecutar strSQL
     Set VSFGRet.DataSource = clsCta.adorec_Def
 'Consulta para el grid sobre las cuentas por pagar del beneficiario seleccionado
    strSQL = " SELECT 0 as sel, cuenta_p_c.cue_p_c_codigo, CONCAT(cue_p_c_fra_cuenta, '/' , cue_p_c_tot_cuenta ) as cue_p_c_fra_cuenta, " & _
             " CONCAT(cue_p_c_serie,FORMAT(cue_p_c_numero,'0000000')) as cue_p_c_egr_codigo, cue_p_c_descripcion, " & _
             " cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor, cue_p_c_valor-COALESCE(sum(pag_monto),0)-COALESCE(comprobante_retencion.com_ret_total,0) as saldo,'' as pag," & _
             " COALESCE(comprobante_retencion.com_ret_total,0) as ret,iif(COALESCE(sum(pag_monto),0)>0,'1','0') as mon," & _
             " COALESCE(cta_codigo,'" & CtaBlanco & "') as cta_codigo " & _
             " FROM  (cuenta_p_c LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo)" & _
             " LEFT JOIN det_asiento ON cuenta_p_c.emp_codigo=det_asiento.emp_codigo " & _
             " AND cuenta_p_c.asi_numasiento=det_asiento.asi_numasiento " & _
             " AND det_asiento.det_asi_haber!=0" & _
             " AND ROUND(cuenta_p_c.cue_p_c_valor-COALESCE(com_ret_total,0),2)=ROUND(COALESCE(det_asiento.det_asi_haber,cuenta_p_c.cue_p_c_valor),2) " & _
             " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo" & _
             " WHERE per_codigo IN (" & strPersona & ") " & _
             " AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'P' " & _
             " AND cue_p_c_pagado='0'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cue_p_c_fra_cuenta,cue_p_c_tot_cuenta,cue_p_c_serie,cue_p_c_numero, cue_p_c_descripcion,cue_p_c_fechaemision,cue_p_c_fechapropuesta,cue_p_c_valor,comprobante_retencion.com_ret_total,cta_codigo" & _
             " ORDER BY CONCAT(cue_p_c_serie,left(cue_p_c_numero,'0000000'))"
    
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.EOF = False Then
        VSFG1.Rows = 1
        Set VSFG1.DataSource = clsSql.adorec_Def.DataSource
        'ponerBotones
    Else
        VSFG1.Clear 1
        VSFG1.Rows = 2
    
    End If
    
  'Consulta el saldo de la cuenta
  n = VSFG1.Rows - 1
  Call CxC
End Sub

Private Sub CxC()
    Dim Per_Ruc As String
    
    strPersona = "'" & dcmbBeneficiario.BoundText & "'"
    strSQL = " SELECT per_ruc " & _
             " FROM persona " & _
             " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        Per_Ruc = clsSql.adorec_Def("per_ruc")
    End If
    
    If Len(Per_Ruc) > 0 Then
        'Consulta para el grid sobre las cuentas por cobrar del beneficiario seleccionado

strSQL = "SELECT ' ' as a,'0' as b, cuenta_p_c.cue_p_c_codigo, CONCAT(cue_p_c_fra_cuenta, '/' , cue_p_c_tot_cuenta ) as cue_p_c_fra_cuenta, cue_p_c_egr_codigo,"
strSQL = strSQL + " cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,DATEDIFF(DAY, cue_p_c_fechapropuesta,CURRENT_TIMESTAMP) AS dven,"
strSQL = strSQL + " cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d, ' ' as e,iif(com_ret_total IS NULL,' ','Tot.Ret.'),"
strSQL = strSQL + " iif(com_ret_total IS NULL,'0',com_ret_total),' ',' ',iif(comprobante_retencion.com_ret_total IS NULL,'1','0') as f"
strSQL = strSQL + " From"
strSQL = strSQL + " persona P inner join"
strSQL = strSQL + " cuenta_p_c on"
strSQL = strSQL + " p.emp_codigo = cuenta_p_c.emp_codigo and"
strSQL = strSQL + " p.per_codigo = cuenta_p_c.per_codigo LEFT JOIN"
strSQL = strSQL + " pago ON"
strSQL = strSQL + " cuenta_p_c.emp_codigo=pago.emp_codigo AND"
strSQL = strSQL + " cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND"
strSQL = strSQL + " cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo  LEFT JOIN"
strSQL = strSQL + " comprobante_retencion ON"
strSQL = strSQL + " cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND"
strSQL = strSQL + " cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND"
strSQL = strSQL + " cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo"
strSQL = strSQL + " where"
strSQL = strSQL + " per_ruc ='" & Per_Ruc & "' AND"
strSQL = strSQL + " cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND"
strSQL = strSQL + " cuenta_p_c.cue_p_c_tipo = 'C' AND"
strSQL = strSQL + " cue_p_c_pagado='0'"
strSQL = strSQL + " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,pag_monto,cue_p_c_fra_cuenta,cue_p_c_tot_cuenta,cue_p_c_egr_codigo,cue_p_c_descripcion,cue_p_c_fechaemision,cue_p_c_fechapropuesta,cue_p_c_valor,com_ret_total"
strSQL = strSQL + " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)!=0"
strSQL = strSQL + " ORDER BY cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo"
 
 
 
 
        
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Valor = clsSql.adorec_Def("cue_p_c_valor")
            Set vsfgCXC.DataSource = clsSql.adorec_Def.DataSource
             vsfgCXC.ColDataType(1) = flexDTBoolean
        End If
    End If

    vsfgCXC.SubtotalPosition = flexSTBelow
    vsfgCXC.Subtotal flexSTClear         ' remove old values
    vsfgCXC.Subtotal flexSTSum, -1, 9, , , vbRed, True
    vsfgCXC.Subtotal flexSTSum, -1, 10, , , vbRed, True
    vsfgCXC.Subtotal flexSTSum, -1, 11, , , vbRed, True
    
End Sub
Private Sub Form_Activate()
 Dim strComparar As String

    strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'D'" & _
             " ORDER BY tip_not_codigo"
    clsBan.Ejecutar strSQL
    
    Set dcmbTipo.RowSource = clsBan.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
    dcmbTipo.Text = clsBan.adorec_Def("tip_not_nombre")

'     consulta para saber los  bancos existentes
    strSQL = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL

    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        dcmbBanco = ""
    End If
     
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
    'Inicializa las clases para hacer distintas consultas
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsCta1.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsCod.Inicializar AdoConn, AdoConnMaster
    clsPgd.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsRet.Inicializar AdoConn, AdoConnMaster
    txtr = 0
    txtD = 0
    txtp = 0
    
    If dcmbBeneficiario.Text = "" Then
        cmdAceptar.Enabled = False
        VSFG1.Enabled = False
    End If
    strSQL = " SELECT ret_nombre,' ',ret_ctaconta,ret_codigo" & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY ret_ctaconta"
     clsCta.Ejecutar strSQL
     Set VSFGRet.DataSource = clsCta.adorec_Def
     strSQL = " SELECT ret_ctaconta,cta_nombre,0,0" & _
                 " FROM retencion INNER JOIN ctaconta ON retencion.emp_codigo=ctaconta.emp_codigo AND retencion.ret_ctaconta=ctaconta.cta_codigo " & _
                 " WHERE retencion.emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY ret_ctaconta"
     clsCta.Ejecutar strSQL
     'Set VSFG.DataSource = clsCta.adorec_Def
    strSQL = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsCta.Ejecutar strSQL

     VSFG.ColComboList(5) = VSFG.BuildComboList(clsCta.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
     'Set VSFG.DataSource = clsCta.adorec_Def
    strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSQL

     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")

    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaDoc.Value = Format(HoyDia, "yyyy-mm-dd")

    optproveedor_Click
    
End Sub

Private Sub optcheque_Click()
    txtDocumento = ""
    dcmbNota.Enabled = False
    dcmbBanco.Enabled = True
    dcmbCuenta.Enabled = True
    txtDocumento.Enabled = False
    dcmbTipo.Enabled = False
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(1, 4) = txtValor
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub


Private Sub optTransferenciaAuto_Click()
    txtDocumento = ""
    dcmbNota.Enabled = False
    dcmbBanco.Enabled = True
    dcmbCuenta.Enabled = True
    txtDocumento.Enabled = True
    dcmbTipo.Enabled = False
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(1, 4) = txtValor
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub

Private Sub OptCliente_Click()
  
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiario.Text = ""
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub optefectivo_Click()
    dcmbBanco.Enabled = False
    dcmbBanco = ""
    dcmbCuenta.Enabled = False
    dcmbCuenta = ""
    dcmbNota.Enabled = False
    txtDocumento.Enabled = True
    txtDocumento = ""
    dcmbTipo.Enabled = False
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(1, 4) = txtValor
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub

Private Sub optNCredito_Click()
    dcmbBanco.Enabled = False
    dcmbBanco = ""
    dcmbCuenta.Enabled = False
    dcmbCuenta = ""
    txtDocumento.Enabled = False
    txtDocumento = ""
    dcmbTipo.Enabled = False
    dcmbNota.Enabled = True
    clsNot.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT egr_codigo,CONCAT(egr_serie,'-',egr_numero) as num,egr_saldo,egr_numasiento,egr_total-egr_saldo as sal " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='DPV' " & _
             " AND per_codigo IN (" & strPersona & ") " & _
             " AND egr_anulado=0 " & _
             " AND egr_total-egr_saldo!=0 " & _
             " ORDER BY egr_codigo"
    clsNot.Ejecutar strSQL
    dcmbNota.ListField = "num"
    dcmbNota.BoundColumn = "egr_codigo"
    Set dcmbNota.RowSource = clsNot.adorec_Def.DataSource
    LlenarVariableDescripcion
End Sub

Private Sub optNDebito_Click()
    dcmbBanco.Enabled = False
    dcmbBanco = ""
    dcmbCuenta.Enabled = False
    dcmbCuenta = ""
    txtDocumento.Enabled = True
    txtDocumento = ""
    dcmbTipo.Enabled = False
    VSFG.TextMatrix(1, 3) = txtValor
    VSFG.TextMatrix(1, 4) = 0
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub

Private Sub optOtros_Click()
    dcmbBanco.Enabled = False
    dcmbBanco = ""
    dcmbCuenta.Enabled = False
    dcmbCuenta = ""
    dcmbNota.Enabled = False
    txtDocumento.Enabled = True
    txtDocumento = ""
    dcmbTipo.Enabled = False
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(1, 4) = txtValor
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub

Private Sub optTransferencia_Click()
    txtDocumento = ""
    txtDocumento.Enabled = True
    dcmbBanco.Enabled = True
    dcmbCuenta.Enabled = True
    dcmbTipo.Enabled = True
    dcmbNota.Enabled = False
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(1, 4) = txtValor
    CalcuTotal
    LlenarVariableDescripcion
    CargarCuentaAsiento
End Sub

Private Sub optproveedor_Click()
    p = 1
    Frame1.Caption = "Proveedor"
    dcmbBeneficiario.Text = ""
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub txtDocumento_Change()
    LlenarVariableDescripcion
End Sub

Private Sub txtTotal_Change()
    TxtTotal = FormatoD2(TxtTotal)
End Sub

Private Sub txtTotalDebe_Change()
    txtTotalDebe = FormatoD2(txtTotalDebe)
End Sub

Private Sub txtTotalHaber_Change()
 txtTotalHaber = FormatoD2(txtTotalHaber)
End Sub


Private Sub txtValor_Change()
    CargarCuentaAsiento
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

    ' only interesetd in left button
    If Button <> 1 Then Exit Sub

    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub

    If (c <> 0 Or r = (VSFG.Rows)) Then Exit Sub

    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        If r > 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pagos"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)

        'Recorro el FlexGrid para poner números a las filas

        If respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            PonerBotones
            CalcuTotal
        Else
            VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub dcmbCuenta_Change()
    
    If Trim(dcmbCuenta) <> "" And (optcheque.Value = True Or optTransferenciaAuto.Value = True Or optTransferencia.Value = True) Then
        
        strSQL = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre,cta_ban_ch_ultimo as cheque,cta_ban_saldoreal,cta_ban_saldoprevisto " & _
                 " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
        clsctc.Ejecutar strSQL
        If Not clsctc.adorec_Def.EOF Then
            If optcheque.Value = True Then
                If IsNull(clsctc.adorec_Def("cheque")) Then
                    txtDocumento.Enabled = False
                    txtDocumento.Text = 1
                Else
                    txtDocumento.Enabled = False
                    Cheque = clsctc.adorec_Def("cheque") + 1
                    txtDocumento.Text = Cheque
                End If
                Dim strNCH As String
                Dim booPasar As Boolean
                booPasar = False
                strNCH = txtDocumento.Text
                numComp = ""
                numAsi = ""
                While booPasar = False
                    txtDocumento.Text = InputBox("No. de cheque", "Comprobante de Egreso", strNCH)
                    strSQL = " SELECT count(*) as Num FROM comp_egreso " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ban_codigo='" & dcmbBanco.BoundText & "'" & _
                             " AND cta_ban_numero = '" & dcmbCuenta & "' AND com_egr_proceso_cash=0" & _
                             " AND com_egr_ch_num='" & txtDocumento.Text & "'"
                    clsSql.Ejecutar strSQL
                    If clsSql.adorec_Def("Num") <> 0 Then
                        strSQL = " SELECT com_egr_codigo,asi_numasiento,com_egr_ch_estado " & _
                                 " FROM comp_egreso " & _
                                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                                 " AND ban_codigo='" & dcmbBanco.BoundText & "'" & _
                                 " AND cta_ban_numero = '" & dcmbCuenta & "'" & _
                                 " AND com_egr_ch_num='" & txtDocumento.Text & "' AND com_egr_proceso_cash=0"
                        clsSql.Ejecutar strSQL
                        If clsSql.adorec_Def("com_egr_ch_estado") = "ANULADO" Then
                            If MsgBox("El cheque tiene estado Anulado." & vbNewLine & "Desea reutilizar el cheque y el compobante?", vbYesNo + vbQuestion, "Comprobante de Egreso") = vbYes Then
                                numComp = clsSql.adorec_Def("com_egr_codigo")
                                numAsi = clsSql.adorec_Def("asi_numasiento")
                                booPasar = True
                            End If
                        Else
                            MsgBox "Ese cheque ya ha sido emitido", vbCritical, "Comprobante de Egreso"
                            txtDocumento.Text = strNCH
                            numComp = ""
                            numAsi = ""
                        End If
                    Else
                        booPasar = True
                    End If
                Wend
                If Format(txtDocumento.Text, "0000000000") >= Format(strNCH, "0000000000") Then
                    booCambiar = True
                Else
                    booCambiar = False
                End If
            Else
                txtDocumento.Enabled = True
            End If
            txtSaldoReal = clsctc.adorec_Def("cta_ban_saldoreal")
            txtPrevisto = clsctc.adorec_Def("cta_ban_saldoprevisto")
            txtp = clsctc.adorec_Def("cta_ban_saldoprevisto")
            saldodisponible
            If clsctc.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(1, 1) = clsctc.adorec_Def("cta_ban_ctaconta")
            End If
            VSFG.TextMatrix(1, 4) = txtValor
        End If
    Else
        txtSaldoReal = 0
        txtPrevisto = 0
        txtDisponible = 0
        txtp = 0
        txtD = 0
        txtDocumento = ""
        'VSFG.Clear 1
        'VSFG.Rows = 2
        'VSFG.Rows = VSFGRet.Rows + 1
        VSFG.TextMatrix(1, 1) = ""
        VSFG.TextMatrix(1, 2) = ""
    End If
    LlenarVariableDescripcion
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <> OldRow Then
        If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
            If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes

    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
                If Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "4" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "5" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "6" Then
                    VSFG.AddItem ""
                ElseIf VSFG.TextMatrix(VSFG.Row, 5) <> "" Then
                    VSFG.AddItem ""
                End If
            Else
                VSFG.AddItem ""
            End If
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Row <= 1 Then
            If dcmbCuenta.Enabled = True Then
                If Col = 1 Then
                    Cancel = True
                End If
                If Col = 2 Then
                    Cancel = True
                End If
            Else
                If Row < 1 Then
                    If Col = 1 Then
                        Cancel = True
                    End If
                    If Col = 2 Then
                        Cancel = True
                    End If
                End If
            End If
            If Col = 3 Then
               Cancel = True
            End If
            If Col = 4 Then
                Cancel = True
            End If
        End If
    
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    'Verifica que se ingrese la cuenta contable en el grid
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
        MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
        VSFG.TextMatrix(Row, 3) = ""
        VSFG.TextMatrix(Row, 4) = ""
    ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe
        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = intDato
        End If
        'Verifica que solo se ingresen números tanto en el Debe como en el Haber
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
        CalcuTotal
    End If
End Sub


Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    d = Format(HoyDia, "yyyy-MM-dd")
    dia = Mid(d, 9, 2)
    Mes = Mid(d, 6, 2)
    Año = Mid(d, 1, 4)
    ffch = Format(dtpFecha.Value, "yyyy-mm-dd")
    m = Mid(ffch, 6, 2)

    If Val(Format(dtpFecha.Value, "dd")) > dia Or m > Mes Or Val(Format(dtpFecha.Value, "yyyy")) > Año Then
            txtDisponible.Text = txtD
    Else
            txtDisponible.Text = Val(txtD) - Val(VSFG.TextMatrix(1, 4))
    End If
    txtPrevisto.Text = txtp - Val(VSFG.TextMatrix(1, 4))
If Row > 0 Then
    
' Asigna codigos de cuenta y nombres en el grid
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                     .TextMatrix(Row, 2) = .TextMatrix(Row, 1)
             End If

             If Col = 2 Then
                     .TextMatrix(Row, 1) = .TextMatrix(Row, 2)
             End If
         End If
    End With
End If
CalcuTotal
End Sub

Private Sub VSFG_Validate(Cancel As Boolean)
    If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
        If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 10 Then
        'Verifica que solo se ingresen números en el campo Debe
        If Not IsNumeric(VSFG1.TextMatrix(Row, 10)) And VSFG1.TextMatrix(Row, 10) <> "" Then
            MsgBox "Ingrese solo números en el Valor de Pago.", vbInformation, "Pagos"
            VSFG1.TextMatrix(Row, 10) = 0
        End If
    End If
    
    If Val(VSFG1.TextMatrix(Row, 10)) > Val(VSFG1.TextMatrix(Row, 9)) And optNDebito.Value = False Then
        MsgBox "El valor a pagar es mayor al Saldo", vbCritical, "Pagos"
        VSFG1.Select Row, 10
        VSFG1.TextMatrix(Row, 10) = 0
    End If
    pagos
    
    Dim strSQL As String
    Dim i As Long
    If Col = 1 And Row > 0 Then
        If VSFG1.TextMatrix(Row, 1) = "-1" Then
            VSFG1.Select Row, 1, Row, 11
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HC0FFFF
            VSFG1.Select Row, 10
            If VSFG1.TextMatrix(Row, 2) <> "" And Val(VSFG1.TextMatrix(Row, 12)) = 0 Then
                strSQL = " SELECT ret_codigo, det_com_ret_valor,det_com_ret_porcentaje " & _
                         " FROM det_comp_ret " & _
                         " WHERE emp_codigo='" & strEmpresa & _
                         "' AND cue_p_c_codigo=" & VSFG1.TextMatrix(Row, 2) & _
                         " AND cue_p_c_tipo='P'"
                clsRet.Ejecutar strSQL
                If clsRet.adorec_Def.RecordCount > 0 Then
                    While clsRet.adorec_Def.EOF = False
                        For i = 1 To VSFGRet.Rows - 1
                            If VSFGRet.TextMatrix(i, 4) = clsRet.adorec_Def("ret_codigo") Then
                                VSFGRet.TextMatrix(i, 2) = Val(VSFGRet.TextMatrix(i, 2)) + Val(clsRet.adorec_Def("det_com_ret_valor")) * Val(clsRet.adorec_Def("det_com_ret_porcentaje")) / 100#
                                i = VSFGRet.Rows
                            End If
                        Next i
                        clsRet.adorec_Def.MoveNext
                    Wend
                End If
            End If
            CargarCuentaAsiento
        ElseIf VSFG1.TextMatrix(Row, 1) = "0" Then
            VSFG1.Select Row, 1, Row, 11
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HFFFFFF
            VSFG1.Select Row, 10
            VSFG1.TextMatrix(Row, 10) = "0"
            If VSFG1.TextMatrix(Row, 2) <> "" And Val(VSFG1.TextMatrix(Row, 12)) = 0 Then
                strSQL = " SELECT ret_codigo, det_com_ret_valor*det_com_ret_porcentaje/100.00 as valor " & _
                         " FROM det_comp_ret " & _
                         " WHERE emp_codigo='" & strEmpresa & _
                         "' AND cue_p_c_codigo=" & VSFG1.TextMatrix(Row, 2) & _
                         " AND cue_p_c_tipo='P'"
                clsRet.Ejecutar strSQL
                If clsRet.adorec_Def.RecordCount > 0 Then
                    While clsRet.adorec_Def.EOF = False
                        For i = 1 To VSFGRet.Rows - 1
                            If VSFGRet.TextMatrix(i, 4) = clsRet.adorec_Def("ret_codigo") Then
                                VSFGRet.TextMatrix(i, 2) = Val(VSFGRet.TextMatrix(i, 2)) - Val(clsRet.adorec_Def("valor"))
                                i = VSFGRet.Rows
                            End If
                        Next i
                        clsRet.adorec_Def.MoveNext
                    Wend
                End If
            End If
            CargarCuentaAsiento
        End If
    End If
    
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFG1.TextMatrix(Row, 1) = "0" Or VSFG1.TextMatrix(Row, 1) = "" Then
        If Col = 10 Then
            Cancel = True
        End If
               
    ElseIf VSFG1.TextMatrix(Row, 1) = "-1" Then
        If Col = 10 Then
            Cancel = False
        End If
    End If
    
  
End Sub

Private Sub VSFG1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 2 Or NewCol = 3 Or NewCol = 4 Or NewCol = 5 Or NewCol = 6 Or NewCol = 7 Or NewCol = 8 Or NewCol = 9 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub CargarCuentaAsiento()
    Dim boo As Boolean
    Dim i As Long
    Dim j As Long
    VSFG.Rows = 3
    If VSFG.Rows > 1 Then
        If optNDebito.Value = True Then
            VSFG.TextMatrix(1, 3) = txtValor
            VSFG.TextMatrix(1, 4) = 0
            For j = 1 To VSFG1.Rows - 1
                If Abs(FormatoD0(VSFG1.TextMatrix(j, 1))) = 1 Then
                    boo = False
                    For i = 1 To VSFG.Rows - 1
                        If VSFG.TextMatrix(i, 1) = VSFG1.TextMatrix(j, 13) Then
                            VSFG.TextMatrix(i, 4) = FormatoD2(VSFG.TextMatrix(i, 4)) + FormatoD2(VSFG1.TextMatrix(j, 10))
                            boo = True
                            Exit For
                        End If
                    Next i
                    If boo = False Then
                        VSFG.AddItem ""
                        VSFG.TextMatrix(VSFG.Rows - 1, 1) = VSFG1.TextMatrix(j, 13)
                        VSFG.TextMatrix(VSFG.Rows - 1, 4) = VSFG1.TextMatrix(j, 10)
                    End If
                End If
            Next j
        Else
            VSFG.TextMatrix(1, 3) = 0
            VSFG.TextMatrix(1, 4) = txtValor
            For j = 1 To VSFG1.Rows - 1
                If Abs(FormatoD0(VSFG1.TextMatrix(j, 1))) = 1 Then
                    boo = False
                    For i = 1 To VSFG.Rows - 1
                        If VSFG.TextMatrix(i, 1) = VSFG1.TextMatrix(j, 13) Then
                            VSFG.TextMatrix(i, 3) = FormatoD2(VSFG.TextMatrix(i, 3)) + FormatoD2(VSFG1.TextMatrix(j, 10))
                            boo = True
                            Exit For
                        End If
                    Next i
                    If boo = False Then
                        VSFG.AddItem ""
                        VSFG.TextMatrix(VSFG.Rows - 1, 1) = VSFG1.TextMatrix(j, 13)
                        VSFG.TextMatrix(VSFG.Rows - 1, 3) = VSFG1.TextMatrix(j, 10)
                    End If
                End If
            Next j
        End If
    End If
End Sub
Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (VSFG.TextMatrix(VSFG.Row, 3) = "") Then
                VSFG.TextMatrix(VSFG.Row, 3) = 0
     ElseIf VSFG.TextMatrix(VSFG.Row, 4) = "" Then
                VSFG.TextMatrix(VSFG.Row, 4) = 0
     End If
End Sub
