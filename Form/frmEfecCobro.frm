VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEfecCobro 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efectivización de Cobros"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "frmEfecCobro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11535
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11295
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   360
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExportarFC_ContacPoint 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Contabilización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   28
         Top             =   420
         Width           =   3375
         Begin NEED2.dtpFecha dtpFechaConta 
            Height          =   285
            Left            =   840
            TabIndex        =   37
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
         End
         Begin VB.OptionButton optBanco 
            BackColor       =   &H00DDDDDD&
            Caption         =   "En Banco o Cajas"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optPosfechado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "En Cheque PostFechado"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label2 
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
            Left            =   240
            TabIndex        =   31
            Top             =   1260
            Width           =   495
         End
      End
      Begin VB.Frame frmBanco 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Para depositar en :"
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
         Left            =   3840
         TabIndex        =   21
         Top             =   420
         Width           =   3735
         Begin VB.TextBox txtNumero 
            Height          =   285
            Left            =   1530
            TabIndex        =   22
            Top             =   1080
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1530
            TabIndex        =   23
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   1530
            TabIndex        =   24
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            TabIndex        =   27
            Top             =   405
            Width           =   510
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº. Documento:"
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
            Top             =   1110
            Width           =   1125
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
            TabIndex        =   25
            Top             =   765
            Width           =   1245
         End
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4800
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Fecha de Consulta"
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
         Height          =   2055
         Left            =   7800
         TabIndex        =   16
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox chkfechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Todas las Fechas"
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
            Left            =   195
            TabIndex        =   0
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton cmdConsultar 
            Caption         =   "Consultar"
            Height          =   375
            Left            =   2040
            TabIndex        =   17
            Top             =   1560
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
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
            Left            =   930
            TabIndex        =   33
            Top             =   840
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   2621443
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
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
            Left            =   930
            TabIndex        =   34
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   2621443
            CurrentDate     =   37463
         End
         Begin MSDataListLib.DataCombo dcmbTipoDoc 
            Height          =   315
            Left            =   930
            TabIndex        =   35
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Docs:"
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
            TabIndex        =   36
            Top             =   405
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            TabIndex        =   19
            Top             =   1230
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            TabIndex        =   18
            Top             =   870
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos para el ingreso de Asiento"
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
         Left            =   720
         TabIndex        =   12
         Top             =   5160
         Width           =   9975
         Begin VB.TextBox txtTotalDebe 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txtTotalHaber 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtDescripciont 
            Enabled         =   0   'False
            Height          =   525
            Left            =   5040
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtnota 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFG 
            Height          =   1095
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   9480
            _cx             =   16722
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
            Rows            =   1
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmEfecCobro.frx":030A
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
         Begin MSDataListLib.DataCombo dcmbTipo 
            Height          =   315
            Left            =   1245
            TabIndex        =   3
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Image imgBtnDn 
            Height          =   210
            Left            =   1920
            Picture         =   "frmEfecCobro.frx":03E4
            Top             =   2160
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Image imgBtnUp 
            Height          =   210
            Left            =   1680
            Picture         =   "frmEfecCobro.frx":0510
            ToolTipText     =   "Elimina una Fila"
            Top             =   2160
            Visible         =   0   'False
            Width           =   225
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
            Left            =   2880
            TabIndex        =   15
            Top             =   2055
            Width           =   855
         End
         Begin VB.Label lblBeneficiario 
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
            TabIndex        =   14
            Top             =   405
            Width           =   930
         End
         Begin VB.Label lbldescripcion1 
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
            Left            =   4080
            TabIndex        =   13
            Top             =   405
            Width           =   900
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   11055
         _cx             =   19500
         _cy             =   4260
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
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEfecCobro.frx":0646
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   7200
         TabIndex        =   20
         Top             =   4830
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5700
      TabIndex        =   10
      Top             =   8040
      Width           =   1575
   End
End
Attribute VB_Name = "frmEfecCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para realizar la efectivización de Cobros de CxC                      #
'#  frmEfecCobro V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'# -Ventana que permite visualizar y escoger los cobros a depositarse en la     #
'#  cuenta Bancaria  que sea escogida por el usuario                            #
'# -Realiza el asiento de las cuentas afectadas con el pago del documento       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  doc_pago: Contiene todo los datos del los cobros realizados hasta la fecha  #
'#  Banco: Contiene los nombres de los bancos utilizados por la empresa         #
'#  Cta_banco: Contiene las cuentas bancarias de la empresa                     #
'#  aiento: Tabla que almacena todos los asientos de las transacciones realizadas#
'#  det_asiento: Tiene lo detalles de los asientos realizados                   #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsBan: Objeto para consultar a la base de datos                          #
'#    clsCta: Objeto para consultar a la base de datos                          #
'#    clsPag: Objeto para consultar a la base de datos                          #
'#    clsSQL: Objeto para consultar a la base de datos                          #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsPag As New clsConsulta
Private clsSql As New clsConsulta
Private clsAsi As New clsConsulta
Private clsDet As New clsConsulta
Private clsTip As New clsConsulta
Private strSql As String

Private Sub cmdExportarFC_ContacPoint_Click()
    ExportarVSFG VSFG1
End Sub
Private Sub ExportarVSFG(objetoVSFG As VSFlexGrid)
    Dim sDir As String
    sDir = CurDir
    cdArchivo.ShowSave
    If cdArchivo.FileName <> "" Then
        objetoVSFG.SaveGrid cdArchivo.FileName, flexFileCommaText, True
    End If
    ChDir sDir
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
    Set clsPag = Nothing
    Set clsSql = Nothing
    Set clsAsi = Nothing
    Set clsDet = Nothing
    Set clsTip = Nothing
End Sub

Private Sub PonerBotones()
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub numerargrid()
    For i = 1 To (VSFG1.Rows - 1)
        VSFG1.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub Llenar_Grid(ByVal Row)
'limpia el grid para los asientos
VSFG.Clear 1
VSFG.Rows = 1
If optBanco.value = True Then
    strSql = " SELECT cta_codigo,cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE cta_codigo = '" & dcmbCuenta.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
Else
    strSql = " SELECT cta_codigo,cta_nombre " & _
             " FROM parametro INNER JOIN ctaconta ON parametro.emp_codigo=ctaconta.emp_codigo AND parametro.par_texto=ctaconta.cta_codigo " & _
             " WHERE par_codigo = 'CHP' AND parametro.emp_codigo = '" & strEmpresa & "' "
End If
'Cuenta del Banco
clsCta.Ejecutar strSql
'coloca los datos de los asientos de los documentos seleccionados
For i = 1 To VSFG1.Rows - 1
    
    If VSFG1.TextMatrix(i, 1) = "-1" Then ' -1 significa que la fila ha sido seleccionada
        strSql = " SELECT det_doc_pago.cta_codigo as codigo, if(det_doc_pago.cta_codigo= '*' , 'CAJA', ctaconta.cta_nombre) as nombre, det_doc_pag_debe as debe, det_doc_pag_haber as haber,COALESCE(cen_cos_codigo,'') as cen_cos_codigo " & _
                 " FROM doc_pago INNER JOIN det_doc_pago ON doc_pago.emp_codigo=det_doc_pago.emp_codigo AND doc_pago.doc_pag_codigo=det_doc_pago.doc_pag_codigo " & _
                 " LEFT JOIN ctaconta ON det_doc_pago.emp_codigo =ctaconta.emp_codigo " & _
                 " AND det_doc_pago.cta_codigo = ctaconta.cta_codigo " & _
                 " WHERE doc_pago.doc_pag_codigo = '" & VSFG1.TextMatrix(i, 2) & "' " & _
                 " AND det_doc_pago.emp_codigo = '" & strEmpresa & "' " & _
                 " AND if(doc_pag_estado='GIRADO',det_doc_pag_n=0,det_doc_pag_n=1) "
        clsDet.Ejecutar strSql
        'coloca los valores en el grid
        While Not clsDet.adorec_Def.EOF
            With VSFG
                'Añade una fila antes de ingresar los datos
                .AddItem ""
                'Pone los datos consultados en el Grid
                .TextMatrix(.Rows - 1, 1) = clsDet.adorec_Def("codigo")
                .TextMatrix(.Rows - 1, 2) = clsDet.adorec_Def("nombre")
                .TextMatrix(.Rows - 1, 3) = clsDet.adorec_Def("debe")
                .TextMatrix(.Rows - 1, 4) = clsDet.adorec_Def("haber")
                .TextMatrix(.Rows - 1, 5) = clsDet.adorec_Def("cen_cos_codigo")
                clsDet.adorec_Def.MoveNext
            End With
        Wend
        'En las cuentas con * se coloca la cuenta del banco escogido para realizar el depósito
        For j = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(j, 1) = "*" Then
                If clsCta.adorec_Def.RecordCount > 0 Then
                     VSFG.TextMatrix(j, 1) = clsCta.adorec_Def("cta_codigo")
                     VSFG.TextMatrix(j, 2) = clsCta.adorec_Def("cta_nombre")
                Else
                    VSFG.TextMatrix(j, 1) = ""
                    VSFG.TextMatrix(j, 2) = ""
                End If
            End If
        Next j
    End If
Next
'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd
 
    a = VSFG.Rows - 1
    For i = 1 To a
        For j = i + 1 To a
            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) And VSFG.TextMatrix(i, 5) = VSFG.TextMatrix(j, 5) Then
                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
                VSFG.RemoveItem j
                a = a - 1
                j = j - 1
            End If
            If j >= a Then
                Exit For
            End If
        Next j
    Next i
    CalcuTotal
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

End Sub

Private Sub pagos()
    txtValor = ""
    For i = 1 To VSFG1.Rows - 1
        If VSFG1.TextMatrix(i, 1) = "-1" Then
            Suma = Suma + Val(VSFG1.TextMatrix(i, 8))
        End If
    Next i
    txtValor = Format(Suma, "##0.00")
End Sub
Private Function AutoNumero_Cuenta() As String
    'Pone el número de cuenta por cobrar / pagar siguiente
    Set clsNumCuentas = New clsConsulta
    clsNumCuentas.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT COALESCE(max(not_d_c_codigo),0) as num_cuenta" & _
             " FROM nota_d_c " & _
             " WHERE tip_not_d_c = 'C' AND emp_codigo = '" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsNumCuentas.Ejecutar strSql
    If Not clsNumCuentas.adorec_Def.EOF Then
        Var_NumCuenta = Val(clsNumCuentas.adorec_Def("num_cuenta")) + 1
    Else
        Var_NumCuenta = 1
    End If
    AutoNumero_Cuenta = Var_NumCuenta
End Function

Private Sub Limpiar()
    VSFG1.Clear 1
    VSFG1.Rows = 2
    VSFG.Clear 1
    VSFG.Rows = 2
    dcmbBanco.Text = ""
    txtNumero = ""
    txtTotalHaber = 0
    txtTotalDebe = 0
    txtValor = 0
    dcmbTipo.Text = ""
    'cmdConsultar.Enabled = False
End Sub


Private Sub chkFechas_Click()
    If chkFechas = 1 Then
        dtpFechaDesde.Enabled = False
        dtpFechaHasta.Enabled = False
    Else
        dtpFechaDesde.Enabled = True
        dtpFechaHasta.Enabled = True
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim Descripcion As String
    Dim FechaD As String
    Dim fechah As String
    Dim fechac As String
    Dim Pendiente As Integer
    Dim EstadoCH As String
    Dim campoAsiento As String
    Dim CHPost As String
'Comprueba que todos los datos esten ingresados
    If VSFG.Rows = 1 Then Exit Sub
    FechaD = Format(dtpFechaDesde.value, "yyyy-mm-dd")
    fechah = Format(dtpFechaHasta.value, "yyyy-mm-dd")
    fechac = Format(dtpFechaConta.value, "yyyy-mm-dd")
    If (IsDate(FechaD) = False) Then
        MsgBox "La fecha de asiento no es válida", vbInformation, "Efectivización de Cobros"
        Exit Sub
    End If

    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "Efectivización de Cobros"
        Exit Sub
    Else
        'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If VSFG.TextMatrix(1, 1) = "" Or (optBanco.value = True And (dcmbBanco = "" Or dcmbTipo.Text = "" Or txtNumero = "")) Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "Efectivización de Cobros"
            txtNumero.SetFocus
            Exit Sub
        Else
            'descripcion de asiento y nota de crédito
            Dim nl As String
            If optBanco.value = True Then
                Descripcion = UCase("Depósito Banco: " + " " + dcmbBanco.Text + " " + "No. de Documento:" + " " + txtNumero + " " + "Cantidad: " + " " + txtValor)
                Pendiente = 0
                EstadoCH = "COBRADO"
                campoAsiento = "asi_numasiento"
            Else
                Descripcion = UCase("Cheques Post Fechados por: " + " " + txtValor)
                Pendiente = 1
                EstadoCH = "POSTFECHADO"
                campoAsiento = "asi_numasiento2"
                strSql = " SELECT par_texto " & _
                         " FROM parametro " & _
                         " WHERE par_codigo = 'CHP' AND parametro.emp_codigo = '" & strEmpresa & "' "
                clsPag.Ejecutar strSql
                CHPost = clsPag.adorec_Def("par_texto")
            End If
            Dim clsAsiento As New clsContable
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            clsAsiento.NuevoAsiento "I", fechac, 0, 0, FormatoD2(txtTotalDebe), Descripcion
            For i = 1 To VSFG1.Rows - 1
                If Abs(VSFG1.TextMatrix(i, 1)) = 1 Then
                    Descripcion = Descripcion & vbNewLine & "CLI: " & VSFG1.TextMatrix(i, 7) & " DOC: " & VSFG1.TextMatrix(i, 3) & "(" & VSFG1.TextMatrix(i, 4) & ") No:" & VSFG1.TextMatrix(i, 5) & " ABONO: " & VSFG1.TextMatrix(i, 8) & Replace(" OBS:" & VSFG1.TextMatrix(i, 9), vbNewLine, " ")
                    'Actualiza asientos en pagos
                    strSql = " UPDATE pago " & _
                             " SET asi_numasiento='" & clsAsiento.NumAsiento & _
                             "' , pag_fechamod= CURRENT_TIMESTAMP, pag_usumod='" & strUsuario & "' " & _
                             " WHERE doc_pag_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND emp_codigo = '" & strEmpresa & "' " & _
                             " AND cue_p_c_tipo='C' "
                    clsPag.Ejecutar strSql, "M"
                    'Actualiza la tabla doc_pago
                    strSql = " UPDATE doc_pago " & _
                             " SET doc_pag_fecha_efec='" & fechac & _
                             "'," & campoAsiento & "='" & clsAsiento.NumAsiento & _
                             "',doc_pag_pendiente='" & Pendiente & "', doc_pag_estado = '" & EstadoCH & "' , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                             " WHERE doc_pag_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsPag.Ejecutar strSql, "M"
                    If optBanco.value = True Then
                        strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo," & _
                                 " max(doc_pago.doc_pag_fecha_doc) as fecha,cuenta_p_c.cue_p_c_valor,COALESCE(sum(p2.pag_monto),0),COALESCE(com_ret_total,0)," & _
                                 " cuenta_p_c.cue_p_c_valor-COALESCE(sum(p2.pag_monto),0)-COALESCE(com_ret_total,0) as saldo " & _
                                 " FROM pago as p1 INNER JOIN cuenta_p_c ON cuenta_p_c.cue_p_c_codigo=p1.cue_p_c_codigo " & _
                                 " AND cuenta_p_c.cue_p_c_tipo=p1.cue_p_c_tipo " & _
                                 " AND cuenta_p_c.emp_codigo=p1.emp_codigo " & _
                                 " INNER JOIN pago as p2 ON cuenta_p_c.cue_p_c_codigo=p2.cue_p_c_codigo " & _
                                 " AND cuenta_p_c.cue_p_c_tipo=p2.cue_p_c_tipo " & _
                                 " AND cuenta_p_c.emp_codigo=p2.emp_codigo " & _
                                 " INNER JOIN doc_pago ON p2.doc_pag_codigo=doc_pago.doc_pag_codigo " & _
                                 " AND p2.emp_codigo=doc_pago.emp_codigo " & _
                                 " AND doc_pago.doc_pag_pendiente=0 AND doc_pago.doc_pag_estado!='ANULADO' " & _
                                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                                 " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                                 " AND cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                                 " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                                 " AND p1.doc_pag_codigo='" & VSFG1.TextMatrix(i, 2) & "' " & _
                                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_valor "
                        clsPag.Ejecutar strSql, "M"
                        While Not clsPag.adorec_Def.EOF
                            If (FormatoD2(clsPag.adorec_Def("saldo")) <= 0) Then
                                strSql = " UPDATE cuenta_p_c " & _
                                         " SET cue_p_c_fechapago='" & clsPag.adorec_Def("fecha") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                         " WHERE cue_p_c_tipo= 'C' " & _
                                         " AND cue_p_c_codigo= '" & clsPag.adorec_Def("cue_p_c_codigo") & _
                                         "' AND cue_p_c_egr_codigo = '" & clsPag.adorec_Def("cue_p_c_egr_codigo") & _
                                         "' AND emp_codigo = '" & strEmpresa & "' "
                                clsSql.Ejecutar strSql, "M"
                            End If
                            clsPag.adorec_Def.MoveNext
                        Wend
                    Else
                        strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                                 " VALUES ('" & strEmpresa & "','" & VSFG1.TextMatrix(i, 2) & "',1, '*','" & FormatoD2(VSFG1.TextMatrix(i, 8)) & "', '0' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                        clsSql.Ejecutar strSql, "M"
                        strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                                 " VALUES ('" & strEmpresa & "','" & VSFG1.TextMatrix(i, 2) & "',1, '" & CHPost & "','0','" & FormatoD2(VSFG1.TextMatrix(i, 8)) & "', CURRENT_TIMESTAMP, '" & strUsuario & "') "
                        clsSql.Ejecutar strSql, "M"
                    End If
                End If
            Next i
            
            
            clsAsiento.ModificarAsiento FormatoD2(txtTotalDebe), FormatoD2(txtTotalHaber), , , , Descripcion
            'ingreso del detalle del asiento
            With VSFG
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) = "" Then
                        Exit For
                    Else
                        clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                    End If
                Next i
            End With
            
            If optBanco.value = True Then
                'GENERACION DE LA NOTA DE CREDITO
                'Calcula el código de la Nota de Crédito
                strSql = " SELECT cta_ban_saldoreal, cta_ban_saldoprevisto " & _
                          " FROM cta_banco " & _
                          " WHERE cta_ban_numero = '" & dcmbCuenta.Text & "' AND emp_codigo = '" & strEmpresa & "' "
                clsCta.Ejecutar strSql
                If Not clsCta.adorec_Def.EOF Then
                    saldoreal = clsCta.adorec_Def("cta_ban_saldoreal") + txtValor
                    saldoPrevisto = clsCta.adorec_Def("cta_ban_saldoprevisto") + txtValor
                Else
                    saldoreal = txtValor
                    saldoPrevisto = txtValor
                End If
                'Guarda los datos de la Nota de Crédito
                
                strSql = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto,asi_numasiento,not_d_c_conciliado , not_d_c_fechamod, not_d_c_usumod) " & _
                         " VALUES ('C','" & AutoNumero_Cuenta & "', '" & dcmbCuenta.Text & "', '" & dcmbBanco.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtNumero.Text & "','" & fechac & "','" & Descripcion & "','" & txtValor.Text & "','" & clsAsiento.NumAsiento & "',0, CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsSql.Ejecutar strSql, "M"
                'Actualiza los valores de los saldos
                strSql = " UPDATE cta_banco " & _
                         " SET cta_ban_saldoreal= '" & saldoreal & "',cta_ban_saldoprevisto= '" & saldoPrevisto & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                         " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
                clsSql.Ejecutar strSql, "M"

            End If
            MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
            
''            'Impresion de Comprobante de Ingreso
''            Dim rptNuevo As New frmReporte
''            rptNuevo.strAsiento = clsAsiento.NumAsiento
''            rptNuevo.strReporte = "rptAsiento"
''            rptNuevo.Show
            
            Dim rptCompIng As New frmReporte
            rptCompIng.strAsiento = clsAsiento.NumAsiento
            rptCompIng.strReporte = "rptComprobanteIngreso"
            rptCompIng.Show
            
            
            
            Set clsAsiento = Nothing
            
            dtpFechaConta.value = HoyDia
            dtpFechaDesde.value = HoyDia
            dtpFechaHasta.value = HoyDia
        End If
    End If
    Limpiar
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConsultar_Click()
    Dim strSqlEstados As String
    Dim strSqlDoc As String
    If optBanco = True Then
        strSqlEstados = "(doc_pag_estado = 'GIRADO' OR doc_pag_estado = 'POSTFECHADO')"
    Else
        strSqlEstados = "doc_pag_estado = 'GIRADO'"
    End If
    'fechas desde y hasta
    fechadesde = Format(dtpFechaDesde.value, "yyyy-mm-dd")
    fechahasta = Format(dtpFechaHasta.value, "yyyy-mm-dd")
    
    If dcmbTipoDoc.BoundText = "E%E" Then
        strSqlDoc = " doc_pago.tip_doc_pag_codigo is null "
    ElseIf dcmbTipoDoc.BoundText = "%" Then
        strSqlDoc = " (doc_pago.tip_doc_pag_codigo LIKE '" & dcmbTipoDoc.BoundText & "' OR doc_pago.tip_doc_pag_codigo is null) "
    Else
        strSqlDoc = " doc_pago.tip_doc_pag_codigo LIKE '" & dcmbTipoDoc.BoundText & "' "
    End If
    
    ' Consulta los datos del los documentos cobrados para todas las fechas
    If chkFechas.value = False Then
        
        strSql = " SELECT '0', doc_pag_codigo, if(isnull(doc_pago.tip_doc_pag_codigo) or doc_pago.tip_doc_pag_codigo ='' ,'EFECTIVO',tipo_doc_pago.tip_doc_pag_nombre)as tip_doc_pag_nombre, COALESCE(banco.ban_nombre,'-') as ban_nombre, doc_pag_numero, doc_pag_fecha_doc, CONCAT(per_apellido,' ',per_nombre), doc_pag_valor, doc_pag_observacion,doc_pag_estado " & _
                 " FROM ((doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo) " & _
                 " LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo) " & _
                 " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' " & _
                 " AND " & strSqlDoc & " " & _
                 " AND doc_pag_fecha_doc between '" & fechadesde & "' AND '" & fechahasta & "' AND " & strSqlEstados
    Else
        strSql = " SELECT '0', doc_pag_codigo, if(isnull(doc_pago.tip_doc_pag_codigo) or doc_pago.tip_doc_pag_codigo ='' ,'EFECTIVO',tipo_doc_pago.tip_doc_pag_nombre)as tip_doc_pag_nombre, COALESCE(banco.ban_nombre,'-') as ban_nombre, doc_pag_numero, doc_pag_fecha_doc, CONCAT(per_apellido,' ',per_nombre), doc_pag_valor, doc_pag_observacion,doc_pag_estado " & _
                 " FROM ((doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo) " & _
                 " LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo) " & _
                 " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' " & _
                 " AND " & strSqlDoc & " " & _
                 " AND " & strSqlEstados
    End If
    clsPag.Ejecutar strSql
    cmdConsultar.Tag = "N"
    If Not clsPag.adorec_Def.EOF Then
        Set VSFG1.DataSource = clsPag.adorec_Def.DataSource
        VSFG1.ColDataType(1) = flexDTBoolean
    Else
        Set VSFG1.DataSource = Nothing
        VSFG1.Clear 1
        VSFG1.Rows = 2
    End If
    cmdConsultar.Tag = "S"
    numerargrid
End Sub

Private Sub dcmbBanco_Change()
dcmbCuenta = ""
    If dcmbBanco.Text = "" Then
        dcmbCuenta.Text = ""
        Exit Sub
    Else
        strSql = " SELECT cta_ban_numero, cta_ban_ctaconta, ban_codigo " & _
                 " FROM cta_banco " & _
                 " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
        clsCta.Ejecutar strSql
        
        If clsCta.adorec_Def.EOF = False Then
            Set dcmbCuenta.RowSource = clsCta.adorec_Def.DataSource
            dcmbCuenta.ListField = "cta_ban_numero"
            dcmbCuenta.BoundColumn = "cta_ban_ctaconta"
            'dcmbCuenta.Tag = clsCta.adorec_Def("cta_ban_ctaconta")
            'dcmbCuenta.Text = clsCta.adorec_Def("cta_ban_numero")
            
        Else
            Set dcmbCuenta.RowSource = Nothing
        End If
        
    End If
End Sub


Private Sub dcmbCuenta_Change()
    Dim j As Long
    'cmdConsultar.Enabled = True
    If VSFG.Rows > 1 Then
        Llenar_Grid (1)
    End If
End Sub

Private Sub dcmbTipo_Change()
 If dcmbTipo = "" Then
    txtDescripciont = ""
    Exit Sub
 End If
strSql = " SELECT CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
         " FROM tipo_nota " & _
         " WHERE tip_not_d_c = 'C' AND  tip_not_codigo = '" & dcmbTipo.BoundText & "' "
clsTip.Ejecutar strSql

If clsTip.adorec_Def.EOF = False Then
    txtDescripciont = clsTip.adorec_Def("descripcion")
Else
    txtDescripciont.Text = ""
End If

End Sub


Private Sub Form_Activate()
'     consulta para saber los  bancos existentes
    strSql = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSql

    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        dcmbBanco = ""
    End If
    
    'cmdConsultar.Enabled = False
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
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsTip.Inicializar AdoConn, AdoConnMaster
    'Pone la fecha actual en los combos
    
    dtpFechaConta.value = HoyDia
    dtpFechaDesde.value = HoyDia
    dtpFechaHasta.value = HoyDia
    
    'cmdConsultar.Enabled = False
    
    'Consulta el tipo documento de pago
    strSql = " SELECT tip_doc_pag_codigo, tip_doc_pag_nombre " & _
             " FROM tipo_doc_pago " & _
             " UNION " & _
             " SELECT '%',' TODOS ' " & _
             " UNION " & _
             " SELECT 'E%E','FACT. EFECTIVO' " & _
             " ORDER BY tip_doc_pag_nombre"
    clsTip.Ejecutar strSql
    If clsTip.adorec_Def.EOF = False Then
        Set dcmbTipoDoc.RowSource = clsTip.adorec_Def.DataSource
        dcmbTipoDoc.ListField = "tip_doc_pag_nombre"
        dcmbTipoDoc.BoundColumn = "tip_doc_pag_codigo"
        
        dcmbTipoDoc.Text = " TODOS "
        dcmbTipoDoc.BoundText = "%"
    Else
        dcmbTipoDoc.Text = " TODOS "
        dcmbTipoDoc.BoundText = "%"
    End If
    
    'Consulta el tipo de nota de crédito
    strSql = " SELECT cen_cos_codigo, cen_cos_nombre " & _
             " FROM centro_costo " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY cen_cos_nombre"
    clsTip.Ejecutar strSql
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsTip.adorec_Def, "cen_cos_codigo,*cen_cos_nombre", "cen_cos_codigo")
    'Consulta el tipo de nota de crédito
    strSql = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'C'" & _
             " ORDER BY tip_not_codigo"
    clsTip.Ejecutar strSql
    If clsTip.adorec_Def.EOF = False Then
    Set dcmbTipo.RowSource = clsTip.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
'    dcmbTipo.Text = clsTip.adorec_Def("tip_not_nombre")
'    txtDescripciont.Text = clsTip.adorec_Def("descripcion")
    Else
        dcmbTipo.Text = ""
        dcmbTipo.BoundText = ""
        txtDescripcion = ""
    End If
    
End Sub

Private Sub optBanco_Click()
    frmBanco.Enabled = True
    dcmbTipo.Enabled = True
    If VSFG1.Rows > 1 Then
        cmdConsultar_Click
        Llenar_Grid 1
    End If
End Sub

Private Sub optPosfechado_Click()
    frmBanco.Enabled = False
    dcmbTipo.Enabled = False
    If VSFG1.Rows > 1 Then
        cmdConsultar_Click
        Llenar_Grid 1
    End If
End Sub

Private Sub txtTotalDebe_Change()
    txtTotalDebe = FormatoD2(txtTotalDebe)
End Sub

Private Sub txtTotalHaber_Change()
 txtTotalHaber = FormatoD2(txtTotalHaber)
End Sub


Private Sub txtValor_Change()
    txtValor = FormatoD2(txtValor)
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

        If Row = 1 Then
            If Col = 1 Then
                Cancel = True
            End If
            If Col = 2 Then
                Cancel = True
            End If
            If Col = 3 Then
               Cancel = True
            End If
            If Col = 4 Then
                Cancel = True
            End If
        End If

End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If cmdConsultar.Tag = "S" Then
        If Col = 1 And Row > 0 Then
            If VSFG1.TextMatrix(Row, 1) = "-1" Then
                VSFG1.Select Row, 1, Row, 9
                VSFG1.FillStyle = flexFillRepeat
                VSFG1.CellBackColor = &HC0FFFF
                'VSFG1.Select Row, 9
                Llenar_Grid (Row)
                pagos
            ElseIf VSFG1.TextMatrix(Row, 1) = "0" Then
                  VSFG1.Select Row, 1, Row, 9
                  VSFG1.FillStyle = flexFillRepeat
                  VSFG1.CellBackColor = &HFFFFFF
                  'VSFG1.Select Row, 9
              '    VSFG1.TextMatrix(Row, 7) = "0"
                  Llenar_Grid (Row)
                 pagos
            End If
        End If
    End If
End Sub


