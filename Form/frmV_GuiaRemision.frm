VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmV_GuiaRemision 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantener Guías y Reservas"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmV_GuiaRemision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   12285
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerarPedido 
      Caption         =   "Pasar a pedido"
      Height          =   375
      Left            =   480
      TabIndex        =   47
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBaja 
      Caption         =   "Baja"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   37
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdDevolver 
      Caption         =   "Devolver"
      Height          =   375
      Left            =   3727
      TabIndex        =   31
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8715
      TabIndex        =   30
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton CmdFacturar 
      Caption         =   "Facturar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2115
      TabIndex        =   29
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   7155
      TabIndex        =   28
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton CmdFacturarDevolver 
      Caption         =   "Devolver y Facturar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5235
      TabIndex        =   27
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Frame frmDoc 
      BackColor       =   &H00DDDDDD&
      Caption         =   "DATOS DE GUIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   75
      TabIndex        =   18
      Top             =   120
      Width           =   12135
      Begin VB.TextBox txtTotalFacturar 
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
         Height          =   285
         Left            =   8715
         TabIndex        =   43
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDevolver 
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
         Height          =   285
         Left            =   6135
         TabIndex        =   41
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtLector 
         Height          =   285
         Left            =   9600
         TabIndex        =   39
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optReserva 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Reservas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optGuia 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Guías"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGGuia 
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   11895
         _cx             =   20981
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_GuiaRemision.frx":030A
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   1815
         TabIndex        =   19
         Top             =   600
         Width           =   6720
         _ExtentX        =   11853
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
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   6720
         TabIndex        =   45
         Top             =   240
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
      Begin VB.Label Label14 
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
         Left            =   5880
         TabIndex        =   46
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot. Facturar:"
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
         Left            =   7590
         TabIndex        =   44
         Top             =   3405
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot. Devolver:"
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
         Left            =   4980
         TabIndex        =   42
         Top             =   3405
         Width           =   1005
      End
      Begin VB.Label Label11 
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
         Left            =   8880
         TabIndex        =   40
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el tipo de documento"
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
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un Cliente:"
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
         TabIndex        =   20
         Top             =   645
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "DETALLE A FACTURAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   1035
      TabIndex        =   0
      Top             =   3960
      Width           =   10215
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   7680
         TabIndex        =   38
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Value           =   42024.4754398148
      End
      Begin VB.TextBox TxtSubTotal 
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
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox TxtIva 
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
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox TxtRecargo 
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
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TxtDesc 
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
         Height          =   285
         Left            =   8640
         TabIndex        =   2
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TxtObserv 
         Height          =   405
         Left            =   360
         MaxLength       =   250
         TabIndex        =   1
         Top             =   4560
         Width           =   9615
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   9960
         _cx             =   17568
         _cy             =   2990
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_GuiaRemision.frx":0514
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
         SubtotalPosition=   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
         Height          =   855
         Left            =   390
         TabIndex        =   8
         Top             =   3360
         Width           =   4305
         _cx             =   7594
         _cy             =   1508
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
         FormatString    =   $"frmV_GuiaRemision.frx":065D
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
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   7680
         TabIndex        =   21
         Top             =   645
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   705
         Width           =   4920
         _ExtentX        =   8678
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
      Begin MSDataListLib.DataCombo cmbCliente2 
         Height          =   315
         Left            =   1095
         TabIndex        =   32
         Top             =   240
         Width           =   4920
         _ExtentX        =   8678
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   465
         TabIndex        =   33
         Top             =   285
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   225
         TabIndex        =   25
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Factura:"
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
         Left            =   6270
         TabIndex        =   24
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pago:"
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
         Left            =   6720
         TabIndex        =   22
         Top             =   690
         Width           =   900
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label LblDetalle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
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
         Left            =   405
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total pedido:"
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
         Left            =   7590
         TabIndex        =   15
         Top             =   4230
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   7755
         TabIndex        =   14
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Left            =   7665
         TabIndex        =   13
         Top             =   3885
         Width           =   825
      End
      Begin VB.Label LblIva 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA X%:"
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
         Left            =   7920
         TabIndex        =   12
         Top             =   3390
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   7380
         TabIndex        =   11
         Top             =   3150
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   10
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         TabIndex        =   9
         Top             =   4320
         Width           =   1185
      End
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   0
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   375
      Picture         =   "frmV_GuiaRemision.frx":06DD
      ToolTipText     =   "Elimina una Fila"
      Top             =   5880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   615
      Picture         =   "frmV_GuiaRemision.frx":0813
      Top             =   5880
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmV_GuiaRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################
'#  Forma para ver y facturar un proyecto de venta, en base a los ingresos y
'#  egresos que se han realizado en el proyecto.
'#  frmV_FacProVenta V1.0
'#  Copyright (C) 2002
'#
'#  Opciones que permite:
'#  *   En una lista se despliegan los datos del los distintos proyectos de
'#      trabajo de una emprea como el cliente y el vendedor que lo atiende y el
'#      estado del mismo.
'#  *   De igual manera es necesario seleccionar el tipo de facturación que se
'#      va a aplicar al proyecto y la fecha en que se lo factura.
'#  *   Es necesario también seleccionar la forma de pago.
'#  *   El usuario puede seleccionar los posibles recargos que puede generar
'#      la facturación de proyecto.
'#
'#  Procesos internos que maneja:
'#  *   La lista que muestra los distintos proyectos, se refresca automáticamente
'#      cada 20 segundos para buscar un nuevo proyecto de trabajo.
'#  *   Al dar un click en la lista de proyectos, automáticamente se cargan los
'#      detalles de los movimientos del mismo en un segundo grid.
'#  *   Una vez que el proyecto ha sido facturado su estado pasa a vendido. Al
'#      igual que su cotización relacionada.
'#  *   Se pueden ver solo los proyectos que no están facturados y los que ya
'#      se han facturado el día de hoy.
'#  *   Una vez que se va a facturar el proyecto se generan automáticamente las
'#      respectivas retenciones que puede tener el cliente del mismo.
'#
'#  Tablas que maneja:
'#
'#  persona:
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el
'#      proyecto que se está facturando.
'#  *   También se extrae el nombre del vendedor asignado al pedido.
'#  persona_ret:
'#  *   De esta tala se extraen las diferentes retenciones que puede tener un
'#      cliente determinado para luego aplicarlas a esta factura.
'#  retencion:
'#  *   De aquí se extraen los valores y descripciones de las retenciones, que
'#      se aplicarán posteriormente.
'#  det_egreso:
'#  *   En esta tabla se guardan los detalles del nuevo documento de egreso de
'#      productos.
'#  ocargo:
'#  *   De esta tabla se extraen los diferentes recargos que se puede manejar
'#      al realizar un nuevo egreso de productos de bodega, como pueden ser:
'#      transporte, fletes, etc.
'#  det_egreso_c:
'#  *   En esta tabla se guardan los diferentes recargos que puede tener esta
'#      nueva compra o egreso de productos.
'#  det_egreso_ret:
'#  *   En esta tabla se guardan los valores de las retenciones aplicadas a este
'#      ingreso de productos a bodega.
'#
'################################################################################

Private clsClie As New clsConsulta
Private clsSql As New clsConsulta
Private clsFPago As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsPrds As New clsConsulta
Private clsBods As New clsConsulta
Private clsLstPrds As New clsConsulta
Private clsCos As New clsCostear
Private strTipEgr As String
Private strTipIng As String
Private IVA As Double
Private strClaveMAESTRA As String
Private FechaUltFac As String
Private strForma As String
Private MINCredito As Double
Private CodPer As String
Private SecPublico As Boolean
Private SinIVA As Boolean

Private Sub cmbCliente2_Change()
    Dim strSql As String
    Dim Forma_Defecto As String, codDC As String
        FechaUltFac = ""
    If cmbCliente2.MatchedWithList = True Then
        CodPer = cmbCliente2.BoundText
        strSql = " SELECT COALESCE(persona.for_pag_codigo,'CONT') as for_pag_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_sec_publico,per_siniva " & _
                " FROM persona " & _
                " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente2.BoundText & "' "
        clsSql.Ejecutar strSql
        Forma_Defecto = clsSql.adorec_Def("for_pag_codigo")
        If FormatoD0(clsSql.adorec_Def("per_bloqueado")) = 1 Then
            MsgBox "Cliente BLOQUEADO por cartera." & vbNewLine & vbNewLine & "No podrá hacer pedido hasta resolver el problema en CARTERA", vbCritical, "Cartera"
            CmdFacturar.Enabled = False
            CmdFacturarDevolver.Enabled = False
        Else
            CmdFacturar.Enabled = True
            CmdFacturarDevolver.Enabled = True
        End If
        If Abs(FormatoD0(clsSql.adorec_Def("per_sec_publico"))) = 1 Then
            SecPublico = True
        Else
            SecPublico = False
        End If
        If Abs(FormatoD0(clsSql.adorec_Def("per_siniva"))) = 1 Then
            SinIVA = True
        Else
            SinIVA = False
        End If
        codDC = ""
        strForma = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo  " & _
                  " FROM forma_pago " & _
                  " WHERE emp_codigo='" & strEmpresa & "' " & _
                  " AND for_pag_codigo IN ('CONT','TAR'"
                 
        If Val(TxtTotal.Text) < Val(MINCredito) Then
            strSql = " SELECT COALESCE(egreso.egr_fecha,'" & HoyDia & "') as fecha, COALESCE(egreso.for_pag_codigo,'CONT') as codigo, COALESCE(forma_pago.for_pag_tiempo,0) as tiempo  " & _
                " FROM egreso " & _
                " INNER JOIN forma_pago ON forma_pago.emp_codigo=egreso.emp_codigo AND egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                " WHERE egreso.emp_codigo = '" & strEmpresa & "' AND egr_anulado=0 AND egreso.tip_egr_codigo='FAC' " & _
                " AND egreso.per_codigo = '" & CodPer & "' ORDER BY egreso.egr_fecha DESC LIMIT 1 "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                FechaUltFac = Format(DateAdd("d", CDbl(clsSql.adorec_Def("tiempo")), clsSql.adorec_Def("fecha")), "yyyy-MM-dd")
            
                If CStr(dtpFecha.Value) < FechaUltFac Then
                    'coloca el tiempo
                    codDC = Right(clsSql.adorec_Def("codigo"), 1)
                    Dim Diferencia As Long
                    Diferencia = CLng(DateDiff("d", dtpFecha.Value, FechaUltFac))
                    strForma = strForma & ",'" & Format(Diferencia, "00") & codDC & "')"
                    Forma_Defecto = Format(Diferencia, "00") & codDC
                Else
                     strForma = strForma & ") "
                     Forma_Defecto = "CON"
                End If
            Else
                strForma = strForma & ") "
                Forma_Defecto = "CON"
            End If
        Else
            strForma = strForma & ",'" & Forma_Defecto & "') "
        End If
        strForma = strForma & " ORDER BY for_pag_nombre "
        clsFPago.Ejecutar strForma

        Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
        CmbFpago.ListField = "for_pag_nombre"
        CmbFpago.BoundColumn = "for_pag_codigo"
        CmbFpago.BoundText = Forma_Defecto
    End If
        
        
                
        '******************************************


End Sub


Private Sub cmbNegocio_Change()
    Dim strCli As String
    cmdLimpiar_Click
    If cmbNegocio.BoundText <> "" Then
        
            strSql = " SELECT CONCAT(per_apellido,' ',per_nombre) as nombC, CONCAT(ven_apellido,' ',ven_nombre) as nombV, " & _
                     " cat_p_nombre, lis_pre_codigo, per_codigo, vendedor.ven_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                     " FROM (vendedor INNER JOIN persona ON (vendedor.ven_codigo = persona.ven_codigo) " & _
                     " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
                     " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
                     " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
                     " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
                     " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                     " ORDER BY nombC "
            clsClie.Ejecutar (strSql)
                'Coloca los datos del primer cliente de la lista
            Set cmbCliente.RowSource = clsClie.adorec_Def.DataSource
            Set cmbCliente2.RowSource = clsClie.adorec_Def.DataSource
            cmbCliente.ListField = "nombC"
            cmbCliente.BoundColumn = "per_codigo"
            cmbCliente2.ListField = "nombC"
            cmbCliente2.BoundColumn = "per_codigo"
            
    End If
    
End Sub

Private Sub cmdBaja_Click()
    If Me.cmbVendedor = "" Then
        MsgBox "Seleccione un Vendedor", , "Guias"
        Exit Sub
    Else
        'Verifica que se haya seleccionado un tipo de forma de pago
        If CmbFpago = "" Then
            MsgBox "Seleccione un tipo de forma de pago por favor.", vbInformation, "Forma de Pago"
            CmbFpago.SetFocus
            Exit Sub
        End If
        If (Bajar = True) Then
            BajarE
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCargar_Click()
    Dim sDir As String
    Dim strCodigo As String
    Dim lngCantidad As Long
    Dim blanco As Integer
    If cmbCliente = "" Then
        MsgBox "Primero seleccione un Cliente", vbInformation, "Cliente"
        cmbCliente.SetFocus
        Exit Sub
    End If
    sDir = CurDir
    cdArchivo.ShowOpen
    'cdArchivo.FileName
    If cdArchivo.FileName <> "" Then
        CargarArchivo cdArchivo.FileName
    End If
    ChDir sDir
End Sub

Private Sub CargarArchivo(pathArchivo As String)
    Dim CadenaRetorno() As Variant
    Dim i As Long
    Dim j As Long
    VSFG.Clear 1
    VSFG.Rows = 1
    'Llamada a la ultra función de lujo llamada LeerArchivo, archivo RRHH
    CadenaRetorno() = LeerArchivo(pathArchivo, 1, False, vbTab, vbNewLine)
    For i = 0 To UBound(CadenaRetorno) - 1
        For j = 1 To CadenaRetorno(i)(1)
            txtLector.Text = CadenaRetorno(i)(0)
            txtLector_KeyDown vbKeyReturn, 1
        Next j
    Next i
    CalcularTotales
    'CONTROLAR
End Sub


Private Sub cmdGenerarPedido_Click()
    Dim i As Long
    Dim num As String
    Dim clsSql As New clsConsulta
    Dim clsSqlDet As New clsConsulta
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    
    Dim Fact As String
    strSql = " SELECT tip_ped_ptofac,tip_ped_factura_directo " & _
             " FROM persona INNER JOIN tipo_pedido " & _
             " ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo='" & cmbCliente.BoundText & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        Fact = clsSql.adorec_Def(0)
    End If
    clsSqlDet.Inicializar AdoConn, AdoConnMaster
    strSql = " LOCK TABLES pedido WRITE "
    clsSql.Ejecutar strSql, "M"
    strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(strSucursal & Fact & "0000001") & "') as num " & _
             " From pedido " & _
             " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(strSucursal & Fact) & "%'" & _
             " GROUP BY emp_codigo"
    clsSql.Ejecutar (strSql), "M"
    num = clsSql.adorec_Def("num")
    strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
         " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega, ped_fechamod, ped_usumod) " & _
         " VALUES ('" & strEmpresa & "'," & num & ",'" & cmbCliente.BoundText & "','" & cmbVendedor.BoundText & "', " & _
         " 'SINTC','0'," & _
         " CURRENT_TIMESTAMP,'1',0,'GUIA', " & _
         " '0',1,'0',CURRENT_TIMESTAMP, '" & strUsuario & "') "
    clsSql.Ejecutar (strSql), "M"
    strSql = " UNLOCK TABLES"
    clsSql.Ejecutar (strSql), "M"
    For i = 1 To VSFGGuia.Rows - 1
        VSFGGuia.TextMatrix(i, 0) = 1
        VSFGGuia.TextMatrix(i, 9) = VSFGGuia.TextMatrix(i, 7)
        
        strSql = " SELECT * FROM det_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo='" & num & "' " & _
                 " AND prd_codigo='" & VSFGGuia.TextMatrix(i, 4) & "' " & _
                 " AND dep_codigo='" & VSFGGuia.TextMatrix(i, 3) & "' "
        clsSql.Ejecutar (strSql), "M"
        If clsSql.adorec_Def.RecordCount = 0 Then
            strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                     " det_ped_cant_entregada,det_ped_cant_confirmada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
                     " VALUES ('" & strEmpresa & "'," & num & ",'" & VSFGGuia.TextMatrix(i, 4) & "','" & VSFGGuia.TextMatrix(i, 3) & "'," & VSFGGuia.TextMatrix(i, 7) & "," & VSFGGuia.TextMatrix(i, 7) & "," & VSFGGuia.TextMatrix(i, 7) & ", " & _
                     " " & VSFGGuia.TextMatrix(i, 8) & ",0, CURRENT_TIMESTAMP, '" & strUsuario & "') "
        Else
            strSql = " UPDATE det_pedido " & _
                     " SET det_ped_cant_pedida=det_ped_cant_pedida+'" & VSFGGuia.TextMatrix(i, 7) & "'," & _
                     " det_ped_cant_entregada=det_ped_cant_entregada+'" & VSFGGuia.TextMatrix(i, 7) & "'," & _
                     " det_ped_cant_confirmada=det_ped_cant_confirmada+'" & VSFGGuia.TextMatrix(i, 7) & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ped_codigo='" & num & "' " & _
                     " AND prd_codigo='" & VSFGGuia.TextMatrix(i, 4) & "' " & _
                     " AND dep_codigo='" & VSFGGuia.TextMatrix(i, 3) & "' "
        End If
        clsSql.Ejecutar (strSql), "M"
        
    Next i
    MsgBox "Pedido gerado No. " & num
    cmdDevolver_Click
End Sub

Private Sub dtpFecha_LostFocus()
    If dtpFecha.Value <> HoyDia Then
        frmClave.strClaveMAESTRA = strClaveMAESTRA
        frmClave.dblPrecio = "Fecha"
        frmClave.Show vbModal
        If frmClave.Ret = False Then
            dtpFecha.Value = HoyDia
        End If
    Else
        dtpFecha.Value = HoyDia
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsClie = Nothing
    Set clsSql = Nothing
    Set clsFPago = Nothing
    Set clsRecargos = Nothing
    Set clsPrds = Nothing
    Set clsBods = Nothing
    Set clsLstPrds = Nothing
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
Private Sub PonerBotonesFac(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With VSFG
        For i = 1 To (.Rows - 1)
            '.TextMatrix(i, 0) = i
            If conBot = True And Val(.TextMatrix(i, 7)) <> 1 Then
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
    TxtTotal = Format(Suma + Val(TxtIva) + Val(TxtSubTotal), "####0.00")
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
        If Abs(FormatoD0(VSFG.TextMatrix(i, 10))) = 1 Then
            Suma = Suma + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
            sumaSDcto = sumaSDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
        Else
            SumaIVA = SumaIVA + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
            SumaIVASDcto = SumaIVASDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
        End If
        SumaDcto = SumaDcto + 0
    Next i
    TxtRecargo.Tag = FormatoD2(TxtRecargo.Text) + FormatoD2(SumaIVA)
    TxtRecargo.Text = FormatoD2(TxtRecargo.Text) + FormatoD2(SumaIVASDcto)
    'Coloca los totales parciales de la factura
    TxtDesc.Text = FormatoD2(SumaDcto)
    TxtSubTotal = FormatoD2(sumaSDcto)
    If SinIVA = False Then
        TxtIva = FormatoD2((Suma) * IVA / 100)
    Else
        TxtIva = 0#
    End If
    TxtTotal = FormatoD2(Suma + TxtIva + Val(TxtRecargo.Tag))
End Sub


Private Sub cmbCliente_Validate(Cancel As Boolean)
    CmdLimpiar = True
    'Cargar datos de Guias
    If cmbCliente.MatchedWithList = True Then
        CargarGuias
        strSql = " SELECT existencia.dep_codigo, producto.prd_codigo, sum(existencia.exi_cantidad) as exi_cantidad, " & _
                 " producto.prd_nombre, (producto.prd_costo/(1 - 0.064)) as prd_costo, lis_pre_p_precio,prd_iva " & _
                 " FROM ((((producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                 " AND producto.emp_codigo=lista_precio_p.emp_codigo) INNER JOIN existencia " & _
                 " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
                 " INNER JOIN categoria_p ON lista_precio_p.lis_pre_codigo=categoria_p.lis_pre_codigo AND lista_precio_p.emp_codigo=categoria_p.emp_codigo) " & _
                 " INNER JOIN persona ON categoria_p.cat_p_tipo=persona.cat_p_tipo AND categoria_p.cat_p_codigo=persona.cat_p_codigo) " & _
                 " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' " & _
                 " GROUP BY dep_codigo, prd_codigo " & _
                 " ORDER BY existencia.dep_codigo, producto.prd_codigo "
        'Ejecuta la consulta de lista de precios
        'clsLstPrds.Ejecutar (strSql)
        cmbCliente2.Text = cmbCliente.Text
        cmbCliente2.BoundText = cmbCliente.BoundText
    Else
        VSFGGuia.Clear 1
        VSFGGuia.Rows = 1
    End If
    
End Sub

Private Sub CmbFpago_Change()
'    CmdLimpiar = True
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Sub cmdDevolver_Click()
    If (Devolver = True) Then Unload Me
End Sub

Private Sub CmdFacturar_Click()
    If Me.cmbVendedor = "" Then
        MsgBox "Seleccione un Vendedor", , "Guias"
        Exit Sub
    Else
        'Verifica que se haya seleccionado un tipo de forma de pago
        If CmbFpago = "" Then
            MsgBox "Seleccione un tipo de forma de pago por favor.", vbInformation, "Forma de Pago"
            CmbFpago.SetFocus
            Exit Sub
        End If
        If (Facturar = True) Then
            FacturarE
            Unload Me
        End If
    End If
End Sub

Private Sub CmdFacturarDevolver_Click()
    Dim i As Long
    Dim num As String
    Dim clsSql As New clsConsulta
    Dim clsSqlDet As New clsConsulta
    clsSql.Inicializar AdoConn, AdoConnMaster
    strContenedorRecurrente = "111"
    Dim Fact As String
    strSql = " SELECT tip_ped_ptofac,tip_ped_factura_directo " & _
             " FROM persona INNER JOIN tipo_pedido " & _
             " ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo='" & cmbCliente.BoundText & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        Fact = clsSql.adorec_Def(0)
    End If
    clsSqlDet.Inicializar AdoConn, AdoConnMaster
    strSql = " BEGIN TRAN "
    clsSql.Ejecutar strSql, "M"
    strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(strSucursal & Fact & "0000001") & "') as num " & _
             " From pedido WITH (TABLOCKX) " & _
             " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(strSucursal & Fact) & "%'" & _
             " GROUP BY emp_codigo"
    clsSql.Ejecutar (strSql), "M"
    num = clsSql.adorec_Def("num")
    strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
         " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega, ped_fechamod, ped_usumod) " & _
         " VALUES ('" & strEmpresa & "'," & num & ",'" & cmbCliente.BoundText & "','" & cmbVendedor.BoundText & "', " & _
         " 'SINTC','0'," & _
         " CURRENT_TIMESTAMP,'1',0,'GUIA', " & _
         " '0',1,'0',CURRENT_TIMESTAMP, '" & strUsuario & "') "
    clsSql.Ejecutar (strSql), "M"
    strSql = " COMMIT TRAN "
    clsSql.Ejecutar (strSql), "M"
    For i = 1 To VSFGGuia.Rows - 1
        If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 Then
            'VSFGGuia.TextMatrix(i, 9) = VSFGGuia.TextMatrix(i, 7)
            
            strSql = " SELECT emp_codigo FROM det_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ped_codigo='" & num & "' AND prd_codigo='" & VSFGGuia.TextMatrix(i, 4) & "' AND dep_codigo='" & VSFGGuia.TextMatrix(i, 3) & "'"
            clsSql.Ejecutar (strSql), "M"
            If clsSql.adorec_Def.RecordCount = 0 Then
                strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                         " det_ped_cant_entregada,det_ped_cant_confirmada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
                         " VALUES ('" & strEmpresa & "'," & num & ",'" & VSFGGuia.TextMatrix(i, 4) & "','" & VSFGGuia.TextMatrix(i, 3) & "'," & VSFGGuia.TextMatrix(i, 9) & "," & VSFGGuia.TextMatrix(i, 9) & "," & VSFGGuia.TextMatrix(i, 9) & ", " & _
                         " " & VSFGGuia.TextMatrix(i, 8) & ",0, CURRENT_TIMESTAMP, '" & strUsuario & "') "
            Else
                strSql = " UPDATE det_pedido " & _
                         " SET det_ped_cant_pedida=det_ped_cant_pedida+" & VSFGGuia.TextMatrix(i, 9) & ", " & _
                         " det_ped_cant_entregada=det_ped_cant_entregada+" & VSFGGuia.TextMatrix(i, 9) & ", " & _
                         " det_ped_cant_confirmada=det_ped_cant_confirmada+" & VSFGGuia.TextMatrix(i, 9) & " " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_codigo='" & num & "' AND prd_codigo='" & VSFGGuia.TextMatrix(i, 4) & "' AND dep_codigo='" & VSFGGuia.TextMatrix(i, 3) & "'"
            End If
            clsSql.Ejecutar (strSql), "M"
           End If
    Next i
    MsgBox "Pedido gerado No. " & num
    cmdDevolver_Click

End Sub

Private Sub cmdLimpiar_Click()
    'Muestra el formulario como si se hubiera cargado por primera vez
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGReca.Clear 1
    VSFGReca.Rows = 2
    VSFGReca.Enabled = True
    'CmdConfirmar.Enabled = False
    'CmdDeBaja.Enabled = False
    TxtSubTotal = ""
    TxtTotal = ""
    TxtRecargo = ""
    TxtIva = ""
    TxtDesc = ""
    fila = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" Then
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
    'Inicializa las clases para hacer distintas consultas
    clsClie.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsPrds.Inicializar AdoConn, AdoConnMaster
    clsBods.Inicializar AdoConn, AdoConnMaster
    clsLstPrds.Inicializar AdoConn, AdoConnMaster
    clsCos.Inicializar AdoConn, AdoConnMaster
    
    '****** CREDITO
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT par_numero " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'MIC' "
    clsSql.Ejecutar (strSql)
    MINCredito = clsSql.adorec_Def("par_numero")
    
    '****** CLAVE
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSql)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")
    
    IVA = PorIVA
    LblIva = "IVA " & IVA & " %:"
    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
'****** VENDEDORES
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as nombV " & _
             " FROM vendedor " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY nombV "
    clsSql.Ejecutar (strSql)
    Set cmbVendedor.RowSource = clsSql.adorec_Def.DataSource
    cmbVendedor.ListField = "nombV"
    cmbVendedor.BoundColumn = "ven_codigo"
'****** CLIENTES
    cargarTipoPedido
    'Obtiene todos los clientes de una empresa con su respectiva lista de precios y vendedor asociado
    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre) as nombC, CONCAT(ven_apellido,' ',ven_nombre) as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, vendedor.ven_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
             " FROM (vendedor INNER JOIN persona ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY nombC "
    clsClie.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsClie.adorec_Def.DataSource
    Set cmbCliente2.RowSource = clsClie.adorec_Def.DataSource
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
        cmbCliente2.ListField = "nombC"
        cmbCliente2.BoundColumn = "per_codigo"
    
    '****** PRODUCTOS
    'Recupera todos los productos de una empresa
    strSql = " SELECT prd_codigo, prd_nombre " & _
             " FROM producto " & _
             " Where emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
             " ORDER BY prd_codigo "
    clsPrds.Ejecutar (strSql)
    'Carga los productos en el combo de la columna 2 del flexGrid
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
'****** BODEGAS
    'Recupera todas las bodegas de una empresa
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By dep_nombre "
    clsBods.Ejecutar (strSql)
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsBods.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    
    PonerBotonesFac
    
    'Selecciona el primer elemento del combo de cotizaciones
    dtpFecha.Value = HoyDia
    TipoDoc
End Sub

Private Sub optGuia_Click()
    TipoDoc
End Sub

Private Sub optReserva_Click()
    TipoDoc
End Sub

Private Sub TxtDesc_Change()
    'TxtDesc = Replace(TxtDesc, ",", ".")
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    'Valida que solo se ingresen números en el campo de descuento
    If KeyAscii < 44 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDesc_LostFocus()
    'Calcula el total de la factura
    CalcuTotal
End Sub

Private Sub TxtIva_Change()
    'TxtIva = Replace(TxtIva, ",", ".")
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProdDevolucion UCase(txtLector.Text)
        txtLector.Text = ""
    ElseIf KeyCode = vbKeyF9 Then
        Dim i As Long
        For i = 1 To Me.VSFGGuia.Rows - 1
            If VSFGGuia.TextMatrix(i, 1) = txtLector.Text Then
                VSFGGuia.TextMatrix(i, 0) = 1
                VSFGGuia.ShowCell i, 7
                VSFGGuia.TextMatrix(i, 9) = VSFGGuia.TextMatrix(i, 7)
            End If
        Next i
    End If
End Sub

Private Sub AgregarProdDevolucion(strProd As String)
    Dim i As Long
    Dim Encontro As Boolean
    Encontro = False
    For i = 1 To VSFGGuia.Rows - 1
        VSFGGuia.ShowCell i, 9
        If VSFGGuia.TextMatrix(i, 4) = strProd Then
            If FormatoD0(VSFGGuia.TextMatrix(i, 7)) > FormatoD0(VSFGGuia.TextMatrix(i, 9)) Then
                VSFGGuia.TextMatrix(i, 0) = 1
                VSFGGuia.TextMatrix(i, 9) = FormatoD0(VSFGGuia.TextMatrix(i, 9)) + 1
                Encontro = True
                Exit For
            End If
        End If
    Next i
    If Encontro = False Then
        MsgBox "Producto " & strProd & " no encontrado o ya no cuadra la cantidad", vbInformation, "Dev. Guias"
    End If
'    CalcularTotales
End Sub

Private Sub CalcularTotales()
    Dim i As Long
    txtTotalDevolver.Text = 0
    txtTotalFacturar.Text = 0
    For i = 1 To VSFGGuia.Rows - 1
        txtTotalDevolver.Text = FormatoD0(txtTotalDevolver.Text) + FormatoD0(VSFGGuia.TextMatrix(i, 9))
        txtTotalFacturar.Text = FormatoD0(txtTotalFacturar.Text) + FormatoD0(VSFGGuia.TextMatrix(i, 10))
    Next i
End Sub

Private Sub TxtRecargo_Change()
    'TxtRecargo = Replace(TxtRecargo, ",", ".")
End Sub

Private Sub TxtSubTotal_Change()
    'TxtSubTotal = Replace(TxtSubTotal, ",", ".")
End Sub

Private Sub TxtSubTotal_KeyPress(KeyAscii As Integer)
    'Valida que solo se ingresen números en el campo de subtotal
    If KeyAscii < 44 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSubTotal_LostFocus()
    'Calcula el total de la factura
    CalcuTotal
End Sub

Private Sub txtTotal_Change()
    TxtTotal = Format(TxtTotal, "###0.00")
    cmbCliente2_Change
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Verifica que solo se ingresen números tanto en la cantidad como en el precio
    If Col = 4 Or Col = 5 Then
        'Verifica que solo se ingresen números en el campo cantidad
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en la cantidad.", vbInformation, "Cantidad"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
        'Verifica que solo se ingresen números en el campo precio
        If Not IsNumeric(VSFG.TextMatrix(Row, 5)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el precio.", vbInformation, "Precio"
            VSFG.TextMatrix(Row, 5) = intDato
        End If
        'Verifica que no se esté pidiendo más productos de los que hay en existencia
        If Val(VSFG.TextMatrix(Row, 4)) > Val(VSFG.TextMatrix(Row, 8)) And Val(VSFG.TextMatrix(Row, 7)) <> 1 And Left(VSFG.TextMatrix(Row, 2), 3) <> "PR-" Then
            If Val(VSFG.TextMatrix(Row, 8)) = 0 Then
                MsgBox "No hay existencia del este producto en la bodega.", vbInformation, "Existencia"
                VSFG.TextMatrix(Row, 4) = 0
            Else
                MsgBox "Solo hay diponible " & VSFG.TextMatrix(Row, 7) & " unidades de este producto en esta bodega.", vbInformation, "Cantidad"
                VSFG.TextMatrix(Row, 4) = VSFG.TextMatrix(Row, 7)
            End If
        End If
        'Verifica que el precio de venta del producto no sea menor al costo
        If Val(VSFG.TextMatrix(Row, 5)) < Val(VSFG.TextMatrix(Row, 9)) And tipoPed <> 1 Then
            If MsgBox("El precio mínimo de venta de este producto es: " & VSFG.TextMatrix(Row, 9) & vbNewLine & vbNewLine & "Desea Factrurar a otro precio?", vbQuestion + vbYesNo, "Precio") = vbYes Then
                frmClave.dblPrecio = Val(VSFG.TextMatrix(Row, 5))
                frmClave.Show vbModal
                If frmClave.Ret = False Then
                    VSFG.TextMatrix(Row, 5) = Format(VSFG.TextMatrix(Row, 9), "####0.00")
                VSFG.TextMatrix(Row, 6) = VSFG.TextMatrix(Row, 5) * VSFG.TextMatrix(Row, 4)
                End If
            Else
                VSFG.TextMatrix(Row, 5) = Format(VSFG.TextMatrix(Row, 9), "####0.00")
                VSFG.TextMatrix(Row, 6) = VSFG.TextMatrix(Row, 5) * VSFG.TextMatrix(Row, 4)
            End If
        End If
        'Actualiza el total del producto pedido
        VSFG.TextMatrix(Row, 6) = Val(VSFG.TextMatrix(Row, 5)) * Val(VSFG.TextMatrix(Row, 4))
        CalcuTotal
    End If
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid en caso de ser necesario
    If VSFG.Rows > 1 And OldRow <= VSFG.Rows - 1 Then
        If OldCol = 5 And OldRow = VSFG.Rows - 1 And NewCol = 6 And VSFG.TextMatrix(OldRow, 2) <> "" Then
            VSFG.AddItem ""
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, 7)) = 1 And Col <> 5 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol
    If r <= 0 Then Exit Sub
    If Val(VSFG.TextMatrix(r, 7)) = 1 Then Exit Sub
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         Dim h As Long
         VSFG.RemoveItem (r)
         For h = 1 To VSFGGuia.Rows - 1
            If VSFGGuia.TextMatrix(h, 12) > r Then
                VSFGGuia.TextMatrix(h, 12) = VSFGGuia.TextMatrix(h, 12) - 1
            End If
         Next h
         PonerBotonesFac
         CalcuTotal
    Else
        VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    'No permite entrar en las celdas de las columnas siguientes
    If NewCol = 3 Or NewCol = 6 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Coloca la descripción del producto en caso que se haga un pedido manual y el usuario haya seleccionado un código de producto
    If Col = 1 Or Col = 2 Then
        If VSFG.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero una bodega", vbInformation, "Bodega"
            VSFG.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
        'Verifica que no se seleccione más de una vez el mismo producto en la misma bodega
'        For i = 1 To VSFG.Rows - 1
'            If VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(i, 2) And VSFG.TextMatrix(Row, 1) = VSFG.TextMatrix(i, 1) And Row <> i Then
'                MsgBox "Ese producto ya fue seleccionado en la bodega " & VSFG.TextMatrix(i, 2) & ", solo cambie la candidad del mismo.", vbInformation, "Producto"
'                VSFG.RemoveItem Row
'                PonerBotones
'                VSFG.Row = i
'                VSFG.Col = 2
'                Exit Sub
'            End If
'        Next i
        'Coloca los datos de un producto seleccionado
        If VSFG.TextMatrix(Row, 2) <> "" Then
            'Busca el producto seleccionado y coloca sus datos respectivos
            clsLstPrds.adorec_Def.MoveFirst
            clsLstPrds.Filtrar "dep_codigo='" & VSFG.TextMatrix(Row, 1) & "' AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'"
            If Not clsLstPrds.adorec_Def.EOF Then
                VSFG.TextMatrix(Row, 3) = clsLstPrds.adorec_Def("prd_nombre")
                'Coloca el costo del producto en una columna oculta
                VSFG.TextMatrix(Row, 9) = clsLstPrds.adorec_Def("prd_costo")
                VSFG.TextMatrix(Row, 7) = 0
                'Verifica que el precio de la lista no sea menor al costo del producto y tampoco sea una cotización
                If clsLstPrds.adorec_Def("prd_costo") > clsLstPrds.adorec_Def("lis_pre_p_precio") And tipoPed <> 1 Then
                    VSFG.TextMatrix(Row, 5) = Format(clsLstPrds.adorec_Def("prd_costo"), "####0.00")
                Else
                    VSFG.TextMatrix(Row, 5) = Format(clsLstPrds.adorec_Def("lis_pre_p_precio"), "####0.00")
                End If
                'Verifica que la existencia del producto sea mayor que cero
                If clsLstPrds.adorec_Def("exi_cantidad") > 0 Then
                    VSFG.TextMatrix(Row, 4) = 1
                Else
                    VSFG.TextMatrix(Row, 4) = 0
                End If
                VSFG.TextMatrix(Row, 6) = VSFG.TextMatrix(Row, 5) * VSFG.TextMatrix(Row, 4)
                VSFG.TextMatrix(Row, 8) = clsLstPrds.adorec_Def("exi_cantidad")
                VSFG.TextMatrix(Row, 10) = clsLstPrds.adorec_Def("prd_iva")
            End If
            clsLstPrds.QuitarFiltro
            CalcuTotal
        End If
    End If
    If Col = 5 Then
        CalcuTotal
    End If
End Sub

Private Sub VSFGGuia_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'If NewCol = 0 Or ((NewCol = 9 Or NewCol = 10) And Abs(VSFGGuia.TextMatrix(NewRow, 0)) = 1) Then
        VSFGGuia.Editable = flexEDKbdMouse
    'Else
    '    VSFGGuia.Editable = flexEDNone
    'End If
End Sub

Private Sub VSFGGuia_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 9 Then
        If VSFGGuia.TextMatrix(Row, Col) = "" Then
            VSFGGuia.TextMatrix(Row, Col) = 0
        End If
    ElseIf Col = 10 Then
        If VSFGGuia.TextMatrix(Row, Col) = "" Then
            VSFGGuia.TextMatrix(Row, Col) = 0
        End If
    End If
    If VSFGGuia.Tag <> "A" Then
        If Col = 0 And Row > 0 Then
            If Abs(VSFGGuia.TextMatrix(Row, 0)) = 1 Then
                VSFGGuia.Select Row, 0, Row, 13
                VSFGGuia.FillStyle = flexFillRepeat
                VSFGGuia.CellBackColor = &HC0FFFF
                VSFGGuia.Select Row, 0
            ElseIf Abs(VSFGGuia.TextMatrix(Row, 0)) = 0 Then
                VSFGGuia.Select Row, 0, Row, 13
                VSFGGuia.FillStyle = flexFillRepeat
                VSFGGuia.CellBackColor = &HFFFFFF
                
                If FormatoD2(VSFGGuia.TextMatrix(Row, 11)) > 0 Then
                    For i = 1 To VSFGGuia.Rows - 1
                        If VSFGGuia.TextMatrix(i, 12) > VSFGGuia.TextMatrix(Row, 12) Then
                            VSFGGuia.TextMatrix(i, 12) = VSFGGuia.TextMatrix(i, 12) - 1
                        End If
                    Next i
                    VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                End If
                VSFGGuia.TextMatrix(Row, 12) = 0
                VSFGGuia.TextMatrix(Row, 11) = 0
                VSFGGuia.TextMatrix(Row, 10) = 0
                VSFGGuia.TextMatrix(Row, 9) = 0
                VSFGGuia.Select Row, 0
            End If
        ElseIf (Col = 9 Or Col = 10) And Row > 0 And VSFGGuia.TextMatrix(Row, 9) <> "" And VSFGGuia.TextMatrix(Row, 10) <> "" Then
            If CDbl(VSFGGuia.TextMatrix(Row, 9)) + CDbl(VSFGGuia.TextMatrix(Row, 10)) > CDbl(VSFGGuia.TextMatrix(Row, 7)) Or CDbl(VSFGGuia.TextMatrix(Row, Col)) < 0 Then
                MsgBox "La cantidad debe mayor a 0 y menor a " & VSFGGuia.TextMatrix(Row, 7) - VSFGGuia.TextMatrix(Row, IIf(Col = 10, 9, 10)), vbCritical, "ERROR"
                VSFGGuia.TextMatrix(Row, Col) = VSFGGuia.TextMatrix(Row, Col) - 1
                If CDbl(VSFGGuia.TextMatrix(Row, 11)) > 0 Then
                    For i = 1 To VSFGGuia.Rows - 1
                        If VSFGGuia.TextMatrix(i, 12) > VSFGGuia.TextMatrix(Row, 12) Then
                            VSFGGuia.TextMatrix(i, 12) = VSFGGuia.TextMatrix(i, 12) - 1
                        End If
                    Next i
                    VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                    VSFGGuia.TextMatrix(Row, 12) = 0
                End If
            End If
            VSFGGuia.TextMatrix(Row, 11) = VSFGGuia.TextMatrix(Row, 8) * VSFGGuia.TextMatrix(Row, 10)
            If Col = 10 And VSFGGuia.TextMatrix(Row, 12) <> "" Then
                If Val(VSFGGuia.TextMatrix(Row, 12)) = 0 And FormatoD2(VSFGGuia.TextMatrix(Row, 10)) > 0 Then
                    If VSFG.Rows > 1 Then
                        If VSFG.TextMatrix(VSFG.Rows - 1, 2) = "" Then
                            VSFG.RemoveItem VSFG.Rows - 1
                        End If
                    End If
                    VSFG.AddItem vbTab & VSFGGuia.TextMatrix(Row, 3) & vbTab & VSFGGuia.TextMatrix(Row, 4) & vbTab & VSFGGuia.TextMatrix(Row, 5) & vbTab & VSFGGuia.TextMatrix(Row, 10) & vbTab & VSFGGuia.TextMatrix(Row, 8) & vbTab & VSFGGuia.TextMatrix(Row, 11) & vbTab & 1 & vbTab & vbTab & VSFGGuia.TextMatrix(Row, 14) & vbTab & VSFGGuia.TextMatrix(Row, 16)
                    VSFGGuia.TextMatrix(Row, 12) = VSFG.Rows - 1
                ElseIf Val(VSFGGuia.TextMatrix(Row, 12)) > 0 Then
                    If CDbl(VSFGGuia.TextMatrix(Row, 10)) = 0 Then
                        For i = 1 To VSFGGuia.Rows - 1
                            If CDbl(VSFGGuia.TextMatrix(i, 12)) > CDbl(VSFGGuia.TextMatrix(Row, 12)) Then
                                VSFGGuia.TextMatrix(i, 12) = CDbl(VSFGGuia.TextMatrix(i, 12)) - 1
                            End If
                        Next i
                        VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                        VSFGGuia.TextMatrix(Row, 12) = 0
                    Else
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 4) = VSFGGuia.TextMatrix(Row, 10)
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 5) = VSFGGuia.TextMatrix(Row, 8)
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 6) = VSFGGuia.TextMatrix(Row, 8) * VSFGGuia.TextMatrix(Row, 10)
                    End If
                End If
            End If
        End If
'        CalcuTotal
    End If
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
             CalcuReca
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
        CalcuReca
    
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'Verifica que solo se ingresen números en el grid de recargos en caso de ser necesario
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CargarGuias()
    Dim strSql As String
    
    strSql = " SELECT '0',egreso.egr_codigo,egr_factura,det_egreso.dep_codigo,det_egreso.prd_codigo,producto.prd_nombre,det_egr_cantidad,det_egr_cantidad - COALESCE(sum(det_ing_cantidad),0),det_egr_precio,'0' as devo,'0' as fact,'0.00' as tota,'0' as linea,egr_observacion,(producto.prd_costo/(1 - IIF(producto.prd_codigo='BEL1583A',0.0715,0.15))) as prd_costo,det_egr_costo,prd_iva " & _
             " FROM (((egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo AND egreso.emp_codigo=det_egreso.emp_codigo AND egreso.egr_anulado=0) " & _
             " INNER JOIN producto ON det_egreso.prd_codigo=producto.prd_codigo AND det_egreso.emp_codigo=producto.emp_codigo) " & _
             " LEFT JOIN ingreso ON CAST(egreso.egr_codigo as varchar)=ingreso.ing_factura AND  egreso.emp_codigo=ingreso.emp_codigo AND ingreso.tip_ing_codigo='" & strTipIng & "' AND ingreso.ing_anulado=0) " & _
             " LEFT JOIN  det_ingreso ON ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.emp_codigo=det_ingreso.emp_codigo AND det_egreso.prd_codigo=det_ingreso.prd_codigo " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egreso.per_codigo='" & cmbCliente.BoundText & "' " & _
             " AND egreso.tip_egr_codigo='" & strTipEgr & "' " & _
             " GROUP BY egreso.egr_codigo,egr_factura,det_egreso.dep_codigo,det_egreso.prd_codigo,producto.prd_codigo,producto.prd_nombre,det_egr_cantidad,det_egr_precio,egr_observacion,producto.prd_costo,det_egr_costo,prd_iva " & _
             " HAVING det_egr_cantidad - COALESCE(sum(det_ing_cantidad),0)>0 "
    clsSql.Ejecutar strSql
    'VSFGGuia.AllowUserResizing = flexResizeNone
    VSFGGuia.Tag = "A"
    Set VSFGGuia.DataSource = clsSql.adorec_Def.DataSource
    VSFGGuia.Tag = ""
    'VSFGGuia.AllowUserResizing = flexResizeColumns
End Sub
Private Sub FacturarE()
    Dim prdFactura() As Variant
    Dim i As Long
    Dim j As Long
    Dim maxj As Long
    Dim booBandera As Boolean
    Dim booPasar As Boolean
    Dim booGuardar As Boolean
    Dim booPed As Boolean
    Dim codEgr As Double
    Dim strNumero As String
    Dim strSql As String
    Dim codVen As String
    Dim CodPer As String
    Dim strGuia As String
    Dim strGuias As String
    
    Dim GuiaAutomatica As Boolean
    
    Dim clsAsiento As New clsContable
    Dim clsEgreso As New clsInventario
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    
    If VSFG.TextMatrix(VSFG.Rows - 1, 4) = "" And VSFG.TextMatrix(VSFG.Rows - 1, 5) = "" Then
        VSFG.RemoveItem VSFG.Rows - 1
    End If
    strGuias = ""
'    For i = 1 To VSFGGuia.Rows - 1
'        If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 2) Then
'            If VSFGGuia.TextMatrix(i, 2) = "" Then
'                strGuias = strGuias & VSFGGuia.TextMatrix(i, 1) & "/"
'                strGuia = VSFGGuia.TextMatrix(i, 1)
'            Else
'                strGuias = strGuias & VSFGGuia.TextMatrix(i, 2) & "/"
'                strGuia = VSFGGuia.TextMatrix(i, 2)
'            End If
'        End If
'    Next i
    For i = 1 To VSFGGuia.Rows - 1
        If VSFGGuia.TextMatrix(i, 2) <> "" Then
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 2) Then
                strGuias = strGuias & VSFGGuia.TextMatrix(i, 2) & "/"
                strGuia = VSFGGuia.TextMatrix(i, 2)
            End If
        Else
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 1) Then
                strGuias = strGuias & VSFGGuia.TextMatrix(i, 1) & "/"
                strGuia = VSFGGuia.TextMatrix(i, 1)
            End If
        End If
    Next i
    TxtObserv = "Guia:" & strGuias & "  - " & TxtObserv
    'compactar en matriz la factura
    maxj = 0
    booPed = False
    ReDim prdFactura(4, maxj)
    prdFactura(0, maxj) = VSFG.TextMatrix(1, 1)
    prdFactura(1, maxj) = VSFG.TextMatrix(1, 2)
    prdFactura(2, maxj) = CDbl(VSFG.TextMatrix(1, 4))
    prdFactura(3, maxj) = CDbl(VSFG.TextMatrix(1, 5))
    If VSFG.TextMatrix(1, 7) <> 1 Then
        prdFactura(4, maxj) = CDbl(VSFG.TextMatrix(1, 4))
        booPed = True
    Else
        prdFactura(4, maxj) = 0
    End If
    
    For i = 2 To VSFG.Rows - 1
        booBandera = False
        For j = 0 To maxj
            ' si encontro repetido
            If prdFactura(0, j) = VSFG.TextMatrix(i, 1) And prdFactura(1, j) = VSFG.TextMatrix(i, 2) Then
                ' prcio promedio para no afectar total
                prdFactura(3, j) = (CDbl(VSFG.TextMatrix(i, 5)) * CDbl(VSFG.TextMatrix(i, 4)) + prdFactura(3, j) * prdFactura(2, j)) / (CDbl(prdFactura(2, j)) + CDbl(VSFG.TextMatrix(i, 4)))
                prdFactura(2, j) = CDbl(prdFactura(2, j)) + CDbl(VSFG.TextMatrix(i, 4))
                If VSFG.TextMatrix(i, 7) <> 1 Then
                    prdFactura(4, j) = CDbl(prdFactura(4, j)) + CDbl(VSFG.TextMatrix(i, 4))
                    booPed = True
                End If
                booBandera = True
                Exit For
            End If
        Next j
        'no encontro igual
        If booBandera = False Then
            'inserta en matriz item para facturar
            maxj = maxj + 1
            ReDim Preserve prdFactura(4, maxj)
            prdFactura(0, maxj) = VSFG.TextMatrix(i, 1)
            prdFactura(1, maxj) = VSFG.TextMatrix(i, 2)
            prdFactura(2, maxj) = CDbl(VSFG.TextMatrix(i, 4))
            prdFactura(3, maxj) = CDbl(VSFG.TextMatrix(i, 5))
            If VSFG.TextMatrix(i, 7) <> 1 Then
                prdFactura(4, maxj) = CDbl(VSFG.TextMatrix(i, 4))
                booPed = True
            Else
                prdFactura(4, maxj) = 0
            End If
        End If
    Next i
    
    'manda a la base de datos
'****** EGRESO
    'Genera un egreso de mercadería
    'Obtiene el código con el que se debe insertar el nuevo egreso de productos
    codVen = cmbVendedor.BoundText
    CodPer = cmbCliente2.BoundText
    booGuardar = clsEgreso.NuevoEgr(True, "FAC", True, , , , Me.CmbFpago.BoundText, CodPer, Format(dtpFecha.Value, "yyyy-MM-dd"), , codVen, TxtObserv, , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotal), FormatoD2(TxtRecargo), FormatoD2(TxtDesc), FormatoD2(TxtIva), FormatoD2(TxtTotal), 0, SecPublico, SinIVA, CodigoIVA)
    If booGuardar = True Then
        'Inserta la cabecera del egreso
        codEgr = clsEgreso.strDoc
        clsAsiento.Inicializar AdoConn, AdoConnMaster
        clsAsiento.NuevoAsiento "F", dtpFecha.Value, 0, 0, TxtTotal.Text, "FACTURA " & codEgr
        
        clsEgreso.ModificaEgr , , , , , , clsAsiento.NumAsiento
    '****** CABECERA PEDIDO
        'Obtiene el código del pedido a ingresar
        If booPed = True Then
            Dim num As Double
            strSql = " LOCK TABLES pedido WRITE "
            clsSql.Ejecutar strSql, "M"
            strSql = " Select COALESCE(max(ped_codigo),0) as num " & _
                     " From pedido " & _
                     " Where emp_codigo='" & strEmpresa & "' AND ped_codigo like '" & strSucursal + 0 & "%'" & _
                     " GROUP BY emp_codigo"
            clsSql.Ejecutar (strSql), "M"
            num = clsSql.adorec_Def("num") + 1
            'Inserta la cabecera del pedido
            strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo, ped_fecha, " & _
                     " ped_estado, ped_subtotal, ped_observacion,ped_tip_egr_codigo,ped_egr_codigo, tipo_fac_codigo, ped_fechamod, ped_usumod) " & _
                     " VALUES ('" & strEmpresa & "'," & num & ",'" & CodPer & "','" & codVen & "', " & _
                     " '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "',2,'" & Format(TxtSubTotal, "####0.00") & "','COMPLEMENTO A GUIAS " & strGuias & "', " & _
                     " 'FAC','" & codEgr & "',1,CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsSql.Ejecutar (strSql), "M"
            strSql = " UNLOCK TABLES"
            clsSql.Ejecutar (strSql), "M"
        End If
        
        For j = 0 To maxj
            strSql = " SELECT prd_costo " & _
                     " FROM producto " & _
                     " WHERE prd_codigo='" & prdFactura(1, j) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql)
            'Inserta los detalles de egreso
            clsEgreso.NuevoDetEgr (prdFactura(1, j)), (prdFactura(0, j)), (prdFactura(2, j)), (prdFactura(3, j)), clsSql.adorec_Def("prd_costo"), 0
            'Inserta los detalles del pedido
            If prdFactura(4, j) > 0 And booPed = True Then
                strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                         " det_ped_cant_entregada, det_ped_precio, det_ped_fechamod, det_ped_usumod) " & _
                         " VALUES ('" & strEmpresa & "'," & num & ",'" & prdFactura(1, j) & "','" & prdFactura(0, j) & "'," & prdFactura(2, j) & ", " & _
                         prdFactura(2, j) & "," & Format(prdFactura(3, j), "####.0000") & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsSql.Ejecutar (strSql), "M"
            End If
        Next j
    '****** RETENCIONES
        clsEgreso.DetRetenciones
    '****** RECARGOS
        'Genera los posibles recargos que podujo esta factura
        For i = 1 To VSFGReca.Rows - 1
            If VSFGReca.TextMatrix(i, 1) <> "" Then
                clsEgreso.NuevoDetEgrRecargo VSFGReca.TextMatrix(i, 1), FormatoD2(VSFGReca.TextMatrix(i, 3))
            End If
        Next i
        clsFPago.adorec_Def.MoveFirst
        strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
        clsFPago.adorec_Def.Find strComparar
        'Recupera el nuevo código con el cual se debe ingresar la nueva cuenta por cobrar
        Dim clsCta As New clsCtaXx
        clsCta.Inicializar AdoConn, AdoConnMaster
        clsCta.NuevaCta "C", "1", "00", Format(dtpFecha.Value, "yyyy-mm-dd"), Format(DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha.Value), "yyyy-MM-dd"), CodPer, "Factura # " & codEgr & " - " & TxtObserv, strSucursal & strPtoFactura, Right(codEgr, 7), strAutorFactura, strCaducaFactura, clsEgreso.dblTotalProd, clsEgreso.dblTotalServ, clsEgreso.dblTotalProdIVA, clsEgreso.dblTotalServIVA, 2, clsEgreso.dblIVA, clsEgreso.dblSubTotal0, 0, 0, 0, clsEgreso.dblTotal, clsAsiento.NumAsiento
        clsCta.IngAsientoEgr clsAsiento, clsEgreso
        Set clsCta = Nothing
        Set clsAsiento = Nothing
        
    End If
    'Escribe la Factura en la guia
    strGuia = ""
    For i = 1 To VSFGGuia.Rows - 1
        If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 2) Then
            strSql = " UPDATE egreso " & _
                     " SET egr_observacion = CONCAT('F:" & codEgr & " / ',egr_observacion) " & _
                     " WHERE tip_egr_codigo='" & strTipEgr & "' and egr_codigo='" & VSFGGuia.TextMatrix(i, 1) & _
                     "' AND egr_codigo='" & VSFGGuia.TextMatrix(i, 1) & "'"
            clsSql.Ejecutar strSql, "M"
            strGuia = VSFGGuia.TextMatrix(i, 2)
        End If
    Next i
    
    If booGuardar = True Then
        Dim rptFa As New frmReporte
        rptFa.strNumero = clsEgreso.strDoc
        'no usa
        GuiaAutomatica = False
        rptFa.strReporte = IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola")
        rptFa.Show
    End If
End Sub
Private Sub BajarE()
    Dim prdFactura() As Variant
    Dim i As Long
    Dim j As Long
    Dim maxj As Long
    Dim booBandera As Boolean
    Dim booPasar As Boolean
    Dim booGuardar As Boolean
    Dim booPed As Boolean
    Dim codEgr As Double
    Dim strNumero As String
    Dim strSql As String
    Dim codVen As String
    Dim CodPer As String
    Dim strGuia As String
    Dim strGuias As String
    Dim clsAsiento As New clsContable
    Dim clsEgreso As New clsInventario
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    
    If VSFG.TextMatrix(VSFG.Rows - 1, 4) = "" And VSFG.TextMatrix(VSFG.Rows - 1, 5) = "" Then
        VSFG.RemoveItem VSFG.Rows - 1
    End If
    strGuias = ""
    For i = 1 To VSFGGuia.Rows - 1
        If VSFGGuia.TextMatrix(i, 2) <> "" Then
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 2) Then
                strGuias = strGuias & VSFGGuia.TextMatrix(i, 2) & "/"
                strGuia = VSFGGuia.TextMatrix(i, 2)
            End If
        Else
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 1) Then
                strGuias = strGuias & VSFGGuia.TextMatrix(i, 1) & "/"
                strGuia = VSFGGuia.TextMatrix(i, 1)
            End If
        End If
    Next i
    TxtObserv = "Guia:" & strGuias & "  - " & TxtObserv
    'compactar en matriz la factura
    maxj = 0
    booPed = False
    ReDim prdFactura(4, maxj)
    prdFactura(0, maxj) = VSFG.TextMatrix(1, 1)
    prdFactura(1, maxj) = VSFG.TextMatrix(1, 2)
    prdFactura(2, maxj) = CDbl(VSFG.TextMatrix(1, 4))
    prdFactura(3, maxj) = CDbl(VSFG.TextMatrix(1, 5))
    If VSFG.TextMatrix(1, 7) <> 1 Then
        prdFactura(4, maxj) = CDbl(VSFG.TextMatrix(1, 4))
        booPed = True
    Else
        prdFactura(4, maxj) = 0
    End If
    
    For i = 2 To VSFG.Rows - 1
        booBandera = False
        For j = 0 To maxj
            ' si encontro repetido
            If prdFactura(0, j) = VSFG.TextMatrix(i, 1) And prdFactura(1, j) = VSFG.TextMatrix(i, 2) Then
                ' prcio promedio para no afectar total
                prdFactura(3, j) = (CDbl(VSFG.TextMatrix(i, 5)) * CDbl(VSFG.TextMatrix(i, 4)) + prdFactura(3, j) * prdFactura(2, j)) / (CDbl(prdFactura(2, j)) + CDbl(VSFG.TextMatrix(i, 4)))
                prdFactura(2, j) = CDbl(prdFactura(2, j)) + CDbl(VSFG.TextMatrix(i, 4))
                If VSFG.TextMatrix(i, 7) <> 1 Then
                    prdFactura(4, j) = CDbl(prdFactura(4, j)) + CDbl(VSFG.TextMatrix(i, 4))
                    booPed = True
                End If
                booBandera = True
                Exit For
            End If
        Next j
        'no encontro igual
        If booBandera = False Then
            'inserta en matriz item para facturar
            maxj = maxj + 1
            ReDim Preserve prdFactura(4, maxj)
            prdFactura(0, maxj) = VSFG.TextMatrix(i, 1)
            prdFactura(1, maxj) = VSFG.TextMatrix(i, 2)
            prdFactura(2, maxj) = CDbl(VSFG.TextMatrix(i, 4))
            prdFactura(3, maxj) = CDbl(VSFG.TextMatrix(i, 5))
            If VSFG.TextMatrix(i, 7) <> 1 Then
                prdFactura(4, maxj) = CDbl(VSFG.TextMatrix(i, 4))
                booPed = True
            Else
                prdFactura(4, maxj) = 0
            End If
        End If
    Next i
    
    'manda a la base de datos
'****** EGRESO
    'Genera un egreso de mercadería
    'Obtiene el código con el que se debe insertar el nuevo egreso de productos
    codVen = cmbVendedor.BoundText
    CodPer = cmbCliente2.BoundText
    booGuardar = clsEgreso.NuevoEgr(True, "BAU", True, , , , Me.CmbFpago.BoundText, CodPer, Format(dtpFecha.Value, "yyyy-MM-dd"), , codVen, TxtObserv, , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotal), FormatoD2(TxtRecargo), FormatoD2(TxtDesc), FormatoD2(TxtIva), FormatoD2(TxtTotal), 0, SecPublico, SinIVA)
    If booGuardar = True Then
        'Inserta la cabecera del egreso
        codEgr = clsEgreso.strDoc
    '****** CABECERA PEDIDO
        'Obtiene el código del pedido a ingresar
        If booPed = True Then
            Dim num As Double
            strSql = " LOCK TABLES pedido WRITE "
            clsSql.Ejecutar strSql, "M"
            strSql = " Select COALESCE(max(ped_codigo),0) as num " & _
                     " From pedido " & _
                     " Where emp_codigo='" & strEmpresa & "' AND ped_codigo like '" & strSucursal + 0 & "%'" & _
                     " GROUP BY emp_codigo"
            clsSql.Ejecutar (strSql), "M"
            num = clsSql.adorec_Def("num") + 1
            'Inserta la cabecera del pedido
            strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo, ped_fecha, " & _
                     " ped_estado, ped_subtotal, ped_observacion,ped_tip_egr_codigo,ped_egr_codigo, tipo_fac_codigo, ped_fechamod, ped_usumod) " & _
                     " VALUES ('" & strEmpresa & "'," & num & ",'" & CodPer & "','" & codVen & "', " & _
                     " '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "',2,'" & Format(TxtSubTotal, "####0.00") & "','COMPLEMENTO A GUIAS " & strGuias & "', " & _
                     " 'FAC','" & codEgr & "',1,CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsSql.Ejecutar (strSql), "M"
            strSql = " UNLOCK TABLES"
            clsSql.Ejecutar (strSql), "M"
        End If
        
        For j = 0 To maxj
            strSql = " SELECT prd_costo " & _
                     " FROM producto " & _
                     " WHERE prd_codigo='" & prdFactura(1, j) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql)
            'Inserta los detalles de egreso
            clsEgreso.NuevoDetEgr (prdFactura(1, j)), (prdFactura(0, j)), (prdFactura(2, j)), (prdFactura(3, j)), clsSql.adorec_Def("prd_costo"), 0
            'Inserta los detalles del pedido
            If prdFactura(4, j) > 0 And booPed = True Then
                strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                         " det_ped_cant_entregada, det_ped_precio, det_ped_fechamod, det_ped_usumod) " & _
                         " VALUES ('" & strEmpresa & "'," & num & ",'" & prdFactura(1, j) & "','" & prdFactura(0, j) & "'," & prdFactura(2, j) & ", " & _
                         prdFactura(2, j) & "," & Format(prdFactura(3, j), "####.0000") & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsSql.Ejecutar (strSql), "M"
            End If
        Next j
        
    End If
    'Escribe la Factura en la guia
    strGuia = ""
    For i = 1 To VSFGGuia.Rows - 1
        If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 And VSFGGuia.TextMatrix(i, 10) <> 0 And strGuia <> VSFGGuia.TextMatrix(i, 2) Then
            strSql = " UPDATE egreso " & _
                     " SET egr_observacion = CONCAT('B:" & codEgr & " / ',egr_observacion) " & _
                     " WHERE tip_egr_codigo='" & strTipEgr & "' and egr_codigo='" & VSFGGuia.TextMatrix(i, 1) & _
                     "' AND egr_codigo='" & VSFGGuia.TextMatrix(i, 1) & "'"
            clsSql.Ejecutar strSql, "M"
            strGuia = VSFGGuia.TextMatrix(i, 2)
        End If
    Next i
    
    If booGuardar = True Then
        Dim rptFa As New frmReporte
        
        
        rptFa.strNumero = clsEgreso.strDoc
        rptFa.strTipo = "BAU"
        rptFa.strReporte = "rptEgresoMercaderia"
        rptFa.Show
        
    End If
End Sub
Private Function Devolver() As Boolean
    Dim i As Integer
    Dim guia_actual As Double
    Dim ultima_guia As Double
    Dim numero_ingreso As Double
    Dim strNDev As String
    Dim operacion As Boolean
    Dim MAguia As String
    Dim ex As Boolean
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim rpMov As New frmReporte
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    ex = False
    operacion = False
    strContenedorRecurrente = "111"
    If VSFGGuia.Tag <> "A" And MsgBox("Desea hacer la devolución de todas las guias seleccionadas?", vbYesNo + vbQuestion, "Devolución de Guias") = vbYes Then
        For i = 1 To VSFGGuia.Rows - 1
            'Verifica que la casilla esté seleccionada
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 Then
                'Verifica si hay un valor en el campo Devolución
                If (Val(VSFGGuia.TextMatrix(i, 9)) > 0 And VSFGGuia.TextMatrix(i, 9) <> "") Then
                    guia_actual = Val(VSFGGuia.TextMatrix(i, 1))
                    operacion = True
                    If (guia_actual = ultima_guia) Then
                        'Añadir DET_INGRESO en último INGRESO
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 9)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                    Else
                        If ex = True Then
                            rpMov.strNumero = clsIngreso.strDoc
                            rpMov.strTipo = clsIngreso.strTipo
                            rpMov.strReporte = "rptIngresoMercaderia"
                            rpMov.Show
                        End If
                        ex = True
                        'Crear nuevo INGRESO con su DET_INGRESO
                        'Obtiene el código con el que se debe insertar el nuevo ingreso
                        booGuardar = clsIngreso.NuevoIng(True, strTipIng, False, strSucursal, strPtoFactura, , , cmbCliente.BoundText, Format(dtpFecha.Value, "yyyy-MM-dd"), (guia_actual), cmbVendedor.BoundText, TxtObserv)
                        numero_ingreso = clsIngreso.strDoc
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 9)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                        MAguia = "0"
                        While MAguia = "0"
                            MAguia = guia_actual
                            If Trim(MAguia) = "" Or Trim(MAguia) = "0" Then
                                MAguia = "0"
                            End If
                        Wend
                        strSql = " SELECT CONCAT('" & UCase(MAguia) & "',' - ',egr_observacion) as obs " & _
                                 " FROM egreso " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND tip_egr_codigo='" & strTipEgr & "' AND egr_codigo='" & guia_actual & "' "
                        clsSql.Ejecutar (strSql)
                        clsEgreso.strTipo = strTipEgr
                        clsEgreso.strDoc = guia_actual
                        clsEgreso.strIE = "E"
                        clsEgreso.strFecha = Format(dtpFecha.Value, "yyyy-MM-dd")
                        clsEgreso.ModificaEgr , , , , , clsSql.adorec_Def("obs")
                    End If
                    ultima_guia = guia_actual
                End If
            End If
        Next i
        InicializarContenedorRecurrente
        If operacion = True Then
            rpMov.strNumero = clsIngreso.strDoc
            rpMov.strTipo = clsIngreso.strTipo
            rpMov.strReporte = "rptIngresoMercaderia"
            rpMov.Show
            MsgBox "Devolución realizada con éxito", vbInformation
        End If
    End If
    Devolver = operacion
End Function

Private Function Facturar() As Boolean
    Dim i As Integer
    Dim guia_actual As Double
    Dim ultima_guia As Double
    Dim numero_ingreso As Double
    Dim strNDev As String
    Dim operacion As Boolean
    Dim clsIngreso As New clsInventario
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    
    operacion = False
    If VSFGGuia.Tag <> "A" And MsgBox("Esta seguro de realizar la Facturación?", vbYesNo + vbQuestion, "Facturación") = vbYes Then
        For i = 1 To VSFGGuia.Rows - 1
            'Verifica que la casilla esté seleccionada
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 Then
                'Verifica si hay un valor en el campo Facturar
                If (FormatoD2(VSFGGuia.TextMatrix(i, 10)) > 0 And VSFGGuia.TextMatrix(i, 10) <> "") Then
                    guia_actual = Val(VSFGGuia.TextMatrix(i, 1))
                    operacion = True
                    If (guia_actual = ultima_guia) Then
                        'Añadir DET_INGRESO en último INGRESO
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 10)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                    Else
                        'Crear nuevo INGRESO con su DET_INGRESO
                        'Obtiene el código con el que se debe insertar el nuevo ingreso
                        'Inserta la cabecera del ingreso
                        booGuardar = clsIngreso.NuevoIng(True, strTipIng, False, strSucursal, strPtoFactura, , , cmbCliente.BoundText, Format(dtpFecha.Value, "yyyy-MM-dd"), (guia_actual), cmbVendedor.BoundText, TxtObserv)
                        numero_ingreso = clsIngreso.strDoc
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 10)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                    End If
                    ultima_guia = guia_actual
                End If
            End If
        Next i
        InicializarContenedorRecurrente
        If operacion = True Then
            'MsgBox "Facturación realizada con éxito", vbInformation
        End If
    End If
    Facturar = operacion
End Function
Private Function Bajar() As Boolean
    Dim i As Integer
    Dim guia_actual As Double
    Dim ultima_guia As Double
    Dim numero_ingreso As Double
    Dim strNDev As String
    Dim operacion As Boolean
    Dim clsIngreso As New clsInventario
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    
    operacion = False
    If VSFGGuia.Tag <> "A" And MsgBox("Esta seguro de realizar la BAJA?", vbYesNo + vbQuestion, "BAJA") = vbYes Then
        For i = 1 To VSFGGuia.Rows - 1
            'Verifica que la casilla esté seleccionada
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 Then
                'Verifica si hay un valor en el campo Facturar
                If (FormatoD2(VSFGGuia.TextMatrix(i, 10)) > 0 And VSFGGuia.TextMatrix(i, 10) <> "") Then
                    guia_actual = Val(VSFGGuia.TextMatrix(i, 1))
                    operacion = True
                    If (guia_actual = ultima_guia) Then
                        'Añadir DET_INGRESO en último INGRESO
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 10)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                    Else
                        'Crear nuevo INGRESO con su DET_INGRESO
                        'Obtiene el código con el que se debe insertar el nuevo ingreso
                        'Inserta la cabecera del ingreso
                        booGuardar = clsIngreso.NuevoIng(True, strTipIng, False, strSucursal, strPtoFactura, , , cmbCliente.BoundText, Format(dtpFecha.Value, "yyyy-MM-dd"), (guia_actual), cmbVendedor.BoundText, TxtObserv)
                        numero_ingreso = clsIngreso.strDoc
                        'Inserta el detalle de ingreso al proyecto
                        clsIngreso.NuevoDetIng VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD2(VSFGGuia.TextMatrix(i, 10)), FormatoD8(VSFGGuia.TextMatrix(i, 8)), VSFGGuia.TextMatrix(i, 15)
                    End If
                    ultima_guia = guia_actual
                End If
            End If
        Next i
        InicializarContenedorRecurrente
        If operacion = True Then
            'MsgBox "Facturación realizada con éxito", vbInformation
        End If
    End If
    Bajar = operacion
End Function

Private Sub TipoDoc()
    If optGuia.Value = True Then
        strTipEgr = "GRE"
        strTipIng = "DRE"
        frmDoc.Caption = "DATOS DE GUIAS"
    Else
        strTipEgr = "ERE"
        strTipIng = "IRE"
        frmDoc.Caption = "DATOS DE RESERVAS"
    End If
    cmdLimpiar_Click
    cmbCliente_Validate False
End Sub
