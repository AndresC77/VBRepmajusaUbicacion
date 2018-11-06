VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmV_PedBod 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar pedido a Bodega"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_PedBod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   9690
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   3990
      TabIndex        =   40
      Top             =   120
      Width           =   5535
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   255
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
         Left            =   120
         TabIndex        =   41
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos del Cliente"
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
      Left            =   150
      TabIndex        =   23
      Top             =   2640
      Width           =   9495
      Begin VB.CommandButton cmdDirEnvio 
         Caption         =   "Dir. Envio"
         Height          =   285
         Left            =   7440
         TabIndex        =   53
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCredito 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDisponible 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtDcto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtCategoria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtTF 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtDireccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtRuc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
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
         TabIndex        =   39
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
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
         Left            =   7320
         TabIndex        =   38
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dcto:"
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
         Left            =   4335
         TabIndex        =   33
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría Cliente:"
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
         Top             =   990
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telf/Fax:"
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
         TabIndex        =   29
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   660
         TabIndex        =   27
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CI/RUC:"
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
         Left            =   840
         TabIndex        =   25
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo de Pedido:"
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
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3735
      Begin VSFlex8LCtl.VSFlexGrid VSFGTPeds 
         Height          =   525
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3300
         _cx             =   1922111165
         _cy             =   1922106270
         Appearance      =   2
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
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_PedBod.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
   End
   Begin VB.Frame FraDetalle 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle de Pedido"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   9495
      Begin VB.TextBox TxtTotalConIVA 
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
         Height          =   345
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2400
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.TextBox txtComi 
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
         Height          =   360
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
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
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
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
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtTotDcto 
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
         Height          =   345
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtObser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   8040
      End
      Begin VB.TextBox TxtTotal 
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
         Height          =   345
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2055
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   9000
         _cx             =   1913208323
         _cy             =   1913196073
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_PedBod.frx":03F1
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
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   1440
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7455
         TabIndex        =   55
         Top             =   2475
         Width           =   450
      End
      Begin VB.Label lblComi 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6105
         TabIndex        =   46
         Top             =   3240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cantidad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1890
         TabIndex        =   43
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dcto:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5010
         TabIndex        =   35
         Top             =   2835
         Width           =   855
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4800
         TabIndex        =   16
         Top             =   2475
         Width           =   1065
      End
      Begin VB.Label LblObser 
         Alignment       =   1  'Right Justify
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
         Left            =   375
         TabIndex        =   15
         Top             =   3000
         Width           =   1155
      End
   End
   Begin VB.Frame FraPedido 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cotización Nº "
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
      Height          =   1815
      Left            =   720
      TabIndex        =   13
      Top             =   3960
      Width           =   8295
      Begin VB.CommandButton cmdCargaBackCot 
         Caption         =   "Cargar"
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
         Left            =   6000
         TabIndex        =   49
         Top             =   960
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGCot 
         Height          =   1455
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   5655
         _cx             =   1913202423
         _cy             =   1913195014
         Appearance      =   2
         BorderStyle     =   0
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
         Rows            =   5
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_PedBod.frx":057F
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
      Begin MSDataListLib.DataCombo CmbBodega 
         Height          =   330
         Left            =   6000
         TabIndex        =   47
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega de despacho"
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
         TabIndex        =   48
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame FraBotones 
      BackColor       =   &H00DDDDDD&
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   9360
      Width           =   9495
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
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
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   5663
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdPedido 
         Caption         =   "Ejecutar Pedido"
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
         Left            =   2783
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos del Pedido:"
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
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   9495
      Begin VSFlex8Ctl.VSFlexGrid VSFRRed 
         Height          =   1335
         Left            =   5640
         TabIndex        =   52
         Top             =   240
         Width           =   3735
         _cx             =   116070844
         _cy             =   116066611
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Rows            =   9
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_PedBod.frx":0634
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
         Left            =   1170
         TabIndex        =   4
         Top             =   240
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
         Format          =   16711683
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   975
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1305
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo cmbTC 
         Height          =   315
         Left            =   1200
         TabIndex        =   50
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA:"
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
         Left            =   150
         TabIndex        =   51
         Top             =   645
         Width           =   480
      End
      Begin VB.Label LblCliente 
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
         Left            =   150
         TabIndex        =   22
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label2 
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
         Left            =   150
         TabIndex        =   21
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Pedido:"
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
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   8880
      Picture         =   "frmV_PedBod.frx":0695
      ToolTipText     =   "Elimina una Fila"
      Top             =   3120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   9120
      Picture         =   "frmV_PedBod.frx":07CB
      Top             =   3120
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmV_PedBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para generar un pedido a bodega en forma manual, o a partir de una    #
'#  cotización o un backorder.                                                  #
'#  frmV_PedBod V1.0                                                            #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Opciones que permite:                                                       #
'#  *   Se muestra dos combo box en la parte superior en los cuales se puede    #
'#      seleccionar la forma en la que se va a realizar el nuevo pedido a       #
'#      bodega es decir a parti de una contización, un backorder o manual.      #
'#  *   Cuando se selecciona en el primer combo desde una cotización o un back_ #
'#      order se debe seleccionar en el segundo combo el código del mismo       #
'#  *   En el detalle del pedido a generar aparecerá los detalles de la cotiza_ #
'#      ción o el backorder si se eligió uno.                                   #
'#                                                                              #
'#  Procesos internos que maneja:                                               #
'#  *   Se consulta la tabla depósitos para obtener y generar los combos que    #
'#      aparecerán en el grid de detalles.                                      #
'#  *   Se consulta la tabla de productos para obtener y generar los combos     #
'#      que aparecerán en el grid de detalles.                                  #
'#  *   Controla que el usuario pida como máximo la existencia de productos     #
'#      de la bodega seleccionada.                                              #
'#  *   Se controla que el precio del producto pedido no se menor al costo del  #
'#      mismo.                                                                  #
'#  *   Los totales por producto y total del pedido se calculan automáticamente.#
'#                                                                              #
'#  Tablas que maneja:                                                          #
'#                                                                              #
'#  deposito:                                                                   #
'#  *   Se consultan todas las bodegas de productos de una empresa.             #
'#  producto:                                                                   #
'#  *   Se consultan todas los productos de una empresa.                        #
'#  persona:                                                                    #
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el  #
'#      pedido que se está realizando.                                          #
'#  *   También se extrae el nombre del vendedor asignado al pedido a realizar. #
'#  cotizacion:                                                                 #
'#  *   De aquí se consultan los datos que servirán como plantilla para crear   #
'#      un nuevo pedido a partir de una cotizacón.                              #
'#  backOrder:                                                                  #
'#  *   De aquí se consultan los datos que servirán como plantilla para crear   #
'#      un nuevo pedido a partir de un backorder.                               #
'#  pedido:                                                                     #
'#  *   Aquí se almacenan los datos de la cabecera de un nuevo pedido.          #
'#  det_pedido:                                                                 #
'#  *   Aquí se almacenan los datos de los detalles de los productos que in_    #
'#      tervienen en el nuevo pedido, así como su precio y cantidad.            #
'#                                                                              #
'################################################################################

Private clsCots As New clsConsulta
Private clsBods As New clsConsulta
Private clsDet As New clsConsulta
Private clsPrds As New clsConsulta
Private clsLstPrds As New clsConsulta
Private clsClie As New clsConsulta
Private clsSql As New clsConsulta
Private clsBack As New clsConsulta
Private clsFacAnu As New clsConsulta
Private clsTC As New clsConsulta
Private clsSqlNum As New clsConsulta
Private strSql As String
Private strSqlPrd As String
Private intDato As Variant
Private TipoPed As Integer
Private dblComision As Double
Private lngCod As String
Private clsProm As New clsPrePromCot
Private strClaveMAESTRA As String
Private booDcto As Boolean
Public Controlado As Boolean
Private dctoMax As Double
Private conCupon As Boolean
Private strBodegaPedido As String
Private CodigoListaPrecio As String
Private StockMin As Integer
Private FacDirecto As Boolean
Private FacTicket As Boolean

Private Sub cmbCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        CargaClientes "N"
    End If
End Sub

Private Sub cmbNegocio_Change()
    Dim fact1 As String
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac,dep_codigo,tip_ped_factura_directo,tip_ped_facturaticket " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD0(clsSql.adorec_Def("tip_ped_factura_directo")) = 1 Then
                FacDirecto = True
            Else
                FacDirecto = False
            End If
            If FormatoD0(clsSql.adorec_Def("tip_ped_facturaticket")) = 1 Then
                FacTicket = True
            Else
                FacTicket = False
            End If
            
            
            fact1 = clsSql.adorec_Def(0)
            strBodegaPedido = clsSql.adorec_Def(1)
            'LimpiarTodo
             cmbCliente.BoundText = ""
            cmbVendedor.Text = ""
            cmbVendedor.Text = ""
            'txtRuc.Text = ""
            txtDireccion.Text = ""
            TxtCategoria.Text = ""
            txtCantidad.Text = ""
            txtCredito.Text = ""
            txtDcto.Text = ""
            txtDisponible.Text = ""
            TxtObser.Text = ""
            txtTF.Text = ""
            txtTotDcto.Text = ""
            TxtTotal.Text = ""
            TxtTotalConIVA.Text = ""
            TxtTotal.Tag = ""
            VSFG.Clear 1
            VSFG.Rows = 2
            VSFGCot.Clear 1
        '    VSFGTPeds.Cell(flexcpText, 1, 0) = "Hacer Pedido Manualmente"
        
            VSFGTPeds.Col = 0
            VSFGTPeds.Row = 1
            VSFGTPeds.ColComboList(1) = ""
        End If
    Else
        Exit Sub
    End If
    '****** TARJETAS
    
    strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
             " FROM tarjeta_credito " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND (tip_ped_codigo='%' OR tip_ped_codigo='" & cmbNegocio.BoundText & "') " & _
             " ORDER BY tar_cre_nombre "
    clsTC.Ejecutar (strSql)
    Set cmbTC.RowSource = clsTC.adorec_Def.DataSource
    cmbTC.ListField = "tar_cre_nombre"
    cmbTC.BoundColumn = "tar_cre_codigo"
    cmbTC.BoundText = "SINTC"
    '****** CLIENTES
    cmbCliente.BoundText = ""
'    strSql = " SELECT CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
'             " cat_p_nombre, lis_pre_codigo, persona.per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,persona.per_ruc,persona.per_direccion, " & _
'             " COALESCE(CONCAT(persona.per_telf,'/',persona.per_fax),'') as per_tf,persona.per_celular,persona.per_email,persona.per_observacion,cat_p_dcto,persona.per_dcto,persona.per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,persona.per_codigo_ref,persona.per_codigo_ref2, " & _
'             " CONCAT(COALESCE(GZ.per_apellido,'-'),' ',COALESCE(GZ.per_nombre,'-')) as GERZON,CONCAT(COALESCE(DIR.per_apellido,'-'),' ',COALESCE(DIR.per_nombre,'-')) as DIRE" & _
'             " FROM usuario_gerente INNER JOIN persona ON usuario_gerente.per_codigo=persona.per_codigo_ref AND usuario_gerente.emp_codigo=persona.emp_codigo " & _
'             " INNER JOIN categoria_p " & _
'             " ON persona.cat_p_tipo = categoria_p.cat_p_tipo AND persona.cat_p_codigo = categoria_p.cat_p_codigo " & _
'             " AND persona.emp_codigo = categoria_p.emp_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
'             " AND vendedor.emp_codigo = persona.emp_codigo " & _
'             " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
'             " AND persona.per_codigo_ref = GZ.per_codigo AND GZ.per_es_gz=1 " & _
'             " LEFT JOIN persona as DIR ON persona.emp_codigo = DIR.emp_codigo " & _
'             " AND persona.per_codigo_ref2 = DIR.per_codigo AND DIR.per_es_di=1 " & _
'             " Where usuario_gerente.emp_codigo='" & strEmpresa & "' " & _
'             " AND usuario_gerente.usu_codigo='" & strUsuario & "'" & _
'             " AND categoria_p.cat_p_tipo='C' " & _
'             " AND persona.per_inactivo=0 " & _
'             " ORDER BY nombC "
'    clsClie.Ejecutar (strSql)
'    If clsClie.adorec_Def.RecordCount > 0 Then
'        strSql = " SELECT CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
'                 " cat_p_nombre, lis_pre_codigo, persona.per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,persona.per_ruc,persona.per_direccion, " & _
'                 " COALESCE(CONCAT(persona.per_telf,'/',persona.per_fax),'') as per_tf,persona.per_celular,persona.per_email,persona.per_observacion,cat_p_dcto,persona.per_dcto,persona.per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado," & _
'                 " persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6, " & _
'                 " persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9, " & _
'                 " CONCAT(COALESCE(GZ.per_apellido,'-'),' ',COALESCE(GZ.per_nombre,'-')) as GERZON," & _
'                 " CONCAT(COALESCE(DIR.per_apellido,'-'),' ',COALESCE(DIR.per_nombre,'-')) as DIRE," & _
'                 " CONCAT(COALESCE(EMP.per_apellido,'-'),' ',COALESCE(EMP.per_nombre,'-')) as EMPR," & _
'                 " CONCAT(COALESCE(EJE.per_apellido,'-'),' ',COALESCE(EJE.per_nombre,'-')) as EJES," & _
'                 " CONCAT(COALESCE(N5.per_apellido,'-'),' ',COALESCE(N5.per_nombre,'-')) as NN5," & _
'                 " CONCAT(COALESCE(N6.per_apellido,'-'),' ',COALESCE(N6.per_nombre,'-')) as NN6," & _
'                 " CONCAT(COALESCE(N7.per_apellido,'-'),' ',COALESCE(N7.per_nombre,'-')) as NN7," & _
'                 " CONCAT(COALESCE(N8.per_apellido,'-'),' ',COALESCE(N8.per_nombre,'-')) as NN8," & _
'                 " CONCAT(COALESCE(N9.per_apellido,'-'),' ',COALESCE(N9.per_nombre,'-')) as NN9" & _
'                 " FROM usuario_gerente INNER JOIN persona ON usuario_gerente.per_codigo=persona.per_codigo_ref AND usuario_gerente.emp_codigo=persona.emp_codigo " & _
'                 " INNER JOIN categoria_p " & _
'                 " ON persona.cat_p_tipo = categoria_p.cat_p_tipo AND persona.cat_p_codigo = categoria_p.cat_p_codigo " & _
'                 " AND persona.emp_codigo = categoria_p.emp_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
'                 " AND vendedor.emp_codigo = persona.emp_codigo "
'        strSql = strSql & " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
'                 " AND persona.per_codigo_ref = GZ.per_codigo AND GZ.per_es_gz=1 " & _
'                 " LEFT JOIN persona as DIR ON persona.emp_codigo = DIR.emp_codigo " & _
'                 " AND persona.per_codigo_ref2 = DIR.per_codigo AND DIR.per_es_di=1 " & _
'                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
'                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
'                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
'                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
'                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
'                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
'                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
'                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
'                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
'                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
'                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
'                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
'                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
'                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 "
'        strSql = strSql & " Where usuario_gerente.emp_codigo='" & strEmpresa & "' " & _
'                 " AND usuario_gerente.usu_codigo='" & strUsuario & "'" & _
'                 " AND categoria_p.cat_p_tipo='C' " & _
'                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
'                 " AND persona.per_inactivo=0 " & _
'                 " ORDER BY nombC "
'        clsClie.Ejecutar (strSql)
'        VSFGTPeds.Enabled = False
'    Else
        'Obtiene todos los clientes de una empresa con su respectiva lista de precios y vendedor asociado
        
        CargaClientes "R"
        If txtRuc.Text <> "" Then
            txtRuc_Validate True
        End If
        
End Sub

Private Sub CargaClientes(Forma As String)
    Dim strFiltroForma As String
        strSql = " SELECT CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
                 " cat_p_nombre, lis_pre_codigo, persona.per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,persona.per_ruc,persona.per_direccion, " & _
                 " COALESCE(CONCAT(persona.per_telf,'/',persona.per_fax),'') as per_tf,persona.per_celular,persona.per_email,persona.per_observacion,cat_p_dcto,persona.per_dcto,persona.per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado," & _
                 " persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6, " & _
                 " persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9, " & _
                 " CONCAT(COALESCE(GZ.per_apellido,'-'),' ',COALESCE(GZ.per_nombre,'-')) as GERZON," & _
                 " CONCAT(COALESCE(DIR.per_apellido,'-'),' ',COALESCE(DIR.per_nombre,'-')) as DIRE," & _
                 " CONCAT(COALESCE(EMP.per_apellido,'-'),' ',COALESCE(EMP.per_nombre,'-')) as EMPR," & _
                 " CONCAT(COALESCE(EJE.per_apellido,'-'),' ',COALESCE(EJE.per_nombre,'-')) as EJES," & _
                 " CONCAT(COALESCE(N5.per_apellido,'-'),' ',COALESCE(N5.per_nombre,'-')) as NN5," & _
                 " CONCAT(COALESCE(N6.per_apellido,'-'),' ',COALESCE(N6.per_nombre,'-')) as NN6," & _
                 " CONCAT(COALESCE(N7.per_apellido,'-'),' ',COALESCE(N7.per_nombre,'-')) as NN7," & _
                 " CONCAT(COALESCE(N8.per_apellido,'-'),' ',COALESCE(N8.per_nombre,'-')) as NN8," & _
                 " CONCAT(COALESCE(N9.per_apellido,'-'),' ',COALESCE(N9.per_nombre,'-')) as NN9" & _
                 " FROM persona INNER JOIN categoria_p " & _
                 " ON persona.cat_p_tipo = categoria_p.cat_p_tipo AND persona.cat_p_codigo = categoria_p.cat_p_codigo " & _
                 " AND persona.emp_codigo = categoria_p.emp_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
                 " AND vendedor.emp_codigo = persona.emp_codigo "
        strSql = strSql & " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
                 " AND persona.per_codigo_ref = GZ.per_codigo AND GZ.per_es_gz=1 AND GZ.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as DIR ON persona.emp_codigo = DIR.emp_codigo " & _
                 " AND persona.per_codigo_ref2 = DIR.per_codigo AND DIR.per_es_di=1 AND DIR.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 AND EMP.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 AND EJE.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 AND N5.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 AND N6.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 AND N7.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 AND N8.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 AND N9.tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        If Forma = "R" Then
            strFiltroForma = " AND persona.per_ruc='" & txtRuc.Text & "'"
        ElseIf Forma = "C" Then
            strFiltroForma = " AND persona.per_codigo='" & cmbCliente.BoundText & "'"
        Else
            strFiltroForma = " AND CONCAT(persona.per_apellido,' ',persona.per_nombre) LIKE '" & cmbCliente.Text & "%'"
        End If
        strSql = strSql & " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND persona.per_inactivo=0 " & _
                 strFiltroForma & _
                 " ORDER BY nombC "
        clsClie.Ejecutar (strSql)
        VSFGTPeds.Enabled = True
    'End If
    
    
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsClie.adorec_Def.DataSource
    If clsClie.adorec_Def.RecordCount > 0 Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If


End Sub

Private Sub cmbNegocio_LostFocus()
    If cmbNegocio.BoundText = "" Then
        MsgBox "Primero seleccione un Tipo de Negocio", vbInformation, "Tipo de Negocio"
        cmbNegocio.SetFocus
    End If
End Sub

Private Sub cmbProducto_Validate(Cancel As Boolean)
    VSFG.TextMatrix(VSFG.Row, 2) = cmbProducto.BoundText
    cmbProducto.Visible = False
    VSFG.SetFocus
    VSFG.Col = 2
    VSFG.EditCell
End Sub

Private Sub cmbTC_Change()
    Dim i As Long, comis As Double
    Dim ProdAux As String
    Dim CantAux As Long
    Dim Tipo As Integer
    comis = 0
    If cmbTC.MatchedWithList = True And cmbCliente.MatchedWithList = True Then
            '****** TARJETAS
        strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
                 " FROM tarjeta_credito " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' AND '" & cmbNegocio.BoundText & "' LIKE (tip_ped_codigo) " & _
                 " ORDER BY tar_cre_nombre "
        clsTC.Ejecutar strSql
    
        clsTC.Filtrar "tar_cre_codigo='" & cmbTC.BoundText & "' "
        Tipo = FormatoD0(clsTC.adorec_Def("tip_com_codigo"))
        
        If Tipo = 1 Then 'No comision
            ''dblComision = clsTC.adorec_Def("tar_cre_porcentaje")
            dblComision = "0"
            TxtObser.Width = 8040
            txtComi.Visible = False
            lblComi.Visible = False
            txtComi = ""
        ElseIf Tipo = 2 Then
            dblComision = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
            TxtObser.Width = 8040
            txtComi.Visible = False
            lblComi.Visible = False
            txtComi = ""
        ElseIf Tipo = 3 Then
            dblComision = "0"
            comis = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
            TxtObser.Width = 5400
            txtComi.Visible = True
            lblComi.Visible = True
            txtComi = ""
        End If
        
'''''''        If tipoPed = 1 And VSFGTPeds.ComboIndex <> -1 Then
'''''''            'Crea una tabla temporal de precios promedio de la cotización seleccionada
'''''''            clsProm.crearTabla VSFGTPeds.ComboItem(VSFGTPeds.ComboIndex), clsClie.adorec_Def("per_codigo"), strEmpresa
'''''''            If Controlado = True Then
'''''''            'Consulta que obtiene la lista de precios promedio de una cotización
'''''''            strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0)-COALESCE(TempReser.cant,0) as exi_cantidad, " & _
'''''''                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, (PromPre*'" & 1 + dblComision / 100.00 & "') as lis_pre_p_precio " & _
'''''''                     " FROM ((producto INNER JOIN PrePromCot ON producto.prd_codigo=PrePromCot.prd_codigo) " & _
'''''''                     " INNER JOIN existencia ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
'''''''                     " LEFT JOIN TempReser ON existencia.prd_codigo=TempReser.prd_codigo AND existencia.dep_codigo=TempReser.dep_codigo " & _
'''''''                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 " & _
'''''''                     " GROUP BY dep_codigo, prd_codigo " & _
'''''''                     " ORDER BY existencia.dep_codigo, producto.prd_nombre "
'''''''            End If
'''''''        Else
            'Obtiene todas los productos de una empresa con respecto a la lista de precio del cliente
'            strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0) as exi_cantidad, " & _
'                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio*'" & 1 + dblComision / 100.00 & "' as lis_pre_p_precio,prd_cambia_precio " & _
'                     " FROM ((producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
'                     " AND producto.emp_codigo=lista_precio_p.emp_codigo) INNER JOIN existencia " & _
'                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
'                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 " & _
'                     " AND lista_precio_p.lis_pre_codigo=" & clsClie.adorec_Def("lis_pre_codigo") & " " & _
'                     " GROUP BY dep_codigo, prd_codigo " & _
'                     " ORDER BY existencia.dep_codigo, producto.prd_nombre "
''''''''        End If 'Fin verificar tipo pedido
'        'Ejecuta la consulta de lista de precios
'        clsLstPrds.Ejecutar strSqlPrd

        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 2) <> "" Then
                ProdAux = VSFG.TextMatrix(i, 2)
                CantAux = VSFG.TextMatrix(i, 4)
                VSFG.TextMatrix(i, 2) = ""
                VSFG.TextMatrix(i, 2) = ProdAux
                'VSFG.TextMatrix(i, 4) = CantAux
            End If
        Next i
    End If
    'CalcuTotal
    txtComi.Text = FormatoD2(FormatoD4(TxtTotal.Text) * FormatoD4(comis / 100#))
End Sub


Private Sub cmdCargaBackCot_Click()
    Dim i As Long
    For i = 1 To VSFGCot.Rows - 1
        VSFG.AddItem ""
        VSFG.TextMatrix(i, 1) = strBodegaPedido
        VSFG.TextMatrix(i, 2) = VSFGCot.TextMatrix(i, 0)
        VSFG.TextMatrix(i, 4) = VSFGCot.TextMatrix(i, 2)
        VSFG_AfterEdit i, 4
    Next i
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
    PonerBotones
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
        'VSFG.Rows = i + 2
        If VSFG.TextMatrix(VSFG.Rows - 1, 2) <> "" Then
            VSFG.AddItem ""
        End If
        j = VSFG.Rows - 1
        VSFG.TextMatrix(j, 1) = strBodegaPedido
        VSFG.TextMatrix(j, 2) = Trim(CadenaRetorno(i)(0))
        If VSFG.TextMatrix(j, 3) <> "" Then
            VSFG.TextMatrix(j, 4) = Val(CadenaRetorno(i)(1))
            VSFG_AfterEdit j, 4
        Else
            MsgBox "El producto " & Trim(CadenaRetorno(i)(0)) & " no fue encontrado." & vbNewLine & "Cantidad pedida " & Val(CadenaRetorno(i)(1)), vbInformation, "Cantidad"
            VSFG.RemoveItem j
        End If
        If UBound(CadenaRetorno(0)) > 2 Then
            TxtObser.Text = CadenaRetorno(i)(2)
        End If
    Next i
    'CONTROLAR
End Sub



Private Sub cmdDirEnvio_Click()
    frmDireccionEnvio.strCliente = cmbCliente.BoundText
    frmDireccionEnvio.Show vbModal
    frmDireccionEnvio.strPedido = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    'Unload frmCartera
    On Error Resume Next
    'Elimina la tabla temporal de precios promedio de productos de una cotización
    clsProm.elimTabla
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSqlPrd = ""
    Set clsCots = Nothing
    Set clsBods = Nothing
    Set clsDet = Nothing
    Set clsPrds = Nothing
    Set clsLstPrds = Nothing
    Set clsClie = Nothing
    Set clsSql = Nothing
    Set clsBack = Nothing
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
End Sub

Private Sub CalcuTotal()
    'Calcula el total del pedido
    Dim Suma As Double
    Dim SumaDcto As Double
    Dim SumaCant As Double
    
    For i = 1 To VSFG.Rows - 1
        Suma = Suma + FormatoD4(VSFG.TextMatrix(i, 7))
        SumaDcto = SumaDcto + FormatoD4(VSFG.TextMatrix(i, 6))
        SumaCant = SumaCant + FormatoD4(VSFG.TextMatrix(i, 4))
    Next i
    
    TxtTotal = FormatoD2(Suma)
    TxtTotal.Tag = FormatoD4(Suma)
    txtTotDcto = FormatoD2(SumaDcto)
    TxtTotalConIVA = FormatoD2((FormatoD2(Suma)) * (1 + PorIVA / 100))
    txtCantidad = FormatoD2(SumaCant)
    If FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 7)) > 0 Then
        VSFG_AfterRowColChange VSFG.Row, VSFG.Col, VSFG.Row, VSFG.Col + 1
    End If
End Sub

'Función que verifica que cantidad de un producto ya ha sido pedida
Private Function cantXPed(CodPrd As String, fila As Long) As Double
        'Obtiene que cantidad de un producto se ha cotizado
        Dim cantPrd As Double
        For i = 1 To VSFGCot.Rows - 1
            If CodPrd = VSFGCot.TextMatrix(i, 0) Then
                cantPrd = VSFGCot.TextMatrix(i, 2)
                Exit For
            End If
        Next i
        'Encuetra el total de la cantidad pedida
        Dim sumPrd As Double
        For i = 1 To VSFG.Rows - 1
            If CodPrd = VSFG.TextMatrix(i, 2) And fila <> i Then
                sumPrd = sumPrd + FormatoD4(VSFG.TextMatrix(i, 4))
            End If
        Next i
        'Devuelve la cantidad a pedir del producto
        If (cantPrd - sumPrd > 0) Then
            cantXPed = cantPrd - sumPrd
        Else
            cantXPed = 0
        End If
End Function

Private Sub cmbCliente_Validate(Cancel As Boolean)
'Coloca los datos del cliente seleccionado en el combo
    Dim strSqlPrdTemp As String
    Dim strMensajeCartera As String
    Dim dblDeuda As Double
    Dim Gerente As String, Director As String, entra As Boolean
    Dim strEMail As String
    Dim strCelular As String
    entra = False
    conCupon = False
    txtDisponible.Text = 0
    If cmbCliente.BoundText <> "" Then
        If clsClie.adorec_Def.RecordCount > 0 Then
            clsClie.adorec_Def.MoveFirst
            clsClie.adorec_Def.Find "nombC='" & cmbCliente & "'"
        End If
        If Not clsClie.adorec_Def.EOF Then
        
        If CodigoListaPrecio <> clsClie.adorec_Def("lis_pre_codigo") Then
            CodigoListaPrecio = clsClie.adorec_Def("lis_pre_codigo")
        '****** PRODUCTOS
            'Recupera todos los productos de una empresa
'            strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
'                     " FROM producto " & _
'                     " INNER JOIN lista_precio_p " & _
'                     " ON lista_precio_p.emp_codigo=producto.emp_codigo " & _
'                     " AND lista_precio_p.prd_codigo=producto.prd_codigo " & _
'                     " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
'                     " AND lista_precio_p.lis_pre_codigo=" & CodigoListaPrecio & " " & _
'                     " AND lista_precio_p.lis_pre_p_precio!=0 " & _
'                     " ORDER BY producto.prd_nombre "
'            clsPrds.Ejecutar (strSql)
'AQUIIIIIIIIIIIII
            'Carga los productos en el combo de la columna 2 del flexGrid
            'VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
            'VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
        End If
        
        cmbVendedor.BoundText = IIf(IsNull(clsClie.adorec_Def("ven_codigo")), "", clsClie.adorec_Def("ven_codigo"))
        TxtCategoria = clsClie.adorec_Def("cat_p_nombre")
        Me.txtDireccion.Text = clsClie.adorec_Def("per_direccion")
        Me.txtRuc.Text = clsClie.adorec_Def("per_ruc")
        Me.txtTF.Text = clsClie.adorec_Def("per_tf")
        Me.txtDcto.Text = clsClie.adorec_Def("per_dcto")
        strEMail = clsClie.adorec_Def("per_email")
        strCelular = clsClie.adorec_Def("per_celular")
        
        If cmbNegocio.BoundText = "PRO" Or cmbNegocio.BoundText = "GYE" Then
            strEMail = InputBox("Ingrese el email:", "Clientes", strEMail)
            strCelular = InputBox("Ingreso OBLIGATORIAMENTE un Celular ", "Cliente", strCelular)
            If Trim(strEMail) & Trim(strCelular) <> "" Then
                strSql = " UPDATE persona " & _
                         " SET per_celular='" & strCelular & "'," & _
                         " per_email='" & strEMail & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND per_codigo='" & cmbCliente.BoundText & "'" & _
                         " AND cat_p_tipo='C'"
                clsSql.Ejecutar strSql, "M"
            End If
        End If
        
        Gerente = clsClie.adorec_Def("per_codigo_ref")
        Director = clsClie.adorec_Def("per_codigo_ref2")
        VSFRRed.TextMatrix(0, 1) = clsClie.adorec_Def("GERZON")
        VSFRRed.TextMatrix(1, 1) = clsClie.adorec_Def("DIRE")
        VSFRRed.TextMatrix(2, 1) = clsClie.adorec_Def("EMPR")
        VSFRRed.TextMatrix(3, 1) = clsClie.adorec_Def("EJES")
        VSFRRed.TextMatrix(4, 1) = clsClie.adorec_Def("NN5")
        VSFRRed.TextMatrix(5, 1) = clsClie.adorec_Def("NN6")
        VSFRRed.TextMatrix(6, 1) = clsClie.adorec_Def("NN7")
        VSFRRed.TextMatrix(7, 1) = clsClie.adorec_Def("NN8")
        VSFRRed.TextMatrix(8, 1) = clsClie.adorec_Def("NN9")
        VSFRRed.ShowCell 8, 1
        strSql = " SELECT SUM(per_credito) " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' AND tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " AND per_ruc='" & clsClie.adorec_Def("per_ruc") & "' " & _
                 " GROUP BY emp_codigo "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            '''Me.txtCredito.Text = clsClie.adorec_Def("per_credito")
            Me.txtCredito.Text = FormatoD4(clsSql.adorec_Def(0))
        Else
            Me.txtCredito.Text = "0"
        End If
        'frmCartera.Persona = cmbCliente.BoundText
        'frmCartera.Carga
        If clsClie.adorec_Def("cat_p_dcto") = 1 Then
            booDcto = True
        Else
            booDcto = False
        End If
        If Not IsNull(clsClie.adorec_Def("per_observacion")) And clsClie.adorec_Def("per_observacion") <> "" Then
            MsgBox clsClie.adorec_Def("per_observacion"), vbInformation, "Observaciones"
        End If
        
        If FormatoD0(clsClie.adorec_Def("per_bloqueado")) = 1 Then
            MsgBox "Cliente BLOQUEADO por cartera." & vbNewLine & vbNewLine & "No podrá hacer pedido hasta resolver el problema en CARTERA", vbCritical, "Cartera"
            CmdPedido.Enabled = False
        Else
'            strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
'                     " FROM persona " & _
'                     " WHERE emp_codigo='" & strEmpresa & "' " & _
'                     " AND cat_p_tipo='C' " & _
'                     " AND per_codigo='" & Gerente & "' " & _
'                     " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
'            clsSql.Ejecutar strSql
'            If clsSql.adorec_Def.RecordCount > 0 Then
'                If FormatoD0(clsSql.adorec_Def(0)) = 1 Then
'                    MsgBox "El Gerente de Zona del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "No podrá hacer pedido hasta resolver el problema", vbCritical, "Bloqueado"
'                    CmdPedido.Enabled = False
'                    entra = True
'                End If
'            End If
            
'            If entra = False Then
'                strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
'                         " FROM persona " & _
'                         " WHERE emp_codigo='" & strEmpresa & "' " & _
'                         " AND cat_p_tipo='C' " & _
'                         " AND per_codigo='" & Director & "' " & _
'                         " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
'                clsSql.Ejecutar strSql
'                If clsSql.adorec_Def.RecordCount > 0 Then
'                    If FormatoD0(clsSql.adorec_Def(0)) = 1 Then
'                        MsgBox "El Director del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "No podrá hacer pedido hasta resolver el problema", vbCritical, "Bloqueado"
'                        CmdPedido.Enabled = False
'                        entra = True
'                    End If
'                End If
'            End If
            
            If entra = False Then
                CmdPedido.Enabled = True
                VSFG.SetFocus
            Else
                cmbCliente.SetFocus
            End If
        End If
        'Verifica que tipo de pedido se está haciendo para obtener su respectiva lista de precios de productos
'''        strSqlPrdTemp = "DROP TABLE IF EXISTS TempReser"
'''        clsLstPrds.Ejecutar strSqlPrdTemp
'''        strSqlPrdTemp = " CREATE TEMPORARY TABLE TempReser ( " & _
'''                        " prd_codigo varchar(20) NOT NULL, " & _
'''                        " dep_codigo char(3) NOT NULL, " & _
'''                        " cant decimal(14,4), " & _
'''                        " PRIMARY KEY(prd_codigo,dep_codigo)) "
'''        clsLstPrds.Ejecutar strSqlPrdTemp
'''        strSqlPrdTemp = " INSERT INTO TempReser SELECT prd_codigo,dep_codigo,sum(IIF(ped_estado=1,det_ped_cant_entregada,det_ped_cant_pedida)) as cant " & _
'''                        " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
'''                        " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
'''                        " AND ped_estado<=1 GROUP BY prd_codigo,dep_codigo"
'''        clsLstPrds.Ejecutar strSqlPrdTemp
        
        
'''''''        If tipoPed = 1 And VSFGTPeds.ComboIndex <> -1 Then
'''''''            'Crea una tabla temporal de precios promedio de la cotización seleccionada
'''''''            clsProm.crearTabla VSFGTPeds.ComboItem(VSFGTPeds.ComboIndex), clsClie.adorec_Def("per_codigo"), strEmpresa
'''''''            'Consulta que obtiene la lista de precios promedio de una cotización
'''''''            strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0)-COALESCE(TempReser.cant,0) as exi_cantidad, " & _
'''''''                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, (PromPre*'" & 1 + dblComision / 100.00 & "') as lis_pre_p_precio " & _
'''''''                     " FROM ((producto INNER JOIN PrePromCot ON producto.prd_codigo=PrePromCot.prd_codigo) " & _
'''''''                     " INNER JOIN existencia ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
'''''''                     " LEFT JOIN TempReser ON existencia.prd_codigo=TempReser.prd_codigo AND existencia.dep_codigo=TempReser.dep_codigo " & _
'''''''                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 " & _
'''''''                     " GROUP BY dep_codigo, prd_codigo " & _
'''''''                     " ORDER BY existencia.dep_codigo, producto.prd_nombre "
'''''''        Else
            'Obtiene todas los productos de una empresa con respecto a la lista de precio del cliente
'            strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0) as exi_cantidad, " & _
'                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio*'" & 1 + dblComision / 100.00 & "' as lis_pre_p_precio,prd_cambia_precio " & _
'                     " FROM ((producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
'                     " AND producto.emp_codigo=lista_precio_p.emp_codigo) INNER JOIN existencia " & _
'                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
'                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 " & _
'                     " AND lista_precio_p.lis_pre_codigo=" & clsClie.adorec_Def("lis_pre_codigo") & " " & _
'                     " GROUP BY dep_codigo, prd_codigo " & _
'                     " ORDER BY existencia.dep_codigo, producto.prd_nombre "
'''''''        End If 'Fin verificar tipo pedido
'        'Ejecuta la consulta de lista de precios
'        clsLstPrds.Ejecutar (strSqlPrd)
    End If
        If TipoPed <> 3 Then
            VSFG.Clear 1
            TxtTotal = 0
            TxtTotal.Tag = 0
            TxtTotalConIVA = 0
            txtTotDcto = 0
            txtCantidad = 0
            txtComi = 0
        End If
    End If
'    If FormatoD2(txtCredito.Text) <> 0 Then
'        RevisarCupo
'        If FormatoD2(txtDisponible.Text) <= 0 Then
'            CmdPedido.Enabled = False
'        End If
'    End If
End Sub

Private Sub RevisarCupo()
    strSql = " CREATE TABLE #Abo ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " abono decimal(14,2) default NULL," & _
             " abonoNC decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #Abo " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(sum(IIF(pag_observacion!='NOTA DE CREDITO' AND pag_observacion not like '%ABONA - NOTA DE CR%DITO%',pag_monto,0)),0.000) as abono,COALESCE(sum(IIF(pag_observacion='NOTA DE CREDITO' OR pag_observacion like '%ABONA - NOTA DE CR%DITO%',pag_monto,0)),0.000) as abonoNC " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = pago.emp_codigo AND pago.pag_fecha <='" & HoyDia & "') " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " GROUP BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsSql.Ejecutar strSql
    strSql = " CREATE TABLE #RetFech ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " reten decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #RetFech " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(comprobante_retencion.com_ret_total,0.000) as reten " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo AND comprobante_retencion.com_ret_fecha <='" & HoyDia & "') " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsSql.Ejecutar strSql
    strSql = " CREATE TABLE #Ret ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " reten decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #Ret " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(#RetFech.reten,0.000) as reten " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             " LEFT JOIN #RetFech ON cuenta_p_c.cue_p_c_codigo = #RetFech.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = #RetFech.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = #RetFech.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo='C' AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsSql.Ejecutar strSql
    strSql = " CREATE TABLE #Cob ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " abono decimal(14,2) default NULL," & _
             " abonoNC decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #Cob " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(#Abo.abono,0.000) as abono,COALESCE(#Abo.abonoNC,0.000) as abonoNC " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             " LEFT JOIN #Abo ON cuenta_p_c.cue_p_c_codigo = #Abo.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = #Abo.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = #Abo.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " ORDER BY cue_p_c_codigo "
    clsSql.Ejecutar strSql
    strSql = " EXEC Sp_Drop_Table_if_Exist '#Abo' "
    clsSql.Ejecutar strSql
    strSql = " CREATE TABLE #Cuentas " & _
             " SELECT COALESCE(cue_p_c_valor,0.000) - COALESCE(#Cob.abono,0.000) - COALESCE(#Cob.abonoNC,0.000) - COALESCE(#Ret.reten,0.000) as cartera " & _
             " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo like 'C' " & _
             " INNER JOIN #Ret ON cuenta_p_c.cue_p_c_codigo = #Ret.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = #Ret.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = #Ret.emp_codigo " & _
             " INNER JOIN #Cob ON cuenta_p_c.cue_p_c_codigo = #Cob.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = #Cob.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = #Cob.emp_codigo " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " AND ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(#Cob.abono,0.000) - COALESCE(#Cob.abonoNC,0.000) - COALESCE(#Ret.reten,0.000),2)>0 "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #Cuentas " & _
             " SELECT COALESCE(-1 * ing_total,0.000) + COALESCE(-1 * ing_saldo,0.000) " & _
             " FROM ingreso INNER JOIN persona ON ingreso.per_codigo = persona.per_codigo " & _
             " AND ingreso.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo like 'C' " & _
             " WHERE ingreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND ingreso.tip_ing_codigo='DCL' AND  persona.per_ruc = '" & txtRuc.Text & "' " & _
             " AND ingreso.ing_anulado=0" & _
             " AND ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2)>0 "
    clsSql.Ejecutar strSql
    strSql = " INSERT INTO #Cuentas " & _
             " SELECT SUM(doc_pag_valor) " & _
             " FROM doc_pago INNER JOIN persona ON doc_pago.per_codigo = persona.per_codigo " & _
             " AND doc_pago.emp_codigo = persona.emp_codigo AND doc_pago.doc_pag_estado!='ANULADO' " & _
             " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' AND doc_pag_fecha_doc>'" & HoyDia & "'" & _
             " AND persona.per_ruc = '" & txtRuc.Text & "' GROUP BY doc_pago.emp_codigo"
    clsSql.Ejecutar strSql
    strSql = " SELECT COALESCE(SUM(cartera),0) as car " & _
             " FROM #Cuentas "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
    txtDisponible.Text = txtCredito.Text - clsSql.adorec_Def("car")
    Else
    txtDisponible.Text = txtCredito.Text
    End If
    strSql = " Sp_Drop_Table_if_Exist '#Ret' "
    
    clsSql.Ejecutar strSql
    strSql = " EXEC Sp_Drop_Table_if_Exist '#RetFech' "
    clsSql.Ejecutar strSql
    strSql = " EXEC Sp_Drop_Table_if_Exist '#Cob' "
    clsSql.Ejecutar strSql
    strSql = " EXEC Sp_Drop_Table_if_Exist '#Cuentas' "
    clsSql.Ejecutar strSql
    
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdLimpiar_Click()
    VSFG.Clear 1
    VSFG.Rows = 2
    cmbCliente.BoundText = ""
    cmbVendedor.Text = ""
    cmbVendedor.Text = ""
    txtRuc.Text = ""
    txtDireccion.Text = ""
    TxtCategoria.Text = ""
    txtCantidad.Text = ""
    txtCredito.Text = ""
    txtDcto.Text = ""
    txtDisponible.Text = ""
    TxtObser.Text = ""
    txtTF.Text = ""
    txtTotDcto.Text = ""
    TxtTotal.Text = ""
    TxtTotalConIVA = ""
    TxtTotal.Tag = ""
End Sub

Private Sub CmdPedido_Click()
    Dim Confirmado As Boolean
    Dim dblDeuda As Double
    Dim dblDcto As Double
    Dim strDireccion As String
    Dim clsPedido As New clsPedidos
    clsPedido.Inicializar AdoConn, AdoConnMaster
    
    FacDirecto = False
'****** NEGOCIO
    If cmbNegocio = "" Or cmbNegocio.MatchedWithList = False Then
        MsgBox "Seleccione un tipo de negocio primero.", vbInformation, "Tipo de Negocio"
        cmbNegocio.SetFocus
        Exit Sub
    End If
'****** CLIENTE
    If cmbCliente = "" Or cmbCliente.MatchedWithList = False Then
        MsgBox "Seleccione un cliente primero.", vbInformation, "Cliente"
        cmbCliente.SetFocus
        Exit Sub
    End If
'****** VENDEDOR
    If cmbVendedor = "" Or cmbVendedor.MatchedWithList = False Then
        MsgBox "Seleccione un vendedor primero.", vbInformation, "Cliente"
        cmbVendedor.SetFocus
        Exit Sub
    End If
'****** TOTAL PEDIDO
    'Verifica que por lo menos se haya seleccionado una cantidad para un producto
    If Val(FormatoD4(TxtTotal.Tag)) = 0 And FormatoD2(txtCantidad) = 0 Then
        If MsgBox("No ha seleccionado ningún producto para el pedido." & vbNewLine & "Desea generar un pedido confirmado en blanco?", vbQuestion + vbYesNo, "Pedido") = vbNo Then
            Exit Sub
        End If
    End If
'****** ACTUALIZAR TIPO PEDIDO

    
    Dim Fact As String
    strSql = " SELECT tip_ped_ptofac,tip_ped_factura_directo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        Fact = clsSql.adorec_Def(0)
        If FormatoD0(clsSql.adorec_Def(1)) = 1 Then
            FacDirecto = True
        Else
            FacDirecto = False
        End If
    End If
    
    'Verifica que tipo de pedido se hizo para actulizar datos en la tabla correspondiente
    Select Case TipoPed
        Case 1 'Cotización
            'Actualiza el estado de la cotización a pedida
            strSql = " UPDATE cotizacion SET cot_estado=4 " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & VSFGTPeds.TextMatrix(1, 1) & "' "
            clsSql.Ejecutar (strSql), "M"
            'Actualiza el estado del proyecto de venta
            strSql = " SELECT pro_ven_codigo FROM cotizacion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & VSFGTPeds.TextMatrix(1, 1) & "' "
            clsSql.Ejecutar (strSql)
            strSql = " UPDATE proyecto_venta SET pro_ven_estado=3 " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND pro_ven_codigo='" & clsSql.adorec_Def("pro_ven_codigo") & "'"
            clsSql.Ejecutar (strSql), "M"
        Case 2 'BackOrder
            'Actualiza el estado del backOrder como de baja
            strSql = " UPDATE backorder SET bac_baja= 1 " & _
                     " WHERE bac_codigo='" & VSFGTPeds.TextMatrix(1, 1) & "' AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql), "M"
        Case 4 'Modificar Pedido
            'Actualiza el estado del backOrder como de baja
            strSql = " SELECT est_descripcion,ped_fecha " & _
                     " FROM pedido INNER JOIN est_pedido ON pedido.ped_estado=est_pedido.est_codigo " & _
                     " WHERE ped_codigo='" & lngCod & "' AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql)
            TxtObser.Text = TxtObser.Text & vbNewLine & "Estado Ant: " & clsSql.adorec_Def("est_descripcion") & " / Fecha: " & clsSql.adorec_Def("ped_fecha")
            'Actualiza el estado del backOrder como de baja
            strSql = " DELETE FROM pedido " & _
                     " WHERE ped_codigo='" & lngCod & "' AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql), "M"
            strSql = " DELETE FROM det_pedido " & _
                     " WHERE ped_codigo='" & lngCod & "' AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar (strSql), "M"
    End Select
'****** CABECERA PEDIDO
    'Obtiene el código del pedido a ingresar
    Dim num As Double
    'Inserta la cabecera del pedido
    '''If MsgBox("Desea Facturar Sin Confirmar?", vbQuestion + vbYesNo, "Pedido") = vbNo Then
        Confirmado = False
        If Val(FormatoD4(TxtTotal.Tag)) = 0 And Val(FormatoD4(txtCantidad.Text)) = 0 Then
            Confirmado = True
        End If
        
        If TipoPed <> 4 Then
            Do
                strSql = " BEGIN TRAN "
                clsSqlNum.Ejecutar strSql, "M"
                strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(strSucursal & Fact & "0000001") & "') as num " & _
                         " From pedido WITH (TABLOCKX) " & _
                         " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(strSucursal & Fact) & "%'" & _
                         " GROUP BY emp_codigo"
                clsSqlNum.Ejecutar (strSql), "M"
                strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
                     " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega, ped_fechamod, ped_usumod) " & _
                     " VALUES ('" & strEmpresa & "'," & clsSqlNum.adorec_Def("num") & ",'" & clsClie.adorec_Def("per_codigo") & "','" & cmbVendedor.BoundText & "', " & _
                     " '" & cmbTC.BoundText & "','" & dblComision & "'," & _
                     " CURRENT_TIMESTAMP,'" & IIf(Confirmado = False, 0, 1) & "'," & FormatoD2(TxtTotalConIVA) & ",'" & UCase(TxtObser) & "', " & _
                     " '" & FormatoD0(VSFGTPeds.TextMatrix(1, 1)) & "',1,'" & IIf(conCupon = True, txtDcto.Text, 0) & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsSql.Ejecutar (strSql), "M"
                num = clsSqlNum.adorec_Def("num")
                strSql = " COMMIT TRAN "
                clsSqlNum.Ejecutar strSql, "M"
                strSql = " Select COUNT(*) as n " & _
                         " From pedido " & _
                         " Where emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_codigo = '" & num & "'" & _
                         " AND per_codigo = '" & clsClie.adorec_Def("per_codigo") & "'"
                clsSql.Ejecutar (strSql), "M"
                If FormatoD0(clsSql.adorec_Def("n")) = 0 Then
                    MsgBox "Se Asignará un nuevo numero de pedido", vbInformation, "PEDIDOS"
                End If
            Loop Until FormatoD0(clsSql.adorec_Def("n")) <> 0
        
        Else
            num = lngCod
            strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
                 " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo, ped_fechamod, ped_usumod) " & _
                 " VALUES ('" & strEmpresa & "'," & num & ",'" & clsClie.adorec_Def("per_codigo") & "','" & cmbVendedor.BoundText & "', " & _
                 " '" & cmbTC.BoundText & "','" & dblComision & "'," & _
                 " CURRENT_TIMESTAMP,'" & IIf(Confirmado = False, 0, 1) & "'," & FormatoD2(TxtTotalConIVA) & ",'" & UCase(TxtObser) & "', " & _
                 " '" & VSFGTPeds.TextMatrix(1, 1) & "',1,CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsSql.Ejecutar (strSql), "M"
        
        End If
        
        
        
        '''********Cuando ponia sin confirmar
    '''Else
       ''' If tipoPed <> 4 Then
        '''    strSql = " Select COALESCE(max(ped_codigo),0) as num " & _
        '''            " From pedido " & _
        '''             " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE CONCAT('" & strSucursal & "'+0,'%')"
        '''    clsSql.Ejecutar (strSql), "M"
        '''    num = clsSql.adorec_Def("num") + 1
       ''' Else
        '''    num = lngCod
        ''' End If
        '''Confirmado = True
        '''strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
        '''         " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo, ped_fechamod, ped_usumod) " & _
        '''         " VALUES ('" & strEmpresa & "'," & num & ",'" & clsClie.adorec_Def("per_codigo") & "','" & cmbVendedor.BoundText & "', " & _
        '''         " '" & cmbTC.BoundText & "','" & dblComision & "'," & _
        '''         " '" & Format(dtpFecha, "yyyy-MM-dd") & "',1," & FormatoD2(txtTotal) & ",'" & UCase(txtObser) & "', " & _
        '''         " '" & VSFGTPeds.TextMatrix(1, 1) & "',1,CURRENT_TIMESTAMP, '" & strUsuario & "') "
    '''End If
    
'****** DETALLES PEDIDO
    'Inserta los detalles del pedido
    With VSFG
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" And Val(FormatoD4(.TextMatrix(i, 4))) > 0 And Confirmado = False Then
                strSql = " SELECT det_ped_cant_pedida,det_ped_precio,det_ped_dcto " & _
                         " FROM det_pedido " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_codigo=" & num & " " & _
                         " AND prd_codigo='" & .TextMatrix(i, 2) & "' " & _
                         " AND dep_codigo='" & .TextMatrix(i, 1) & "' "
                clsSql.Ejecutar (strSql), "M"
                If clsSql.adorec_Def.RecordCount = 0 Then
                    strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                             " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod,det_ped_incentivo) " & _
                             " VALUES ('" & strEmpresa & "'," & num & ",'" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 1) & "'," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 11) & ", " & _
                             " " & .TextMatrix(i, 5) & "," & .TextMatrix(i, 6) & ", CURRENT_TIMESTAMP, '" & strUsuario & "','" & .TextMatrix(i, 12) & "') "
                    clsSql.Ejecutar (strSql), "M"
                Else
                    strSql = " UPDATE det_pedido " & _
                             " SET det_ped_cant_pedida=det_ped_cant_pedida+" & .TextMatrix(i, 4) & "," & _
                             " det_ped_cant_entregada=det_ped_cant_entregada+" & .TextMatrix(i, 11) & "," & _
                             " det_ped_precio=" & (clsSql.adorec_Def("det_ped_precio") * clsSql.adorec_Def("det_ped_cant_pedida") + .TextMatrix(i, 4) * .TextMatrix(i, 5)) / (clsSql.adorec_Def("det_ped_cant_pedida") + .TextMatrix(i, 4)) & "," & _
                             " det_ped_dcto=det_ped_dcto+" & .TextMatrix(i, 6) & " " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND ped_codigo=" & num & " " & _
                             " AND prd_codigo='" & .TextMatrix(i, 2) & "' " & _
                             " AND dep_codigo='" & .TextMatrix(i, 1) & "' "
                    clsSql.Ejecutar (strSql), "M"
                
                End If
'            ElseIf .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" And Val(FormatoD4(.TextMatrix(i, 4))) > 0 And Confirmado = True Then
'                strSQL = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
'                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
'                         " VALUES ('" & strEmpresa & "'," & num & ",'" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 1) & "'," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 4) & " " & _
'                         "," & .TextMatrix(i, 5) & "," & .TextMatrix(i, 6) & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
'                clsSql.Ejecutar (strSQL), "M"
            End If
        Next i
    End With
    
    'PROMO PRENDA PRECIO'
    dblDcto = FormatoD2(PromoPrendaPrecio(cmbNegocio.BoundText, num, " FORMAT(CURRENT_TIMESTAMP,'yyyyMMdd') ", True))
    If FormatoD2(dblDcto) > 0 Then
        MsgBox "El pedido tiene con un Dcto adicional de $" & FormatoD2(dblDcto), vbInformation, "Advertencia"
    End If
    If FacDirecto = True Then
        
        strSql = " UPDATE det_pedido " & _
                 " SET det_ped_cant_entregada=det_ped_cant_pedida," & _
                 " det_ped_cant_confirmada=det_ped_cant_pedida " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo=" & num & " "
        clsSql.Ejecutar (strSql), "M"
        strSql = " UPDATE pedido " & _
                 " SET ped_estado=1" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo=" & num & " "
        clsSql.Ejecutar (strSql), "M"
    End If
    clsPedido.RecalculoTotal (num)
    If MsgBox("El Pedido sera por le valor de: " & clsPedido.dblSubtotalMasIva & vbNewLine & _
              "Desea continuar?", vbQuestion + vbYesNo, "Ventas") = vbNo Then
        clsPedido.Eliminar
    Else
    
        If Me.cmbNegocio.BoundText = "JON" Or Me.cmbNegocio.BoundText = "LEM" Then
            frmDireccionEnvio.strCliente = cmbCliente.BoundText
            frmDireccionEnvio.strPedido = num
            frmDireccionEnvio.Show vbModal
        End If
        If FacDirecto = False Then
            Dim RepPed As New frmReporte
            RepPed.strNumero = num
            RepPed.strTipo = 2
            RepPed.strReporte = "rptPedido"
            RepPed.Show
        Else
            frmV_VerPedConfirm.Show
            frmV_VerPedConfirm.cmdNotaEntrega.Visible = False
            frmV_VerPedConfirm.cmbNegocio.BoundText = cmbNegocio.BoundText
            frmV_VerPedConfirm.txtPedido = FormatoD0(num)
            frmV_VerPedConfirm.txtPedido_KeyDown vbKeyReturn, 0
            frmV_VerPedConfirm.SetFocus
            frmV_VerPedConfirm.CmdConfirmar.SetFocus
        End If
    
    End If
    

    TipoPed = 0
    cmdLimpiar_Click
    VSFGTPeds.ComboIndex = 0
    VSFGTPeds.TextMatrix(1, 0) = "Hacer Pedido Manualmente"
    VSFGTPeds.TextMatrix(1, 1) = ""
End Sub

Private Sub Form_Activate()
    If strSqlPrd <> "" Then
        'clsLstPrds.Ejecutar strSqlPrd
    End If
End Sub

'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
    cmbVendedor.Text = ""
    cmbVendedor.Text = ""
    txtRuc.Text = ""
    txtDireccion.Text = ""
    TxtCategoria.Text = ""
    txtCantidad.Text = ""
    txtCredito.Text = ""
    txtDcto.Text = ""
    txtDisponible.Text = ""
    TxtObser.Text = ""
    txtTF.Text = ""
    txtTotDcto.Text = ""
    TxtTotal.Text = ""
    TxtTotalConIVA.Text = ""
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGCot.Clear 1
'    VSFGTPeds.Cell(flexcpText, 1, 0) = "Hacer Pedido Manualmente"

    VSFGTPeds.Col = 0
    VSFGTPeds.Row = 1
    VSFGTPeds.ColComboList(1) = ""

    VSFGTPeds_ValidateEdit 1, 0, False
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    If Trim(strPtoFactura) = "" Then
        frmSelNegocio.Show vbModal
    End If
    
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
    Controlado = True
    'frmCartera.Show
    'Inicializa las clases para hacer distintas consultas
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsBods.Inicializar AdoConn, AdoConnMaster
    clsPrds.Inicializar AdoConn, AdoConnMaster
    clsClie.Inicializar AdoConn, AdoConnMaster
    clsLstPrds.Inicializar AdoConn, AdoConnMaster
    clsCots.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsBack.Inicializar AdoConn, AdoConnMaster
    clsFacAnu.Inicializar AdoConn, AdoConnMaster
    clsTC.Inicializar AdoConn, AdoConnMaster
    clsSqlNum.Inicializar AdoConn, AdoConnMaster
    dblComision = 0
    
    cargarTipoPedido
'****** CLAVE
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSql)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")
'****** STOCK DE SEGURIDAD
    strSql = " SELECT par_web_valor " & _
             " FROM parametro_web " & _
             " WHERE par_web_id = 'STOCK_SEGURO_MINIMO' "
    clsSql.Ejecutar (strSql)
    StockMin = FormatoD0(clsSql.adorec_Def("par_web_valor"))
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
    
'****** TARJETAS
    strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
             " FROM tarjeta_credito " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND (tip_ped_codigo='%' OR tip_ped_codigo='" & cmbNegocio.BoundText & "') " & _
             " ORDER BY tar_cre_nombre "
    clsTC.Ejecutar (strSql)
    Set cmbTC.RowSource = clsTC.adorec_Def.DataSource
    cmbTC.ListField = "tar_cre_nombre"
    cmbTC.BoundColumn = "tar_cre_codigo"
    cmbTC.BoundText = "SINTC"
    'Selecciona el primer elemento del combo de cotizaciones
    strSql = " SELECT CURRENT_TIMESTAMP as dh"
    clsBods.Ejecutar (strSql)
    dtpFecha.Value = clsBods.adorec_Def("dh")
    Me.Height = 8960
    
'****** BODEGAS
    'Recupera todas las bodegas de una empresa
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By dep_nombre "
    clsBods.Ejecutar (strSql)
    Set cmbBodega.RowSource = clsBods.adorec_Def.DataSource
    cmbBodega.ListField = "dep_nombre"
    cmbBodega.BoundColumn = "dep_codigo"
    
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsBods.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    'Coloca los botones del grid en la primera columna
    PonerBotones
    If clsClie.adorec_Def.RecordCount > 0 Then
        CodigoListaPrecio = clsClie.adorec_Def("lis_pre_codigo")
    '****** PRODUCTOS
        'Recupera todos los productos de una empresa
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " INNER JOIN lista_precio_p " & _
                 " ON lista_precio_p.emp_codigo=producto.emp_codigo " & _
                 " AND lista_precio_p.prd_codigo=producto.prd_codigo " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " AND lista_precio_p.lis_pre_codigo=" & CodigoListaPrecio & " " & _
                 " AND lista_precio_p.lis_pre_p_precio!=0 " & _
                 " ORDER BY producto.prd_nombre "
        clsPrds.Ejecutar (strSql)
            'Carga los productos en el combo de la columna 2 del flexGrid
            'VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
            'VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    End If
    'cmbTC_Change
End Sub

Private Sub Form_Resize()
    'Coloca los objetos del formulario en su posición correcta
    If Me.Height = 10800 Then
        FraDetalle.Top = 5840
        FraBotones.Top = 9460
    ElseIf Me.Height = 8960 Then
        FraDetalle.Top = 3990 '3880
        FraBotones.Top = 7590 '6720
    End If
End Sub

Private Sub txtDcto_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmbNegocio.BoundText = "PRO" Then
        If KeyCode = vbKeyF5 Then
            If VSFG.TextMatrix(1, 2) = "" And Me.txtCantidad.Text = 0 And Me.TxtTotal.Text = 0 Then
                txtDcto.Text = 5
                conCupon = True
                MsgBox "Aplica 5% por CUPON, debe adjuntar el cupon a la factura"
            Else
                MsgBox "Debe registrar el cupón antes de ingresar las prendas"
            End If
        ElseIf KeyCode = vbKeyF6 Then
            If VSFG.Rows <= 2 And VSFG.TextMatrix(1, 2) = "" And Me.txtCantidad.Text = 0 And Me.TxtTotal.Text = 0 Then
                txtDcto.Text = 0
                conCupon = False
                MsgBox "Sin CUPON"
            Else
                MsgBox "Debe volver a pasar las prendas sin el descuento"
            End If
        End If
    End If
End Sub

Private Sub txtRuc_Validate(Cancel As Boolean)
    
    If Trim(txtRuc.Text) <> "" Then
        CargaClientes "R"
        
        If clsClie.adorec_Def.RecordCount > 0 Then
            If cmbCliente.BoundText <> clsClie.adorec_Def("per_codigo") Then
                cmbCliente.BoundText = clsClie.adorec_Def("per_codigo")
                cmbCliente_Validate False
            End If
        Else
            If FacDirecto = True And FacTicket = True Then
                If MsgBox("No se encontró un cliente con CI/RUC " & txtRuc.Text & vbNewLine & "Desea Ingresar el nuevo cliente?", vbQuestion + vbYesNo, "CI/RUC") = vbYes Then
                    frmClientePromo.strTipoPed = cmbNegocio.BoundText
                    frmClientePromo.txtCIRUC.Text = txtRuc.Text
                    frmClientePromo.Show vbModal
                    Cancel = False
                Else
                    If cmbCliente.Text <> "" Then
                        LimpiarTodo
                    Else
                        txtRuc.Text = ""
                    End If
                End If
            Else
                MsgBox "No se encontró un cliente con CI/RUC " & txtRuc.Text, vbInformation, "CI/RUC"
                If cmbCliente.Text <> "" Then
                    LimpiarTodo
                Else
                    txtRuc.Text = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Verifica que solo se ingresen números tanto en la cantidad como en el precio
    If Col = 4 Or Col = 5 Or Col = 6 Then
        'Verifica que solo se ingresen números en el campo cantidad
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en la cantidad.", vbInformation, "Cantidad"
            VSFG.TextMatrix(Row, 4) = FormatoD4(intDato)
            VSFG.TextMatrix(Row, 11) = FormatoD4(intDato)
        Else
            VSFG.TextMatrix(Row, 4) = FormatoD4(VSFG.TextMatrix(Row, 4))
            VSFG.TextMatrix(Row, 11) = FormatoD4(VSFG.TextMatrix(Row, 4))
        End If
        'Verifica que solo se ingresen números en el campo precio
        If Not IsNumeric(VSFG.TextMatrix(Row, 5)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el precio.", vbInformation, "Precio"
            VSFG.TextMatrix(Row, 5) = intDato
        End If
        '****** Controla la EXISTENCIA
        'Verifica que no se esté pidiendo más productos de los que hay en existencia
        
        If Col = 4 Then
            
            If Val(FormatoD4(VSFG.TextMatrix(Row, 4))) > Val(FormatoD4(VSFG.TextMatrix(Row, 8))) And Left(VSFG.TextMatrix(Row, 2), 3) <> "PR-" Then
                If FormatoD2(VSFG.TextMatrix(Row, 8)) = 0 Then
                        MsgBox "No hay existencia del producto " & VSFG.Cell(flexcpTextDisplay, Row, 3) & " en la bodega.", vbInformation, "Existencia"
                        VSFG.TextMatrix(Row, 11) = 0
                Else
                    'MsgBox "Solo hay diponible " & VSFG.TextMatrix(Row, 8) & " unidades de este producto en esta bodega.", vbInformation, "Cantidad"
                    MsgBox "Solo hay diponible X unidades de este producto en esta bodega.", vbInformation, "Cantidad"
                    VSFG.TextMatrix(Row, 11) = IIf(FormatoD4(VSFG.TextMatrix(Row, 8)) > 0, VSFG.TextMatrix(Row, 8), 0)
                End If
            End If
        End If
        '*****************************
        'Verifica que no se pidan más productos de los cotizados en caso de una cotización
        If TipoPed > 0 And TipoPed < 3 Then
            If Col = 4 Then
                Dim cntP As Double
                cntP = Val(FormatoD4(VSFG.TextMatrix(Row, 8))) 'cantXPed(VSFG.TextMatrix(Row, 2), Row)
                If Val(FormatoD4(VSFG.TextMatrix(Row, 4))) > cntP Then
                    'MsgBox "Solo puede pedir " & cntP & " unidades de este producto.", vbInformation, "Unidades"
                    MsgBox "Solo puede pedir X unidades de este producto.", vbInformation, "Unidades"
                    '*****VSFG.TextMatrix(row, col) = cntP
                End If
            End If
        End If
        'Verifica que el precio de venta del producto no sea menor al costo
        If Val(FormatoD4(VSFG.TextMatrix(Row, 5))) <= 0 Then ' Val(FormatoD4(VSFG.TextMatrix(Row, 9))) Then
            If MsgBox("El precio mínimo de venta de este producto es: " & VSFG.TextMatrix(Row, 9) & vbNewLine & vbNewLine & "Desea Factrurar a otro precio?", vbQuestion + vbYesNo, "Precio") = vbYes Then
                frmClave.strClaveMAESTRA = strClaveMAESTRA
                frmClave.dblPrecio = Val(FormatoD4(VSFG.TextMatrix(Row, 5)))
                frmClave.Show vbModal
                If frmClave.Ret = False Then
                    VSFG.TextMatrix(Row, 5) = FormatoD4(VSFG.TextMatrix(Row, 9))
                    VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
                End If
            Else
                VSFG.TextMatrix(Row, 5) = FormatoD4(VSFG.TextMatrix(Row, 9))
                VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
            End If
        End If
        'Actualiza el total del producto pedido
        dctoMax = 0
        dctoMax = FormatoD2(txtDcto.Text)
        strSql = " SELECT prd_pro_porcentaje " & _
                 " FROM producto_promo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "' AND producto_promo.tip_ped_codigo='" & Me.cmbNegocio.BoundText & "'" & _
                 " AND CURRENT_TIMESTAMP BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD2(clsSql.adorec_Def(0)) > FormatoD2(txtDcto.Text) Then
                dctoMax = FormatoD2(clsSql.adorec_Def(0))
            End If
        End If
'''
        VSFG.TextMatrix(Row, 6) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD4(dctoMax) / 100#)
        VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
        CalcuTotal
    End If
    
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid en caso de ser necesario
    'If VSFG.Rows - 1 >= OldRow And VSFG.Rows - 1 >= NewRow Then
    If FormatoD2(FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 4)) * FormatoD4(VSFG.TextMatrix(VSFG.Rows - 1, 5))) <> 0 Then
     '   If OldRow = VSFG.Rows - 1 And OldCol = 5 And VSFG.TextMatrix(OldRow, 8) <> "" Then
            VSFG.AddItem ""
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            If VSFG.Rows > 2 Then
                VSFG.TextMatrix(VSFG.Rows - 1, 1) = strBodegaPedido
            End If
            VSFG.Row = VSFG.Rows - 1
            VSFG.Col = 2
      '  End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Captura el dato ya almacenado en una celda antes de ser modificado
    If cmbCliente = "" Then
        MsgBox "Primero seleccione un Cliente", vbInformation, "Cliente"
        Cancel = True
        cmbCliente.SetFocus
    End If
    If Col = 4 Or Col = 5 Then
        If Col = 4 And FormatoD4(VSFG.TextMatrix(Row, 5)) = 0 Then
            'Cancel = True
        Else
            intDato = VSFG.TextMatrix(Row, Col)
        End If
    End If
    If Col = 5 Then
        If Abs(FormatoD0(VSFG.TextMatrix(Row, 10))) = 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    'No permite entrar en las celdas de las columnas siguientes
    If NewCol = 7 Or NewCol = 8 Or NewCol = 9 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
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
    
    If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub
     
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
         VSFG.RemoveItem (r)
         PonerBotones
         CalcuTotal
    Else
        VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub


Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Coloca la descripción del producto en caso que se haga un pedido manual y el usuario haya seleccionado un código de producto
    If Row > 0 And VSFG.Tag <> "N" Then
        If Col = 4 Then
        VSFG.TextMatrix(Row, 11) = VSFG.TextMatrix(Row, 4)
        End If
        If Col = 1 Or Col = 2 Then
            If Col = 2 Then
                
                strSqlPrd = " SELECT producto.prd_codigo,producto.prd_baja,producto.prd_nombre " & _
                             " FROM producto " & _
                             " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'"
    
                clsLstPrds.Ejecutar (strSqlPrd)
                If clsLstPrds.adorec_Def.RecordCount > 0 Then
                    If clsLstPrds.adorec_Def("prd_baja") = 1 Then
                        MsgBox "El producto " & clsLstPrds.adorec_Def("prd_codigo") & " " & clsLstPrds.adorec_Def("prd_nombre") & vbNewLine & _
                                "esta de baja y no se podra despachar"
                        Exit Sub
                    End If
                Else
                    MsgBox "El producto " & VSFG.TextMatrix(Row, 2) & vbNewLine & _
                                "no existe"
                    Exit Sub
                End If
                
            End If
            If VSFG.TextMatrix(Row, 1) = "" Then
                MsgBox "Seleccione primero una bodega", vbInformation, "Bodega"
                VSFG.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            If Col = 1 And VSFG.TextMatrix(Row, 1) <> strBodegaPedido Then
                frmClave.strClaveMAESTRA = strClaveMAESTRA
                frmClave.dblPrecio = "Bodega"
                frmClave.Show vbModal
                If frmClave.Ret = False Then
                    VSFG.TextMatrix(Row, 1) = strBodegaPedido
                End If
            End If
            If Col = 3 Then
                VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(Row, 3)
            End If
'            'Verifica que no se seleccione más de una vez el mismo producto en la misma bodega
'            For i = 1 To VSFG.Rows - 1
'                If VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(i, 2) And VSFG.TextMatrix(Row, 1) = VSFG.TextMatrix(i, 1) And Row <> i Then
'                    MsgBox "Ese producto ya fue seleccionado en la bodega " & VSFG.TextMatrix(i, 2) & ", solo cambie la candidad del mismo.", vbInformation, "Producto"
'                    VSFG.RemoveItem Row
'                    PonerBotones
'                    If i < VSFG.Rows Then
'                        VSFG.Row = i
'                    Else
'                        VSFG.Row = 1
'                    End If
'                    VSFG.Col = 2
'                    Exit Sub
'                End If
'            Next i
            'Coloca los datos de un producto seleccionado
            If VSFG.TextMatrix(Row, 2) <> "" Then
                'Busca el producto seleccionado y coloca sus datos respectivos
            strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0)-COALESCE(rr.res,0) as exi_cantidad, " & _
                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio*'" & 1 + dblComision / 100# & "' as lis_pre_p_precio,prd_cambia_precio,prd_incentivo " & _
                     " FROM ((producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                     " AND producto.emp_codigo=lista_precio_p.emp_codigo) INNER JOIN existencia " & _
                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo) " & _
                     " LEFT JOIN (SELECT det_pedido.emp_codigo,det_pedido.dep_codigo,det_pedido.prd_codigo,SUM(det_ped_cant_entregada) as res" & _
                     " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                     " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                     " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                     " AND det_pedido.dep_codigo='" & VSFG.TextMatrix(Row, 1) & "' AND det_pedido.prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'" & _
                     " AND pedido.ped_estado in (0,1)" & _
                     " GROUP BY det_pedido.emp_codigo,det_pedido.dep_codigo,det_pedido.prd_codigo) as rr" & _
                     " ON existencia.emp_codigo=rr.emp_codigo AND existencia.dep_codigo=rr.dep_codigo" & _
                     " AND existencia.prd_codigo=rr.prd_codigo" & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND producto.prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'" & _
                     " AND existencia.dep_codigo='" & VSFG.TextMatrix(Row, 1) & "' AND lista_precio_p.lis_pre_codigo=" & clsClie.adorec_Def("lis_pre_codigo") & " " & _
                     " GROUP BY existencia.dep_codigo, producto.prd_codigo, rr.res, producto.prd_nombre, producto.prd_costo, lis_pre_p_precio, prd_cambia_precio, prd_incentivo " & _
                     " ORDER BY existencia.dep_codigo, producto.prd_nombre "

            clsLstPrds.Ejecutar (strSqlPrd)
                clsLstPrds.adorec_Def.MoveFirst
                clsLstPrds.Filtrar "dep_codigo='" & VSFG.TextMatrix(Row, 1) & "' AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'"
                If Not clsLstPrds.adorec_Def.EOF Then
                    VSFG.TextMatrix(Row, 3) = clsLstPrds.adorec_Def("prd_nombre")
                    ''''VSFG.TextMatrix(Row, 3) = clsLstPrds.adorec_Def("prd_nombre")
                    'Coloca el costo del producto en una columna oculta
                    VSFG.TextMatrix(Row, 9) = clsLstPrds.adorec_Def("prd_costo")
                    VSFG.TextMatrix(Row, 10) = Abs(FormatoD0(clsLstPrds.adorec_Def("prd_cambia_precio")))
                    VSFG.TextMatrix(Row, 12) = Abs(FormatoD0(clsLstPrds.adorec_Def("prd_incentivo")))
                    VSFG.TextMatrix(Row, 6) = 0#
                    'Verifica que el precio de la lista no sea menor al costo del producto y tampoco sea una cotización
                    ''If clsLstPrds.adorec_Def("prd_costo") > clsLstPrds.adorec_Def("lis_pre_p_precio") Then ''''And tipoPed <> 1 Then
                    ''    VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("prd_costo"))
                    ''Else
                        VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("lis_pre_p_precio"))
                    ''End If
                    'Verifica que la existencia del producto sea mayor que cero
                    If clsLstPrds.adorec_Def("exi_cantidad") - IIf(FacDirecto = True, 0, IIf(VSFG.TextMatrix(Row, 1) = "PRI", StockMin, 0)) > 0 And FormatoD4(VSFG.TextMatrix(Row, 5)) <> 0 Then
                        VSFG.TextMatrix(Row, 4) = 1
                        VSFG.TextMatrix(Row, 11) = 1
                    Else
                        If FacDirecto = False Then
                            VSFG.TextMatrix(Row, 4) = 0
                            VSFG.TextMatrix(Row, 11) = 0
                        Else
                            VSFG.TextMatrix(Row, 4) = 1
                            VSFG.TextMatrix(Row, 11) = 1
                        
                        End If
                    End If
                    
                    
                    dctoMax = 0
                    dctoMax = FormatoD2(txtDcto.Text)
                    strSql = " SELECT prd_pro_porcentaje " & _
                             " FROM producto_promo " & _
                             " WHERE emp_codigo = '" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "' AND producto_promo.tip_ped_codigo='" & Me.cmbNegocio.BoundText & "'" & _
                             " AND CURRENT_TIMESTAMP BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                    clsSql.Ejecutar strSql
                    If clsSql.adorec_Def.RecordCount > 0 Then
                        If FormatoD2(clsSql.adorec_Def(0)) > FormatoD2(txtDcto.Text) Then
                            dctoMax = FormatoD2(clsSql.adorec_Def(0))
                        End If
                    End If
                    
                    VSFG.TextMatrix(Row, 6) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD4(dctoMax) / 100#)
                    VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
                    VSFG.TextMatrix(Row, 8) = clsLstPrds.adorec_Def("exi_cantidad") - IIf(VSFG.TextMatrix(Row, 1) = "PRI", StockMin, 0)
                End If
                clsLstPrds.QuitarFiltro
                CalcuTotal
            End If
        End If
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If VSFG.Col = 3 And KeyCode = vbKeyF4 And Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col)) <> "" And Len(Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col))) >= 2 Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " INNER JOIN lista_precio_p " & _
                 " ON lista_precio_p.emp_codigo=producto.emp_codigo " & _
                 " AND lista_precio_p.prd_codigo=producto.prd_codigo " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " AND lista_precio_p.lis_pre_codigo=" & CodigoListaPrecio & " " & _
                 " AND lista_precio_p.lis_pre_p_precio!=0 " & _
                 " AND prd_nombre LIKE '" & Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col)) & "%' " & _
                 " ORDER BY producto.prd_nombre "
        clsAux.Ejecutar strSql
        
        Set cmbProducto.RowSource = clsAux.adorec_Def.DataSource
        cmbProducto.ListField = "prd_nombre"
        cmbProducto.BoundColumn = "prd_codigo"
        cmbProducto.Visible = True
        cmbProducto.SetFocus
    End If
End Sub

Private Sub VSFG_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim booPasa As Boolean
    'Captura el dato ya almacenado en una celda antes de ser modificado
    booPasa = True
    If Col = 6 Then
        If booDcto = False Then
            booPasa = False
        End If
    End If
    If booPasa = False Then
        Cancel = True
    End If
End Sub
Private Sub VSFGTPeds_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim ban As Integer
    VSFGTPeds.TextMatrix(1, 1) = ""
    'Detecta el momento que se selecciona un item del combo de tipo de pedido del grid de tipo de pedido
    If VSFGTPeds.Col = 0 And VSFGTPeds.Row = 1 Then
        VSFG.Rows = 2
        VSFG.Clear 1
        VSFGCot.Rows = 2
        VSFGCot.Clear 1
        VSFG.Editable = flexEDNone
        cmbCliente = ""
        TxtCategoria = ""
        cmbVendedor.BoundText = ""
        
'        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
'                 " FROM producto " & _
'                 " INNER JOIN lista_precio_p " & _
'                 " ON lista_precio_p.emp_codigo=producto.emp_codigo " & _
'                 " AND lista_precio_p.prd_codigo=producto.prd_codigo " & _
'                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
'                 " AND lista_precio_p.lis_pre_codigo=" & clsClie.adorec_Def("lis_pre_codigo") & " " & _
'                 " AND lista_precio_p.lis_pre_p_precio!=0 " & _
'                 " ORDER BY producto.prd_nombre "
'        clsPrds.Ejecutar (strSql)
        
        Select Case VSFGTPeds.ComboIndex
            Case -1 'Pedido Manual
                TipoPed = 0
                VSFGTPeds.ComboIndex = 1
                'Limpia el contenido del grid de pedido
                VSFG.Clear 1
                cmbCliente.Enabled = True
                cmbVendedor.Enabled = True
                Me.Height = 8960 '8010
                TxtObser = ""
                TxtObser.Locked = False
                PonerBotones
                'Carga los productos en el combo de la columna 2 del flexGrid
'                VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
'                VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                'Elimina el combo de códigos
                VSFGTPeds.ColComboList(1) = ""
                VSFG.Editable = flexEDKbdMouse
            Case 0 'Pedido Manual
                TipoPed = 0
                'Limpia el contenido del grid de pedido
                VSFG.Clear 1
                cmbCliente.Enabled = True
                cmbVendedor.Enabled = True
                Me.Height = 8960 '8010
                TxtObser = ""
                TxtObser.Locked = False
                PonerBotones
                'Carga los productos en el combo de la columna 2 del flexGrid
'                VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
'                VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                'Elimina el combo de códigos
                VSFGTPeds.ColComboList(1) = ""
                VSFG.Editable = flexEDKbdMouse
            Case 1 'Cotización
                TipoPed = 1
                '****** COTIZACIONES
                'Realiza la consulta que contiene todas las cotizaciones en la empresa con sus respectivos clientes y vendedores
                strSql = " SELECT cot_codigo, CONCAT(cot_codigo,' - ',SUBSTRING(pro_ven_descricion,1,30),'...') as descrip, " & _
                         " CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, CONCAT(ven_apellido,' ',ven_nombre) as nombV, cot_observacion, vendedor.ven_codigo " & _
                         " FROM ((persona INNER JOIN proyecto_venta ON (persona.emp_codigo = proyecto_venta.emp_codigo) AND (persona.per_codigo = proyecto_venta.per_codigo)) " & _
                         " INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo)) " & _
                         " INNER JOIN vendedor ON (vendedor.emp_codigo = proyecto_venta.emp_codigo) AND (vendedor.ven_codigo = proyecto_venta.ven_codigo) " & _
                         " WHERE proyecto_venta.emp_codigo='" & strEmpresa & "' AND cot_estado=0 " & _
                         " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                         " ORDER BY cot_codigo "
                clsCots.Ejecutar strSql
                'Muestra las cotizaciones en la segunda columna del grid de tipos de pedidos
                VSFGTPeds.ColComboList(1) = VSFGTPeds.BuildComboList(clsCots.adorec_Def, "descrip,*cot_codigo", "cot_codigo")
                Me.Height = 10800 '10200
                FraPedido.Caption = "Cotización Nº "
            Case 2 'BackOrder
                TipoPed = 2
                '****** BACKORDER
                'Consulta todos los backOrder de una empresa con sus respectivo cliente y vendedor que lo generó
                    strSql = " SELECT bac_codigo, TRIM(CONCAT(per_apellido,' ',per_nombre, ' - ',backorder.ped_codigo)) as descrip, TRIM(CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')')) as nombC, " & _
                         " TRIM(CONCAT(ven_apellido,' ',ven_nombre)) as nombV, vendedor.ven_codigo " & _
                         " FROM ((persona INNER JOIN pedido ON (persona.emp_codigo = pedido.emp_codigo) AND (persona.per_codigo = pedido.per_codigo)) " & _
                         " INNER JOIN vendedor ON (vendedor.ven_codigo = pedido.ven_codigo) AND (vendedor.emp_codigo = pedido.emp_codigo)) " & _
                         " INNER JOIN backorder ON (pedido.ped_codigo = backorder.ped_codigo) AND (pedido.emp_codigo = backorder.emp_codigo) " & _
                         " Where backorder.emp_codigo='" & strEmpresa & "' And bac_baja=0 " & _
                         " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                         " ORDER BY descrip "
                clsBack.Ejecutar (strSql)
                'Muestra los backOrders en la segunda columna del grid de tipos de pedidos
                VSFGTPeds.ColComboList(1) = VSFGTPeds.BuildComboList(clsBack.adorec_Def, "descrip,*bac_codigo", "bac_codigo")
                Me.Height = 10800 '10200
                FraPedido.Caption = "BackOrder Nº "
            Case 3 'Factura Anulada
                TipoPed = 3
                'Limpia el contenido del grid de pedido
                VSFG.Clear 1
                cmbCliente.Enabled = True
                cmbVendedor.Enabled = True
                Me.Height = 8960 '8010
                TxtObser = ""
                TxtObser.Locked = False
                PonerBotones
                'Carga los productos en el combo de la columna 2 del flexGrid
'                VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
'                VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                VSFG.Editable = flexEDKbdMouse
                '****** FACTURAS ANULADAS
                'Consulta todas las facturas anuladas de una empresa con sus respectivo cliente y vendedor que lo generó
                strSql = " SELECT egr_codigo, CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, " & _
                         " CONCAT(ven_apellido,' ',ven_nombre) as nombV, vendedor.ven_codigo " & _
                         " FROM ((egreso INNER JOIN persona ON (persona.emp_codigo = egreso.emp_codigo) AND (persona.per_codigo = egreso.per_codigo)) " & _
                         " INNER JOIN vendedor ON (vendedor.ven_codigo = egreso.ven_codigo) AND (vendedor.emp_codigo = egreso.emp_codigo)) " & _
                         " WHERE egreso.emp_codigo='" & strEmpresa & "' And egr_anulado =1 AND tip_egr_codigo='FAC' " & _
                         " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                         " ORDER BY egreso.egr_codigo "
                clsFacAnu.Ejecutar (strSql)
                'Muestra las Facturas anuladas en la segunda columna del grid de tipos de pedidos
                VSFGTPeds.ColComboList(1) = VSFGTPeds.BuildComboList(clsFacAnu.adorec_Def, "nombC,*egr_codigo", "egr_codigo")
                FraPedido.Caption = "Factura Nº "
            Case 4 'Modificar Pedido
                TipoPed = 4
                'Limpia el contenido del grid de pedido
                VSFG.Clear 1
                cmbCliente.Enabled = True
                cmbVendedor.Enabled = True
                Me.Height = 8960 '8010
                TxtObser = ""
                TxtObser.Locked = False
                PonerBotones
                'Carga los productos en el combo de la columna 2 del flexGrid
 '               VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
 '               VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                VSFG.Editable = flexEDKbdMouse
                '****** Pedidos
                'Consulta todas las facturas anuladas de una empresa con sus respectivo cliente y vendedor que lo generó
                strSql = " SELECT ped_codigo,ped_fecha,est_descripcion, CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, " & _
                         " CONCAT(ven_apellido,' ',ven_nombre) as nombV, vendedor.ven_codigo,pedido.per_codigo " & _
                         " FROM ((pedido INNER JOIN est_pedido ON pedido.ped_estado=est_pedido.est_codigo " & _
                         " INNER JOIN persona ON (persona.emp_codigo = pedido.emp_codigo) AND (persona.per_codigo = pedido.per_codigo)) " & _
                         " LEFT JOIN vendedor ON (vendedor.ven_codigo = IIF(pedido.ven_codigo='' OR pedido.ven_codigo IS NULL,persona.ven_codigo,pedido.ven_codigo)) AND (vendedor.emp_codigo = pedido.emp_codigo)) " & _
                         " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_estado <=1 " & _
                         " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                         " ORDER BY nombC,ped_fecha,pedido.ped_codigo "
                clsFacAnu.Ejecutar (strSql)
                'Muestra las Facturas anuladas en la segunda columna del grid de tipos de pedidos
                VSFGTPeds.ColComboList(1) = VSFGTPeds.BuildComboList(clsFacAnu.adorec_Def, "nombC,ped_fecha,est_descripcion,*ped_codigo", "ped_codigo")
                FraPedido.Caption = "Pedido Nº "
        End Select
        ban = 1
        VSFGTPeds.Col = 1
    End If
    'Detecta el momento que se selecciona un item del combo de código del grid de tipo de pedido
    If VSFGTPeds.Col = 1 And VSFGTPeds.Row > 0 And TipoPed > 0 And ban = 0 Then
        If VSFGTPeds.ComboIndex >= 0 Then
            lngCod = VSFGTPeds.ComboItem(VSFGTPeds.ComboIndex)
            VSFG.Editable = flexEDKbdMouse
            VSFG.Rows = 2
            VSFG.Clear 1
            Select Case TipoPed
                Case 1 'Cotización
                '****** DATOS COTIZACION
                    VSFGCot.Clear
                    'Coloca los datos del cliente y vendedor relacionados con una cotización
                    clsCots.adorec_Def.MoveFirst
                    clsCots.adorec_Def.Find "cot_codigo='" & lngCod & "' "
                    cmbCliente = clsCots.adorec_Def("nombC")
                    cmbCliente.Enabled = True
                    cmbCliente_Validate False
                    TxtObser = clsCots.adorec_Def("cot_observacion")
                    TxtObser.Locked = False
                    FraPedido.Caption = "Cotización Nº " & lngCod
                    Me.Height = 10800 '10200
                    cmbVendedor.Enabled = False
                    cmbVendedor.BoundText = clsCots.adorec_Def("ven_codigo")
                '****** DETALLE COTIZACION
                    'Obtiene solo los productos que intervienen en una cotización con sus respectivos datos
                    strSql = " SELECT IIF(producto.prd_codigo<>'',producto.prd_codigo,det_prd_com.prd_codigo)as prd_codigo, " & _
                             " IIF(producto.prd_nombre<>'',producto.prd_nombre,producto_1.prd_nombre) as prd_nombre, " & _
                             " sum(IIF(isnull(det_prd_com.det_prd_com_cantidad),det_cotizacion.det_cot_cantidad, " & _
                             " det_prd_com.det_prd_com_cantidad * det_cotizacion.det_cot_cantidad))as cantidad, " & _
                             " det_cotizacion.det_cot_precio,det_cotizacion.det_cot_precio*det_cotizacion.det_cot_cantidad " & _
                             " FROM (((det_cotizacion LEFT JOIN producto_compuesto ON (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo) " & _
                             " AND (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo)) LEFT JOIN producto ON (det_cotizacion.prd_codigo = producto.prd_codigo) " & _
                             " AND (det_cotizacion.emp_codigo = producto.emp_codigo)) LEFT JOIN det_prd_com ON (producto_compuesto.emp_codigo = det_prd_com.emp_codigo) " & _
                             " AND (producto_compuesto.prd_com_codigo = det_prd_com.prd_com_codigo)) LEFT JOIN producto AS producto_1 ON (det_prd_com.emp_codigo = producto_1.emp_codigo) " & _
                             " AND (det_prd_com.prd_codigo = producto_1.prd_codigo) " & _
                             " WHERE det_cotizacion.cot_codigo='" & lngCod & "' AND det_cotizacion.emp_codigo='" & strEmpresa & "' " & _
                             " GROUP BY prd_codigo "
                    clsDet.Ejecutar (strSql)
                    'Muestra los datos de la cotización en el formulario
                    Set VSFGCot.DataSource = clsDet.adorec_Def.DataSource
                '****** DETALLE PRODUCTOS
                    'Genera un filtro que permite seleccionar solo los productos mostrados en el detalle
''''''''                    Dim prods As String
''''''''                    For i = 1 To VSFGCot.Rows - 1
''''''''                        prods = prods & "prd_codigo='" & VSFGCot.TextMatrix(i, 0) & "' or "
''''''''                    Next i
''''''''                    prods = prods & "prd_codigo='0'"
''''''''
''''''''                    clsPrds.Filtrar prods
                    'Carga los productos en el combo de la columna 2 del flexGrid
                    VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
                    VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
''''''''                    clsPrds.QuitarFiltro
                    VSFG.Rows = 1
                    For i = 1 To VSFGCot.Rows - 1
                        VSFG.AddItem "" & vbTab & strBodegaPedido & vbTab & VSFGCot.TextMatrix(i, 0), i
                        
                        If FormatoD4(VSFGCot.TextMatrix(i, 2)) <= FormatoD4(VSFG.TextMatrix(i, 8)) Then
                            VSFG.TextMatrix(i, 4) = VSFGCot.TextMatrix(i, 2)
                        Else
                            VSFG.TextMatrix(i, 4) = VSFG.TextMatrix(i, 8)
                        End If
                        VSFG.TextMatrix(i, 5) = VSFGCot.TextMatrix(i, 3)
                        VSFG_AfterEdit i, 5
                        VSFG.TextMatrix(i, 0) = i
                        VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
                        VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                    Next i
                    VSFG.AddItem ""
                    VSFG.TextMatrix(i, 0) = i
                    VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
                    VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                Case 2 'BackOrder
                '****** DATOS BACKORDER
                    VSFGCot.Clear
                    'Coloca los datos del cliente y vendedor relacionados con una cotización
                    clsBack.adorec_Def.MoveFirst
                    clsBack.adorec_Def.Find "bac_codigo=" & lngCod
                    cmbCliente = clsBack.adorec_Def("nombC")
                    cmbCliente.Enabled = False
                    cmbCliente_Validate False
                    TxtObser = clsBack.adorec_Def("descrip")
                    TxtObser.Locked = False
                    FraPedido.Caption = "BackOrder Nº " & lngCod
                    Me.Height = 10800 '10200
                    cmbVendedor.Enabled = False
                    cmbVendedor.BoundText = clsBack.adorec_Def("ven_codigo")
                '****** DETALLE COTIZACION
                    'Obtiene solo los productos que intervienen en un backOrder con sus respectivos datos
                    strSql = " SELECT producto.prd_codigo, prd_nombre, det_bac_cantidad " & _
                             " FROM producto INNER JOIN det_backorder ON (producto.emp_codigo = det_backorder.emp_codigo) " & _
                             " AND (producto.prd_codigo = det_backorder.prd_codigo) " & _
                             " Where producto.emp_codigo='" & strEmpresa & "' And bac_codigo=" & lngCod & _
                             " ORDER BY producto.prd_codigo "
                    clsDet.Ejecutar (strSql)
                    'Muestra los datos de la cotización en el formulario
                    Set VSFGCot.DataSource = clsDet.adorec_Def.DataSource
                    'Carga los productos en el combo de la columna 2 del flexGrid
                    VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
                    VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                Case 3 'Factura Anulada
                '****** DATOS Factura
                    VSFGCot.Clear
                    'Coloca los datos del cliente y vendedor relacionados con una cotización
                    clsFacAnu.adorec_Def.MoveFirst
                    clsFacAnu.adorec_Def.Find "egr_codigo=" & lngCod
                    cmbCliente = clsFacAnu.adorec_Def("nombC")
                    cmbCliente.Enabled = True
                    cmbCliente_Validate False
                    TxtObser = "FACTURA ANULADA " & clsFacAnu.adorec_Def("egr_codigo")
                    TxtObser.Locked = False
                    FraPedido.Caption = "Factura Nº " & lngCod
                    cmbVendedor.Enabled = False
                    cmbVendedor.BoundText = clsFacAnu.adorec_Def("ven_codigo")
                '****** DETALLE FACTURA
'''''                     strSqlPrdTemp = "DROP TABLE IF EXISTS TempReser"
'''''                    clsDet.Ejecutar strSqlPrdTemp
'''''                    strSqlPrdTemp = " CREATE TEMPORARY TABLE TempReser ( " & _
'''''                                    " prd_codigo varchar(20) NOT NULL, " & _
'''''                                    " dep_codigo char(3) NOT NULL, " & _
'''''                                    " cant decimal(14,4), " & _
'''''                                    " PRIMARY KEY(prd_codigo,dep_codigo)) "
'''''                    clsDet.Ejecutar strSqlPrdTemp
'''''                    strSqlPrdTemp = " INSERT INTO TempReser SELECT prd_codigo,dep_codigo,sum(IIF(ped_estado=1,det_ped_cant_entregada,det_ped_cant_pedida)) as cant " & _
'''''                                    " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
'''''                                    " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
'''''                                    " AND ped_estado<=1 GROUP BY prd_codigo,dep_codigo"
'''''                    clsDet.Ejecutar strSqlPrdTemp
                    
                    'Obtiene solo los productos que intervienen en la factura con sus respectivos datos
                    strSql = " SELECT det_egreso.dep_codigo,det_egreso.prd_codigo, prd_nombre, IIF(det_egr_cantidad<SUM(exi_cantidad),det_egr_cantidad,SUM(exi_cantidad)),det_egr_precio,0,IIF(det_egr_cantidad<SUM(exi_cantidad),det_egr_cantidad,SUM(exi_cantidad))*det_egr_precio as tot,COALESCE(SUM(exi_cantidad),0) as exist,(producto.prd_costo/(1 - 0.1)) as prd_costo,prd_cambia_precio " & _
                             " FROM (det_egreso INNER JOIN producto ON producto.emp_codigo = det_egreso.emp_codigo  AND producto.prd_codigo = det_egreso.prd_codigo) " & _
                             " INNER JOIN existencia ON existencia.emp_codigo = det_egreso.emp_codigo AND existencia.prd_codigo = det_egreso.prd_codigo AND existencia.dep_codigo = det_egreso.dep_codigo " & _
                             " Where det_egreso.emp_codigo='" & strEmpresa & "' And egr_codigo=" & lngCod & _
                             " AND tip_egr_codigo='FAC' " & _
                             " GROUP BY dep_codigo,prd_codigo, prd_nombre, det_egr_cantidad,det_egr_precio " & _
                             " ORDER BY producto.prd_codigo "
                    clsDet.Ejecutar (strSql)
                    'Muestra los datos de la cotización en el formulario
                    VSFG.Tag = "N"
                    Set VSFG.DataSource = clsDet.adorec_Def.DataSource
                    VSFG.Tag = ""
                    PonerBotones
                    VSFG.ColComboList(1) = VSFG.BuildComboList(clsBods.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
                    VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
                    VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
                Case 4 'Modificar Pedido
                '****** DATOS Factura
                    VSFGCot.Clear
                    'Coloca los datos del cliente y vendedor relacionados con una cotización
                    clsFacAnu.adorec_Def.MoveFirst
                    clsFacAnu.adorec_Def.Find "ped_codigo=" & lngCod
                    cmbCliente.BoundText = clsFacAnu.adorec_Def("per_codigo")
                    CargaClientes "C"
                    cmbCliente.BoundText = clsFacAnu.adorec_Def("per_codigo")
                    
                    cmbCliente.Enabled = True
                    cmbCliente_Validate False
                    FraPedido.Caption = "Pedido Nº " & lngCod
                    cmbVendedor.Enabled = False
                    'cmbVendedor.BoundText = clsFacAnu.adorec_Def("ven_codigo")
                '****** DETALLE FACTURA
''''                    strSqlPrdTemp = "DROP TABLE IF EXISTS TempReser"
''''                    clsDet.Ejecutar strSqlPrdTemp
''''                    strSqlPrdTemp = " CREATE TEMPORARY TABLE TempReser ( " & _
''''                                    " prd_codigo varchar(20) NOT NULL, " & _
''''                                    " dep_codigo char(3) NOT NULL, " & _
''''                                    " cant decimal(14,4), " & _
''''                                    " PRIMARY KEY(prd_codigo,dep_codigo)) "
''''                    clsDet.Ejecutar strSqlPrdTemp
''''                    strSqlPrdTemp = " INSERT INTO TempReser SELECT prd_codigo,dep_codigo,sum(IIF(ped_estado=1,det_ped_cant_entregada,det_ped_cant_pedida)) as cant " & _
''''                                    " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
''''                                    " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
''''                                    " AND ped_estado<=1 GROUP BY prd_codigo,dep_codigo"
''''                    clsDet.Ejecutar strSqlPrdTemp
                    'Obtiene solo los productos que intervienen en la factura con sus respectivos datos
                    strSql = " SELECT det_pedido.dep_codigo,det_pedido.prd_codigo, prd_nombre, det_ped_cant_pedida,det_ped_precio,det_ped_dcto as dcto,det_ped_cant_pedida*det_ped_precio as tot,COALESCE(SUM(exi_cantidad),0)+det_ped_cant_pedida as exist,(producto.prd_costo/(1 - 0.1)) as prd_costo,prd_cambia_precio " & _
                             " FROM (det_pedido INNER JOIN producto ON producto.emp_codigo = det_pedido.emp_codigo  AND producto.prd_codigo = det_pedido.prd_codigo) " & _
                             " INNER JOIN existencia ON existencia.emp_codigo = det_pedido.emp_codigo AND existencia.prd_codigo = det_pedido.prd_codigo AND existencia.dep_codigo = det_pedido.dep_codigo " & _
                             " Where det_pedido.emp_codigo='" & strEmpresa & "' And ped_codigo=" & lngCod & _
                             " GROUP BY det_pedido.dep_codigo,det_pedido.prd_codigo, prd_nombre, det_ped_cant_pedida,det_ped_precio,det_ped_dcto,producto.prd_costo,prd_cambia_precio " & _
                             " ORDER BY det_pedido.prd_codigo "
                    clsDet.Ejecutar (strSql)
                    'Muestra los datos de la cotización en el formulario
                    VSFG.Tag = ""
                    '''''Set VSFG.DataSource = clsDet.adorec_Def.DataSource 'CAMBIO 2008/DIC/01
                    VSFG.Rows = 1
                    i = 1
                    While Not clsDet.adorec_Def.EOF
                        VSFG.AddItem "", i
                        VSFG.ShowCell i, 2
                        VSFG.TextMatrix(i, 1) = clsDet.adorec_Def("dep_codigo")
                        VSFG.TextMatrix(i, 2) = clsDet.adorec_Def("prd_codigo")
                        If VSFG.TextMatrix(i, 3) <> "" Then
                        VSFG.TextMatrix(i, 4) = clsDet.adorec_Def("det_ped_cant_pedida")
                        VSFG_AfterEdit i, 4
                        Else
                        VSFG.RemoveItem i
                        i = i - 1
                        End If

                        i = i + 1
                        clsDet.adorec_Def.MoveNext
                    Wend
                    VSFG.Tag = ""
                    PonerBotones
                    VSFG.ColComboList(1) = VSFG.BuildComboList(clsBods.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
                    'VSFG.ColComboList(2) = VSFG.BuildComboList(clsPrds.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
                    'VSFG.ColComboList(3) = VSFG.BuildComboList(clsPrds.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
            End Select
        End If 'Fin selección item de código
    End If 'Fin combo columna 1
    CalcuTotal
End Sub
