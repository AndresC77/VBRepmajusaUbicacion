VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmListaEmbarque 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Embarque"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   3405
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListaEmbarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8940
   Begin MSCommLib.MSComm MSCommBalanza 
      Left            =   8280
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.OptionButton optTodos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Todos"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   42
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optNs 
      BackColor       =   &H00DDDDDD&
      Caption         =   "N#'s"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   41
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton optSegunPed 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Segun Ped."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   40
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7890
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2925
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4485
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox TxtObserv 
      Height          =   525
      Left            =   1290
      MaxLength       =   250
      TabIndex        =   5
      Top             =   1320
      Width           =   7575
   End
   Begin MSDataListLib.DataCombo cmbCliente 
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   75
      Width           =   4185
      _ExtentX        =   7382
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
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   285
      Left            =   5850
      TabIndex        =   19
      Top             =   90
      Width           =   1335
      _ExtentX        =   3201
      _ExtentY        =   503
      Value           =   41836.5404166667
      Enabled         =   0   'False
   End
   Begin MSDataListLib.DataCombo cmbCourier 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   3225
      _ExtentX        =   5689
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
   Begin VB.TextBox txtGuia 
      Height          =   315
      Left            =   5010
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   25
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "PEDIDOS"
      TabPicture(0)   =   "frmListaEmbarque.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "VSFG"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLector"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtmail"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "OTROS"
      TabPicture(1)   =   "frmListaEmbarque.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAgregar"
      Tab(1).Control(1)=   "txtDescripcion"
      Tab(1).Control(2)=   "cmbCliente2"
      Tab(1).Control(3)=   "VSFG2"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label5"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "PESOS"
      TabPicture(2)   =   "frmListaEmbarque.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(2)=   "VSFGCajas"
      Tab(2).Control(3)=   "txtNoCajas"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Impresora"
      TabPicture(3)   =   "frmListaEmbarque.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtImpresora"
      Tab(3).Control(1)=   "cmdCambiar"
      Tab(3).Control(2)=   "chkImprimirSTK"
      Tab(3).Control(3)=   "Label9"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtNoCajas 
         Height          =   315
         Left            =   -73950
         TabIndex        =   11
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtImpresora 
         Height          =   525
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   13
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar Impresora"
         Height          =   375
         Left            =   -73080
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkImprimirSTK 
         Caption         =   "NO Imprimir Stiker"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -74520
         TabIndex        =   15
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtmail 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   480
         Width           =   4215
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   -71280
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
         _cx             =   5953
         _cy             =   661
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   -2.58236865538736
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   3
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
         AutoLinkNavigate=   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -69000
         TabIndex        =   10
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   525
         Left            =   -73800
         MaxLength       =   250
         TabIndex        =   9
         Top             =   1020
         Width           =   4575
      End
      Begin VB.TextBox txtLector 
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   420
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3855
         Left            =   120
         TabIndex        =   26
         Top             =   780
         Width           =   8655
         _cx             =   116079522
         _cy             =   116071056
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmListaEmbarque.frx":037A
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
      Begin MSDataListLib.DataCombo cmbCliente2 
         Height          =   330
         Left            =   -74160
         TabIndex        =   8
         Top             =   540
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   30
         Top             =   1740
         Width           =   8655
         _cx             =   116079522
         _cy             =   116069362
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmListaEmbarque.frx":046A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGCajas 
         Height          =   4095
         Left            =   -72960
         TabIndex        =   12
         Top             =   480
         Width           =   5175
         _cx             =   116073384
         _cy             =   116071479
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmListaEmbarque.frx":052A
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cajas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   38
         Top             =   532
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos:"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   -74880
         TabIndex        =   37
         Top             =   4740
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos:"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   -74880
         TabIndex        =   36
         Top             =   4740
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Etiquetas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74640
         TabIndex        =   35
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos:"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   4740
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   29
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   28
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   27
         Top             =   495
         Width           =   795
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7440
      TabIndex        =   39
      Top             =   855
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   885
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "No.Guia:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4320
      TabIndex        =   21
      Top             =   855
      Width           =   615
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   5280
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   540
      Width           =   405
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Negocio:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmListaEmbarque"
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
Private strSql As String
Private clsSql As New clsConsulta
Private emailLista As String
Private emailPapaLista As String
Private PedPendiente As String

Private Sub chkImprimirSTK_Click()
    If chkImprimirSTK.Value = 1 Then
        ImprimirEtiquetaDespacho = True
    Else
        ImprimirEtiquetaDespacho = False
    End If
    
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.MatchedWithList = True Then
        LimpiarForm
        If PedPendiente <> "" Then
            txtLector.Text = PedPendiente
            PedPendiente = ""
            txtLector_KeyDown vbKeyReturn, 0
        End If
    End If
End Sub

Private Sub LimpiarForm()
    Dim strCli As String
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    VSFG2.Clear flexClearScrollable
    VSFG2.Rows = 1
    TxtTotal.Text = 0
    txtmail.Text = ""
    If PedPendiente <> "" And optSegunPed.Value = True Then
        strCli = ""
        strSql = " SELECT per_codigo " & _
                 " FROM pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo='" & PedPendiente & "'"
        clsSql.Ejecutar (strSql)
        If clsSql.adorec_Def.RecordCount > 0 Then
            strCli = clsSql.adorec_Def("per_codigo")
        End If
        strSql = " SELECT per_codigo,nombC " & _
                 " FROM Fn_RedDeCliente('" & strEmpresa & "','" & strCli & "') "
    ElseIf optSegunPed.Value = True Then
        strSql = " SELECT per_codigo,nombC " & _
                 " FROM Fn_RedDeCliente('" & strEmpresa & "','') "
    ElseIf optSegunPed.Value = False Then
        strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
                 " FROM persona " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        If optNs.Value = True Then
            strSql = strSql & " AND (per_es_gz=1 " & _
                            " OR per_es_di=1 " & _
                            " OR per_es_em=1 " & _
                            " OR per_es_ee=1 " & _
                            " OR per_es_n5=1 " & _
                            " OR per_es_n6=1 " & _
                            " OR per_es_n7=1 " & _
                            " OR per_es_n8=1 " & _
                            " OR per_es_n9=1 " & _
                            " OR per_es_n10=1) "
        End If
    End If
    strSql = strSql & " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente2.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente2.ListField = "nombC"
    cmbCliente2.BoundColumn = "per_codigo"
    
End Sub


Private Sub cmbCourier_Validate(Cancel As Boolean)
    txtGuia.Locked = False
    txtGuia.Text = ""
    txtGuia.Tag = ""
    If cmbCourier.MatchedWithList = True Then
        strSql = " SELECT cou_prefijo_secuencial,cou_secuencial_actual,cou_secuencial_mascara" & _
                 " FROM courier " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & cmbCourier.BoundText & "'"
        clsSql.Ejecutar strSql, "M"
        If clsSql.adorec_Def("cou_secuencial_mascara") <> "" Then
            txtGuia.Locked = True
            txtGuia.Text = clsSql.adorec_Def("cou_prefijo_secuencial") & Format(clsSql.adorec_Def("cou_secuencial_actual"), clsSql.adorec_Def("cou_secuencial_mascara"))
            txtGuia.Tag = "T"
        End If
    End If
End Sub

Private Sub cmbNegocio_Change()
    Dim strCli As String
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
        End If
    Else
        Exit Sub
    End If
    LimpiarForm
    cmbCliente.BoundText = ""
    If PedPendiente <> "" And optSegunPed.Value = True Then
        strCli = ""
        strSql = " SELECT per_codigo " & _
                 " FROM pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo='" & PedPendiente & "'"
        clsSql.Ejecutar (strSql)
        If clsSql.adorec_Def.RecordCount > 0 Then
            strCli = clsSql.adorec_Def("per_codigo")
        End If
        strSql = " SELECT per_codigo,nombC " & _
                 " FROM Fn_RedDeCliente('" & strEmpresa & "','" & strCli & "') "
    ElseIf optSegunPed.Value = True Then
        strSql = " SELECT per_codigo,nombC " & _
                 " FROM Fn_RedDeCliente('" & strEmpresa & "','') "
    ElseIf optSegunPed.Value = False Then
        strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
                 " FROM persona " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        If optNs.Value = True Then
            strSql = strSql & " AND (per_es_gz=1 " & _
                     " OR per_es_di=1 " & _
                     " OR per_es_em=1 " & _
                     " OR per_es_ee=1 " & _
                     " OR per_es_n5=1 " & _
                     " OR per_es_n6=1 " & _
                     " OR per_es_n7=1 " & _
                     " OR per_es_n8=1 " & _
                     " OR per_es_n9=1 " & _
                     " OR per_es_n10=1) "
        End If
    End If
    strSql = strSql & " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub cmdAceptar_Click()
    Dim lngContenedor As Long
    Dim lngGuia As Long
    Dim i As Long
    Dim Sec As Long
    Dim lngNumPedidos As Long
    
    Dim emailLista As String
    Dim emailPapaLista As String
    Dim emailInmediato As String
    
    Dim RepStk As New frmReporte
    Dim RepStk2 As New frmReporte
    Dim RepEmpaque As New frmReporte
    
    Dim clsAux As New clsGuiaUrbano
    
    If FormatoD0(VSFG.Rows) + FormatoD0(VSFG2.Rows) <= 2 Then
        MsgBox "No ha ingresado ningun item a despachar", vbInformation, "Despacho"
        Exit Sub
    ElseIf cmbCliente.MatchedWithList = False Or cmbCliente.BoundText = "" Then
        MsgBox "No ha seleccionado un lider", vbInformation, "Despacho"
        Exit Sub
    ElseIf cmbCourier.MatchedWithList = False Or cmbCourier.BoundText = "" Then
        MsgBox "No ha seleccionado un Operador", vbInformation, "Despacho"
        Exit Sub
    ElseIf RevisarPesosYTam = False Or FormatoD2(txtPeso.Text) <= 0 Then
        MsgBox "No ha ingresado el Peso", vbInformation, "Despacho"
        Exit Sub
    End If
    strSql = " BEGIN TRAN "
    clsSql.Ejecutar strSql, "M"
    If txtGuia.Locked = True Then
        strSql = " SELECT cou_prefijo_secuencial,cou_secuencial_actual,cou_secuencial_mascara" & _
                 " FROM courier WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & cmbCourier.BoundText & "'"
        clsSql.Ejecutar strSql, "M"
        If clsSql.adorec_Def("cou_secuencial_mascara") <> "" Then
            txtGuia.Text = clsSql.adorec_Def("cou_prefijo_secuencial") & Format(clsSql.adorec_Def("cou_secuencial_actual"), clsSql.adorec_Def("cou_secuencial_mascara"))
            txtGuia.Tag = "T"
            Sec = clsSql.adorec_Def("cou_secuencial_actual") + 1
        End If
        
    End If
    
    strSql = " SELECT COALESCE(MAX(con_codigo),0) as n" & _
             " FROM contenedor WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSql, "M"
    If clsSql.adorec_Def.RecordCount > 0 Then
        lngContenedor = FormatoD0(clsSql.adorec_Def("n")) + 1
    Else
        lngContenedor = 1
    End If
    
    If txtGuia.Text = "" Then
        strSql = " COMMIT TRAN"
        clsSql.Ejecutar strSql, "M"
        MsgBox "No ha ingresado el numero de Guia", vbInformation, "Despacho"
        Exit Sub
    End If
    strSql = " SELECT persona.per_codigo,persona.per_email,COALESCE(p1.per_email,'') as emailpapa, IIF(p10.per_email<>'' AND p10.per_email is not null,p10.per_email,IIF(p9.per_email<>'' AND p9.per_email is not null,p9.per_email,IIF(p8.per_email<>'' AND p8.per_email is not null,p8.per_email,IIF(p7.per_email<>'' AND p7.per_email is not null,p7.per_email,IIF(p6.per_email<>'' AND p6.per_email is not null,p6.per_email,IIF(p5.per_email<>'' AND p5.per_email is not null,p5.per_email,IIF(p4.per_email<>'' AND p4.per_email is not null,p4.per_email,IIF(p3.per_email<>'' AND p3.per_email is not null,p3.per_email,IIF(p2.per_email<>'' AND p2.per_email is not null,p2.per_email,''))))))))) as emailinmediato " & _
             " FROM persona LEFT JOIN persona as p1 ON persona.emp_codigo=p1.emp_codigo AND persona.per_codigo_ref=p1.per_codigo AND p1.per_es_gz=1" & _
             " LEFT JOIN persona as p2 ON persona.emp_codigo=p2.emp_codigo AND persona.per_codigo_ref2=p2.per_codigo AND p2.per_es_di=1" & _
             " LEFT JOIN persona as p3 ON persona.emp_codigo=p3.emp_codigo AND persona.per_codigo_ref3=p3.per_codigo AND p3.per_es_em=1" & _
             " LEFT JOIN persona as p4 ON persona.emp_codigo=p4.emp_codigo AND persona.per_codigo_ref4=p4.per_codigo AND p4.per_es_ee=1" & _
             " LEFT JOIN persona as p5 ON persona.emp_codigo=p5.emp_codigo AND persona.per_codigo_ref5=p5.per_codigo AND p5.per_es_n5=1" & _
             " LEFT JOIN persona as p6 ON persona.emp_codigo=p6.emp_codigo AND persona.per_codigo_ref6=p6.per_codigo AND p6.per_es_n6=1" & _
             " LEFT JOIN persona as p7 ON persona.emp_codigo=p7.emp_codigo AND persona.per_codigo_ref7=p7.per_codigo AND p7.per_es_n7=1" & _
             " LEFT JOIN persona as p8 ON persona.emp_codigo=p8.emp_codigo AND persona.per_codigo_ref8=p8.per_codigo AND p8.per_es_n8=1" & _
             " LEFT JOIN persona as p9 ON persona.emp_codigo=p9.emp_codigo AND persona.per_codigo_ref9=p9.per_codigo AND p9.per_es_n9=1" & _
             " LEFT JOIN persona as p10 ON persona.emp_codigo=p10.emp_codigo AND persona.per_codigo_ref10=p10.per_codigo AND p10.per_es_n10=1"
    strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_codigo='" & cmbCliente.BoundText & "' "
    clsSql.Ejecutar (strSql)
    emailLista = ""
    emailPapaLista = ""
    emailInmediato = ""
    txtmail = ""
    If clsSql.adorec_Def.RecordCount > 0 Then
        emailLista = Trim(clsSql.adorec_Def("per_email"))
        emailPapaLista = Trim(clsSql.adorec_Def("emailpapa"))
        emailInmediato = Trim(clsSql.adorec_Def("emailinmediato"))
        txtmail = emailLista & IIf(emailLista <> "", "; ", "") & emailInmediato & IIf(emailInmediato <> "", "; ", "") & emailPapaLista
    End If
    
    strSql = " INSERT INTO contenedor (emp_codigo, con_codigo, cou_codigo, per_codigo, con_fecha," & _
             " con_guia,con_peso, con_observacion, con_fechamod, con_usumod) " & _
             " VALUES('" & strEmpresa & "','" & lngContenedor & "','" & cmbCourier.BoundText & "','" & cmbCliente.BoundText & "','" & dtpFecha.Value & "', " & _
             " '" & UCase(txtGuia.Text) & "','" & FormatoD4(txtPeso.Text) & "','" & UCase(TxtObserv.Text) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsSql.Ejecutar strSql, "M"
    
    If txtGuia.Locked = True Then
        strSql = " UPDATE courier " & _
                 " SET cou_secuencial_actual='" & FormatoD0(Sec) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & cmbCourier.BoundText & "'"
        clsSql.Ejecutar strSql, "M"
    End If
    lngNumPedidos = 0
    strSql = " COMMIT TRAN "
    clsSql.Ejecutar strSql, "M"
    For i = 1 To VSFG.Rows - 1
        strSql = " INSERT INTO det_contenedor (emp_codigo, con_codigo, tip_egr_codigo, egr_codigo, " & _
                 " det_con_fechamod, det_con_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & lngContenedor & "','" & VSFG.TextMatrix(i, 7) & "','" & VSFG.TextMatrix(i, 1) & "', " & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
        lngNumPedidos = lngNumPedidos + 1
    Next i
    For i = 1 To VSFG2.Rows - 1
        strSql = " SELECT det_con_per_detalle " & _
                 " FROM det_contenedor_per " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND con_codigo='" & lngContenedor & "'" & _
                 " AND per_codigo='" & VSFG2.TextMatrix(i, 0) & "'"
        clsSql.Ejecutar strSql, "M"
        If clsSql.adorec_Def.RecordCount = 0 Then
            strSql = " INSERT INTO det_contenedor_per (emp_codigo, con_codigo, per_codigo, " & _
                     " det_con_per_detalle,det_con_per_fechamod, det_con_per_usumod) " & _
                     " VALUES('" & strEmpresa & "','" & lngContenedor & "','" & VSFG2.TextMatrix(i, 0) & "', " & _
                     " '" & UCase(VSFG2.TextMatrix(i, 2)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        Else
            strSql = " UPDATE det_contenedor_per " & _
                     " SET det_con_per_detalle=CONCAT(det_con_per_detalle,', ','" & UCase(VSFG2.TextMatrix(i, 2)) & "') " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND con_codigo='" & lngContenedor & "' " & _
                     " AND per_codigo='" & VSFG2.TextMatrix(i, 0) & "'"
        End If
        clsSql.Ejecutar strSql, "M"
        lngNumPedidos = lngNumPedidos + 1
    Next i
    For i = 1 To VSFGCajas.Rows - 1
        strSql = " INSERT INTO det_contenedor_caja (emp_codigo, con_codigo, paq_env_codigo," & _
                 " det_con_caj_codigo, det_con_caj_peso, det_con_caj_fechamod,det_con_caj_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & lngContenedor & "','" & VSFGCajas.TextMatrix(i, 2) & "'," & _
                 " '" & VSFGCajas.TextMatrix(i, 0) & "','" & VSFGCajas.TextMatrix(i, 1) & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
    Next i
    
    If UCase(Left(txtGuia.Text, 9)) = "URBANO-WS" Then
        txtGuia.Text = clsAux.Generar(Format(lngContenedor, "0"))
    End If
    If txtGuia.Text = "NO GENERA" Then
        MsgBox "Listada de Embarque Guardado No. " & lngContenedor & _
            vbNewLine & "NO ENVIADA A OPERADOR" & lngContenedor & _
            vbNewLine & vbNewLine & "ERROR EN GUIA DE REMISION" & _
            vbNewLine & vbNewLine & "Comuniquese con SISTEMAS con el numero de contenedor", vbCritical, "Contenedores"
        Exit Sub
    End If
    MsgBox "Listada de Embarque Guardado No. " & lngContenedor, vbInformation, "Contenedores"
    
    If ImpresoraEtiqueta = "" Then
        RepStk.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = RepStk.VSPrint.Device
        txtImpresora.Text = ImpresoraEtiqueta
        GuardarImpresoras
    End If
    
    If chkImprimirSTK.Value = 0 Then
'        If txtGuia.Tag = "T" Then
            RepStk.VSPrint.Device = ImpresoraEtiqueta
'            RepStk.VSPrint.PaperWidth = 7669.292
'            RepStk.VSPrint.PaperHeight = 8885.039
            RepStk.VSPrint.PaperWidth = 7669.292
            RepStk.VSPrint.PaperHeight = 3150.039

            RepStk.strNumero = lngContenedor
            RepStk.strReporte = "rptSTKGuia"
            RepStk.Show
            RepStk.Form_Activate
            RepStk.VSPrint.Copies = 2
            RepStk.VSPrint.PrintDoc
            Unload RepStk
            
'            RepStk2.VSPrint.Device = ImpresoraEtiqueta
'            RepStk2.VSPrint.PaperWidth = 7669.292
'            RepStk2.VSPrint.PaperHeight = 3150.039
'        '    RepStk.VSPrint.PaperWidth = 7669.292
'        '    RepStk.VSPrint.PaperHeight = 8885.039
'            RepStk2.strNumero = lngContenedor
'            RepStk2.strReporte = "rptSTKGuiaBlanco"
'            RepStk2.Show
'            RepStk2.Form_Activate
'            RepStk2.VSPrint.PrintDoc
'            Unload RepStk2
'        Else
'            RepStk.VSPrint.Device = ImpresoraEtiqueta
'            RepStk.VSPrint.PaperWidth = 7669.292
'            RepStk.VSPrint.PaperHeight = 3885.039
'            RepStk.strNumero = lngContenedor
'            RepStk.strReporte = "rptSTKListaEmbarque"
'            RepStk.Show
'            RepStk.Form_Activate
'            RepStk.VSPrint.PrintDoc
'            Unload RepStk
'
'            RepStk2.VSPrint.Device = ImpresoraEtiqueta
'            RepStk2.VSPrint.PaperWidth = 7669.292
'            RepStk2.VSPrint.PaperHeight = 3885.039
'            RepStk2.strNumero = cmbCliente.BoundText
'            RepStk2.strReporte = "rptSTKIdCaja"
'            RepStk2.Show
'            RepStk2.Form_Activate
'            RepStk2.VSPrint.PrintDoc
'            Unload RepStk2
'        End If
    End If
    DefinirImpresoraPorDefecto ImpresoraPorDefecto
    RepEmpaque.strNumero = lngContenedor
    RepEmpaque.strReporte = "rptListaEmbarque"
    'RepEmpaque.Show
    RepEmpaque.Form_Activate
    RepEmpaque.VSRpt.RenderToFile "LE" & FormatoD0(lngContenedor) & ".pdf", vsrPDF
        'MsgBox "archiv generado"
'    If MsgBox("Imprime Lista de Embarque?", vbQuestion + vbYesNo, "Impresion") = vbYes Then
'        RepEmpaque.VSPrint.PrintDoc
'    End If
    If lngNumPedidos > 1 Then
        frmImpresionDirecta.strReporte = "rptListaEmbarque"
        frmImpresionDirecta.strNumero = lngContenedor
        frmImpresionDirecta.Show
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.cmdImprimir_Click
        frmImpresionDirecta.CmdCerrar_Click
    End If

On Error GoTo errhandler
    If Trim(txtmail) <> "" Then
        EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, cmbCliente.Text, Trim(txtmail), CorreoSupervisorDeTransportes, "Lista de Embarque " & lngContenedor, _
                "Estimad@" & vbNewLine & _
                cmbCliente.Text & vbNewLine & _
                "Adjunto encontrará la lista de embarque despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                "Saludos Cordiales" & vbNewLine & _
                "Departamento de Despachos" & vbNewLine & _
                NombreComercial, "LE" & FormatoD0(lngContenedor) & ".pdf"
        Kill "LE" & FormatoD0(lngContenedor) & ".pdf"
        If Trim(emailLista) = "" Then
            EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, "Supervidor de Transportes", CorreoSupervisorDeTransportes, "", "Lista de Embarque " & lngContenedor & "Cliente sin mails", _
                    "Estimad@" & vbNewLine & _
                    "El lider: " & cmbCliente.Text & vbNewLine & _
                    "No recibio la Lista de Embarque adjunta, despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                    "Ya que no tiene ingresado Email" & vbNewLine & _
                    "Saludos Cordiales" & vbNewLine & _
                    "Departamento de Despachos" & vbNewLine & _
                    NombreComercial
        
        ElseIf Trim(emailPapaLista) = "" Then
            EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, "Supervidor de Transportes", CorreoSupervisorDeTransportes, "", "Lista de Embarque " & lngContenedor & "Lider sin mails", _
                    "Estimad@" & vbNewLine & _
                    "El N1 del cliente: " & cmbCliente.Text & vbNewLine & _
                    "No recibio la Lista de Embarque adjunta, despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                    "Ya que no tiene ingresado Email" & vbNewLine & _
                    "Saludos Cordiales" & vbNewLine & _
                    "Departamento de Despachos" & vbNewLine & _
                    NombreComercial
        End If
    Else
        EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, "Supervidor de Transportes", CorreoSupervisorDeTransportes, "", "Lista de Embarque " & lngContenedor & "Lider sin mails", _
                "Estimad@" & vbNewLine & _
                "La guia del cliente : " & cmbCliente.Text & vbNewLine & _
                "Nadie recibio la Lista de Embarque adjunta, despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                "Ya que no tiene ingresado Email" & vbNewLine & _
                "Y el N1 de este tampoco recibira pues no tiene email ingresado" & vbNewLine & _
                "Saludos Cordiales" & vbNewLine & _
                "Departamento de Despachos" & vbNewLine & _
                NombreComercial, "LE" & FormatoD0(lngContenedor) & ".pdf"
        Kill "LE" & FormatoD0(lngContenedor) & ".pdf"
    End If
    Unload RepEmpaque
    
    frmVerListaEmbarque.Show
    Unload Me
    Exit Sub
errhandler:
    MsgBox "[" & Err.Number & "] " & Err.Description

    Unload RepEmpaque
    
    frmVerListaEmbarque.Show
    frmVerListaEmbarque.cmdRecibirLista.Visible = False
    frmVerListaEmbarque.cmdNuevo.Visible = True
    frmVerListaEmbarque.cmdCambiarOperador.Visible = True
    frmVerListaEmbarque.cmdImprimirListado.Visible = True
    frmVerListaEmbarque.cmdImprimirEtiqueta.Visible = True
    frmVerListaEmbarque.cmdEnviarCorreo.Visible = True
    
    Unload Me

End Sub

Private Sub cmdAgregar_Click()
    AgregarDetalle cmbCliente2.BoundText, UCase(txtDescripcion.Text)
    cmbCliente2.BoundText = ""
    txtDescripcion.Text = ""
End Sub

Private Sub cmdCambiar_Click()
    VSPrinterAUX.PrintDialog pdPrint
    ImpresoraEtiqueta = VSPrinterAUX.Device
    txtImpresora.Text = ImpresoraEtiqueta
End Sub

Private Sub Command1_Click()
    Dim clsAux As New clsGuiaUrbano
    clsAux.Generar "212539"
'Generar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    MSCommBalanza_DesConectar
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If ImprimirEtiquetaDespacho = True Then
        chkImprimirSTK.Value = 1
    Else
        chkImprimirSTK.Value = 0
    End If
End Sub

Private Sub cargarCombos()
    
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
    
    
    strSql = " SELECT cou_codigo, cou_nombre " & _
             " FROM courier " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbCourier.RowSource = clsSql.adorec_Def.DataSource
    cmbCourier.ListField = "cou_nombre"
    cmbCourier.BoundColumn = "cou_codigo"
       
    strSql = " SELECT paq_env_codigo, paq_env_nombre " & _
             " FROM paquete_envio " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By paq_env_codigo "
    clsSql.Ejecutar (strSql)
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFGCajas.ColComboList(2) = VSFGCajas.BuildComboList(clsSql.adorec_Def, "paq_env_codigo, *paq_env_nombre", "paq_env_codigo")
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" And Screen.ActiveControl.Name <> "TxtObserv" Then
        KeyCode = 0
        SendKeys vbKeyTab
    ElseIf (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And Screen.ActiveControl.Name = "TxtObserv" Then
        txtLector.SetFocus
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    txtImpresora.Text = ImpresoraEtiqueta
    PedPendiente = ""
    cargarCombos
    dtpFecha.Value = HoyDia
End Sub

Private Sub optNs_Click()
    cmbNegocio_Change
End Sub

Private Sub optSegunPed_Click()
    cmbNegocio_Change
End Sub

Private Sub optTodos_Click()
    cmbNegocio_Change
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    If SSTab.Tab = 3 Then
        TxtTotal.Visible = False
        cmdAceptar.Visible = False
        cmdSalir.Visible = False
    Else
        TxtTotal.Visible = True
        cmdAceptar.Visible = True
        cmdSalir.Visible = True
    End If
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtLector.Text <> "" Then
        If Me.cmbCliente.BoundText = "" And PedPendiente = "" Then
            PedPendiente = txtLector.Text
            cmbNegocio_Change
            cmbCliente.SetFocus
            txtLector.Text = ""
        Else
            AgregarPedido UCase(txtLector.Text)
            txtLector.Text = ""
        End If
    End If
End Sub

Private Sub AgregarPedido(strPedido As String)
    Dim clsSqlAux As New clsConsulta
    Dim j As Long
    Dim booAgrega As Boolean
    Dim strCXC As String
    Dim dblCXC As Double
    Dim dblRet As Double
    Dim dblCob As Double
    Dim noPasa As Boolean
    strPedido = Trim(strPedido)
    noPasa = EgresoRet(strPedido)
    If noPasa = False Then
        For j = 0 To VSFG.Rows - 1
            If strPedido = VSFG.TextMatrix(j, 0) Then
                MsgBox "El ya esta cargado en este mismo Contenedor." & vbNewLine & _
                       "Verifiquelo. ", vbInformation, "Despachos"
                Exit Sub
            End If
        Next j
        clsSqlAux.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT pedido.ped_codigo,egreso.egr_codigo,egr_fecha,CONCAT(per_apellido,' ',per_nombre) as cli, " & _
                 " LEFT(pedido.ped_direccion_envio,8) as ciu,per_direccion2, ciu_nombre, zon_nombre,for_pag_revisiondespacho,ped_estado,persona.for_pag_codigo,egreso.tip_egr_codigo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND (persona.per_codigo_ref='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref2='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref3='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref4='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref5='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref6='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref7='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref8='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref9='" & cmbCliente.BoundText & "'" & _
                 " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "'"
        If optTodos.Value = True Or optSegunPed.Value = True Then
            strSql = strSql & " OR persona.per_codigo='" & cmbCliente.BoundText & "'"
        End If
        strSql = strSql & ") INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
                 " AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
                 " INNER JOIN egreso ON pedido.emp_codigo=egreso.emp_codigo " & _
                 " AND pedido.ped_egr_codigo=egreso.egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=egreso.tip_egr_codigo " & _
                 " AND pedido.per_codigo=egreso.per_codigo " & _
                 " AND egreso.egr_anulado=0 " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo='" & strPedido & "'"
        clsSql.Ejecutar strSql
        
        If clsSql.adorec_Def.RecordCount > 0 Then
            If clsSql.adorec_Def("ped_estado") = 2 Or clsSql.adorec_Def("ped_estado") = 10 Then
                If FormatoD0(clsSql.adorec_Def("for_pag_revisiondespacho")) = 1 And clsSql.adorec_Def("tip_egr_codigo") <> "NET" Then
                    strSql = " SELECT egr_codigo " & _
                             " FROM egreso_despacho " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND tip_egr_codigo='" & clsSql.adorec_Def("tip_egr_codigo") & "' " & _
                             " AND egr_des_anulado=0 " & _
                             " AND egr_codigo='" & clsSql.adorec_Def("egr_codigo") & "' "
                    clsSqlAux.Ejecutar strSql
                    If clsSqlAux.adorec_Def.RecordCount <= 0 Then
                        strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cue_p_c_valor " & _
                                 " FROM cuenta_p_c " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND cue_p_c_tipo='C' " & _
                                 " AND cue_p_c_egr_codigo='" & clsSql.adorec_Def("egr_codigo") & "' "
                        clsSqlAux.Ejecutar strSql
                        strCXC = clsSqlAux.adorec_Def("cue_p_c_codigo")
                        dblCXC = clsSqlAux.adorec_Def("cue_p_c_valor")
                        strSql = " SELECT COALESCE(com_ret_total,0) as com_ret_total " & _
                                 " FROM comprobante_retencion " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND cue_p_c_tipo='C' " & _
                                 " AND cue_p_c_codigo='" & strCXC & "' "
                        clsSqlAux.Ejecutar strSql
                        If clsSqlAux.adorec_Def.RecordCount > 0 Then
                            dblRet = clsSqlAux.adorec_Def("com_ret_total")
                        Else
                            dblRet = 0
                        End If
                        strSql = " SELECT emp_codigo,cue_p_c_codigo,cue_p_c_tipo,SUM(pag_monto) as monto " & _
                                 " FROM pago " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND cue_p_c_tipo='C' " & _
                                 " AND cue_p_c_codigo='" & strCXC & "' " & _
                                 " GROUP BY emp_codigo,cue_p_c_codigo,cue_p_c_tipo "
                        clsSqlAux.Ejecutar strSql
                        If clsSqlAux.adorec_Def.RecordCount > 0 Then
                            dblCob = clsSqlAux.adorec_Def("monto")
                        Else
                            dblCob = 0
                        End If
                        If FormatoD2(FormatoD2(dblCXC) - FormatoD2(dblRet) - FormatoD2(dblCob)) > 1.5 Then
                            MsgBox "El pedido no puede ser despachado." & vbNewLine & _
                                    "La factura tiene un saldo de " & FormatoD2(FormatoD2(dblCXC) - FormatoD2(dblRet) - FormatoD2(dblCob)) & vbNewLine & _
                                    "O en cartera aun no dan la autorizacion para que salga.", vbInformation, "Despachos"
                            Exit Sub
                        End If
                    End If
                End If
                strSql = " SELECT contenedor.con_codigo,con_fecha " & _
                         " FROM contenedor INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo AND contenedor.con_codigo=det_contenedor.con_codigo" & _
                         " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                         " AND egr_codigo='" & clsSql.adorec_Def("egr_codigo") & "' AND tip_egr_codigo='" & clsSql.adorec_Def("tip_egr_codigo") & "'" & _
                         " AND det_con_estado=0"
                clsSqlAux.Ejecutar strSql
                If clsSqlAux.adorec_Def.RecordCount > 0 Then
                    MsgBox "El pedido no puede ser despachado DOS VECES." & vbNewLine & _
                        "Esta ya incluido en el Contenedor No. " & clsSqlAux.adorec_Def("con_codigo") & vbNewLine & _
                        "realizado el " & clsSqlAux.adorec_Def("con_fecha"), vbInformation, "Despachos"
                Else
                    If clsSql.adorec_Def("for_pag_codigo") = "CONT" And clsSql.adorec_Def("ciu") <> "DIRECTOR" Then
                        If VSFG.Rows = 1 Then
                            booAgrega = True
                        Else
                            booAgrega = False
                            MsgBox "Un pedido de contado debe irse en una sola guia", vbInformation, "Envio"
                        End If
                    Else
                        booAgrega = True
                    End If
                    If booAgrega = True Then
                        VSFG.AddItem clsSql.adorec_Def("ped_codigo") & vbTab & _
                                     clsSql.adorec_Def("egr_codigo") & vbTab & _
                                     clsSql.adorec_Def("egr_fecha") & vbTab & _
                                     clsSql.adorec_Def("cli") & vbTab & _
                                     clsSql.adorec_Def("per_direccion2") & vbTab & _
                                     clsSql.adorec_Def("ciu_nombre") & vbTab & _
                                     clsSql.adorec_Def("zon_nombre") & vbTab & _
                                     clsSql.adorec_Def("tip_egr_codigo")
                        VSFG.ShowCell VSFG.Rows - 1, VSFG.Col
                    End If
                    booAgrega = False
                End If
                TxtTotal.Text = VSFG.Rows - 1 + VSFG2.Rows - 1
            Else
                MsgBox "El pedido no puede ser despachado." & vbNewLine & _
                        "No ha sido confirmado el pedido y escaneadas las prendas ", vbInformation, "Despachos"
            End If
        Else
            MsgBox "El pedido no puede ser despachado." & vbNewLine & _
                    "El cliente no es de la red, error en zona, error en ciudad " & vbNewLine & _
                    "o no tiene factura válida ", vbInformation, "Despachos"
        End If
    Else
        MsgBox "Este pedido no puede salir, FACTURA RETENIDA", vbCritical
    End If
    
End Sub

Private Sub AgregarDetalle(strCliente As String, strDescripcion As String)
    Dim clsSqlAux As New clsConsulta
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT persona.per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli,'" & strDescripcion & "' as descr, " & _
             " per_direccion2, ciu_nombre, zon_nombre " & _
             " FROM persona " & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
             " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND (persona.per_codigo_ref='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref2='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref3='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref4='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref5='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref6='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref7='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref8='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref9='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "'"
    If optTodos.Value = True Or optSegunPed.Value = True Then
        strSql = strSql & " OR persona.per_codigo='" & cmbCliente.BoundText & "'"
    End If
    strSql = strSql & ")" & _
             " AND persona.per_codigo='" & strCliente & "'"
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG2.AddItem clsSql.adorec_Def("per_codigo") & vbTab & _
                     clsSql.adorec_Def("cli") & vbTab & _
                     clsSql.adorec_Def("descr") & vbTab & _
                     clsSql.adorec_Def("per_direccion2") & vbTab & _
                     clsSql.adorec_Def("ciu_nombre") & vbTab & _
                     clsSql.adorec_Def("zon_nombre")
        TxtTotal.Text = VSFG.Rows - 1 + VSFG2.Rows - 1
    Else
        MsgBox "El cliente no es de la red, error en zona, error en ciudad ", vbInformation, "Despachos"
    End If
    
End Sub

Private Sub txtNoCajas_Validate(Cancel As Boolean)
    Dim i As Long
    If IsNumeric(txtNoCajas.Text) = False Then
        txtNoCajas.Text = 1
        MsgBox "Deben ingresar solo numeros", vbCritical + vbOKOnly, "Contenedor"
    End If
    txtNoCajas.Text = Abs(FormatoD0(txtNoCajas.Text))
    VSFGCajas.Rows = txtNoCajas.Text + 1
    For i = 1 To VSFGCajas.Rows - 1
        VSFGCajas.TextMatrix(i, 0) = i
        VSFGCajas.TextMatrix(i, 1) = 0
        VSFGCajas.TextMatrix(i, 2) = ""
    Next i
    
End Sub

Private Function RevisarPesosYTam() As Boolean
    Dim i As Long
    RevisarPesosYTam = True
    For i = 1 To VSFGCajas.Rows - 1
        If FormatoD2(VSFGCajas.TextMatrix(i, 1)) <= 0 Then
            RevisarPesosYTam = False
            VSFGCajas.ShowCell i, 1
            VSFGCajas.Select i, 1
            SSTab.Tab = 2
            VSFGCajas.SetFocus
            Exit For
        End If
        If VSFGCajas.TextMatrix(i, 2) = "" Then
            RevisarPesosYTam = False
            VSFGCajas.ShowCell i, 2
            VSFGCajas.Select i, 1
            SSTab.Tab = 2
            VSFGCajas.SetFocus
            Exit For
        End If
    Next i
End Function


Private Sub txtPeso_GotFocus()
    MSCommBalanza_Conectar
End Sub

Private Sub txtPeso_LostFocus()
    MSCommBalanza_DesConectar
End Sub

Private Sub txtPeso_Validate(Cancel As Boolean)
    If IsNumeric(Replace(txtPeso.Text, ",", ".")) = True Then
        SSTab.Tab = 2
        txtPeso.Text = FormatoD2(Replace(txtPeso.Text, ",", "."))
        txtNoCajas.Text = 1
        VSFGCajas.ShowCell 1, 1
        VSFGCajas.Select 1, 1
        VSFGCajas.TextMatrix(1, 1) = FormatoD2(txtPeso.Text)
        VSFGCajas.SetFocus
    Else
        Cancel = True
    End If
End Sub

Private Sub VSFGCajas_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        VSFGCajas.TextMatrix(Row, Col) = FormatoD2(Replace(VSFGCajas.TextMatrix(Row, Col), ",", "."))
        strSql = " SELECT paq_env_codigo " & _
                 " FROM paquete_envio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND '" & VSFGCajas.TextMatrix(Row, Col) & "' BETWEEN paq_env_peso_min AND paq_env_peso_max " & _
                 " Order By paq_env_codigo "
        clsSql.Ejecutar (strSql)
        If clsSql.adorec_Def.RecordCount > 0 Then
            VSFGCajas.TextMatrix(Row, Col + 1) = clsSql.adorec_Def("paq_env_codigo")
        End If
    End If
End Sub


Private Sub VSFGCajas_GotFocus()
    MSCommBalanza_Conectar
End Sub

Private Sub VSFGCajas_LostFocus()
    MSCommBalanza_DesConectar
End Sub

Private Sub MSCommBalanza_Conectar()
    MSCommBalanza.RThreshold = 1
    MSCommBalanza.RTSEnable = True
    MSCommBalanza.SThreshold = 1
    If MSCommBalanza.PortOpen = False Then
    'determina el puerto que hemos seleccionado
        MSCommBalanza.CommPort = PuertoBalanza
    'determina: 9600-Velocidad en Baudios, N-No utiliza ninguna paridad,
    '8-Cantidad de bits de envio y recepcion por paquete,
    '1-Determina los bits de parada
        MSCommBalanza.Settings = "4800,E,7,1"
    'lee todo el buffer de entrada para que quede vacio
        MSCommBalanza.InputLen = 0
    'Abre el puerto seleccionado
        On Error Resume Next
        MSCommBalanza.PortOpen = True
    'Me.Caption = "Conectado por el puerto " & MSComm1.CommPort
    End If
End Sub

Private Sub MSCommBalanza_DesConectar()
    If MSCommBalanza.PortOpen Then
        'cierra el puerto
        MSCommBalanza.PortOpen = False
    End If
End Sub

Private Sub MSCommBalanza_OnComm()
    Dim i As Integer
    Dim Valor As String
    Dim cadena As String
    
    'recoge el valor de entrada
    Valor = MSCommBalanza.Input
    'busca la posicion del caracter de salto de linea
    i = InStrRev(UCase(Valor), "KG")
    If i = 0 Then
        i = InStrRev(UCase(Valor), "LB")
    End If
    'si no hay ningun salto de linea, quiere decir que la informacion que recibe
    'es parte de una cadena recibida con anterioridad.
    If i < 3 Then
        cadena = cadena & Valor
    Else
        cadena = Left(Valor, i + 1)
        cadena = Right(cadena, Len(cadena) - (InStr(1, cadena, "NET:") + 3))
        If UCase(Right(cadena, 2)) = "KG" Then
            If Screen.ActiveControl.Name = "txtPeso" Then
                txtPeso.Text = FormatoD2(Left(cadena, Len(cadena) - 2))
            Else
                VSFGCajas.TextMatrix(VSFGCajas.Row, 1) = FormatoD2(Left(cadena, Len(cadena) - 2))
            End If
            
        Else
            MsgBox "La balanza esta en " & UCase(Right(cadena, 2)), vbInformation, "Unidad de Medida"
            If Screen.ActiveControl.Name = "txtPeso" Then
                txtPeso.Text = ""
            Else
                VSFGCajas.TextMatrix(VSFGCajas.Row, 1) = ""
            End If
        End If
        cadena = ""
    End If
End Sub

