VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPermisoDespacho 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones de Despacho"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPermisoDespacho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4485
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6135
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Facturas A Despachar"
      TabPicture(0)   =   "frmPermisoDespacho.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "ucrtVSFG"
      Tab(0).Control(4)=   "VSFG"
      Tab(0).Control(5)=   "cmdAgregar"
      Tab(0).Control(6)=   "txtDescripcion"
      Tab(0).Control(7)=   "txtLector"
      Tab(0).Control(8)=   "txtTotal"
      Tab(0).Control(9)=   "cmdAceptar"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Anular Autorizaciones de Despacho"
      TabPicture(1)   =   "frmPermisoDespacho.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "ucrtVSFG2"
      Tab(1).Control(4)=   "VSFG2"
      Tab(1).Control(5)=   "txtLector2"
      Tab(1).Control(6)=   "txtDescripcion2"
      Tab(1).Control(7)=   "cmdAgregar2"
      Tab(1).Control(8)=   "txtTotal2"
      Tab(1).Control(9)=   "cmdAceptar2"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Pedidos a Facturar"
      TabPicture(2)   =   "frmPermisoDespacho.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ucrtVSFG3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "VSFG3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdAceptar3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtTotal3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtLector3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtDescripcion3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdAgregar3"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtValorTotal"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtValorTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar3 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion3 
         Height          =   315
         Left            =   1200
         MaxLength       =   250
         TabIndex        =   30
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtLector3 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar3 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2925
         TabIndex        =   27
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   -72075
         TabIndex        =   26
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -69000
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion2 
         Height          =   315
         Left            =   -73800
         MaxLength       =   250
         TabIndex        =   18
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtLector2 
         Height          =   315
         Left            =   -73800
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   -72075
         TabIndex        =   16
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox txtLector 
         Height          =   315
         Left            =   -73800
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   -73800
         MaxLength       =   250
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -69000
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   -71280
         TabIndex        =   6
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   9
         Top             =   1680
         Width           =   7935
         _cx             =   13996
         _cy             =   6588
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
         FormatString    =   $"frmPermisoDespacho.frx":035E
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
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   1320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   20
         Top             =   1680
         Width           =   7935
         _cx             =   13996
         _cy             =   6588
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
         FormatString    =   $"frmPermisoDespacho.frx":045D
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
      Begin NEED2.uctrVSFG ucrtVSFG2 
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   1320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG3 
         Height          =   3735
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   7935
         _cx             =   13996
         _cy             =   6588
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
         FormatString    =   $"frmPermisoDespacho.frx":051B
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
      Begin NEED2.uctrVSFG ucrtVSFG3 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total pedidos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   5520
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   270
         TabIndex        =   34
         Top             =   555
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Facturas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74880
         TabIndex        =   24
         Top             =   5520
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   23
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Factura:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   22
         Top             =   555
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Factura:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   13
         Top             =   555
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74760
         TabIndex        =   11
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Facturas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74880
         TabIndex        =   10
         Top             =   5520
         Width           =   1080
      End
   End
   Begin MSDataListLib.DataCombo cmbCliente 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   75
      Width           =   4185
      _ExtentX        =   7382
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
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmPermisoDespacho"
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

Private Sub LimpiarForm()
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    txtDescripcion.Text = ""
    TxtTotal.Text = 0
    VSFG2.Clear flexClearScrollable
    VSFG2.Rows = 1
    txtDescripcion2.Text = ""
    txtTotal2.Text = 0
End Sub

Private Sub cmbNegocio_Change()
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
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND (per_es_gz=1 " & _
             " OR per_es_di=1 " & _
             " OR per_es_em=1 " & _
             " OR per_es_ee=1 " & _
             " OR per_es_n5=1 " & _
             " OR per_es_n6=1 " & _
             " OR per_es_n7=1 " & _
             " OR per_es_n8=1 " & _
             " OR per_es_n9=1 " & _
             " OR per_es_n10=1) " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Long
    If FormatoD0(VSFG.Rows) > 1 Then
    
        For i = 1 To VSFG.Rows - 1
            strSql = " SELECT egr_codigo,egr_des_anulado,egr_des_observacion," & _
                     " egr_des_fechamod,egr_des_usumod " & _
                     " FROM egreso_despacho " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND tip_egr_codigo='FAC' " & _
                     " AND egr_codigo='" & VSFG.TextMatrix(i, 0) & "'"
            clsSql.Ejecutar strSql, "M"
            If clsSql.adorec_Def.RecordCount <= 0 Then
                strSql = " INSERT INTO egreso_despacho (emp_codigo,tip_egr_codigo,egr_codigo," & _
                         " egr_des_observacion,egr_des_anulado,egr_des_fechamod,egr_des_usumod) " & _
                         " VALUES('" & strEmpresa & "','FAC','" & VSFG.TextMatrix(i, 0) & "', " & _
                         " '" & VSFG.TextMatrix(i, 5) & "','0',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    clsSql.Ejecutar strSql, "M"
            Else
                If FormatoD0(clsSql.adorec_Def("egr_des_anulado")) = 0 Then
                    MsgBox "La factura " & clsSql.adorec_Def("egr_codigo") & _
                            " ya esta autorizada el " & clsSql.adorec_Def("egr_des_fechamod") & vbNewLine & _
                            "por el usuario " & clsSql.adorec_Def("egr_des_usumod") & vbNewLine & _
                            "Motivo: " & clsSql.adorec_Def("egr_des_observacion")
                Else
                
                    strSql = " UPDATE egreso_despacho " & _
                             " SET egr_des_anulado=0," & _
                             " egr_des_observacion=CONCAT('" & VSFG.TextMatrix(i, 5) & "','\n',egr_des_observacion)," & _
                             " egr_des_fechamod=CURRENT_TIMESTAMP,egr_des_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' AND tip_egr_codigo='FAC' " & _
                             " AND egr_codigo='" & VSFG.TextMatrix(i, 0) & "'"
                    clsSql.Ejecutar strSql, "M"
                End If
            End If
        Next i
        MsgBox "Autorizaciones registradas", vbInformation, "Autorización"
        
        Unload Me
    End If
End Sub

Private Sub cmdAceptar2_Click()
    Dim i As Long
    If FormatoD0(VSFG2.Rows) > 1 Then
    
        For i = 1 To VSFG2.Rows - 1
            strSql = " SELECT egr_codigo,egr_des_anulado,egr_des_observacion," & _
                     " egr_des_fechamod,egr_des_usumod " & _
                     " FROM egreso_despacho " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND tip_egr_codigo='FAC' " & _
                     " AND egr_codigo='" & VSFG2.TextMatrix(i, 0) & "'"
            clsSql.Ejecutar strSql, "M"
            If clsSql.adorec_Def.RecordCount <= 0 Then
                MsgBox "La factura " & VSFG2.TextMatrix(i, 0) & _
                       " NO esta autorizada para el despacho", vbInformation, "Autorización de despacho"
            Else
                If FormatoD0(clsSql.adorec_Def("egr_des_anulado")) = 1 Then
                    MsgBox "La factura " & clsSql.adorec_Def("egr_codigo") & _
                            " ya esta ANULADA la autorizacion para el despacho el " & clsSql.adorec_Def("egr_des_fechamod") & vbNewLine & _
                            "por el usuario " & clsSql.adorec_Def("egr_des_usumod") & vbNewLine & _
                            "Motivo: " & clsSql.adorec_Def("egr_des_observacion")
                Else
                
                    strSql = " UPDATE egreso_despacho " & _
                             " SET egr_des_anulado=1," & _
                             " egr_des_observacion=CONCAT('" & VSFG2.TextMatrix(i, 5) & "','\n',egr_des_observacion)," & _
                             " egr_des_fechamod=CURRENT_TIMESTAMP,egr_des_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' AND tip_egr_codigo='FAC' " & _
                             " AND egr_codigo='" & VSFG2.TextMatrix(i, 0) & "'"
                    clsSql.Ejecutar strSql, "M"
                End If
            End If
        Next i
        MsgBox "Autorizaciones Anuladas", vbInformation, "Autorización"
        
        Unload Me
    End If

End Sub

Private Sub cmdAceptar3_Click()
    Dim i As Long
    If FormatoD0(VSFG3.Rows) > 1 Then
    
        For i = 1 To VSFG3.Rows - 1
            strSql = " SELECT egr_codigo,egr_des_anulado,egr_des_observacion," & _
                     " egr_des_fechamod,egr_des_usumod " & _
                     " FROM egreso_despacho " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND tip_egr_codigo='P' " & _
                     " AND egr_codigo='" & VSFG3.TextMatrix(i, 0) & "'"
            clsSql.Ejecutar strSql, "M"
            If clsSql.adorec_Def.RecordCount <= 0 Then
                strSql = " INSERT INTO egreso_despacho (emp_codigo,tip_egr_codigo,egr_codigo," & _
                         " egr_des_observacion,egr_des_anulado,egr_des_fechamod,egr_des_usumod) " & _
                         " VALUES('" & strEmpresa & "','P','" & VSFG3.TextMatrix(i, 0) & "', " & _
                         " '" & VSFG3.TextMatrix(i, 5) & "','0',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsSql.Ejecutar strSql, "M"
                strSql = " UPDATE pedido " & _
                         " SET ped_estado=1 " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_codigo='" & VSFG3.TextMatrix(i, 0) & "'"
                clsSql.Ejecutar strSql, "M"
            Else
                If FormatoD0(clsSql.adorec_Def("egr_des_anulado")) = 0 Then
                    MsgBox "El pedido " & clsSql.adorec_Def("egr_codigo") & _
                            " ya esta autorizado el " & clsSql.adorec_Def("egr_des_fechamod") & vbNewLine & _
                            "por el usuario " & clsSql.adorec_Def("egr_des_usumod") & vbNewLine & _
                            "Motivo: " & clsSql.adorec_Def("egr_des_observacion")
                Else
                
                    strSql = " UPDATE egreso_despacho " & _
                             " SET egr_des_anulado=0," & _
                             " egr_des_observacion=CONCAT('" & VSFG3.TextMatrix(i, 5) & "','\n',egr_des_observacion)," & _
                             " egr_des_fechamod=CURRENT_TIMESTAMP,egr_des_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' AND tip_egr_codigo='P' " & _
                             " AND egr_codigo='" & VSFG3.TextMatrix(i, 0) & "'"
                    clsSql.Ejecutar strSql, "M"
                    strSql = " UPDATE pedido " & _
                             " SET ped_estado=1 " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND ped_codigo='" & VSFG3.TextMatrix(i, 0) & "'"
                    clsSql.Ejecutar strSql, "M"
                End If
            End If
        Next i
        MsgBox "Autorizaciones registradas", vbInformation, "Autorización"
        Unload Me
    End If

End Sub

Private Sub cmdAgregar_Click()
    AgregarFactura UCase(txtLector.Text), UCase(txtDescripcion.Text)
    txtLector.Text = ""
End Sub

Private Sub cmdAgregar2_Click()
    AgregarFactura2 UCase(txtLector2.Text), UCase(txtDescripcion2.Text)
    txtLector2.Text = ""
End Sub

Private Sub cmdAgregar3_Click()
    AgregarPedido UCase(txtLector3.Text), UCase(txtDescripcion3.Text)
    txtLector3.Text = ""
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    
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
        
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" And Screen.ActiveControl.Name <> "txtLector2" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    Set ucrtVSFG2.VSFGControl = VSFG2
    Set ucrtVSFG3.VSFGControl = VSFG3
    ucrtVSFG.Inicializar False, False, False
    ucrtVSFG2.Inicializar False, False, False
    ucrtVSFG3.Inicializar False, False, False
    
    cargarCombos
End Sub


Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarFactura UCase(txtLector.Text), UCase(txtDescripcion.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub txtLector2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarFactura2 UCase(txtLector2.Text), UCase(txtDescripcion2.Text)
        txtLector2.Text = ""
    End If
End Sub

Private Sub AgregarFactura(strFactura As String, strObservacion As String)
    Dim clsSqlAux As New clsConsulta
    strSql = " SELECT egreso.egr_codigo,egr_fecha,for_pag_nombre," & _
             " CONCAT(per_apellido,' ',per_nombre) as cli, egr_total,per_direccion2 " & _
             " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
             " AND egreso.per_codigo=persona.per_codigo " & _
             " AND (persona.per_codigo_ref='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref2='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref3='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref4='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref5='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref6='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref7='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref8='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref9='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "')" & _
             " AND persona.per_bloqueado_g=0 AND persona.per_bloqueado=0 " & _
             " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_anulado=0 " & _
             " AND egreso.egr_codigo='" & strFactura & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG.AddItem clsSql.adorec_Def("egr_codigo") & vbTab & _
                     clsSql.adorec_Def("egr_fecha") & vbTab & _
                     clsSql.adorec_Def("for_pag_nombre") & vbTab & _
                     clsSql.adorec_Def("cli") & vbTab & _
                     clsSql.adorec_Def("egr_total") & vbTab & _
                     UCase(strObservacion) & vbTab & _
                     cmbCliente.Text & vbTab & _
                     clsSql.adorec_Def("per_direccion2")
        VSFG.ShowCell VSFG.Rows - 1, VSFG.Col
        TxtTotal.Text = VSFG.Rows - 1
    Else
        MsgBox "La factura no puede ser autorizada." & vbNewLine & _
                "El cliente no es de la red, " & vbNewLine & _
                "El cliente esta bloqueado " & vbNewLine & _
                "o no tiene factura válida ", vbInformation, "Autorizacion despacho"
    End If
End Sub


Private Sub AgregarFactura2(strFactura As String, strObservacion As String)
    Dim clsSqlAux As New clsConsulta
    strSql = " SELECT egreso.egr_codigo,egr_fecha,for_pag_nombre," & _
             " CONCAT(per_apellido,' ',per_nombre) as cli, egr_total " & _
             " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
             " AND egreso.per_codigo=persona.per_codigo " & _
             " AND (persona.per_codigo_ref='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref2='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref3='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref4='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref5='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref6='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref7='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref8='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref9='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "')" & _
             " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_anulado=0 " & _
             " AND egreso.egr_codigo='" & strFactura & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG2.AddItem clsSql.adorec_Def("egr_codigo") & vbTab & _
                     clsSql.adorec_Def("egr_fecha") & vbTab & _
                     clsSql.adorec_Def("for_pag_nombre") & vbTab & _
                     clsSql.adorec_Def("cli") & vbTab & _
                     clsSql.adorec_Def("egr_total") & vbTab & _
                     UCase(strObservacion)
        txtTotal2.Text = VSFG2.Rows - 1
    Else
        MsgBox "La factura no puede ser Anulada la autorización." & vbNewLine & _
                "El cliente no es de la red " & vbNewLine & _
                "o no tiene factura válida ", vbInformation, "Autorizacion despacho"
    End If
End Sub

Private Sub AgregarPedido(strPedido As String, strObservacion As String)
    Dim clsSqlAux As New clsConsulta
    Dim i As Long
    strSql = " SELECT pedido.ped_codigo,ped_fechamod,for_pag_nombre," & _
             " CONCAT(per_apellido,' ',per_nombre) as cli, ped_subtotal,ped_direccion_envio " & _
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
             " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "')" & _
             " AND persona.per_bloqueado_g=0 AND persona.per_bloqueado=0 " & _
             " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND pedido.ped_estado=0 " & _
             " AND pedido.ped_codigo='" & strPedido & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG3.AddItem clsSql.adorec_Def("ped_codigo") & vbTab & _
                     clsSql.adorec_Def("ped_fechamod") & vbTab & _
                     clsSql.adorec_Def("for_pag_nombre") & vbTab & _
                     clsSql.adorec_Def("cli") & vbTab & _
                     clsSql.adorec_Def("ped_subtotal") & vbTab & _
                     UCase(strObservacion) & vbTab & _
                     cmbCliente.Text & vbTab & _
                     clsSql.adorec_Def("ped_direccion_envio")
        VSFG3.ShowCell VSFG3.Rows - 1, VSFG3.Col
        txtTotal3.Text = VSFG3.Rows - 1
        txtValorTotal.Text = 0#
        For i = 1 To VSFG3.Rows - 1
            txtValorTotal.Text = FormatoD2(txtValorTotal.Text) + FormatoD2(VSFG3.TextMatrix(i, 4))
        Next i
    Else
        MsgBox "El pedido no puede ser autorizado." & vbNewLine & _
                "El cliente no es de la red, " & vbNewLine & _
                "El cliente esta bloqueado " & vbNewLine & _
                "o no tiene pedido válida ", vbInformation, "Autorizacion despacho"
    End If
End Sub


Private Sub txtLector3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarPedido UCase(txtLector3.Text), UCase(txtDescripcion3.Text)
        txtLector3.Text = ""
    End If
End Sub
