VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmATED 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NEED ANEX                                                               Anexos Transaccionales - ENLACE DIGITAL"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "frmATED.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   12075
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbMesI 
      Height          =   315
      ItemData        =   "frmATED.frx":030A
      Left            =   5760
      List            =   "frmATED.frx":0335
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   240
      Width           =   1425
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   495
      Left            =   3720
      Picture         =   "frmATED.frx":039E
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   2520
      Picture         =   "frmATED.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10800
      Picture         =   "frmATED.frx":09CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   1320
      Picture         =   "frmATED.frx":0AE2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   120
      Picture         =   "frmATED.frx":1014
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin TabDlg.SSTab sstI 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
      Top             =   720
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1. Identificación del Informante"
      TabPicture(0)   =   "frmATED.frx":1546
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdImportarInformante"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExportarInformante"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vsfgDescInformante"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfgInformante"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "2. Transacciones Locales - COMPRAS"
      TabPicture(1)   =   "frmATED.frx":1562
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfgCompras"
      Tab(1).Control(1)=   "vsfgDescCompras"
      Tab(1).Control(2)=   "cmdImportarCompras"
      Tab(1).Control(3)=   "cmdExportarCompras"
      Tab(1).Control(4)=   "chkAutoBusquedaCompras"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "3. Transacciones Locales - VENTAS"
      TabPicture(2)   =   "frmATED.frx":157E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsfgVentas"
      Tab(2).Control(1)=   "vsfgDescVentas"
      Tab(2).Control(2)=   "cmdImportarVentas"
      Tab(2).Control(3)=   "cmdExportarVentas"
      Tab(2).Control(4)=   "cmdJuntar"
      Tab(2).Control(5)=   "chkAutoBusquedaVentas"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "4. Transacciones del Exterior - IMPORTACIONES O PAGOS AL EXTERIOR"
      TabPicture(3)   =   "frmATED.frx":159A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "vsfgDescImportaciones"
      Tab(3).Control(1)=   "vsfgImportaciones"
      Tab(3).Control(2)=   "cmdImportarImportaciones"
      Tab(3).Control(3)=   "cmdExportarImportaciones"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "5. Transacciones del Exterior - EXPORTACIONES O INGRESOS DEL EXTERIOR"
      TabPicture(4)   =   "frmATED.frx":15B6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "vsfgDescExportaciones"
      Tab(4).Control(1)=   "vsfgExportaciones"
      Tab(4).Control(2)=   "cmdImportarExportaciones"
      Tab(4).Control(3)=   "cmdExportarExportaciones"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "6. EMPRESAS EMISORAS DE TARJETAS DE CREDITO"
      TabPicture(5)   =   "frmATED.frx":15D2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "vsfgDescTC"
      Tab(5).Control(1)=   "vsfgTC"
      Tab(5).Control(2)=   "cmdImportarTC"
      Tab(5).Control(3)=   "cmdExportarTC"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "7. Administradoras de Fondos y Fideicomisos: FIDEICOMISOS"
      TabPicture(6)   =   "frmATED.frx":15EE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "vsfgDescFideicomisos"
      Tab(6).Control(1)=   "vsfgFideicomisos"
      Tab(6).Control(2)=   "cmdImportarFideicomisos"
      Tab(6).Control(3)=   "cmdExportarFideicomisos"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "8. COMPROBANTES ANULADOS"
      TabPicture(7)   =   "frmATED.frx":160A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "VSFGAnulados"
      Tab(7).Control(1)=   "vsfgDescAnulados"
      Tab(7).Control(2)=   "cmdImportarAnulados"
      Tab(7).Control(3)=   "cmdExportarAnulados"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "9. RENDIMIENTOS FINANCIEROS"
      TabPicture(8)   =   "frmATED.frx":1626
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "vsfgDescRendimientos"
      Tab(8).Control(1)=   "vsfgRendimientos"
      Tab(8).Control(2)=   "cmdImportarRendimientos"
      Tab(8).Control(3)=   "cmdExportarRendimientos"
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "10. GENERACION Y ADMINISTRACION XML"
      TabPicture(9)   =   "frmATED.frx":1642
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cmdGenerarXML"
      Tab(9).Control(1)=   "WebBrow"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "11. REOC"
      TabPicture(10)  =   "frmATED.frx":165E
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "VSFGREOC"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "ucrtVSFG"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "12. Ventas por Establecimiento"
      TabPicture(11)  =   "frmATED.frx":167A
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "cmdImportarVentasEstablecimiento"
      Tab(11).Control(1)=   "cmdExportarVentasEstablecimiento"
      Tab(11).Control(2)=   "vsfgVentasEstablecimiento"
      Tab(11).Control(3)=   "vsfgDescVentasEstablecimiento"
      Tab(11).ControlCount=   4
      Begin VB.CheckBox chkAutoBusquedaCompras 
         Caption         =   "Auto Busqueda"
         Height          =   255
         Left            =   -65520
         TabIndex        =   55
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoBusquedaVentas 
         Caption         =   "Auto Busqueda"
         Height          =   255
         Left            =   -65520
         TabIndex        =   54
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdImportarVentasEstablecimiento 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":1696
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarVentasEstablecimiento 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":1763
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   -74880
         TabIndex        =   49
         Top             =   2160
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgInformante 
         Height          =   4035
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescInformante 
         Height          =   765
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VB.CommandButton cmdJuntar 
         Caption         =   "Juntar"
         Height          =   495
         Left            =   -72120
         Picture         =   "frmATED.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarRendimientos 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":2074
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarAnulados 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":2453
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarFideicomisos 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":2832
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarTC 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":2C11
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarExportaciones 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":2FF0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarImportaciones 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":33CF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarVentas 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":37AE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarCompras 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   -73680
         Picture         =   "frmATED.frx":3B8D
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdExportarInformante 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   1320
         Picture         =   "frmATED.frx":3F6C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exportar a un Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarRendimientos 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":434B
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarAnulados 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":4418
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarFideicomisos 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":44E5
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarTC 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":45B2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarExportaciones 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":467F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarImportaciones 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":474C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarVentas 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":4819
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarCompras 
         Caption         =   "Importar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":48E6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdImportarInformante 
         Caption         =   "Importar"
         Height          =   495
         Left            =   120
         Picture         =   "frmATED.frx":49B3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Importar de Archivo separado por Tabulaciones"
         Top             =   1860
         Width           =   975
      End
      Begin SHDocVwCtl.WebBrowser WebBrow 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   6
         Top             =   2460
         Width           =   10815
         ExtentX         =   19076
         ExtentY         =   8916
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CommandButton cmdGenerarXML 
         Caption         =   "Generar"
         Height          =   495
         Left            =   -74880
         Picture         =   "frmATED.frx":4A80
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Generación del Archivo XML"
         Top             =   1860
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescCompras 
         Height          =   765
         Left            =   -74880
         TabIndex        =   28
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescVentas 
         Height          =   765
         Left            =   -74880
         TabIndex        =   29
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgImportaciones 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   30
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescImportaciones 
         Height          =   765
         Left            =   -74880
         TabIndex        =   31
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgExportaciones 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   32
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescExportaciones 
         Height          =   765
         Left            =   -74880
         TabIndex        =   33
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgTC 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   34
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescTC 
         Height          =   765
         Left            =   -74880
         TabIndex        =   35
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgFideicomisos 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   36
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescFideicomisos 
         Height          =   765
         Left            =   -74880
         TabIndex        =   37
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescAnulados 
         Height          =   765
         Left            =   -74880
         TabIndex        =   38
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgRendimientos 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   39
         Top             =   3240
         Width           =   10875
         _cx             =   113986286
         _cy             =   113974221
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescRendimientos 
         Height          =   765
         Left            =   -74880
         TabIndex        =   40
         Top             =   2400
         Width           =   10875
         _cx             =   113986286
         _cy             =   113968453
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfgCompras 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   42
         Top             =   3240
         Width           =   10875
         _cx             =   2003782382
         _cy             =   2003770317
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGREOC 
         Height          =   4515
         Left            =   -74880
         TabIndex        =   46
         Top             =   2640
         Width           =   10755
         _cx             =   2003782171
         _cy             =   2003771164
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8Ctl.VSFlexGrid vsfgVentas 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   47
         Top             =   3240
         Width           =   10875
         _cx             =   2003782382
         _cy             =   2003770317
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGAnulados 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   48
         Top             =   3240
         Width           =   10875
         _cx             =   2003782382
         _cy             =   2003770317
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgVentasEstablecimiento 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   50
         Top             =   3240
         Width           =   10875
         _cx             =   51727086
         _cy             =   51715021
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   3
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDescVentasEstablecimiento 
         Height          =   765
         Left            =   -74880
         TabIndex        =   51
         Top             =   2400
         Width           =   10875
         _cx             =   51727086
         _cy             =   51709253
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   9480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker AñoI 
      Height          =   315
      Left            =   7185
      TabIndex        =   44
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
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
      CustomFormat    =   "yyyyXX"
      Format          =   66584579
      UpDown          =   -1  'True
      CurrentDate     =   38054
   End
   Begin VB.Label lblFI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha inicio: "
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4800
      TabIndex        =   45
      Top             =   360
      Width           =   945
   End
   Begin VB.Image imgHelp 
      Height          =   240
      Left            =   8760
      Picture         =   "frmATED.frx":4FEE
      ToolTipText     =   "Elimina una Fila"
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   10200
      Picture         =   "frmATED.frx":50F0
      ToolTipText     =   "Elimina una Fila"
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmATED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CambTam As Boolean
Private clsCon_Def As New clsConsulta
Private strSQL As String
Dim cargado As Boolean

Private Sub chkAutoBusquedaVentas_Click()
    If chkAutoBusquedaVentas.Value = 1 Then
        vsfgVentas.Editable = flexEDNone
        vsfgVentas.AutoSearch = flexSearchFromTop
    Else
        vsfgVentas.Editable = flexEDKbdMouse
        vsfgVentas.AutoSearch = flexSearchNone
    End If
End Sub
Private Sub chkAutoBusquedaCompras_Click()
    If chkAutoBusquedaCompras.Value = 1 Then
        vsfgCompras.Editable = flexEDNone
        vsfgCompras.AutoSearch = flexSearchFromTop
    Else
        vsfgCompras.Editable = flexEDKbdMouse
        vsfgCompras.AutoSearch = flexSearchNone
    End If
End Sub

Private Sub cmdAbrir_Click()
    Dim Archivo As String
    CD.DefaultExt = "ated"
    CD.Filter = "Archivos ATED (*.ated)|*.ated|Todos |*.*"
    CD.FilterIndex = 1
    CD.FileName = ""
    CD.ShowOpen
    Archivo = CD.FileName
    If Archivo <> "" Then
        CambTam = False
        Abrir vsfgInformante, "Informante", Archivo
        Abrir vsfgCompras, "Compras", Archivo
        Abrir vsfgVentas, "Ventas", Archivo
        Abrir vsfgVentasEstablecimiento, "VentasEstablecimiento", Archivo
        Abrir vsfgImportaciones, "Importaciones", Archivo
        Abrir vsfgExportaciones, "Exportaciones", Archivo
        Abrir vsfgTC, "TC", Archivo
        Abrir vsfgFideicomisos, "Fideicomisos", Archivo
        Abrir VSFGAnulados, "Anulados", Archivo
        Abrir vsfgRendimientos, "Rendimientos", Archivo
        CambTam = True
        'Form_Load
    End If
End Sub

Private Sub cmdCargar_Click()
    If sstI.Tab = 0 Or sstI.Tab = 1 Or sstI.Tab = 2 Or sstI.Tab = 7 Or sstI.Tab = 10 Or sstI.Tab = 11 Then
    Dim fechaINI As Date, fechaFIN As Date
    Dim cls_Aux As New clsConsulta
    Dim i As Long, rep As Integer
'    fechaINI = Format(Date, "yyyy-MM-01")
'    fechaFIN = DateAdd("M", 1, fechaINI) - 1
    cls_Aux.Inicializar AdoConn, AdoConnMaster
    On Error GoTo errhandler
    fechaINI = Format(AñoI.Year & "-" & cmbMesI.ListIndex + 1 & "-01", "yyyy-MM-dd")
    fechaFIN = DateAdd("M", 1, fechaINI) - 1
    cargado = True
    If sstI.Tab = 0 Then
        vsfgInformante.Clear 1
        vsfgInformante.Rows = 2
        vsfgInformante.Cell(flexcpBackColor, 1, 0, vsfgInformante.Rows - 1, vsfgInformante.Cols - 1) = vbWhite
        strSQL = " SELECT emp_nombre,emp_ruc " & _
                 " FROM empresa " & _
                 " WHERE emp_codigo='" & strEmpresa & "' "
        clsCon_Def.Ejecutar strSQL
        vsfgInformante.TextMatrix(1, 1) = "R"
        vsfgInformante.TextMatrix(1, 2) = clsCon_Def.adorec_Def("emp_ruc")
        vsfgInformante.TextMatrix(1, 3) = clsCon_Def.adorec_Def("emp_nombre")
        vsfgInformante.TextMatrix(1, 4) = AñoI.Year
        vsfgInformante.TextMatrix(1, 5) = Format(cmbMesI.ListIndex + 1, "00")
        vsfgInformante.TextMatrix(1, 6) = "001"
        vsfgInformante.TextMatrix(1, 7) = 0#
        vsfgInformante.TextMatrix(1, 8) = "IVA"
    ElseIf sstI.Tab = 10 Then
        VSFGREOC.Clear 1
        VSFGREOC.Rows = 2
        VSFGREOC.Cell(flexcpBackColor, 1, 0, VSFGREOC.Rows - 1, VSFGREOC.Cols - 1) = vbWhite

        '" IIF(cue_p_c_st_prod+cue_p_c_st_serv=0 and cue_p_c_st_cero=0,cuenta_p_c.cue_p_c_valor/1.12,cue_p_c_st_prod+cue_p_c_st_serv) as baseImpgrav, " &
        strSQL = " SELECT FORMAT(cuenta_p_c.cue_p_c_fechaemision,'%m/%Y') as anomes, '01' as TipoItem1, " & _
                 " IIF(LEN(per_ruc)=13,'01','02') as tpIdProv, per_ruc as idProv," & _
                 " CONCAT(per_apellido, ' ',per_nombre) as proveedor, tip_doc_cue_codigo as tipoComprobante," & _
                 " cue_p_c_autorizacion as autorizacion, LEFT(cue_p_c_serie,3) as establecimiento," & _
                 " RIGHT(cue_p_c_serie,3) as puntoEmision, cue_p_c_numero as secuencial," & _
                 " FORMAT(cuenta_p_c.cue_p_c_fechaemision,'dd/MM/yyyy') as fechaEmision, COALESCE(com_ret_autorizacion,'') as autRetencion," & _
                 " COALESCE(LEFT(com_ret_serie,3),'') as estabRetencion, COALESCE(RIGHT(com_ret_serie,3),'') as ptoEmiRetencion," & _
                 " COALESCE(com_ret_numero,'') as secRetencion, COALESCE(FORMAT(com_ret_fecha,'dd/MM/yyyy'),'') as fechaEmiRet," & _
                 " COALESCE(com_ret_autorizacion,'') as autRetencion1, COALESCE(LEFT(com_ret_serie,3),'') as estabRetencion1," & _
                 " COALESCE(RIGHT(com_ret_serie,3),'') as ptoEmiRetencion1, COALESCE(com_ret_numero,'') as secRetencion1," & _
                 " COALESCE(FORMAT(com_ret_fecha,'dd/MM/yyyy'),'') as fechaEmiRet1," & _
                 " '02' as TipoItem2, COALESCE(codRetAir,'332') as codRetAir, " & _
                 " COALESCE(porcentajeAir,'0') as porcentajeAir, cue_p_c_st_cero as base0," & _
                 " baseImpAir as baseImpgrav, " & _
                 " '0' as baseImpNOgrav, COALESCE(valRetAir,'0') As valRetAir " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo" & _
                 " INNER JOIN codigo_iva ON cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo " & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo" & _
                 " LEFT JOIN DetRetencionIRREOC ret ON comprobante_retencion.emp_codigo=ret.emp_codigo " & _
                 " AND comprobante_retencion.cue_p_c_codigo=ret.cue_p_c_codigo " & _
                 " AND comprobante_retencion.cue_p_c_tipo=ret.cue_p_c_tipo" & _
                 " AND ret.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "'"
        strSQL = strSQL & "LEFT JOIN DetRetencionIVAREOC compra ON compra.emp_codigo=comprobante_retencion.emp_codigo " & _
                 " AND compra.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " AND compra.cue_p_c_tipo= comprobante_retencion.cue_p_c_tipo" & _
                 " AND compra.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "'" & _
                 " WHERE cuenta_p_c.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "'" & _
                 " AND cuenta_p_c.cue_p_c_tipo='P' AND tip_doc_cue_codigo>0" & _
                 " AND cue_p_c_valor!=0 AND cuenta_p_c.emp_codigo='" & strEmpresa & "' and tip_doc_cue_codigo>0 " & _
                 " ORDER BY FORMAT(cuenta_p_c.cue_p_c_fechaemision,'%Y/%m'),per_ruc,cue_p_c_serie,cue_p_c_numero"
        clsCon_Def.Ejecutar strSQL
        
        Set VSFGREOC.DataSource = clsCon_Def.adorec_Def.DataSource
    
    ElseIf sstI.Tab = 7 Then
    'ANULADOS
            VSFGAnulados.Clear 1
            VSFGAnulados.Rows = 2
            VSFGAnulados.Cell(flexcpBackColor, 1, 0, VSFGAnulados.Rows - 1, VSFGAnulados.Cols - 1) = vbWhite
            
            strSQL = " SELECT '        ' as Item,'01' as tipoComprobante,LEFT(egr_serie,3) as establecimiento,RIGHT(egr_serie,3) as puntoEmision,RIGHT(egr_codigo,7) as secuencialInicio,RIGHT(egr_codigo,7) as secuencialFin,egr_autorizacion as autorizacion" & _
                    " FROM egreso " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND tip_egr_codigo='FAC' AND egr_anulado=1" & _
                    " AND egr_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " UNION " & _
                    " SELECT '        ' as Item,'04' ,LEFT(ing_serie,3) as estab,RIGHT(ing_serie,3) as emision,ing_numero as num,ing_numero,ing_autorizacion" & _
                    " FROM ingreso " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND tip_ing_codigo='DCL' AND ing_anulado=1" & _
                    " AND ing_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " ORDER BY tipoComprobante,establecimiento,puntoEmision,secuencialInicio"
            
            clsCon_Def.Ejecutar strSQL
            
            Set VSFGAnulados.DataSource = clsCon_Def.adorec_Def.DataSource
            
            Boton VSFGAnulados
            NuevaLinea vsfgDescAnulados, VSFGAnulados, VSFGAnulados.Rows - 1

    ElseIf sstI.Tab = 2 Then
        vsfgVentas.Clear 1
        vsfgVentas.Rows = 2
        vsfgVentas.Cell(flexcpBackColor, 1, 0, vsfgVentas.Rows - 1, vsfgVentas.Cols - 1) = vbWhite

            strSQL = " EXEC Sp_Drop_Table_if_Exist '#venta' "
            clsCon_Def.Ejecutar strSQL
            
            strSQL = " CREATE TABLE #venta (emp_codigo char(3),cue_p_c_codigo decimal(11,0),cue_p_c_tipo char(2)," & _
                     " ivaret decimal(14,2)) "
            clsCon_Def.Ejecutar strSQL
            
            strSQL = " INSERT into #venta " & _
                     " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo, " & _
                     " SUM(COALESCE(ROUND(RIVP.det_com_ret_valor*RIVP.det_com_ret_porcentaje/100.00,2),0)) AS ivaret " & _
                     " From cuenta_p_c " & _
                     " INNER JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo  AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                     " AND pago.pag_observacion NOT LIKE '%ANULAD%' " & _
                     " AND pago.pag_observacion LIKE '%(CON RETENCIÓN)%' " & _
                     " AND pago.pag_monto=0 " & _
                     " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo  AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                     " INNER JOIN asiento ON pago.emp_codigo=asiento.emp_codigo " & _
                     " AND pago.asi_numasiento=asiento.asi_numasiento " & _
                     " AND (comprobante_retencion.com_ret_total=asiento.asi_totaldebe or comprobante_retencion.com_ret_total=asiento.asi_totalhaber)" & _
                     " INNER JOIN det_comp_ret as RIVP ON cuenta_p_c.emp_codigo=RIVP.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=RIVP.cue_p_c_codigo  AND cuenta_p_c.cue_p_c_tipo=RIVP.cue_p_c_tipo " & _
                     " AND RIVP.ret_codigo<100 " & _
                     " INNER JOIN retencion AS retencion2 ON retencion2.ret_codigo=RIVP.ret_codigo " & _
                     " AND retencion2.emp_codigo=RIVP.emp_codigo AND retencion2.ret_activo=1 " & _
                     " INNER JOIN codigo_iva on cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo and codigo_iva.cod_iva_codigo='2' " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND asiento.asi_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C'  AND tip_doc_cue_codigo>0  AND cue_p_c_valor!=0 " & _
                     " GROUP BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo"
            clsCon_Def.Ejecutar strSQL
            

            strSQL = " EXEC Sp_Drop_Table_if_Exist '#ret' "
            clsCon_Def.Ejecutar strSQL
            
            strSQL = " CREATE TABLE #ret (emp_codigo char(3),cue_p_c_codigo decimal(11,0),cue_p_c_tipo char(2)," & _
                     " baseImpAir decimal(14,2),valRetAir decimal(16,5)) "
            clsCon_Def.Ejecutar strSQL
            
            strSQL = " insert into #ret " & _
                     " SELECT RIR.emp_codigo,RIR.cue_p_c_codigo,RIR.cue_p_c_tipo, " & _
                     " SUM(COALESCE(RIR.det_com_ret_valor,'')) AS baseImpAir, " & _
                     " SUM(COALESCE(ROUND(RIR.det_com_ret_valor*RIR.det_com_ret_porcentaje/100.00,5),'')) AS valRetAir " & _
                     " From cuenta_p_c " & _
                     " INNER JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo  AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                     " AND pago.pag_observacion NOT LIKE '%ANULAD%' " & _
                     " AND pago.pag_observacion LIKE '%(CON RETENCIÓN)%' " & _
                     " AND pago.pag_monto=0 " & _
                     " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo  AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                     " INNER JOIN asiento ON pago.emp_codigo=asiento.emp_codigo " & _
                     " AND pago.asi_numasiento=asiento.asi_numasiento " & _
                     " AND (comprobante_retencion.com_ret_total=asiento.asi_totaldebe or comprobante_retencion.com_ret_total=asiento.asi_totalhaber)" & _
                     " INNER JOIN det_comp_ret as RIR " & _
                     " ON cuenta_p_c.emp_codigo=RIR.emp_codigo AND cuenta_p_c.cue_p_c_tipo=RIR.cue_p_c_tipo " & _
                     " AND cuenta_p_c.cue_p_c_codigo =RIR.cue_p_c_codigo " & _
                     " INNER JOIN retencion ON retencion.ret_codigo=RIR.ret_codigo AND retencion.emp_codigo=RIR.emp_codigo " & _
                     " AND RIR.ret_codigo>100 AND retencion.ret_activo=1 " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND asiento.asi_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C' AND tip_doc_cue_codigo>0  AND cue_p_c_valor!=0 GROUP BY RIR.emp_codigo,RIR.cue_p_c_codigo,RIR.cue_p_c_tipo"
            clsCon_Def.Ejecutar strSQL

           strSQL = " SELECT cue_p_c_egr_codigo as Item," & _
                    " IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) as tpIdCliente," & _
                    " IIF(LEFT(per_ruc,1)='P',SUBSTRING(per_ruc,2,13),per_ruc) as idCliente,IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('04','05','06'),'NO','') AS parteRel," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),IIF(per_tipo='Natural','01','02'),'')AS tipoCliente," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),CONCAT(per_apellido,' ',per_nombre),'')AS DenoCli," & _
                    " '18' as tipoComprobante, 'E' as tipoEm, 1 as numeroComprobantes, " & _
                    " 0.00 as baseNoGraIva, ROUND(cue_p_c_st_cero+IIF(cue_p_c_st_cero=0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0),0),2) as baseImponible, " & _
                    " ROUND(egr_subtotal-egr_dcto-IIF(cue_p_c_st_cero=0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0),0),2) as baseImpGrav," & _
                    " ROUND(cue_p_c_iva,2) as montoIva,'00' as tipoCompe,0.00 as monto," & _
                    " 0.00 as montoIce, 0.00 as valorRetIva, 0.00 as valorRetRenta, " & _
                    " CAST(COALESCE(for_cob_codigo,'01') as varchar) as formaPago"
            strSQL = strSQL & " From cuenta_p_c INNER JOIN egreso ON cuenta_p_c.emp_codigo=egreso.emp_codigo " & _
                    " AND cuenta_p_c.cue_p_c_egr_codigo=egreso.egr_codigo " & _
                    " AND cuenta_p_c.per_codigo=egreso.per_codigo  " & _
                    " AND egreso.tip_egr_codigo='FAC' AND egreso.egr_anulado=0 " & _
                    " AND egr_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " LEFT JOIN egreso_forma_cobro ON egreso.emp_codigo=egreso_forma_cobro.emp_codigo" & _
                    " AND egreso.tip_egr_codigo=egreso_forma_cobro.tip_egr_codigo" & _
                    " AND egreso.egr_codigo=egreso_forma_cobro.egr_codigo" & _
                    " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                    " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                    " AND egreso.egr_codigo=det_egreso.egr_codigo AND det_egreso.prd_codigo='PR-FLET-DES'" & _
                    " INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                    " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                    " INNER JOIN codigo_iva ON cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo "
            strSQL = strSQL & " WHERE cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " AND cuenta_p_c.cue_p_c_tipo='C'  AND tip_doc_cue_codigo>0  AND cue_p_c_valor!=0 " & _
                    " AND cuenta_p_c.emp_codigo='" & strEmpresa & "'  "
            strSQL = strSQL & " UNION "
            strSQL = strSQL & " SELECT cue_p_c_egr_codigo as Item," & _
                    " IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) as tpIdCliente," & _
                    " IIF(LEFT(per_ruc,1)='P',SUBSTRING(per_ruc,2,13),per_ruc) as idCliente,IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('04','05','06'),'NO','') AS parteRel," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),IIF(per_tipo='Natural','01','02'),'')AS tipoCliente," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),CONCAT(per_apellido,' ',per_nombre),'')AS DenoCli," & _
                    " '18' as tipoComprobante,'E'  as tipoEm, 0 as numeroComprobantes, " & _
                    " 0 as baseNoGraIva, 0 as baseImponible, " & _
                    " 0 as baseImpGrav," & _
                    " 0 as montoIva,'00' as tipoCompe,0.00 as monto, " & _
                    " 0 as montoIce,COALESCE(ivaret,0) as valorRetIva,COALESCE((valRetAir),0) as valorRetRenta, " & _
                    " CAST(COALESCE(for_cob_codigo,'01') as varchar) as formaPago "
            strSQL = strSQL & " From cuenta_p_c INNER JOIN egreso ON cuenta_p_c.emp_codigo=egreso.emp_codigo " & _
                    " AND cuenta_p_c.cue_p_c_egr_codigo=egreso.egr_codigo " & _
                    " AND cuenta_p_c.per_codigo=egreso.per_codigo  " & _
                    " AND egreso.tip_egr_codigo='FAC' AND egreso.egr_anulado=0 " & _
                    " LEFT JOIN egreso_forma_cobro ON egreso.emp_codigo=egreso_forma_cobro.emp_codigo" & _
                    " AND egreso.tip_egr_codigo=egreso_forma_cobro.tip_egr_codigo" & _
                    " AND egreso.egr_codigo=egreso_forma_cobro.egr_codigo" & _
                    " INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                    " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                    " INNER JOIN codigo_iva ON cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo "
            strSQL = strSQL & " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                    " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                    " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                    " AND comprobante_retencion.cue_p_c_tipo='C' " & _
                    " INNER JOIN #ret ret ON comprobante_retencion.emp_codigo=ret.emp_codigo " & _
                    " AND comprobante_retencion.cue_p_c_codigo=ret.cue_p_c_codigo " & _
                    " AND comprobante_retencion.cue_p_c_tipo=ret.cue_p_c_tipo " & _
                    " LEFT JOIN #venta venta ON comprobante_retencion.emp_codigo=venta.emp_codigo " & _
                    " AND comprobante_retencion.cue_p_c_codigo=venta.cue_p_c_codigo " & _
                    " AND comprobante_retencion.cue_p_c_tipo=venta.cue_p_c_tipo "
            strSQL = strSQL & " WHERE cuenta_p_c.cue_p_c_tipo='C'  AND tip_doc_cue_codigo>0  AND cue_p_c_valor!=0 " & _
                    " AND cuenta_p_c.emp_codigo='" & strEmpresa & "'  "
            strSQL = strSQL & " UNION "
            strSQL = strSQL & " SELECT CAST(ing_codigo as varchar) as Item," & _
                    " IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) as tpIdCliente," & _
                    " IIF(LEFT(per_ruc,1)='P',SUBSTRING(per_ruc,2,13),per_ruc) as idCliente,IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('04','05','06'),'NO','') AS parteRel," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),IIF(per_tipo='Natural','01','02'),'')AS tipoCliente," & _
                    " IIF(IIF(LEN(per_ruc)<=8 OR LEFT(per_ruc,1)='P','06',IIF(LEN(per_ruc)=13 AND per_ruc!='9999999999999','04',IIF(LEN(per_ruc)=13 AND per_ruc='9999999999999','07','05'))) in ('06'),CONCAT(per_apellido,' ',per_nombre),'')AS DenoCli," & _
                    " '04' as tipoComprobante, 'E' as tipoEm, 1 as numeroComprobantes, " & _
                    " 0 as baseNoGraIva,  (ing_subtotal_o-IIF(ing_subtotal=0 AND ing_impuesto=0,ing_dcto,0)) as baseImponible,ing_subtotal-IIF(ing_subtotal=0 AND ing_impuesto=0,0,ing_dcto) as baseImpGrav, " & _
                    " ing_impuesto as montoIva,'00' as tipoCompe,0.00 as monto," & _
                    " 0 as montoIce,0 as valorRetIva,0 as valorRetRenta, " & _
                    " '' as formaPago" & _
                    " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo " & _
                    " AND ingreso.per_codigo=persona.per_codigo " & _
                    " WHERE ing_fecha Between '" & fechaINI & "' AND '" & fechaFIN & "'  " & _
                    " AND tip_ing_codigo='DCL'  AND ing_anulado=0 " & _
                    " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                    " ORDER BY idCliente,tipoComprobante "

                    
            clsCon_Def.Ejecutar strSQL
        
            'Set vsfgVentas.DataSource = clsCon_Def.adorec_Def.DataSource
            
            i = 1
            'Set vsfgCompras.DataSource = clsCon_Def.adorec_Def.DataSource
            
             While clsCon_Def.adorec_Def.EOF = False
             
             rep = 0
            If i > 1 Then
                'vsfgCompras.TextMatrix(i, 0) = clsCon_Def.adorec_Def(0)
                For j = 0 To vsfgVentas.Cols - 5
                    If vsfgVentas.TextMatrix(i - 1, j) = clsCon_Def.adorec_Def(j) Then
                        rep = rep + 1
                    Else
                        Exit For
                    End If
                Next j
            End If
            If rep = vsfgVentas.Cols - 5 Then
                For j = 0 To vsfgVentas.Cols - 1
                    If j < vsfgVentas.Cols - 4 Then
                        vsfgVentas.MergeCol(j) = True
                        vsfgVentas.TextMatrix(i, j) = vsfgVentas.TextMatrix(i - 1, j)
                    Else
                        vsfgVentas.MergeCol(j) = False
                    End If
                Next j
                vsfgVentas.TextMatrix(i, 22) = clsCon_Def.adorec_Def(22)
                vsfgVentas.TextMatrix(i, 23) = clsCon_Def.adorec_Def(23)
                vsfgVentas.TextMatrix(i, 24) = clsCon_Def.adorec_Def(24)
                vsfgVentas.TextMatrix(i, 25) = clsCon_Def.adorec_Def(25)
                
            Else
            vsfgVentas.TextMatrix(i, 0) = clsCon_Def.adorec_Def(0)
            vsfgVentas.TextMatrix(i, 1) = clsCon_Def.adorec_Def(1)
            vsfgVentas.TextMatrix(i, 2) = clsCon_Def.adorec_Def(2)
            vsfgVentas.TextMatrix(i, 3) = clsCon_Def.adorec_Def(3)
            vsfgVentas.TextMatrix(i, 4) = clsCon_Def.adorec_Def(4)
            vsfgVentas.TextMatrix(i, 5) = clsCon_Def.adorec_Def(5)
            vsfgVentas.TextMatrix(i, 6) = clsCon_Def.adorec_Def(6)
            vsfgVentas.TextMatrix(i, 7) = clsCon_Def.adorec_Def(7)
            vsfgVentas.TextMatrix(i, 8) = clsCon_Def.adorec_Def(8)
            vsfgVentas.TextMatrix(i, 9) = clsCon_Def.adorec_Def(9)
            vsfgVentas.TextMatrix(i, 10) = clsCon_Def.adorec_Def(10)
            vsfgVentas.TextMatrix(i, 11) = clsCon_Def.adorec_Def(11)
            vsfgVentas.TextMatrix(i, 12) = clsCon_Def.adorec_Def(12)
            vsfgVentas.TextMatrix(i, 13) = clsCon_Def.adorec_Def(13)
            vsfgVentas.TextMatrix(i, 14) = clsCon_Def.adorec_Def(14)
            vsfgVentas.TextMatrix(i, 15) = clsCon_Def.adorec_Def(15)
            vsfgVentas.TextMatrix(i, 16) = clsCon_Def.adorec_Def(16)
            vsfgVentas.TextMatrix(i, 17) = clsCon_Def.adorec_Def(17)
            vsfgVentas.TextMatrix(i, 18) = clsCon_Def.adorec_Def(18)
'            vsfgVentas.TextMatrix(i, 19) = clsCon_Def.adorec_Def(19)
'            vsfgVentas.TextMatrix(i, 20) = clsCon_Def.adorec_Def(20)
'            vsfgVentas.TextMatrix(i, 21) = clsCon_Def.adorec_Def(21)
'            vsfgVentas.TextMatrix(i, 22) = clsCon_Def.adorec_Def(22)
'            vsfgVentas.TextMatrix(i, 23) = clsCon_Def.adorec_Def(23)
'            vsfgVentas.TextMatrix(i, 24) = clsCon_Def.adorec_Def(24)
'            vsfgVentas.TextMatrix(i, 25) = clsCon_Def.adorec_Def(25)
'            vsfgVentas.TextMatrix(i, 26) = clsCon_Def.adorec_Def(26)
'            vsfgVentas.TextMatrix(i, 27) = clsCon_Def.adorec_Def(27)
'            vsfgVentas.TextMatrix(i, 28) = clsCon_Def.adorec_Def(28)
'            vsfgVentas.TextMatrix(i, 29) = clsCon_Def.adorec_Def(29)
'            vsfgVentas.TextMatrix(i, 30) = clsCon_Def.adorec_Def(30)
'            vsfgVentas.TextMatrix(i, 31) = clsCon_Def.adorec_Def(31)
'            vsfgVentas.TextMatrix(i, 32) = clsCon_Def.adorec_Def(32)
'            vsfgVentas.TextMatrix(i, 33) = clsCon_Def.adorec_Def(33)
'            vsfgVentas.TextMatrix(i, 34) = clsCon_Def.adorec_Def(34)
'            vsfgVentas.TextMatrix(i, 35) = clsCon_Def.adorec_Def(35)
'            vsfgVentas.TextMatrix(i, 36) = clsCon_Def.adorec_Def(36)
'            vsfgVentas.TextMatrix(i, 37) = clsCon_Def.adorec_Def(37)
'            vsfgVentas.TextMatrix(i, 38) = clsCon_Def.adorec_Def(38)
'            vsfgVentas.TextMatrix(i, 39) = clsCon_Def.adorec_Def(39)
'            vsfgVentas.TextMatrix(i, 40) = clsCon_Def.adorec_Def(40)
'            vsfgVentas.TextMatrix(i, 41) = clsCon_Def.adorec_Def(41)
'            vsfgVentas.TextMatrix(i, 42) = clsCon_Def.adorec_Def(42)
'            vsfgVentas.TextMatrix(i, 43) = clsCon_Def.adorec_Def(43)
            End If
            'Boton vsfgVentas
            i = vsfgVentas.Rows
            vsfgVentas.AddItem "", i
                clsCon_Def.adorec_Def.MoveNext
            Wend
            Boton vsfgVentas
            CambTam = True
    ElseIf sstI.Tab = 11 Then
        Dim tv As Double
        tv = 0
        vsfgVentasEstablecimiento.Clear 1
        vsfgVentasEstablecimiento.Rows = 2
        vsfgVentasEstablecimiento.Cell(flexcpBackColor, 1, 0, vsfgVentasEstablecimiento.Rows - 1, vsfgVentasEstablecimiento.Cols - 1) = vbWhite

            strSQL = " CREATE TABLE #venEsta (codEstab char(3), ventasEstab decimal(14,2)) "
            clsCon_Def.Ejecutar strSQL
            
           strSQL = " INSERT INTO #venEsta " & _
                    " SELECT RIGHT(cuenta_p_c.cue_p_c_serie,3) as codEstab, " & _
                    " SUM(ROUND(cue_p_c_st_cero+IIF(cue_p_c_st_cero=0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0),0),2) + " & _
                    " ROUND(egr_subtotal-egr_dcto-IIF(cue_p_c_st_cero=0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0),0),2)) as ventasEstab " & _
                    " From cuenta_p_c INNER JOIN egreso ON cuenta_p_c.emp_codigo=egreso.emp_codigo " & _
                    " AND cuenta_p_c.cue_p_c_egr_codigo=egreso.egr_codigo " & _
                    " AND cuenta_p_c.per_codigo=egreso.per_codigo  " & _
                    " AND egreso.tip_egr_codigo='FAC' AND egreso.egr_anulado=0 " & _
                    " AND egr_fecha BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                    " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                    " AND egreso.egr_codigo=det_egreso.egr_codigo AND det_egreso.prd_codigo='PR-FLET-DES'"
            strSQL = strSQL & " WHERE cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                    " AND cuenta_p_c.cue_p_c_tipo='C'  AND tip_doc_cue_codigo>0  AND cue_p_c_valor!=0 " & _
                    " AND cuenta_p_c.emp_codigo='" & strEmpresa & "' GROUP BY RIGHT(cuenta_p_c.cue_p_c_serie,3)  "
            clsCon_Def.Ejecutar strSQL
            strSQL = " INSERT INTO #venEsta " & _
                    " SELECT LEFT(ing_serie,3) as codEstab," & _
                    " -1*SUM((ing_subtotal_o-IIF(ing_subtotal=0 AND ing_impuesto=0,ing_dcto,0)) +" & _
                    " (ing_subtotal-IIF(ing_subtotal=0 AND ing_impuesto=0,0,ing_dcto))) as ventasEstab " & _
                    " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo " & _
                    " AND ingreso.per_codigo=persona.per_codigo " & _
                    " WHERE ing_fecha Between '" & fechaINI & "' AND '" & fechaFIN & "'  " & _
                    " AND tip_ing_codigo='DCL'  AND ing_anulado=0 " & _
                    " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                    " GROUP BY LEFT(ing_serie,3) "
            clsCon_Def.Ejecutar strSQL
        
           strSQL = " SELECT codEstab, SUM(ventasEstab)" & _
                    " FROM #venEsta" & _
                    " GROUP BY codEstab"
            clsCon_Def.Ejecutar strSQL

            i = 1
            
             While clsCon_Def.adorec_Def.EOF = False
             
                 rep = 0
                If i > 1 Then
                    'vsfgCompras.TextMatrix(i, 0) = clsCon_Def.adorec_Def(0)
                    For j = 0 To vsfgVentasEstablecimiento.Cols - 5
                        If vsfgVentasEstablecimiento.TextMatrix(i - 1, j) = clsCon_Def.adorec_Def(j) Then
                            rep = rep + 1
                        Else
                            Exit For
                        End If
                    Next j
                End If
                vsfgVentasEstablecimiento.TextMatrix(i, 0) = clsCon_Def.adorec_Def(0)
                vsfgVentasEstablecimiento.TextMatrix(i, 1) = clsCon_Def.adorec_Def(0)
                vsfgVentasEstablecimiento.TextMatrix(i, 2) = clsCon_Def.adorec_Def(1)
                vsfgVentasEstablecimiento.TextMatrix(i, 3) = 0
                tv = tv + vsfgVentasEstablecimiento.TextMatrix(i, 2)
                'Boton vsfgVentasEstablecimiento
                i = vsfgVentasEstablecimiento.Rows
                vsfgVentasEstablecimiento.AddItem "", i
                clsCon_Def.adorec_Def.MoveNext
            Wend
            
            vsfgInformante.TextMatrix(1, 6) = Format(vsfgVentasEstablecimiento.Rows - 2, "000")
            vsfgInformante.TextMatrix(1, 7) = tv
            Boton vsfgVentasEstablecimiento
            CambTam = True
            strSQL = " EXEC Sp_Drop_Table_if_Exist '#venEsta'"
            clsCon_Def.Ejecutar strSQL

        Else
            
            vsfgCompras.Clear 1
            vsfgCompras.Rows = 2
            vsfgCompras.Cell(flexcpBackColor, 1, 0, vsfgCompras.Rows - 1, vsfgCompras.Cols - 1) = vbWhite

            
            strSQL = " SELECT '        ' as Item,IIF(cod_sus_com_codigo='' OR cod_sus_com_codigo is null,'01',cod_sus_com_codigo) as codSustento," & _
                     " IIF(LEFT(per_ruc,1)='P','03',IIF(LEN(per_ruc)=13,'01','02')) as tpIdProv, " & _
                     " IIF(LEFT(per_ruc,1)='P',SUBSTRING(per_ruc,2,13),per_ruc) as idProv,tip_doc_cue_codigo as tipoComprobante," & _
                     " IIF(LEFT(per_ruc,1)='P','02','') as tipoProv, " & _
                     " IIF(LEFT(per_ruc,1)='P','NO','') as parteRel, IIF(LEFT(per_ruc,1)='P',CONCAT(per_apellido,' ',per_nombre),'') as denopr, " & _
                     " FORMAT(cuenta_p_c.cue_p_c_fechaemision,'dd/MM/yyyy') as fechaRegistro, " & _
                     " LEFT(cue_p_c_serie,3) as establecimiento,RIGHT(cue_p_c_serie,3) as puntoEmision,CAST(cue_p_c_numero as varchar) as secuencial, " & _
                     " FORMAT(cuenta_p_c.cue_p_c_fechaemision,'dd/MM/yyyy') as fechaEmision,cue_p_c_autorizacion as autorizacion, " & _
                     " 0 as baseNoGraIva, cue_p_c_st_cero as baseImponible, " & _
                     " (cue_p_c_st_prod+cue_p_c_st_serv) as baseImpgrav,'0' as baseImpExe,cue_p_c_ice as montoIce,Coalesce(cue_p_c_iva,0) as montoIva, " & _
                     " coalesce(iva10,0) as valRetBien10, " & _
                     " coalesce(iva20,0) as valRetServ20, " & _
                     " coalesce(iva30,0) as valorRetBienes, " & _
                     " coalesce(iva50,0) as valRetServ50, " & _
                     " coalesce(iva70,0) as valorRetServicios, " & _
                     " coalesce(iva100,0) as valRetServ100, "
            strSQL = strSQL & " '01' as pagoLocExt, " & _
                     " IIF(LEFT(per_ruc,1)='P','01','') as tipoRegi, IIF(LEFT(per_ruc,1)='P',pai_numero,'') as paisEfecPagoGen,'' as paisEfecPagoParFis,'' as denopago," & _
                     " 'NA' as paisEfecPago, 'NA' as aplicConvDobTrib, 'NA' as pagExtSujRetNorLeg,'NO' as pagoRegFis," & _
                     " IIF(cue_p_c_st_cero+cue_p_c_st_prod+cue_p_c_st_serv+cue_p_c_ice+cue_p_c_iva>1000,'02','') as formaPag,'' as docModificado, '' as estabModificado,'' as ptoEmiModificado,'' as secModificado," & _
                     " '' as autModificado,'' as tpIdProvReemb, '' as idProvReemb,'' as tipoComprobanteReemb," & _
                     " '' as establecimientoReemb,'' as puntoEmisionReemb,'' as secuencialReemb,'' as fechaEmisionReemb," & _
                     " '' as autorizacionReemb,'' as baseImponibleReemb, '' as baseImpGravReemb,'' as baseNoGraIvaReemb,'' as baseImpExeReemb,'0' as totbasesImpReemb," & _
                     " '' as montoIceReemb,'' as montoIvaRemb, " & _
                     " COALESCE(LEFT(com_ret_serie,3),'') as estabRetencion1,COALESCE(RIGHT(com_ret_serie,3),'') as ptoEmiRetencion1, " & _
                     " COALESCE(CAST(com_ret_numero as varchar),'') as secRetencion1,  COALESCE(doc_ele_autorizacion,com_ret_autorizacion,'') as autRetencion1, " & _
                     " COALESCE(FORMAT(com_ret_fecha,'dd/MM/yyyy'),'') as fechaEmiRet1, " & _
                     " COALESCE(codRetAir,'332') as codRetAir, " & _
                     " IIF(baseImpAir<>0 or baseImpAir IS NOT NULL,baseImpAir ,IIF(cue_p_c_st_prod+cue_p_c_st_serv=0 and cue_p_c_st_cero=0,cuenta_p_c.cue_p_c_valor/1." & PorIVA & ",cue_p_c_st_prod+cue_p_c_st_serv+cue_p_c_st_cero)) as baseImpAir, " & _
                     " COALESCE(porcentajeAir,'0') as porcentajeAir, " & _
                     " COALESCE(valRetAir,'0') As valRetAir, " & _
                     " '' As fechaPagoDiv,'' As imRentaSoc,'' As anioUtDiv,'' as NumCajBan,'' as PrecCajBan, " & _
                     " cuenta_p_c.asi_numasiento as ASIE "
            strSQL = strSQL & " From cuenta_p_c " & _
                     " INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                     " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                     " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                     " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                     " INNER JOIN codigo_iva ON cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo " & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                     " LEFT JOIN DetRetencionIRAT ret ON comprobante_retencion.emp_codigo=ret.emp_codigo " & _
                     " AND comprobante_retencion.cue_p_c_codigo=ret.cue_p_c_codigo " & _
                     " AND comprobante_retencion.cue_p_c_tipo=ret.cue_p_c_tipo " & _
                     " AND ret.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                     " LEFT JOIN DetRetencionIVAAT compra " & _
                     " on compra.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " and compra.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " AND compra.cue_p_c_tipo= comprobante_retencion.cue_p_c_tipo " & _
                     " AND compra.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' "
            strSQL = strSQL & " LEFT JOIN ( SELECT emp_codigo,doc_electronico.doc_ele_codigo,doc_ele_autorizacion " & _
                     " FROM doc_electronico " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND doc_ele_coddoc='07') aut " & _
                     " ON comprobante_retencion.emp_codigo=aut.emp_codigo " & _
                     " AND comprobante_retencion.cue_p_c_codigo=aut.doc_ele_codigo " & _
                     " WHERE cuenta_p_c.cue_p_c_fechaemision BETWEEN '" & fechaINI & "' AND '" & fechaFIN & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='P' AND cuenta_p_c.tip_doc_cue_codigo>0 AND cue_p_c_valor!=0 " & _
                     " AND cuenta_p_c.emp_codigo='" & strEmpresa & "' "
            strSQL = strSQL & " Union " & _
                    " SELECT CAST(egr_numero as varchar) as Item,'06' as codSustento," & _
                    " IIF(LEFT(per_ruc,1)='P','03',IIF(LEN(per_ruc)=13,'01','02')) as tpIdProv," & _
                    " IIF(LEFT(per_ruc,1)='P',SUBSTRING(per_ruc,2,13),per_ruc) as idProv,'04' as tipoComprobante, " & _
                    " IIF(LEFT(per_ruc,1)='P','02','') as tipoProv, " & _
                    " IIF(LEFT(per_ruc,1)='P','NO','') as parteRel, IIF(LEFT(per_ruc,1)='P',CONCAT(per_apellido,' ',per_nombre),'') as denopr, " & _
                    " FORMAT(egr_fecha,'dd/MM/yyyy') as fechaRegistro, " & _
                    " LEFT(egr_serie,3) as establecimiento,RIGHT(egr_serie,3) as puntoEmision,CAST(egr_numero as varchar) as secuencial," & _
                    " FORMAT(egr_fecha,'dd/MM/yyyy') as fechaEmision,egr_autorizacion as autorizacion, " & _
                    " 0 as baseNoGraIva, ROUND((egr_subtotal_o-IIF(egr_subtotal=0 AND egr_impuesto=0,egr_dcto,0)),2) as baseImponible, " & _
                     " ROUND((egr_subtotal-IIF(egr_subtotal=0 AND egr_impuesto=0,0,egr_dcto)),2) as baseImpgrav,'0' as baseImpExe,0 as montoIce,ROUND(egr_impuesto,2) as montoIva, " & _
                     " 0 as valRetBien10, " & _
                     " 0 as valRetServ20, " & _
                     " 0 as valorRetBienes, " & _
                     " 0 as valRetServ50, " & _
                     " 0 as valorRetServicios, " & _
                     " 0 as valRetServ100,"
            strSQL = strSQL & " '01' as pagoLocExt, " & _
                     " IIF(LEFT(per_ruc,1)='P','01','') as tipoRegi, IIF(LEFT(per_ruc,1)='P',pai_numero,'') as paisEfecPagoGen,'' as paisEfecPagoParFis,'' as denopago," & _
                     " 'NA' as paisEfecPago, 'NA' as aplicConvDobTrib, 'NA' as pagExtSujRetNorLeg,'NO' as pagoRegFis," & _
                     " '' as formaPag,'' as docModificado, '' as estabModificado,'' as ptoEmiModificado,'' as secModificado," & _
                     " '' as autModificado,'' as tpIdProvReemb, '' as idProvReemb,'' as tipoComprobanteReemb," & _
                     " '' as establecimientoReemb,'' as puntoEmisionReemb,'' as secuencialReemb,'' as fechaEmisionReemb," & _
                     " '' as autorizacionReemb,'' as baseImponibleReemb, '' as baseImpGravReemb,'' as baseNoGraIvaReemb,'' as baseImpExeReemb,'0' as totbasesImpReemb," & _
                     " '' as montoIceReemb,'' as montoIvaRemb, " & _
                     " '' as estabRetencion1,'' as ptoEmiRetencion1, " & _
                     " '' as secRetencion1, '' as autRetencion1, " & _
                     " '' as fechaEmiRet1, " & _
                     " '' as codRetAir, " & _
                     " 0 as baseImpAir, " & _
                     " 0 as porcentajeAir, " & _
                     " 0 As valRetAir, " & _
                     " '' As fechaPagoDiv,'' As imRentaSoc,'' As anioUtDiv,'' as NumCajBan,'' as PrecCajBan, " & _
                     " egreso.egr_numasiento as ASIE"
            strSQL = strSQL & " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
                    " AND egreso.per_codigo=persona.per_codigo " & _
                    " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                    " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                    " WHERE egr_fecha Between '" & fechaINI & "' AND '" & fechaFIN & "'  " & _
                    " AND tip_egr_codigo='DPV'  AND egr_anulado=0 " & _
                    " AND egreso.emp_codigo='" & strEmpresa & "' " & _
                    " ORDER BY idProv,tipoComprobante,establecimiento,puntoEmision,secuencial "
                
            clsCon_Def.Ejecutar strSQL
            i = 1
            'Set vsfgCompras.DataSource = clsCon_Def.adorec_Def.DataSource
             While clsCon_Def.adorec_Def.EOF = False
             rep = 0
            If i > 1 Then
                'vsfgCompras.TextMatrix(i, 0) = clsCon_Def.adorec_Def(0)
                'If (i = 151 Or i = 152) Then MsgBox i & " - " & rep
                'For j = 1 To vsfgCompras.Cols - 5
                For j = 1 To 14
                    'If i >= 151 Then MsgBox "(" & j & ") " & vsfgCompras.TextMatrix(i - 1, j) & " = " & CStr(clsCon_Def.adorec_Def(j))
                    If CStr(vsfgCompras.TextMatrix(i - 1, j)) = CStr(clsCon_Def.adorec_Def(j)) Then
                        rep = rep + 1
                    Else
                        Exit For
                    End If
                Next j
                'If (i = 151 Or i = 152) Then MsgBox i & " - " & rep
            End If
            
            'If rep = vsfgCompras.Cols - 5 Then
            If rep = 14 Then
                For j = 0 To vsfgCompras.Cols - 1
                    If j < vsfgCompras.Cols - 7 Then
            vsfgCompras.ShowCell i, j
                        'If (i = 151 Or i = 152) Then MsgBox i & " - " & j
                        vsfgCompras.MergeCol(j) = True
                        If j > 0 Then
                            vsfgCompras.TextMatrix(i, j) = CStr(vsfgCompras.TextMatrix(i - 1, j))
                        Else
                            vsfgCompras.TextMatrix(i, j) = vsfgCompras.TextMatrix(i - 1, j)
                        End If
                    Else
                        'vsfgCompras.MergeCol(j) = False
                    End If
                Next j
                vsfgCompras.TextMatrix(i, 59) = CStr(clsCon_Def.adorec_Def(59))
                vsfgCompras.TextMatrix(i, 60) = CStr(clsCon_Def.adorec_Def(60))
                vsfgCompras.TextMatrix(i, 61) = CStr(clsCon_Def.adorec_Def(61))
                vsfgCompras.TextMatrix(i, 62) = CStr(clsCon_Def.adorec_Def(62))
                vsfgCompras.TextMatrix(i, 63) = CStr(clsCon_Def.adorec_Def(63))
                vsfgCompras.TextMatrix(i, 64) = CStr(clsCon_Def.adorec_Def(64))
                vsfgCompras.TextMatrix(i, 65) = CStr(clsCon_Def.adorec_Def(65))
                vsfgCompras.TextMatrix(i, 66) = CStr(clsCon_Def.adorec_Def(66))
                vsfgCompras.TextMatrix(i, 67) = CStr(clsCon_Def.adorec_Def(67))
                vsfgCompras.TextMatrix(i, 68) = CStr(clsCon_Def.adorec_Def(68))
                vsfgCompras.TextMatrix(i, 69) = CStr(clsCon_Def.adorec_Def(69))
                
            Else
            vsfgCompras.ShowCell i, 1
            vsfgCompras.TextMatrix(i, 1) = CStr(clsCon_Def.adorec_Def(1))
            vsfgCompras.ShowCell i, 2
            vsfgCompras.TextMatrix(i, 2) = CStr(clsCon_Def.adorec_Def(2))
            vsfgCompras.ShowCell i, 3
            vsfgCompras.TextMatrix(i, 3) = CStr(clsCon_Def.adorec_Def(3))
            vsfgCompras.ShowCell i, 4
            vsfgCompras.TextMatrix(i, 4) = CStr(clsCon_Def.adorec_Def(4))
            vsfgCompras.ShowCell i, 5
            vsfgCompras.TextMatrix(i, 5) = CStr(clsCon_Def.adorec_Def(5))
            vsfgCompras.ShowCell i, 6
            vsfgCompras.TextMatrix(i, 6) = CStr(clsCon_Def.adorec_Def(6))
            vsfgCompras.ShowCell i, 7
            vsfgCompras.TextMatrix(i, 7) = CStr(clsCon_Def.adorec_Def(7))
            vsfgCompras.ShowCell i, 8
            vsfgCompras.TextMatrix(i, 8) = CStr(clsCon_Def.adorec_Def(8))
            vsfgCompras.ShowCell i, 9
            vsfgCompras.TextMatrix(i, 9) = CStr(clsCon_Def.adorec_Def(9))
            vsfgCompras.ShowCell i, 10
            vsfgCompras.TextMatrix(i, 10) = CStr(clsCon_Def.adorec_Def(10))
            vsfgCompras.ShowCell i, 11
            vsfgCompras.TextMatrix(i, 11) = CStr(clsCon_Def.adorec_Def(11))
            vsfgCompras.ShowCell i, 12
            vsfgCompras.TextMatrix(i, 12) = CStr(clsCon_Def.adorec_Def(12))
            vsfgCompras.ShowCell i, 13
            vsfgCompras.TextMatrix(i, 13) = CStr(clsCon_Def.adorec_Def(13))
            vsfgCompras.ShowCell i, 14
            vsfgCompras.TextMatrix(i, 14) = CStr(clsCon_Def.adorec_Def(14))
            vsfgCompras.ShowCell i, 15
            vsfgCompras.TextMatrix(i, 15) = CStr(clsCon_Def.adorec_Def(15))
            vsfgCompras.ShowCell i, 16
            vsfgCompras.TextMatrix(i, 16) = CStr(clsCon_Def.adorec_Def(16))
            vsfgCompras.ShowCell i, 17
            vsfgCompras.TextMatrix(i, 17) = CStr(clsCon_Def.adorec_Def(17))
            vsfgCompras.ShowCell i, 18
            vsfgCompras.TextMatrix(i, 18) = CStr(clsCon_Def.adorec_Def(18))
            vsfgCompras.ShowCell i, 19
            vsfgCompras.TextMatrix(i, 19) = CStr(clsCon_Def.adorec_Def(19))
            vsfgCompras.ShowCell i, 20
            vsfgCompras.TextMatrix(i, 20) = CStr(clsCon_Def.adorec_Def(20))
            vsfgCompras.ShowCell i, 21
            vsfgCompras.TextMatrix(i, 21) = CStr(clsCon_Def.adorec_Def(21))
            vsfgCompras.ShowCell i, 22
            vsfgCompras.TextMatrix(i, 22) = CStr(clsCon_Def.adorec_Def(22))
            vsfgCompras.ShowCell i, 23
            vsfgCompras.TextMatrix(i, 23) = CStr(clsCon_Def.adorec_Def(23))
            vsfgCompras.ShowCell i, 24
            vsfgCompras.TextMatrix(i, 24) = CStr(clsCon_Def.adorec_Def(24))
            vsfgCompras.ShowCell i, 25
            vsfgCompras.TextMatrix(i, 25) = CStr(clsCon_Def.adorec_Def(25))
            vsfgCompras.ShowCell i, 26
            vsfgCompras.TextMatrix(i, 26) = CStr(clsCon_Def.adorec_Def(26))
            vsfgCompras.ShowCell i, 27
            vsfgCompras.TextMatrix(i, 27) = CStr(clsCon_Def.adorec_Def(27))
            vsfgCompras.ShowCell i, 28
            vsfgCompras.TextMatrix(i, 28) = CStr(clsCon_Def.adorec_Def(28))
            vsfgCompras.ShowCell i, 29
            vsfgCompras.TextMatrix(i, 29) = CStr(clsCon_Def.adorec_Def(29))
            vsfgCompras.ShowCell i, 30
            vsfgCompras.TextMatrix(i, 30) = CStr(clsCon_Def.adorec_Def(30))
            vsfgCompras.ShowCell i, 31
            vsfgCompras.TextMatrix(i, 31) = CStr(clsCon_Def.adorec_Def(31))
            vsfgCompras.ShowCell i, 32
            vsfgCompras.TextMatrix(i, 32) = CStr(clsCon_Def.adorec_Def(32))
            vsfgCompras.ShowCell i, 33
            vsfgCompras.TextMatrix(i, 33) = CStr(clsCon_Def.adorec_Def(33))
            vsfgCompras.ShowCell i, 34
            vsfgCompras.TextMatrix(i, 34) = CStr(clsCon_Def.adorec_Def(34))
            vsfgCompras.ShowCell i, 35
            vsfgCompras.TextMatrix(i, 35) = CStr(clsCon_Def.adorec_Def(35))
            vsfgCompras.ShowCell i, 36
            vsfgCompras.TextMatrix(i, 36) = CStr(clsCon_Def.adorec_Def(36))
            vsfgCompras.ShowCell i, 37
            vsfgCompras.TextMatrix(i, 37) = CStr(clsCon_Def.adorec_Def(37))
            vsfgCompras.ShowCell i, 38
            vsfgCompras.TextMatrix(i, 38) = CStr(clsCon_Def.adorec_Def(38))
            vsfgCompras.ShowCell i, 39
            vsfgCompras.TextMatrix(i, 39) = CStr(clsCon_Def.adorec_Def(39))
            vsfgCompras.ShowCell i, 40
            vsfgCompras.TextMatrix(i, 40) = CStr(clsCon_Def.adorec_Def(40))
            vsfgCompras.ShowCell i, 41
            vsfgCompras.TextMatrix(i, 41) = CStr(clsCon_Def.adorec_Def(41))
            vsfgCompras.ShowCell i, 42
            vsfgCompras.TextMatrix(i, 42) = CStr(clsCon_Def.adorec_Def(42))
            vsfgCompras.ShowCell i, 43
            vsfgCompras.TextMatrix(i, 43) = CStr(clsCon_Def.adorec_Def(43))
            vsfgCompras.ShowCell i, 44
            vsfgCompras.TextMatrix(i, 44) = CStr(clsCon_Def.adorec_Def(44))
            vsfgCompras.ShowCell i, 45
            vsfgCompras.TextMatrix(i, 45) = CStr(clsCon_Def.adorec_Def(45))
            vsfgCompras.ShowCell i, 46
            vsfgCompras.TextMatrix(i, 46) = CStr(clsCon_Def.adorec_Def(46))
            vsfgCompras.ShowCell i, 47
            vsfgCompras.TextMatrix(i, 47) = CStr(clsCon_Def.adorec_Def(47))
            vsfgCompras.ShowCell i, 48
            vsfgCompras.TextMatrix(i, 48) = CStr(clsCon_Def.adorec_Def(48))
            vsfgCompras.ShowCell i, 49
            vsfgCompras.TextMatrix(i, 49) = CStr(clsCon_Def.adorec_Def(49))
            vsfgCompras.ShowCell i, 50
            vsfgCompras.TextMatrix(i, 50) = CStr(clsCon_Def.adorec_Def(50))
            vsfgCompras.ShowCell i, 51
            vsfgCompras.TextMatrix(i, 51) = CStr(clsCon_Def.adorec_Def(51))
            vsfgCompras.ShowCell i, 52
            vsfgCompras.TextMatrix(i, 52) = CStr(clsCon_Def.adorec_Def(52))
            vsfgCompras.ShowCell i, 53
            vsfgCompras.TextMatrix(i, 53) = CStr(clsCon_Def.adorec_Def(53))
            vsfgCompras.ShowCell i, 54
            vsfgCompras.TextMatrix(i, 54) = CStr(clsCon_Def.adorec_Def(54))
            vsfgCompras.ShowCell i, 55
            vsfgCompras.TextMatrix(i, 55) = CStr(clsCon_Def.adorec_Def(55))
            vsfgCompras.ShowCell i, 56
            vsfgCompras.TextMatrix(i, 56) = CStr(clsCon_Def.adorec_Def(56))
            vsfgCompras.ShowCell i, 57
            vsfgCompras.TextMatrix(i, 57) = CStr(clsCon_Def.adorec_Def(57))
            vsfgCompras.ShowCell i, 58
            vsfgCompras.TextMatrix(i, 58) = CStr(clsCon_Def.adorec_Def(58))
            vsfgCompras.ShowCell i, 59
            vsfgCompras.TextMatrix(i, 59) = CStr(clsCon_Def.adorec_Def(59))
            vsfgCompras.ShowCell i, 60
            vsfgCompras.TextMatrix(i, 60) = CStr(clsCon_Def.adorec_Def(60))
            vsfgCompras.ShowCell i, 61
            vsfgCompras.TextMatrix(i, 61) = CStr(clsCon_Def.adorec_Def(61))
            vsfgCompras.ShowCell i, 62
            vsfgCompras.TextMatrix(i, 62) = CStr(clsCon_Def.adorec_Def(62))
            vsfgCompras.ShowCell i, 63
            vsfgCompras.TextMatrix(i, 63) = CStr(clsCon_Def.adorec_Def(63))
            vsfgCompras.ShowCell i, 64
            vsfgCompras.TextMatrix(i, 64) = CStr(clsCon_Def.adorec_Def(64))
            vsfgCompras.ShowCell i, 65
            vsfgCompras.TextMatrix(i, 65) = CStr(clsCon_Def.adorec_Def(65))
            vsfgCompras.ShowCell i, 66
            vsfgCompras.TextMatrix(i, 66) = CStr(clsCon_Def.adorec_Def(66))
            vsfgCompras.ShowCell i, 67
            vsfgCompras.TextMatrix(i, 67) = CStr(clsCon_Def.adorec_Def(67))
            vsfgCompras.ShowCell i, 68
            vsfgCompras.TextMatrix(i, 68) = CStr(clsCon_Def.adorec_Def(68))
            vsfgCompras.ShowCell i, 69
            vsfgCompras.TextMatrix(i, 69) = CStr(clsCon_Def.adorec_Def(69))
            
            vsfgCompras.ShowCell i, 4
            If FormatoD0(vsfgCompras.TextMatrix(i, 4)) = 4 Then
                strSQL = " SELECT cuenta_p_c.* FROM cuenta_p_c INNER JOIN pago " & _
                         " ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                         " AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
                         " AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                         " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "'" & _
                         " AND cuenta_p_c.cue_p_c_tipo='P' " & _
                         " AND pago.asi_numasiento='" & clsCon_Def.adorec_Def("ASIE") & "'" & _
                         " ORDER BY pag_monto DESC"
                cls_Aux.Ejecutar strSQL
                If cls_Aux.adorec_Def.RecordCount > 0 Then
                    vsfgCompras.TextMatrix(i, 36) = 1
                    vsfgCompras.TextMatrix(i, 37) = Left(cls_Aux.adorec_Def("cue_p_c_serie"), 3)
                    vsfgCompras.TextMatrix(i, 38) = Right(cls_Aux.adorec_Def("cue_p_c_serie"), 3)
                    vsfgCompras.TextMatrix(i, 39) = cls_Aux.adorec_Def("cue_p_c_numero")
                    vsfgCompras.TextMatrix(i, 40) = cls_Aux.adorec_Def("cue_p_c_autorizacion")
                End If
            End If
'            vsfgCompras.ShowCell i, 33
'            vsfgCompras.TextMatrix(i, 33) = CStr(clsCon_Def.adorec_Def(33))
'            vsfgCompras.ShowCell i, 34
'            vsfgCompras.TextMatrix(i, 34) = CStr(clsCon_Def.adorec_Def(34))
'            vsfgCompras.ShowCell i, 35
'            vsfgCompras.TextMatrix(i, 35) = CStr(clsCon_Def.adorec_Def(35))
'            vsfgCompras.ShowCell i, 36
'            vsfgCompras.TextMatrix(i, 36) = CStr(clsCon_Def.adorec_Def(36))
'            vsfgCompras.ShowCell i, 37
'            vsfgCompras.TextMatrix(i, 37) = CStr(clsCon_Def.adorec_Def(37))
'            vsfgCompras.ShowCell i, 38
'            vsfgCompras.TextMatrix(i, 38) = CStr(clsCon_Def.adorec_Def(38))
'            vsfgCompras.ShowCell i, 39
'            vsfgCompras.TextMatrix(i, 39) = CStr(clsCon_Def.adorec_Def(39))
'            vsfgCompras.ShowCell i, 40
'            vsfgCompras.TextMatrix(i, 40) = CStr(clsCon_Def.adorec_Def(40))
'            vsfgCompras.ShowCell i, 41
'            vsfgCompras.TextMatrix(i, 41) = CStr(clsCon_Def.adorec_Def(41))
'            vsfgCompras.ShowCell i, 42
'            vsfgCompras.TextMatrix(i, 42) = CStr(clsCon_Def.adorec_Def(42))
'            vsfgCompras.ShowCell i, 43
'            vsfgCompras.TextMatrix(i, 43) = CStr(clsCon_Def.adorec_Def(43))
            End If
            Boton vsfgCompras
            i = vsfgCompras.Rows
            vsfgCompras.AddItem "", i
                clsCon_Def.adorec_Def.MoveNext
            Wend
         
            'JuntarRows vsfgCompras
            Boton vsfgCompras
            'NuevaLinea vsfgDescCompras, vsfgCompras, vsfgCompras.Rows - 1
            
       
            
            
'            Col = 22
'            vsfgCompras.AddItem "", vsfgCompras.Rows
'                For i = 0 To vsfgCompras.Cols - 1
'                    If Not i >= Col Then
'                        vsfgCompras.MergeCol(i) = True
'                        vsfgCompras.TextMatrix(vsfgCompras.Rows - 1, i) = vsfgCompras.TextMatrix(vsfgCompras.Rows - 2, i)
'                    Else
'                        vsfgCompras.MergeCol(i) = False
'                    End If
'                Next i
        End If
        
    End If
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    
    End Select


End Sub

Private Sub cmdExportarInformante_Click()
    Exportar vsfgDescInformante, vsfgInformante, "Informante"
End Sub

Private Sub cmdExportarCompras_Click()
    Exportar vsfgDescCompras, vsfgCompras, "Compras"
End Sub

Private Sub cmdExportarVentas_Click()
    Exportar vsfgDescVentas, vsfgVentas, "Ventas"
End Sub

Private Sub cmdExportarImportaciones_Click()
    Exportar vsfgDescImportaciones, vsfgImportaciones, "Importaciones"
End Sub

Private Sub cmdExportarExportaciones_Click()
    Exportar vsfgDescExportaciones, vsfgExportaciones, "Exportaciones"
End Sub

Private Sub cmdExportarTC_Click()
    Exportar vsfgDescTC, vsfgTC, "TC"
End Sub

Private Sub cmdExportarFideicomisos_Click()
    Exportar vsfgDescFideicomisos, vsfgFideicomisos, "Fideicomisos"
End Sub

Private Sub cmdExportarAnulados_Click()
    Exportar vsfgDescAnulados, VSFGAnulados, "Anulados"
End Sub

Private Sub cmdExportarRendimientos_Click()
    Exportar vsfgDescRendimientos, vsfgRendimientos, "Rendimientos"
End Sub

Private Sub cmdExportarVentasEstablecimiento_Click()
    Exportar vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, "Ventas por Establecimiento"
End Sub

Private Sub cmdGenerarXML_Click()
    CD.DefaultExt = "xml"
    CD.Filter = "Archivos XML (*.xml)|*.xml|Todos |*.*"
    CD.FilterIndex = 1
    CD.FileName = ""
    CD.ShowSave
    If CD.FileName <> "" Then
        SaveValues CD.FileName
        WebBrow.Navigate CD.FileName
    End If
End Sub

Private Sub cmdGuardar_Click()
    Dim Archivo As String
    CD.DefaultExt = "ated"
    CD.Filter = "Archivos ATED (*.ated)|*.ated|Todos |*.*"
    CD.FilterIndex = 1
    CD.FileName = ""
    CD.ShowSave
    Archivo = CD.FileName
    If Archivo <> "" Then
        Guarda vsfgInformante, "Informante", Archivo
        Guarda vsfgCompras, "Compras", Archivo
        Guarda vsfgVentas, "Ventas", Archivo
        Guarda vsfgVentasEstablecimiento, "VentasEstablecimiento", Archivo
        Guarda vsfgImportaciones, "Importaciones", Archivo
        Guarda vsfgExportaciones, "Exportaciones", Archivo
        Guarda vsfgTC, "TC", Archivo
        Guarda vsfgFideicomisos, "Fideicomisos", Archivo
        Guarda VSFGAnulados, "Anulados", Archivo
        Guarda vsfgRendimientos, "Rendimientos", Archivo
    End If
End Sub

Private Sub cmdImportarInformante_Click()
    Importar vsfgDescInformante, vsfgInformante, "Informante"
End Sub

Private Sub cmdImportarCompras_Click()
    Importar vsfgDescCompras, vsfgCompras, "Compras"
End Sub

Private Sub cmdImportarVentas_Click()
    Importar vsfgDescVentas, vsfgVentas, "Ventas"
End Sub

Private Sub cmdImportarImportaciones_Click()
    Importar vsfgDescImportaciones, vsfgImportaciones, "Importaciones"
End Sub

Private Sub cmdImportarExportaciones_Click()
    Importar vsfgDescExportaciones, vsfgExportaciones, "Exportaciones"
End Sub

Private Sub cmdImportarTC_Click()
    Importar vsfgDescTC, vsfgTC, "TC"
End Sub

Private Sub cmdImportarFideicomisos_Click()
    Importar vsfgDescFideicomisos, vsfgFideicomisos, "Fideicomisos"
End Sub

Private Sub cmdImportarAnulados_Click()
    Importar vsfgDescAnulados, VSFGAnulados, "Anulados"
End Sub

Private Sub cmdImportarRendimientos_Click()
    Importar vsfgDescRendimientos, vsfgRendimientos, "Rendimientos"
End Sub

Private Sub cmdImportarVentasEstablecimiento_Click()
    Importar vsfgDescVentasEstablecimiento, vsfgDescVentasEstablecimiento, "Ventas por Establecimiento"
End Sub

Private Sub cmdJuntar_Click()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim ii As Long
    Dim jj As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim rr1 As Long
    Dim rr2 As Long
    Dim cc1 As Long
    Dim cc2 As Long
    Dim Nueva As Boolean
    Dim salir As Boolean
    Dim Salir2 As Boolean
    CambTam = False
    Me.MousePointer = 11
    With vsfgVentas
        .ColDataType(0) = flexDTLong
        .Select 1, 2, 1, 6
        .Sort = flexSortGenericDescending
        i = 1
        While salir = False
            .ShowCell i, 0
            .GetMergedRange i, 0, r1, c1, r2, c2
            i = r2
            .GetMergedRange i + 1, 0, rr1, cc1, rr2, cc2
            ii = rr2
'            If Right(.TextMatrix(i, 2), 10) = "0101213494" Then
'                MsgBox "0101213494"
'            End If
            
            If Format(.TextMatrix(i, 1), "00") = Format(.TextMatrix(ii, 1), "00") And _
               Format(.TextMatrix(i, 2), vsfgDescVentas.TextMatrix(2, 7)) = Format(.TextMatrix(ii, 2), vsfgDescVentas.TextMatrix(2, 7)) And _
               .TextMatrix(i, 3) = .TextMatrix(ii, 3) And _
               .TextMatrix(i, 4) = .TextMatrix(ii, 4) And _
               .TextMatrix(i, 5) = .TextMatrix(ii, 5) And _
               .TextMatrix(i, 6) = .TextMatrix(ii, 6) And _
               1 = 1 And _
               1 = 1 Then
                
                .Cell(flexcpText, r1, 8, r2, 8) = Val(.TextMatrix(i, 8)) + Val(.TextMatrix(ii, 8))
                .Cell(flexcpText, r1, 9, r2, 9) = Val(.TextMatrix(i, 9)) + Val(.TextMatrix(ii, 9))
                .Cell(flexcpText, r1, 10, r2, 10) = Val(.TextMatrix(i, 10)) + Val(.TextMatrix(ii, 10))
                .Cell(flexcpText, r1, 11, r2, 11) = Val(.TextMatrix(i, 11)) + Val(.TextMatrix(ii, 11))
                .Cell(flexcpText, r1, 12, r2, 12) = Val(.TextMatrix(i, 12)) + Val(.TextMatrix(ii, 12))
                .Cell(flexcpText, r1, 12, r2, 14) = Val(.TextMatrix(i, 14)) + Val(.TextMatrix(ii, 14))
                .Cell(flexcpText, r1, 15, r2, 15) = Val(.TextMatrix(i, 15)) + Val(.TextMatrix(ii, 15))
                .Cell(flexcpText, r1, 17, r2, 16) = Val(.TextMatrix(i, 16)) + Val(.TextMatrix(ii, 16))
                .Cell(flexcpText, r1, 17, r2, 17) = Val(.TextMatrix(i, 17)) + Val(.TextMatrix(ii, 17))
'                .Cell(flexcpText, r1, 18, r2, 18) = Val(.TextMatrix(i, 18)) + Val(.TextMatrix(ii, 18))
'                .Cell(flexcpText, r1, 20, r2, 20) = Val(.TextMatrix(i, 20)) + Val(.TextMatrix(ii, 20))
'                If .TextMatrix(i, 16) < .TextMatrix(ii, 16) Then
'                    .Cell(flexcpText, r1, 16, r2, 16) = .TextMatrix(ii, 16)
'                Else
'                    .TextMatrix(ii, 16) = .TextMatrix(i, 16)
'                End If
'                If .TextMatrix(i, 19) < .TextMatrix(ii, 19) Then
'                    .Cell(flexcpText, r1, 19, r2, 19) = .TextMatrix(ii, 19)
'                Else
'                    .TextMatrix(ii, 19) = .TextMatrix(i, 19)
'                End If
'                If .TextMatrix(i, 21) <> .TextMatrix(ii, 21) Then
'                    .Cell(flexcpText, r1, 21, r2, 21) = "S"
'                    .TextMatrix(ii, 21) = "S"
'                End If
                j = rr1
'                Salir2 = False
'                While Salir2 = False
'                    Nueva = True
'                    For k = r1 To r2
'                        If .TextMatrix(k, 22) = "" Then
'                            .TextMatrix(k, 22) = .TextMatrix(j, 22)
'                            .TextMatrix(k, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
'                            .TextMatrix(k, 24) = .TextMatrix(j, 24)
'                            .TextMatrix(k, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
'                            Nueva = False
'                        ElseIf .TextMatrix(j, 22) = "" Then
'                            .TextMatrix(j, 22) = .TextMatrix(k, 22)
'                            .TextMatrix(j, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
'                            .TextMatrix(j, 24) = .TextMatrix(k, 24)
'                            .TextMatrix(j, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
'                            Nueva = False
'                        ElseIf .TextMatrix(k, 22) = .TextMatrix(j, 22) And _
'                           .TextMatrix(k, 24) = .TextMatrix(j, 24) Then
'                            .TextMatrix(k, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
'                            .TextMatrix(k, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
'                            Nueva = False
'                            Exit For
'                        End If
'                    Next k
'                    If Nueva = True Then
'                        For l = 0 To 21
'                            .TextMatrix(j, l) = .TextMatrix(k - 1, l)
'                        Next l
'                    Else
                        .RemoveItem ii
                        i = i - 1
'                        j = j - 1
'                        rr2 = rr2 - 1
'                    End If
'                    j = j + 1
'                    If j > rr2 Then
'                        Salir2 = True
'                    End If
'                Wend
            End If
            i = i + 1
            If i >= .Rows - 2 Then
                salir = True
            End If
        Wend
        .Select 1, 0
        .Sort = flexSortGenericAscending
        'Boton vsfgVentas
    End With
    Me.MousePointer = 0
    CambTam = True
End Sub

Private Sub cmdNuevo_Click()
    Limpiar vsfgInformante
    Limpiar vsfgCompras
    Limpiar vsfgVentas
    Limpiar vsfgImportaciones
    Limpiar vsfgExportaciones
    Limpiar vsfgTC
    Limpiar vsfgFideicomisos
    Limpiar VSFGAnulados
    Limpiar vsfgRendimientos
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set ucrtVSFG.VSFGControl = VSFGREOC
    CambTam = True
    cargado = False
    CargarMes
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    CargaGrid vsfgDescInformante, vsfgInformante, "Informante"
    CargaGrid vsfgDescCompras, vsfgCompras, "Compras"
    CargaGrid vsfgDescVentas, vsfgVentas, "Ventas"
    CargaGrid vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, "VentasEstablecimiento"
    CargaGrid vsfgDescImportaciones, vsfgImportaciones, "Importaciones"
    CargaGrid vsfgDescExportaciones, vsfgExportaciones, "Exportaciones"
    CargaGrid vsfgDescTC, vsfgTC, "TC"
    CargaGrid vsfgDescFideicomisos, vsfgFideicomisos, "Fideicomisos"
    CargaGrid vsfgDescAnulados, VSFGAnulados, "Anulados"
    CargaGrid vsfgDescRendimientos, vsfgRendimientos, "Rendimientos"
    WebBrow.Navigate "about:blank"
    If sstI.Tab = 0 Or sstI.Tab = 1 Or sstI.Tab = 2 Or sstI.Tab = 7 Or sstI.Tab = 10 Or sstI.Tab = 11 Then
        cmdCargar.Enabled = True
        cmbMesI.Enabled = True
        AñoI.Enabled = True
    Else
        cmdCargar.Enabled = False
        cargado = False
        cmbMesI.Enabled = False
        AñoI.Enabled = False
    End If
    AñoI = Date
End Sub

Private Sub CargarMes()
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(Date)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i

End Sub

Private Sub CargaGrid(vsfgDesc As Variant, VSFG As Variant, Arch As String)
    Dim i As Long
    Me.MousePointer = 11
    vsfgDesc.FixedRows = 0
    vsfgDesc.LoadGrid App.Path & "\Datos\Desc" & Arch & ".txt", flexFileTabText
    vsfgDesc.FixedRows = 1
    vsfgDesc.Select 1, vsfgDesc.Cols - 1
    vsfgDesc.Sort = flexSortGenericAscending
    VSFG.Cols = 1
    VSFG.Cols = vsfgDesc.Rows
    For i = 1 To vsfgDesc.Rows - 1
        vsfgDesc.RowHidden(i) = True
        VSFG.Cell(flexcpText, 0, i) = vsfgDesc.TextMatrix(i, 1)
        VSFG.ColDataType(i) = flexDTStringC
        If UCase(Trim(vsfgDesc.TextMatrix(i, 7))) <> "DD/MM/YYYY" Then
            VSFG.ColFormat(i) = vsfgDesc.TextMatrix(i, 7)
            VSFG.ColEditMask(i) = vsfgDesc.TextMatrix(i, 8)
        Else
            VSFG.ColFormat(i) = "00/00/0000"
            VSFG.ColEditMask(i) = Replace(vsfgDesc.TextMatrix(i, 8), "/", "\/")
        End If
        If UCase(Left(vsfgDesc.TextMatrix(i, 6), 3)) = "CON" Then
            VSFG.Cell(flexcpBackColor, 0, i) = vbWhite
        End If
        If UCase(Left(vsfgDesc.TextMatrix(i, 5), 3)) = "TAB" Then
            VSFG.ColComboList(i) = LeeTabla(vsfgDesc.TextMatrix(i, 5), vsfgDesc.TextMatrix(i, 9))
        End If
    Next i
    VSFG.ColDataType(0) = flexDTLong
    vsfgDesc.Cols = vsfgDesc.Cols + 1
    vsfgDesc.ColComboList(vsfgDesc.Cols - 1) = "..."
    vsfgDesc.CellButtonPicture = imgHelp
    For i = 1 To vsfgDesc.Rows - 1
        vsfgDesc.Cell(flexcpPicture, i, vsfgDesc.Cols - 1) = Me.imgHelp
    Next i
    vsfgDesc.TextMatrix(0, vsfgDesc.Cols - 1) = "?"
    ReTam vsfgDesc
    VSFG.TextMatrix(0, 0) = "Item"
    VSFG.FrozenCols = 1
    Boton VSFG
    ReTam VSFG
    For i = 7 To vsfgDesc.Cols - 2
        vsfgDesc.ColHidden(i) = True
    Next i
    Me.MousePointer = 0
End Sub

Private Sub Dependencias(vsfgDesc As Variant, VSFG As Variant, Row As Long, Col As Long)
    Dim strLinea As String
    Dim strLinea2 As String
    Dim i As Long
    Dim fin As Long
    Dim ini As Long
    Dim filas As Long
    Dim Campo As Long
    Dim Depende As Long
    Dim Sale As Boolean
    Dim LeeTabla As String
    Me.MousePointer = 11
    If UCase(Left(vsfgDesc.TextMatrix(Col, 5), 3)) = "TAB" Then
        Open App.Path & "\Datos\" & vsfgDesc.TextMatrix(Col, 5) & ".txt" For Input As #1
        filas = 0
        Depende = -100
        Campo = vsfgDesc.TextMatrix(Col, 9)
        If vsfgDesc.TextMatrix(Col, 11) <> "" And vsfgDesc.TextMatrix(Col, 11) <> "=" Then
            Depende = Val(vsfgDesc.TextMatrix(Col, 10))
        ElseIf vsfgDesc.TextMatrix(Col, 11) = "=" Then
            Depende = -1
        End If
        Line Input #1, strLinea
        Do While Not EOF(1)
            Line Input #1, strLinea
            If Depende > 0 Then
                fin = 0
                ini = 0
                Sale = False
                For i = 0 To Val(vsfgDesc.TextMatrix(Col, 11))
                    ini = fin + 1
                    fin = InStr(ini, strLinea, vbTab)
                Next i
                If fin < ini Then
                    fin = Len(strLinea)
                End If
                If InStr(1, Mid(strLinea, ini, fin - ini + 1), VSFG.TextMatrix(Row, vsfgDesc.TextMatrix(Col, 10))) > 0 Then
                    Sale = True
                End If
            ElseIf Depende = 0 Then
                Sale = False
                If InStr(1, vsfgDesc.TextMatrix(Col, 11), Left(strLinea, InStr(1, strLinea, vbTab) - 1)) > 0 Then
                    Sale = True
                End If
            ElseIf Depende = -1 Then
                Sale = False
                If Trim(VSFG.TextMatrix(Row, vsfgDesc.TextMatrix(Col, 10))) = Trim(Left(strLinea, InStr(1, strLinea, vbTab) - 1)) Then
                    Sale = True
                End If
            Else
                Sale = True
            End If
            If Sale = True Then
                fin = 0
                ini = 0
                For i = 0 To Campo
                    ini = fin + 1
                    fin = InStr(ini, strLinea, vbTab)
                Next i
                If fin < ini Then
                    fin = Len(strLinea)
                End If
                strLinea2 = Right(Left(strLinea, fin - 1), Len(Left(strLinea, fin - 1)) - ini + 1)
                If filas = 0 Then
                    strLinea2 = strLinea2 & "*" & Campo
                End If
                filas = filas + 1
                LeeTabla = LeeTabla & "#" & strLinea2 & ";" & strLinea & "|"
            End If
        Loop
        Close #1
        VSFG.ColComboList(Col) = LeeTabla
    Else
        Dim v1 As Variant
        Dim v2 As Variant
        Dim r1 As Long
        Dim c1 As Long
        Dim r2 As Long
        Dim c2 As Long
        VSFG.GetMergedRange Row, Col, r1, c1, r2, c2
        If vsfgDesc.TextMatrix(Col, 11) = "==" Then
            VSFG.Cell(flexcpText, r1, c1, r2, c2) = VSFG.TextMatrix(Row, vsfgDesc.TextMatrix(Col, 11))
        ElseIf vsfgDesc.TextMatrix(Col, 11) = "+" Then
            v1 = Val(VSFG.TextMatrix(Row, Left(vsfgDesc.TextMatrix(Col, 10), InStr(1, vsfgDesc.TextMatrix(Col, 10), "|") - 1)))
            v2 = Val(VSFG.TextMatrix(Row, Right(vsfgDesc.TextMatrix(Col, 10), Len(vsfgDesc.TextMatrix(Col, 10)) - InStr(1, vsfgDesc.TextMatrix(Col, 10), "|"))))
            VSFG.Cell(flexcpText, r1, c1, r2, c2) = v1 + v2
        ElseIf vsfgDesc.TextMatrix(Col, 11) = "-" Then
            v1 = Val(VSFG.TextMatrix(Row, Left(vsfgDesc.TextMatrix(Col, 10), InStr(1, vsfgDesc.TextMatrix(Col, 10), "|") - 1)))
            v2 = Val(VSFG.TextMatrix(Row, Right(vsfgDesc.TextMatrix(Col, 10), Len(vsfgDesc.TextMatrix(Col, 10)) - InStr(1, vsfgDesc.TextMatrix(Col, 10), "|"))))
            VSFG.Cell(flexcpText, r1, c1, r2, c2) = v1 - v2
        ElseIf vsfgDesc.TextMatrix(Col, 11) = "*" Then
            v1 = Val(VSFG.TextMatrix(Row, Left(vsfgDesc.TextMatrix(Col, 10), InStr(1, vsfgDesc.TextMatrix(Col, 10), "|") - 1)))
            v2 = Val(VSFG.TextMatrix(Row, Right(vsfgDesc.TextMatrix(Col, 10), Len(vsfgDesc.TextMatrix(Col, 10)) - InStr(1, vsfgDesc.TextMatrix(Col, 10), "|"))))
            VSFG.Cell(flexcpText, r1, c1, r2, c2) = v1 * v2
        ElseIf vsfgDesc.TextMatrix(Col, 11) = "/" Then
            v1 = Val(VSFG.TextMatrix(Row, Left(vsfgDesc.TextMatrix(Col, 10), InStr(1, vsfgDesc.TextMatrix(Col, 10), "|") - 1)))
            v2 = Val(VSFG.TextMatrix(Row, Right(vsfgDesc.TextMatrix(Col, 10), Len(vsfgDesc.TextMatrix(Col, 10)) - InStr(1, vsfgDesc.TextMatrix(Col, 10), "|"))))
            VSFG.Cell(flexcpText, r1, c1, r2, c2) = v1 / v2
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Function LeeTabla(Tabla As String, Campo As Long)
    Dim strLinea As String
    Dim strLinea2 As String
    Dim i As Long
    Dim fin As Long
    Dim ini As Long
    Dim filas As Long
    Open App.Path & "\Datos\" & Tabla & ".txt" For Input As #1
    filas = 0
    Line Input #1, strLinea
    Do While Not EOF(1)
        Line Input #1, strLinea
        fin = 0
        ini = 0
        For i = 0 To Campo
            ini = fin + 1
            fin = InStr(ini, strLinea, vbTab)
        Next i
        If fin < ini Then
            fin = Len(strLinea)
        End If
        strLinea2 = Right(Left(strLinea, fin - 1), Len(Left(strLinea, fin - 1)) - ini + 1)
        If filas = 0 Then
            strLinea2 = strLinea2 & "*" & Campo
        End If
        filas = filas + 1
        LeeTabla = LeeTabla & "#" & strLinea2 & ";" & strLinea & "|"
    Loop
    Close #1
End Function

Private Sub sstI_Click(PreviousTab As Integer)
    If sstI.Tab = 0 Or sstI.Tab = 1 Or sstI.Tab = 2 Or sstI.Tab = 7 Or sstI.Tab = 10 Or sstI.Tab = 11 Then
        cmdCargar.Enabled = True
        cmbMesI.Enabled = True
        AñoI.Enabled = True
    Else
        cmdCargar.Enabled = False
        cargado = False
        cmbMesI.Enabled = False
        AñoI.Enabled = False
    End If
End Sub


Private Sub vsfgCompras_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescCompras, vsfgCompras, NewRow
    End If
    NuevaLineaUnida "Compras", vsfgCompras, OldRow, OldCol
End Sub


Private Sub vsfgVentasEstablecimiento_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, NewRow
    End If
    NuevaLineaUnida "VentasEstablecimiento", vsfgVentasEstablecimiento, OldRow, OldCol
End Sub


Private Sub vsfgDescInformante_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescInformante.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescVentasEstablecimiento_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescVentasEstablecimiento.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescCompras_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescCompras.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescVentas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescVentas.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescImportaciones_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescImportaciones.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescExportaciones_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescExportaciones.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescTC_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescTC.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescFideicomisos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescFideicomisos.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescAnulados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescAnulados.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescRendimientos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfgDescRendimientos.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDescInformante_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescInformante, "Informante"
End Sub

Private Sub vsfgDescVentasEstablecimiento_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescVentasEstablecimiento, "Ventas por Establecimiento"
End Sub

Private Sub vsfgDescCompras_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescCompras, "Compras"
End Sub

Private Sub vsfgDescVentas_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescVentas, "Ventas"
End Sub

Private Sub vsfgDescImportaciones_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescImportaciones, "Importaciones"
End Sub

Private Sub vsfgDescExportaciones_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescExportaciones, "Exportaciones"
End Sub

Private Sub vsfgDescTC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescTC, "TC"
End Sub

Private Sub vsfgDescFideicomisos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescFideicomisos, "Fideicomisos"
End Sub

Private Sub vsfgDescAnulados_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescAnulados, "Anulados"
End Sub

Private Sub vsfgDescRendimientos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Ayuda vsfgDescRendimientos, "Rendimientos"
End Sub

Private Sub vsfgVentas_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescVentas, vsfgVentas, NewRow
    End If
    NuevaLineaUnida "Ventas", vsfgVentas, OldRow, OldCol
End Sub

Private Sub vsfgImportaciones_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescImportaciones, vsfgImportaciones, NewRow
    End If
    NuevaLineaUnida "Importaciones", vsfgImportaciones, OldRow, OldCol
End Sub

Private Sub vsfgExportaciones_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescExportaciones, vsfgExportaciones, NewRow
    End If
    NuevaLineaUnida "Exportaciones", vsfgExportaciones, OldRow, OldCol
End Sub

Private Sub vsfgTC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescTC, vsfgTC, NewRow
    End If
    NuevaLineaUnida "TC", vsfgTC, OldRow, OldCol
End Sub

Private Sub vsfgFideicomisos_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescFideicomisos, vsfgFideicomisos, NewRow
    End If
    NuevaLineaUnida "Fideicomisos", vsfgFideicomisos, OldRow, OldCol
End Sub

Private Sub vsfgAnulados_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescAnulados, VSFGAnulados, NewRow
    End If
End Sub

Private Sub vsfgRendimientos_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        NuevaLinea vsfgDescRendimientos, vsfgRendimientos, NewRow
    End If
    NuevaLineaUnida "Rendimientos", vsfgRendimientos, OldRow, OldCol
End Sub

Private Sub vsfgCompras_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgCompras, Row
End Sub

Private Sub vsfgVentas_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgVentas, Row
End Sub

Private Sub vsfgVentasEstablecimiento_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgVentasEstablecimiento, Row
End Sub

Private Sub vsfgImportaciones_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgImportaciones, Row
End Sub

Private Sub vsfgExportaciones_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgExportaciones, Row
End Sub

Private Sub vsfgTC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgTC, Row
End Sub

Private Sub vsfgFideicomisos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgFideicomisos, Row
End Sub

Private Sub vsfgAnulados_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila VSFGAnulados, Row
End Sub

Private Sub vsfgRendimientos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EliminarFila vsfgRendimientos, Row
End Sub

Private Sub vsfgInformante_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgInformante
    Else
        Dependencias vsfgDescInformante, vsfgInformante, Row, Col
    End If
End Sub

Private Sub vsfgCompras_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgCompras
    Else
        Dependencias vsfgDescCompras, vsfgCompras, Row, Col
    End If
End Sub

Private Sub vsfgVentas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgVentas
    Else
        Dependencias vsfgDescVentas, vsfgVentas, Row, Col
    End If
End Sub


Private Sub vsfgVentasEstablecimiento_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgVentasEstablecimiento
    Else
        Dependencias vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, Row, Col
    End If
End Sub

Private Sub vsfgImportaciones_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgImportaciones
    Else
        Dependencias vsfgDescImportaciones, vsfgImportaciones, Row, Col
    End If
End Sub

Private Sub vsfgExportaciones_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgExportaciones
    Else
        Dependencias vsfgDescExportaciones, vsfgExportaciones, Row, Col
    End If
End Sub

Private Sub vsfgTC_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgTC
    Else
        Dependencias vsfgDescTC, vsfgTC, Row, Col
    End If
End Sub

Private Sub vsfgFideicomisos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgFideicomisos
    Else
        Dependencias vsfgDescFideicomisos, vsfgFideicomisos, Row, Col
    End If
End Sub

Private Sub vsfgAnulados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton VSFGAnulados
    Else
        Dependencias vsfgDescAnulados, VSFGAnulados, Row, Col
    End If
End Sub

Private Sub vsfgRendimientos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        CambiaBoton vsfgRendimientos
    Else
        Dependencias vsfgDescRendimientos, vsfgRendimientos, Row, Col
    End If
End Sub

Private Sub vsfgInformante_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgInformante
End Sub

'Private Sub vsfgCompras_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    If CambTam = True Then ReTam vsfgCompras
'End Sub

'Private Sub vsfgVentas_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    If CambTam = True Then ReTam vsfgVentas
'End Sub

Private Sub vsfgImportaciones_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgImportaciones
End Sub

Private Sub vsfgExportaciones_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgExportaciones
End Sub

Private Sub vsfgTC_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgTC
End Sub

Private Sub vsfgFideicomisos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgFideicomisos
End Sub

Private Sub vsfgAnulados_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam VSFGAnulados
End Sub

Private Sub vsfgRendimientos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If CambTam = True Then ReTam vsfgRendimientos
End Sub

Private Sub vsfgInformante_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescInformante, vsfgInformante, NewCol
End Sub

Private Sub vsfgCompras_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescCompras, vsfgCompras, NewCol
End Sub

Private Sub vsfgVentas_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescVentas, vsfgVentas, NewCol
End Sub

Private Sub vsfgVentasEstablecimiento_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, NewCol
End Sub

Private Sub vsfgImportaciones_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescImportaciones, vsfgImportaciones, NewCol
End Sub

Private Sub vsfgExportaciones_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescExportaciones, vsfgExportaciones, NewCol
End Sub

Private Sub vsfgTC_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescTC, vsfgTC, NewCol
End Sub

Private Sub vsfgFideicomisos_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescFideicomisos, vsfgFideicomisos, NewCol
End Sub

Private Sub vsfgAnulados_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescAnulados, VSFGAnulados, NewCol
End Sub

Private Sub vsfgRendimientos_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    CambiarDesc vsfgDescRendimientos, vsfgRendimientos, NewCol
End Sub

Private Sub vsfgInformante_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescInformante, vsfgInformante, Row, Col)
End Sub

Private Sub vsfgCompras_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescCompras, vsfgCompras, Row, Col)
End Sub

Private Sub vsfgVentas_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescVentas, vsfgVentas, Row, Col)
End Sub

Private Sub vsfgImportaciones_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescImportaciones, vsfgImportaciones, Row, Col)
End Sub

Private Sub vsfgExportaciones_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescExportaciones, vsfgExportaciones, Row, Col)
End Sub

Private Sub vsfgTC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescTC, vsfgTC, Row, Col)
End Sub

Private Sub vsfgFideicomisos_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescFideicomisos, vsfgFideicomisos, Row, Col)
End Sub

Private Sub vsfgAnulados_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescAnulados, VSFGAnulados, Row, Col)
End Sub

Private Sub vsfgRendimientos_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Validar(vsfgDescRendimientos, vsfgRendimientos, Row, Col)
End Sub

Private Sub EliminarFila(VSFG As Variant, Row As Long)
    If VSFG.Rows <= 2 Then
        Limpiar VSFG
    Else
        VSFG.RemoveItem Row
        Boton VSFG
    End If
End Sub

Private Sub ReTam(VSFG As Variant)
    VSFG.AutoSize 0, VSFG.Cols - 1, False
End Sub

Private Sub MoverMouse(vsfgDesc As Variant, VSFG As Variant)
    Dim i As Long
    If VSFG.MouseCol <> -1 Then
        For i = 1 To vsfgDesc.Rows - 1
            vsfgDesc.RowHidden(i) = True
        Next i
        vsfgDesc.RowHidden(VSFG.MouseCol) = False
        VSFG.ToolTipText = vsfgDesc.TextMatrix(VSFG.MouseCol, 0)
    End If
End Sub

Private Sub CambiarDesc(vsfgDesc As Variant, VSFG As Variant, Col As Long)
    Dim i As Long
    For i = 1 To vsfgDesc.Rows - 1
        vsfgDesc.RowHidden(i) = True
    Next i
    vsfgDesc.RowHidden(Col) = False
    ReTam vsfgDesc
    VSFG.ToolTipText = vsfgDesc.TextMatrix(Col, 0)
End Sub

Private Function Comprobar(vsfgDesc As Variant, VSFG As Variant, Row As Long, Col As Long) As Variant

Dim Valor As Variant
    
    Valor = VSFG.Cell(flexcpTextDisplay, Row, vsfgDesc.TextMatrix(Col, 13))
    Valor = Left(Trim(Valor), vsfgDesc.TextMatrix(Col, 3))
    If vsfgDesc.TextMatrix(Col, 7) = "##################################" Then
        Valor = Left(Valor, Len(vsfgDesc.TextMatrix(Col, 7)))
    ElseIf UCase(Trim(vsfgDesc.TextMatrix(Col, 7))) <> "DD/MM/YYYY" And UCase(Trim(vsfgDesc.TextMatrix(Col, 7))) <> "" Then
        Valor = Format(Valor, vsfgDesc.TextMatrix(Col, 7))
    End If
    If UCase(Trim(vsfgDesc.TextMatrix(Col, 7))) = "MM/YYYY" Then
        Valor = Format(Valor, Replace(vsfgDesc.TextMatrix(Col, 7), "/", "\/"))
    End If
    Comprobar = Valor
    If UCase(Left(vsfgDesc.TextMatrix(Col, 4), 3)) = "FEC" Then
        Comprobar = Valor
        If Len(Trim(vsfgDesc.TextMatrix(Col, 7))) = 10 And IsDate(Valor) = False Then
            Comprobar = ""
        End If
    ElseIf vsfgDesc.TextMatrix(Col, 7) = "##################################" Then
        Comprobar = Left(Valor, Len(vsfgDesc.TextMatrix(Col, 7)))
    ElseIf vsfgDesc.TextMatrix(Col, 7) = "0000000000###" Then
        If Len(Format(Valor, "#0")) <= 10 Then
            Comprobar = Format(Valor, "0000000000")
        Else
            Comprobar = Format(Valor, "0000000000000")
        End If
    ElseIf InStr(1, vsfgDesc.TextMatrix(Col, 7), "0") > 0 And Valor <> "" Then
        Comprobar = (Format(Left(Val(Valor), vsfgDesc.TextMatrix(Col, 3)), vsfgDesc.TextMatrix(Col, 7)))
    End If
End Function

Private Function Validar(vsfgDesc As Variant, VSFG As Variant, Row As Long, Col As Long) As Boolean
    Dim Valor As Variant
    If Col <> 0 Then
    Valor = VSFG.EditText
    Valor = Left(Trim(Valor), vsfgDesc.TextMatrix(Col, 3))
    If UCase(Trim(vsfgDesc.TextMatrix(Col, 7))) <> "DD/MM/YYYY" Then
        Valor = Format(Valor, vsfgDesc.TextMatrix(Col, 7))
    End If
    Validar = False
    If UCase(Left(vsfgDesc.TextMatrix(Col, 4), 3)) = "FEC" Then
        If IsDate(Valor) = False Then
            Validar = True
            MsgBox "Campo: " & vsfgDesc.TextMatrix(Col, 0) & vbNewLine & "Nombre XML: " & vsfgDesc.TextMatrix(Col, 1) & vbNewLine & "No cumple el requerimeinto de Fecha", vbCritical, "Fecha"
        End If
    ElseIf UCase(Left(vsfgDesc.TextMatrix(Col, 4), 3)) = "NUM" Then
        Valor = (Format(Left(Valor, vsfgDesc.TextMatrix(Col, 3)), Replace(vsfgDesc.TextMatrix(Col, 7), "/", "\/")))
    End If
    If Len(Valor) < Val(vsfgDesc.TextMatrix(Col, 2)) Or Len(Valor) > Val(vsfgDesc.TextMatrix(Col, 3)) Then
        Validar = True
        MsgBox "Campo: " & vsfgDesc.TextMatrix(Col, 0) & vbNewLine & "Nombre XML: " & vsfgDesc.TextMatrix(Col, 1) & vbNewLine & "No cumple el requerimeinto de Longitud", vbCritical, "Longitud"
    End If
    If Trim(Valor) <> "" Then
        VSFG.EditText = Valor
    Else
        Validar = False
    End If
    End If
End Function

Private Sub Guarda(VSFG As Variant, Hoja As String, Archivo As String)
    Me.MousePointer = 11
    Set fs = CreateObject("Scripting.FileSystemObject")
    VSFG.SaveGrid Hoja & ".txt", flexFileAll
    VSFG.Archive Archivo, Hoja & ".txt", arcAdd
    fs.DeleteFile Hoja & ".txt", True
    Set fs = Nothing
    Me.MousePointer = 0
End Sub

Private Sub Abrir(VSFG As Variant, Hoja As String, Archivo As String)
    Me.MousePointer = 11
    Set fs = CreateObject("Scripting.FileSystemObject")
    VSFG.Archive Archivo, Hoja & ".txt", arcExtract
    VSFG.LoadGrid Hoja & ".txt", flexFileAll
    fs.DeleteFile Hoja & ".txt", True
    Set fs = Nothing
    Boton VSFG
    Me.MousePointer = 0
End Sub

Private Sub Limpiar(VSFG As Variant)
    VSFG.Clear 1
    VSFG.Rows = 1
    VSFG.AddItem ""
    Boton VSFG
End Sub

Private Sub Importar(vsfgDesc As Variant, VSFG As Variant, Hoja As String)
    CD.DefaultExt = "txt"
    CD.Filter = "Texto (delimitado por tabulaciones) (*.txt)|*.txt|Todos |*.*"
    CD.FilterIndex = 1
    CD.FileName = ""
    CD.ShowOpen
    If CD.FileName <> "" Then
        Me.MousePointer = 11
        CambTam = False
        VSFG.LoadGrid CD.FileName, flexFileTabText, True
        VSFG.Cols = vsfgDesc.Rows
        CambTam = True
        UnirFilas VSFG, Hoja
        Me.MousePointer = 0
    End If
End Sub

Private Sub Exportar(vsfgDesc As Variant, VSFG As Variant, Hoja As String)
    CD.DefaultExt = "txt"
    CD.Filter = "Texto (delimitado por tabulaciones) (*.txt)|*.txt|Todos |*.*"
    CD.FilterIndex = 1
    CD.FileName = ""
    CD.ShowSave
    If CD.FileName <> "" Then
        Me.MousePointer = 11
        VSFG.SaveGrid CD.FileName, flexFileTabText, True
        Me.MousePointer = 0
    End If
End Sub


Private Sub Boton(VSFG As Variant)
    Dim i As Long
    Dim j As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    VSFG.ColComboList(0) = "..."
    VSFG.ColAlignment(0) = flexAlignLeftCenter
    j = 1
    For i = 1 To VSFG.Rows - 1
        VSFG.Cell(flexcpPictureAlignment, i, 0) = flexPicAlignRightCenter
        VSFG.Cell(flexcpPicture, i, 0) = Me.imgBtnUp
        If VSFG.Rows > 10 And i < VSFG.Rows - 11 Then
            VSFG.ShowCell i + 10, 0
        End If
        VSFG.GetMergedRange i, 0, r1, c1, r2, c2
        If CambTam = True Then VSFG.TextMatrix(i, 0) = j
        
        If i = r2 Then
            j = j + 1
        Else
'            MsgBox r1
        End If
    Next i
End Sub

Private Sub CambiaBoton(VSFG As Variant)
    VSFG.CellButtonPicture = Me.imgBtnUp
End Sub

Private Sub NuevaLinea(vsfgDesc As Variant, VSFG As Variant, Row As Long)
    VerificaLinea vsfgDesc, VSFG, Row
    If VerificaLinea(vsfgDesc, VSFG, VSFG.Rows - 1) = True Then
        VSFG.AddItem ""
        VSFG.Cell(flexcpPictureAlignment, VSFG.Rows - 1, 0) = flexPicAlignRightCenter
        VSFG.Cell(flexcpPicture, VSFG.Rows - 1, 0) = Me.imgBtnUp
        VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
    End If
End Sub
Private Function VerificaLinea(vsfgDesc As Variant, VSFG As Variant, Row As Long) As Boolean
    Dim i As Long
    VerificaLinea = True
    For i = 1 To vsfgDesc.Rows - 1
        If VSFG.TextMatrix(Row, vsfgDesc.TextMatrix(i, 13)) = "" And UCase(Left(vsfgDesc.TextMatrix(i, 6), 3)) = "OBL" Then
            VerificaLinea = False
            VSFG.Cell(flexcpBackColor, Row, vsfgDesc.TextMatrix(i, 13)) = vbYellow
        Else
            VSFG.Cell(flexcpBackColor, Row, vsfgDesc.TextMatrix(i, 13)) = vbWhite
        End If
    Next i
End Function

'****************************************************************
'XML
' Guardar Valores XML.
Private Sub SaveValues(Arch As String)
    Dim xml_doc As DOMDocument
    Dim Nodo As IXMLDOMElement
    Dim Nodo2 As IXMLDOMElement
    Dim Nodo3 As IXMLDOMElement
    Dim Nodo4 As IXMLDOMElement
    Dim Nodo5 As IXMLDOMElement
    Dim Coment As IXMLDOMComment
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Integer
    Dim ll As Boolean
    Dim x As Long
    Dim r1 As Long
    Dim c1 As Long
    Dim r2 As Long
    Dim c2 As Long
    

    Me.MousePointer = 11
    ' Create the XML document.
    Set xml_doc = New DOMDocument

    ' Nodo Principal.
    Set Nodo = xml_doc.createElement("iva")
    Set Coment = xml_doc.createComment("Creado y Editado por ATED v1.0 (http://www.enlace-digital.com) de Andres Cevallos (Enlace Digital)")
    ' Add the Values section node to the document.
    xml_doc.appendChild Coment
    xml_doc.appendChild Nodo

    ' Informante
    vsfgDescInformante.Select 1, 12
    vsfgDescInformante.Sort = flexSortGenericAscending
    For i = 1 To vsfgInformante.Rows - 1
        If LineaAmarilla(vsfgDescInformante, vsfgInformante, i) = False Then
            For j = 1 To vsfgDescInformante.Rows - 1
                If vsfgInformante.TextMatrix(i, vsfgDescInformante.TextMatrix(j, 13)) <> "" Then
                    CreateNode Nodo, vsfgInformante.TextMatrix(0, vsfgDescInformante.TextMatrix(j, 13)), vsfgInformante.TextMatrix(i, vsfgDescInformante.TextMatrix(j, 13))
                End If
            Next j
        End If
    Next i
    vsfgDescInformante.Select 1, 13
    vsfgDescInformante.Sort = flexSortGenericAscending
    ' Compras
    l = 0
    ll = False
    vsfgDescCompras.Select 1, 12
    vsfgDescCompras.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescCompras, vsfgCompras, 1) = False Then
        Set Nodo2 = xml_doc.createElement("compras")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgCompras.Rows - 1
            If LineaAmarilla(vsfgDescCompras, vsfgCompras, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleCompras")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescCompras.Rows - 1
                    If vsfgDescCompras.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("air")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescCompras.TextMatrix(j, 1) = "pagoLocExt" Then
                        Set Nodo4 = xml_doc.createElement("pagoExterior")
                        Nodo3.appendChild Nodo4
                    End If
                    vsfgCompras.ShowCell i, j
                    If vsfgDescCompras.TextMatrix(j, 1) = "formaPago" And Trim(vsfgCompras.TextMatrix(i, vsfgDescCompras.TextMatrix(j, 13))) <> "" Then
                        Set Nodo4 = xml_doc.createElement("formasDePago")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescCompras.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgCompras.TextMatrix(i, vsfgDescCompras.TextMatrix(j, 13))) <> "" Then
                        vsfgCompras.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAir")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 6
                                CreateNode Nodo5, vsfgCompras.TextMatrix(0, vsfgDescCompras.TextMatrix(k, 13)), Comprobar(vsfgDescCompras, vsfgCompras, x, k)
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescCompras.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescCompras.TextMatrix(j, 1) = "pagoLocExt" Then
                        For k = j To j + 8
                            CreateNode Nodo4, vsfgCompras.TextMatrix(0, vsfgDescCompras.TextMatrix(k, 13)), Comprobar(vsfgDescCompras, vsfgCompras, i, k)
                        Next k
                        j = k - 1
                    ElseIf vsfgDescCompras.TextMatrix(j, 1) = "formaPago" Then
                        For k = j To j + 0
                            CreateNode Nodo4, vsfgCompras.TextMatrix(0, vsfgDescCompras.TextMatrix(k, 13)), Comprobar(vsfgDescCompras, vsfgCompras, i, k)
                        Next k
                        j = k - 1
                    ElseIf vsfgDescCompras.TextMatrix(j - l, 1) <> "codRetAir" Then
                        CreateNode Nodo3, vsfgCompras.TextMatrix(0, vsfgDescCompras.TextMatrix(j, 13)), Comprobar(vsfgDescCompras, vsfgCompras, i, j)
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 3 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescCompras.Select 1, 13
    vsfgDescCompras.Sort = flexSortGenericAscending
    
    ' Ventas
    l = 0
    ll = False
    vsfgDescVentas.Select 1, 12
    vsfgDescVentas.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescVentas, vsfgVentas, 1) = False Then
        Set Nodo2 = xml_doc.createElement("ventas")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgVentas.Rows - 1
            If LineaAmarilla(vsfgDescVentas, vsfgVentas, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleVentas")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescVentas.Rows - 1
                    If vsfgDescVentas.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("air")
                        Nodo3.appendChild Nodo4
                    ElseIf vsfgDescVentas.TextMatrix(j, 1) = "tipoCompe" Then
                        Set Nodo4 = xml_doc.createElement("compensaciones")
                        Nodo3.appendChild Nodo4
                        Set Nodo5 = xml_doc.createElement("compensacion")
                        Nodo4.appendChild Nodo5
                    ElseIf vsfgDescVentas.TextMatrix(j, 1) = "formaPago" Then
                        Set Nodo4 = xml_doc.createElement("formadDePago")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescVentas.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgVentas.TextMatrix(i, vsfgDescVentas.TextMatrix(j, 13))) <> "" Then
                        vsfgVentas.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAir")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 3
                                CreateNode Nodo5, vsfgVentas.TextMatrix(0, vsfgDescVentas.TextMatrix(k, 13)), Comprobar(vsfgDescVentas, vsfgVentas, x, k)
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescVentas.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescVentas.TextMatrix(j - l, 1) = "tipoCompe" Then
                        For k = j To j + 1
                            CreateNode Nodo5, vsfgVentas.TextMatrix(0, vsfgDescVentas.TextMatrix(k, 13)), Comprobar(vsfgDescVentas, vsfgVentas, i, k)
                        Next k
                        j = k - 1
                    ElseIf vsfgDescVentas.TextMatrix(j - l, 1) = "formaPago" Then
                        CreateNode Nodo4, vsfgVentas.TextMatrix(0, vsfgDescVentas.TextMatrix(j, 13)), Comprobar(vsfgDescVentas, vsfgVentas, i, j)
                    ElseIf vsfgDescVentas.TextMatrix(j - l, 1) <> "codRetAir" Then
                        CreateNode Nodo3, vsfgVentas.TextMatrix(0, vsfgDescVentas.TextMatrix(j, 13)), Comprobar(vsfgDescVentas, vsfgVentas, i, j)
                        
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 3 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescVentas.Select 1, 13
    vsfgDescVentas.Sort = flexSortGenericAscending
    'ventas por establecimiento
    
    vsfgDescVentasEstablecimiento.Select 1, 12
    vsfgDescVentasEstablecimiento.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, 1) = False Then
        Set Nodo2 = xml_doc.createElement("ventasEstablecimiento")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgVentasEstablecimiento.Rows - 1
            If LineaAmarilla(vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, i) = False Then
                Set Nodo3 = xml_doc.createElement("ventaEst")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescVentasEstablecimiento.Rows - 1
                    CreateNode Nodo3, vsfgVentasEstablecimiento.TextMatrix(0, vsfgDescVentasEstablecimiento.TextMatrix(j, 13)), Comprobar(vsfgDescVentasEstablecimiento, vsfgVentasEstablecimiento, i, j)
                Next j
            End If
        Next i
    End If
    vsfgDescVentasEstablecimiento.Select 1, 13
    vsfgDescVentasEstablecimiento.Sort = flexSortGenericAscending
    
    ' importaciones
    l = 0
    ll = False
    vsfgDescImportaciones.Select 1, 12
    vsfgDescImportaciones.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescImportaciones, vsfgImportaciones, 1) = False Then
        Set Nodo2 = xml_doc.createElement("importaciones")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgImportaciones.Rows - 1
            If LineaAmarilla(vsfgDescImportaciones, vsfgImportaciones, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleImportaciones")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescImportaciones.Rows - 1
                    If vsfgDescImportaciones.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("air")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescImportaciones.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgImportaciones.TextMatrix(i, vsfgDescImportaciones.TextMatrix(j, 13))) <> "" Then
                        vsfgImportaciones.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAir")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 3
                                If (UCase(Left(vsfgDescImportaciones.TextMatrix(k, 6), 3)) = "CON" And Trim(vsfgImportaciones.TextMatrix(x, vsfgDescImportaciones.TextMatrix(k, 13))) <> "") Or UCase(Left(vsfgDescImportaciones.TextMatrix(k, 6), 3)) = "OBL" Then
                                    CreateNode Nodo5, vsfgImportaciones.TextMatrix(0, vsfgDescImportaciones.TextMatrix(k, 13)), Comprobar(vsfgDescImportaciones, vsfgImportaciones, x, k)
                                End If
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescImportaciones.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescImportaciones.TextMatrix(j - l, 1) <> "codRetAir" Then
                        If (UCase(Left(vsfgDescImportaciones.TextMatrix(j, 6), 3)) = "CON" And Trim(vsfgImportaciones.TextMatrix(i, vsfgDescImportaciones.TextMatrix(j, 13))) <> "") Or UCase(Left(vsfgDescImportaciones.TextMatrix(j, 6), 3)) = "OBL" Then
                            CreateNode Nodo3, vsfgImportaciones.TextMatrix(0, vsfgDescImportaciones.TextMatrix(j, 13)), Comprobar(vsfgDescImportaciones, vsfgImportaciones, i, j)
                        End If
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 3 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescImportaciones.Select 1, 13
    vsfgDescImportaciones.Sort = flexSortGenericAscending
    
    ' exportaciones
    l = 0
    ll = False
    vsfgDescExportaciones.Select 1, 12
    vsfgDescExportaciones.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescExportaciones, vsfgExportaciones, 1) = False Then
        Set Nodo2 = xml_doc.createElement("exportaciones")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgExportaciones.Rows - 1
            If LineaAmarilla(vsfgDescExportaciones, vsfgExportaciones, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleExportaciones")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescExportaciones.Rows - 1
                    If vsfgDescExportaciones.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("air")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescExportaciones.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgExportaciones.TextMatrix(i, vsfgDescExportaciones.TextMatrix(j, 13))) <> "" Then
                        vsfgExportaciones.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAir")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 3
                                If (UCase(Left(vsfgDescExportaciones.TextMatrix(k, 6), 3)) = "CON" And Trim(vsfgExportaciones.TextMatrix(x, vsfgDescExportaciones.TextMatrix(k, 13))) <> "") Or UCase(Left(vsfgDescExportaciones.TextMatrix(k, 6), 3)) = "OBL" Then
                                    CreateNode Nodo5, vsfgExportaciones.TextMatrix(0, vsfgDescExportaciones.TextMatrix(k, 13)), Comprobar(vsfgDescExportaciones, vsfgExportaciones, x, k)
                                End If
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescExportaciones.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescExportaciones.TextMatrix(j - l, 1) <> "codRetAir" Then
                        If (UCase(Left(vsfgDescExportaciones.TextMatrix(j, 6), 3)) = "CON" And Trim(vsfgExportaciones.TextMatrix(i, vsfgDescExportaciones.TextMatrix(j, 13))) <> "") Or UCase(Left(vsfgDescExportaciones.TextMatrix(j, 6), 3)) = "OBL" Then
                            CreateNode Nodo3, vsfgExportaciones.TextMatrix(0, vsfgDescExportaciones.TextMatrix(j, 13)), Comprobar(vsfgDescExportaciones, vsfgExportaciones, i, j)
                        End If
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 3 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescExportaciones.Select 1, 13
    vsfgDescExportaciones.Sort = flexSortGenericAscending
    
    ' recap - tarjetas de credito
    l = 0
    ll = False
    vsfgDescTC.Select 1, 12
    vsfgDescTC.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescTC, vsfgTC, 1) = False Then
        Set Nodo2 = xml_doc.createElement("recap")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgTC.Rows - 1
            If LineaAmarilla(vsfgDescTC, vsfgTC, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleRecap")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescTC.Rows - 1
                    If vsfgDescTC.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("air")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescTC.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgTC.TextMatrix(i, vsfgDescTC.TextMatrix(j, 13))) <> "" Then
                        vsfgTC.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAir")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 3
                                CreateNode Nodo5, vsfgTC.TextMatrix(0, vsfgDescTC.TextMatrix(k, 13)), Comprobar(vsfgDescTC, vsfgTC, x, k)
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescTC.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescTC.TextMatrix(j - l, 1) <> "codRetAir" Then
                        CreateNode Nodo3, vsfgTC.TextMatrix(0, vsfgDescTC.TextMatrix(j, 13)), Comprobar(vsfgDescTC, vsfgTC, i, j)
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 3 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescTC.Select 1, 13
    vsfgDescTC.Sort = flexSortGenericAscending
    
    ' fideicomisos
    l = 0
    ll = False
    vsfgDescFideicomisos.Select 1, 12
    vsfgDescFideicomisos.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescFideicomisos, vsfgFideicomisos, 1) = False Then
        Set Nodo2 = xml_doc.createElement("fideicomisos")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgFideicomisos.Rows - 1
            If LineaAmarilla(vsfgDescFideicomisos, vsfgFideicomisos, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleFideicomisos")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescFideicomisos.Rows - 1
                    If vsfgDescFideicomisos.TextMatrix(j, 1) = "tipoFideicomiso" Then
                        Set Nodo4 = xml_doc.createElement("fValor")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescFideicomisos.TextMatrix(j, 1) = "tipoFideicomiso" And Trim(vsfgFideicomisos.TextMatrix(i, vsfgDescFideicomisos.TextMatrix(j, 13))) <> "" Then
                        vsfgFideicomisos.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detallefValor")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 4
                                CreateNode Nodo5, vsfgFideicomisos.TextMatrix(0, vsfgDescFideicomisos.TextMatrix(k, 13)), Comprobar(vsfgDescFideicomisos, vsfgFideicomisos, x, k)
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescFideicomisos.TextMatrix(j, 1) = "tipoFideicomiso" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescFideicomisos.TextMatrix(j - l, 1) <> "tipoFideicomiso" Then
                        CreateNode Nodo3, vsfgFideicomisos.TextMatrix(0, vsfgDescFideicomisos.TextMatrix(j, 13)), Comprobar(vsfgDescFideicomisos, vsfgFideicomisos, i, j)
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 4 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescFideicomisos.Select 1, 13
    vsfgDescFideicomisos.Sort = flexSortGenericAscending
    
    ' anulados
    vsfgDescAnulados.Select 1, 12
    vsfgDescAnulados.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescAnulados, VSFGAnulados, 1) = False Then
        Set Nodo2 = xml_doc.createElement("anulados")
        Nodo.appendChild Nodo2
        For i = 1 To VSFGAnulados.Rows - 1
            If LineaAmarilla(vsfgDescAnulados, VSFGAnulados, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleAnulados")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescAnulados.Rows - 1
                    CreateNode Nodo3, VSFGAnulados.TextMatrix(0, vsfgDescAnulados.TextMatrix(j, 13)), Comprobar(vsfgDescAnulados, VSFGAnulados, i, j)
                Next j
            End If
        Next i
    End If
    vsfgDescAnulados.Select 1, 13
    vsfgDescAnulados.Sort = flexSortGenericAscending
    
    ' rendimientos
    l = 0
    ll = False
    vsfgDescRendimientos.Select 1, 12
    vsfgDescRendimientos.Sort = flexSortGenericAscending
    If LineaAmarilla(vsfgDescRendimientos, vsfgRendimientos, 1) = False Then
        Set Nodo2 = xml_doc.createElement("rendFinancieros")
        Nodo.appendChild Nodo2
        For i = 1 To vsfgRendimientos.Rows - 1
            If LineaAmarilla(vsfgDescRendimientos, vsfgRendimientos, i) = False Then
                Set Nodo3 = xml_doc.createElement("detalleRendFinancieros")
                Nodo2.appendChild Nodo3
                For j = 1 To vsfgDescRendimientos.Rows - 1
                    If vsfgDescRendimientos.TextMatrix(j, 1) = "codRetAir" Then
                        Set Nodo4 = xml_doc.createElement("airRend")
                        Nodo3.appendChild Nodo4
                    End If
                    If vsfgDescRendimientos.TextMatrix(j, 1) = "codRetAir" And Trim(vsfgRendimientos.TextMatrix(i, vsfgDescRendimientos.TextMatrix(j, 13))) <> "" Then
                        vsfgRendimientos.GetMergedRange i, 0, r1, c1, r2, c2
                        For x = i To r2
                            Set Nodo5 = xml_doc.createElement("detalleAirRen")
                            Nodo4.appendChild Nodo5
                            For k = j To j + 4
                                CreateNode Nodo5, vsfgRendimientos.TextMatrix(0, vsfgDescRendimientos.TextMatrix(k, 13)), Comprobar(vsfgDescRendimientos, vsfgRendimientos, x, k)
                            Next k
                        Next x
                        j = k - 1
                        i = r2
                    ElseIf vsfgDescRendimientos.TextMatrix(j, 1) = "codRetAir" Then
                        l = 1
                        ll = True
                    ElseIf vsfgDescRendimientos.TextMatrix(j - l, 1) <> "codRetAir" Then
                        CreateNode Nodo3, vsfgRendimientos.TextMatrix(0, vsfgDescRendimientos.TextMatrix(j, 13)), Comprobar(vsfgDescRendimientos, vsfgRendimientos, i, j)
                    Else
                        If ll = True Then
                            l = l + 1
                        End If
                        If l > 4 Then
                            l = 0
                            ll = False
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    vsfgDescRendimientos.Select 1, 13
    vsfgDescRendimientos.Sort = flexSortGenericAscending

    ' Save the XML document.
    'xml.Doc
    xml_doc.Save Arch
    Me.MousePointer = 0
    Set xml_doc = Nothing
    Set Nodo = Nothing
    Set Nodo2 = Nothing
    Set Nodo3 = Nothing
    Set Nodo4 = Nothing
    Set Nodo5 = Nothing
    Set Coment = Nothing
End Sub
Private Function LineaAmarilla(vsfgDesc As Variant, VSFG As Variant, Linea As Long) As Boolean
    Dim i As Long
    VerificaLinea vsfgDesc, VSFG, Linea
    LineaAmarilla = False
    For i = 1 To VSFG.Cols - 1
        If VSFG.Cell(flexcpBackColor, Linea, i) = vbYellow Then
            LineaAmarilla = True
        End If
    Next i
End Function

' Add a new node to the indicated parent node.
Private Sub CreateNode(ByVal parent As IXMLDOMNode, ByVal node_name As String, ByVal node_value As String)
    Dim new_node As IXMLDOMNode

    ' Create the new node.
    If node_value = "" Then Exit Sub
    Set new_node = parent.ownerDocument.createElement(node_name)

    ' Set the node's text value.
    new_node.Text = node_value

    ' Add the node to the parent.
    parent.appendChild new_node
End Sub

Private Sub Ayuda(vsfgDesc As Variant, Tipo As String)
    frmAyuda.Br.Navigate App.Path & "\Ayuda\" & Tipo & vsfgDesc.TextMatrix(vsfgDesc.Row, 1) & ".htm"
    frmAyuda.Show vbModal
End Sub
Private Sub NuevaLineaUnida(Tipo As String, VSFG As Variant, Row As Long, Col As Long)
    Dim i As Long
    'If cargado = False Then
    If Row < vsfgCompras.Rows And Row > 1 Then
    If Tipo = "Compras" Or Tipo = "Ventas" Or Tipo = "Importaciones" Or Tipo = "Exportaciones" Or Tipo = "TC" Or Tipo = "Rendimientos" Then
        If VSFG.TextMatrix(0, Col) = "codRetAir" And VSFG.TextMatrix(Row, Col) <> "" Then
            If MsgBox("Quiere crear una nueva linea de Retencion para este registro", vbYesNo + vbDefaultButton2, Tipo) = vbYes Then
                VSFG.AddItem "", Row + 1
                For i = 0 To VSFG.Cols - 1
                    If Not i >= Col Then
                        VSFG.MergeCol(i) = True
                        VSFG.TextMatrix(Row + 1, i) = VSFG.TextMatrix(Row, i)
                    Else
                        VSFG.MergeCol(i) = False
                    End If
                Next i
            End If
        End If
    End If
    If Tipo = "Fideicomisos" Then
        If VSFG.TextMatrix(0, Col) = "tipoFideicomiso" And VSFG.TextMatrix(Row, Col) <> "" Then
            If MsgBox("Quiere crear una nueva linea de Retencion para este registro", vbYesNo + vbDefaultButton2, Tipo) = vbYes Then
                VSFG.AddItem "", Row + 1
                For i = 0 To VSFG.Cols - 1
                    If Not i >= Col Then
                        VSFG.MergeCol(i) = True
                        VSFG.TextMatrix(Row + 1, i) = VSFG.TextMatrix(Row, i)
                    Else
                        VSFG.MergeCol(i) = False
                    End If
                Next i
            End If
        End If
    End If
    End If
End Sub

Private Sub UnirFilas(VSFG As Variant, Tipo As String)
    Dim i As Long
    i = 0
    If Tipo = "Compras" Or Tipo = "Ventas" Or Tipo = "Importaciones" Or Tipo = "Exportaciones" Or Tipo = "TC" Or Tipo = "Rendimientos" Then
        Do While (VSFG.TextMatrix(0, i) <> "codRetAir" And i <= VSFG.Cols - 1)
            VSFG.MergeCol(i) = True
            i = i + 1
            If i > VSFG.Cols - 1 Then Exit Do
        Loop
        For i = i To VSFG.Cols - 1
            VSFG.MergeCol(i) = False
        Next i
    ElseIf Tipo = "Fideicomisos" Then
        Do While (VSFG.TextMatrix(0, i) <> "tipoFideicomiso" And i <= VSFG.Cols - 1)
            VSFG.MergeCol(i) = True
            i = i + 1
        Loop
        For i = i To VSFG.Cols - 1
            VSFG.MergeCol(i) = False
        Next i
    End If
    Boton VSFG
End Sub


Private Sub JuntarRows(VSFG As Variant)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim ii As Long
    Dim jj As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim rr1 As Long
    Dim rr2 As Long
    Dim cc1 As Long
    Dim cc2 As Long
    Dim Nueva As Boolean
    Dim salir As Boolean
    Dim Salir2 As Boolean
    CambTam = False
    Me.MousePointer = 11
    With VSFG
        .ColDataType(0) = flexDTLong
        .Select 1, 2
        .Sort = flexSortGenericDescending
        i = 1
        While salir = False
            .GetMergedRange i, 0, r1, c1, r2, c2
            i = r2
            .GetMergedRange i + 1, 0, rr1, cc1, rr2, cc2
            ii = rr2
            If Right(.TextMatrix(i, 2), 6) = "999999" Then
                'MsgBox "390018800001"
            End If
            If Format(.TextMatrix(i, 1), "00") = Format(.TextMatrix(ii, 1), "00") And _
               Format(.TextMatrix(i, 2), vsfgDescVentas.TextMatrix(2, 7)) = Format(.TextMatrix(ii, 2), vsfgDescVentas.TextMatrix(2, 7)) And _
               .TextMatrix(i, 3) = .TextMatrix(ii, 3) And _
               .TextMatrix(i, 4) = .TextMatrix(ii, 4) And _
               .TextMatrix(i, 6) = .TextMatrix(ii, 6) And _
               .TextMatrix(i, 8) = .TextMatrix(ii, 8) And _
               Val(.TextMatrix(i, 10)) = Val(.TextMatrix(ii, 10)) And _
               Val(.TextMatrix(i, 13)) = Val(.TextMatrix(ii, 13)) Then
                
                .Cell(flexcpText, r1, 5, r2, 5) = Val(.TextMatrix(i, 5)) + Val(.TextMatrix(ii, 5))
                .Cell(flexcpText, r1, 7, r2, 7) = Val(.TextMatrix(i, 7)) + Val(.TextMatrix(ii, 7))
                .Cell(flexcpText, r1, 9, r2, 9) = Val(.TextMatrix(i, 9)) + Val(.TextMatrix(ii, 9))
                .Cell(flexcpText, r1, 11, r2, 11) = Val(.TextMatrix(i, 11)) + Val(.TextMatrix(ii, 11))
                .Cell(flexcpText, r1, 12, r2, 12) = Val(.TextMatrix(i, 12)) + Val(.TextMatrix(ii, 12))
                .Cell(flexcpText, r1, 14, r2, 14) = Val(.TextMatrix(i, 14)) + Val(.TextMatrix(ii, 14))
                .Cell(flexcpText, r1, 15, r2, 15) = Val(.TextMatrix(i, 15)) + Val(.TextMatrix(ii, 15))
                .Cell(flexcpText, r1, 17, r2, 17) = Val(.TextMatrix(i, 17)) + Val(.TextMatrix(ii, 17))
                .Cell(flexcpText, r1, 18, r2, 18) = Val(.TextMatrix(i, 18)) + Val(.TextMatrix(ii, 18))
                .Cell(flexcpText, r1, 20, r2, 20) = Val(.TextMatrix(i, 20)) + Val(.TextMatrix(ii, 20))
                If .TextMatrix(i, 16) < .TextMatrix(ii, 16) Then
                    .Cell(flexcpText, r1, 16, r2, 16) = .TextMatrix(ii, 16)
                Else
                    .TextMatrix(ii, 16) = .TextMatrix(i, 16)
                End If
                If .TextMatrix(i, 19) < .TextMatrix(ii, 19) Then
                    .Cell(flexcpText, r1, 19, r2, 19) = .TextMatrix(ii, 19)
                Else
                    .TextMatrix(ii, 19) = .TextMatrix(i, 19)
                End If
                If .TextMatrix(i, 21) <> .TextMatrix(ii, 21) Then
                    .Cell(flexcpText, r1, 21, r2, 21) = "S"
                    .TextMatrix(ii, 21) = "S"
                End If
                j = rr1
                Salir2 = False
                While Salir2 = False
                    Nueva = True
                    For k = r1 To r2
                        If .TextMatrix(k, 22) = "" Then
                            .TextMatrix(k, 22) = .TextMatrix(j, 22)
                            .TextMatrix(k, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
                            .TextMatrix(k, 24) = .TextMatrix(j, 24)
                            .TextMatrix(k, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
                            Nueva = False
                        ElseIf .TextMatrix(j, 22) = "" Then
                            .TextMatrix(j, 22) = .TextMatrix(k, 22)
                            .TextMatrix(j, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
                            .TextMatrix(j, 24) = .TextMatrix(k, 24)
                            .TextMatrix(j, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
                            Nueva = False
                        ElseIf .TextMatrix(k, 22) = .TextMatrix(j, 22) And _
                           .TextMatrix(k, 24) = .TextMatrix(j, 24) Then
                            .TextMatrix(k, 23) = Val(.TextMatrix(k, 23)) + Val(.TextMatrix(j, 23))
                            .TextMatrix(k, 25) = Val(.TextMatrix(k, 25)) + Val(.TextMatrix(j, 25))
                            Nueva = False
                            Exit For
                        End If
                    Next k
                    If Nueva = True Then
                        For l = 0 To 21
                            .TextMatrix(j, l) = .TextMatrix(k - 1, l)
                        Next l
                    Else
                        .RemoveItem j
                        i = i - 1
                        j = j - 1
                        rr2 = rr2 - 1
                    End If
                    j = j + 1
                    If j > rr2 Then
                        Salir2 = True
                    End If
                Wend
            End If
            i = i + 1
            If i >= .Rows - 2 Then
                salir = True
            End If
        Wend
        .Select 1, 0
        .Sort = flexSortGenericAscending
        'Boton vsfgVentas
    End With
    Me.MousePointer = 0
    CambTam = True
End Sub

Private Sub vsfgAnulados_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If UCase(Chr(KeyCode)) = "C" And Shift = 2 Then
        VSFGAnulados.Copy
    ElseIf UCase(Chr(KeyCode)) = "X" And Shift = 2 Then
        VSFGAnulados.Cut
    ElseIf UCase(Chr(KeyCode)) = "V" And Shift = 2 Then
        VSFGAnulados.Paste
    ElseIf KeyCode = vbKeyF4 Then
        VSFGAnulados.Copy
        For i = VSFGAnulados.Row To VSFGAnulados.Rows - 1
            VSFGAnulados.Row = i
            If VSFGAnulados.Text <> "" Then
                VSFGAnulados.Paste
            Else
                Exit For
            End If
        Next i
    End If
    
End Sub
