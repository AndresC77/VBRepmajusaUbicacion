VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIncentivos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incentivos"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIncentivos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10665
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Aplicar Incentivos y Promo"
      TabPicture(0)   =   "frmIncentivos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpFechaFinAplicar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpFechaInicioAplicar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "VSFGAplicar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbNegocioAplicar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAplicarAplicar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdActualizar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optIncentivo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optPromoCombo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optPremio"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optPromoComboPedido"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optDctoCombo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optNPrendasAY"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Cargar Incentivos"
      TabPicture(1)   =   "frmIncentivos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAplicarIncentivo"
      Tab(1).Control(1)=   "txtArchivoIncentivo"
      Tab(1).Control(2)=   "cmdExplorarIncentivo"
      Tab(1).Control(3)=   "txtNombreIncentivo"
      Tab(1).Control(4)=   "VSFGIncentivo"
      Tab(1).Control(5)=   "dtpFechaInicioIncentivo"
      Tab(1).Control(6)=   "dtpFechaFinIncentivo"
      Tab(1).Control(7)=   "cmbNegocioIncentivo"
      Tab(1).Control(8)=   "cdArchivoIncentivo"
      Tab(1).Control(9)=   "Label3"
      Tab(1).Control(10)=   "Label1"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "Label2"
      Tab(1).Control(13)=   "Label11"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Cargar Promo Combo"
      TabPicture(2)   =   "frmIncentivos.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAplicarPromoCombo"
      Tab(2).Control(1)=   "txtArchivoPromoCombo"
      Tab(2).Control(2)=   "cmdExplorarPromoCombo"
      Tab(2).Control(3)=   "txtNombrePromoCombo"
      Tab(2).Control(4)=   "VSFGPromoCombo"
      Tab(2).Control(5)=   "dtpFechaInicioPromoCombo"
      Tab(2).Control(6)=   "dtpFechaFinPromoCombo"
      Tab(2).Control(7)=   "cmbNegocioPromoCombo"
      Tab(2).Control(8)=   "cdArchivoPromoCombo"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "Label10"
      Tab(2).Control(12)=   "Label9"
      Tab(2).Control(13)=   "Label8"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Cargar Premio x Monto"
      TabPicture(3)   =   "frmIncentivos.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label14"
      Tab(3).Control(1)=   "Label15"
      Tab(3).Control(2)=   "Label16"
      Tab(3).Control(3)=   "Label17"
      Tab(3).Control(4)=   "Label18"
      Tab(3).Control(5)=   "cdArchivoPremio"
      Tab(3).Control(6)=   "cmbNegocioPremio"
      Tab(3).Control(7)=   "dtpFechaFinPremio"
      Tab(3).Control(8)=   "dtpFechaInicioPremio"
      Tab(3).Control(9)=   "VSFGPremio"
      Tab(3).Control(10)=   "cmdAplicarPremio"
      Tab(3).Control(11)=   "txtArchivoPremio"
      Tab(3).Control(12)=   "cmdExplorarPremio"
      Tab(3).Control(13)=   "txtNombrePremio"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Cargar Promo Combo Pedido"
      TabPicture(4)   =   "frmIncentivos.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtCantEntPromoComboPedido1"
      Tab(4).Control(1)=   "txtCantMinPromoComboPedido1"
      Tab(4).Control(2)=   "txtArchivoPromoComboPedido2"
      Tab(4).Control(3)=   "cmdExplorarPromoComboPedido2"
      Tab(4).Control(4)=   "cmdAplicarPromoComboPedido"
      Tab(4).Control(5)=   "txtArchivoPromoComboPedido1"
      Tab(4).Control(6)=   "cmdExplorarPromoComboPedido1"
      Tab(4).Control(7)=   "txtNombrePromoComboPedido"
      Tab(4).Control(8)=   "VSFGPromoComboPedido1"
      Tab(4).Control(9)=   "dtpFechaInicioPromoComboPedido"
      Tab(4).Control(10)=   "dtpFechaFinPromoComboPedido"
      Tab(4).Control(11)=   "cmbNegocioPromoComboPedido"
      Tab(4).Control(12)=   "cdArchivoPromoComboPedido"
      Tab(4).Control(13)=   "VSFGPromoComboPedido2"
      Tab(4).Control(14)=   "Label26"
      Tab(4).Control(15)=   "Label25"
      Tab(4).Control(16)=   "Label24"
      Tab(4).Control(17)=   "Label23"
      Tab(4).Control(18)=   "Label22"
      Tab(4).Control(19)=   "Label21"
      Tab(4).Control(20)=   "Label20"
      Tab(4).Control(21)=   "Label19"
      Tab(4).ControlCount=   22
      TabCaption(5)   =   "Cargar Dcto x Combo"
      TabPicture(5)   =   "frmIncentivos.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtNombreDctoCombo"
      Tab(5).Control(1)=   "cmdExplorarDctoCombo"
      Tab(5).Control(2)=   "txtArchivoDctoCombo"
      Tab(5).Control(3)=   "cmdAplicarDctoCombo"
      Tab(5).Control(4)=   "VSFGDctoCombo"
      Tab(5).Control(5)=   "dtpFechaInicioDctoCombo"
      Tab(5).Control(6)=   "dtpFechaFinDctoCombo"
      Tab(5).Control(7)=   "cmbNegocioDctoCombo"
      Tab(5).Control(8)=   "cdArchivoDctoCombo"
      Tab(5).Control(9)=   "Label31"
      Tab(5).Control(10)=   "Label30"
      Tab(5).Control(11)=   "Label29"
      Tab(5).Control(12)=   "Label28"
      Tab(5).Control(13)=   "Label27"
      Tab(5).ControlCount=   14
      Begin VB.OptionButton optNPrendasAY 
         Caption         =   "N prendas a $Y"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   89
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton optDctoCombo 
         Caption         =   "Dcto x Combo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   88
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtNombreDctoCombo 
         Height          =   315
         Left            =   -73875
         TabIndex        =   78
         Top             =   960
         Width           =   3720
      End
      Begin VB.CommandButton cmdExplorarDctoCombo 
         Caption         =   "..."
         Height          =   315
         Left            =   -65235
         TabIndex        =   77
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtArchivoDctoCombo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   76
         Top             =   1320
         Width           =   4080
      End
      Begin VB.CommandButton cmdAplicarDctoCombo 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -71160
         TabIndex        =   75
         Top             =   6840
         Width           =   1455
      End
      Begin VB.OptionButton optPromoComboPedido 
         Caption         =   "Promo Combo Pedido"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   74
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCantEntPromoComboPedido1 
         Height          =   315
         Left            =   -68265
         TabIndex        =   72
         Text            =   "0"
         Top             =   1680
         Width           =   3360
      End
      Begin VB.TextBox txtCantMinPromoComboPedido1 
         Height          =   315
         Left            =   -73875
         TabIndex        =   70
         Text            =   "0"
         Top             =   1680
         Width           =   3360
      End
      Begin VB.TextBox txtArchivoPromoComboPedido2 
         Height          =   315
         Left            =   -68625
         TabIndex        =   68
         Top             =   2040
         Width           =   3360
      End
      Begin VB.CommandButton cmdExplorarPromoComboPedido2 
         Caption         =   "..."
         Height          =   315
         Left            =   -65265
         TabIndex        =   67
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicarPromoComboPedido 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -71280
         TabIndex        =   56
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtArchivoPromoComboPedido1 
         Height          =   315
         Left            =   -73875
         TabIndex        =   55
         Top             =   2040
         Width           =   3360
      End
      Begin VB.CommandButton cmdExplorarPromoComboPedido1 
         Caption         =   "..."
         Height          =   315
         Left            =   -70515
         TabIndex        =   54
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtNombrePromoComboPedido 
         Height          =   315
         Left            =   -73875
         TabIndex        =   53
         Top             =   953
         Width           =   3720
      End
      Begin VB.OptionButton optPremio 
         Caption         =   "Premio x Monto"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtNombrePremio 
         Height          =   315
         Left            =   -73875
         TabIndex        =   42
         Top             =   960
         Width           =   3720
      End
      Begin VB.CommandButton cmdExplorarPremio 
         Caption         =   "..."
         Height          =   315
         Left            =   -65235
         TabIndex        =   41
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtArchivoPremio 
         Height          =   315
         Left            =   -69315
         TabIndex        =   40
         Top             =   1320
         Width           =   4080
      End
      Begin VB.CommandButton cmdAplicarPremio 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -71160
         TabIndex        =   39
         Top             =   6840
         Width           =   1455
      End
      Begin VB.OptionButton optPromoCombo 
         Caption         =   "Promo Combo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optIncentivo 
         Caption         =   "Incentivo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdAplicarPromoCombo 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -71160
         TabIndex        =   27
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtArchivoPromoCombo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   26
         Top             =   1320
         Width           =   4080
      End
      Begin VB.CommandButton cmdExplorarPromoCombo 
         Caption         =   "..."
         Height          =   315
         Left            =   -65235
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtNombrePromoCombo 
         Height          =   315
         Left            =   -73875
         TabIndex        =   24
         Top             =   960
         Width           =   3720
      End
      Begin VB.CommandButton cmdAplicarIncentivo 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -71160
         TabIndex        =   13
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtArchivoIncentivo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   12
         Top             =   1320
         Width           =   4080
      End
      Begin VB.CommandButton cmdExplorarIncentivo 
         Caption         =   "..."
         Height          =   315
         Left            =   -65235
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtNombreIncentivo 
         Height          =   315
         Left            =   -73875
         TabIndex        =   10
         Top             =   960
         Width           =   3720
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "CONSULTAR"
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton cmdAplicarAplicar 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   6840
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cmbNegocioAplicar 
         Height          =   315
         Left            =   885
         TabIndex        =   3
         Top             =   960
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGAplicar 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   10095
         _cx             =   17806
         _cy             =   7858
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":03B2
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
      Begin MSComCtl2.DTPicker dtpFechaInicioAplicar 
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
         Left            =   7485
         TabIndex        =   5
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   67371011
         CurrentDate     =   37463
      End
      Begin MSComCtl2.DTPicker dtpFechaFinAplicar 
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
         Left            =   7485
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   67371011
         CurrentDate     =   37463
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGIncentivo 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   14
         Top             =   1680
         Width           =   10095
         _cx             =   17806
         _cy             =   8916
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":0528
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
      Begin NEED2.dtpFecha dtpFechaInicioIncentivo 
         Height          =   285
         Left            =   -73875
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin NEED2.dtpFecha dtpFechaFinIncentivo 
         Height          =   285
         Left            =   -71595
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin MSDataListLib.DataCombo cmbNegocioIncentivo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   17
         Top             =   960
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
      Begin MSComDlg.CommonDialog cdArchivoIncentivo 
         Left            =   -65280
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPromoCombo 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   28
         Top             =   1680
         Width           =   10095
         _cx             =   17806
         _cy             =   8916
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":05C8
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
      Begin NEED2.dtpFecha dtpFechaInicioPromoCombo 
         Height          =   285
         Left            =   -73875
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin NEED2.dtpFecha dtpFechaFinPromoCombo 
         Height          =   285
         Left            =   -71595
         TabIndex        =   30
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin MSDataListLib.DataCombo cmbNegocioPromoCombo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   31
         Top             =   960
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
      Begin MSComDlg.CommonDialog cdArchivoPromoCombo 
         Left            =   -65280
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPremio 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   43
         Top             =   1680
         Width           =   10095
         _cx             =   17806
         _cy             =   8916
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":0759
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
      Begin NEED2.dtpFecha dtpFechaInicioPremio 
         Height          =   285
         Left            =   -73875
         TabIndex        =   44
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin NEED2.dtpFecha dtpFechaFinPremio 
         Height          =   285
         Left            =   -71595
         TabIndex        =   45
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin MSDataListLib.DataCombo cmbNegocioPremio 
         Height          =   315
         Left            =   -69315
         TabIndex        =   46
         Top             =   960
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
      Begin MSComDlg.CommonDialog cdArchivoPremio 
         Left            =   -65280
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPromoComboPedido1 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   57
         Top             =   2400
         Width           =   4935
         _cx             =   8705
         _cy             =   7646
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":083F
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
      Begin NEED2.dtpFecha dtpFechaInicioPromoComboPedido 
         Height          =   285
         Left            =   -73875
         TabIndex        =   58
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin NEED2.dtpFecha dtpFechaFinPromoComboPedido 
         Height          =   285
         Left            =   -71595
         TabIndex        =   59
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin MSDataListLib.DataCombo cmbNegocioPromoComboPedido 
         Height          =   315
         Left            =   -69315
         TabIndex        =   60
         Top             =   953
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
      Begin MSComDlg.CommonDialog cdArchivoPromoComboPedido 
         Left            =   -69960
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPromoComboPedido2 
         Height          =   4335
         Left            =   -69840
         TabIndex        =   66
         Top             =   2400
         Width           =   4935
         _cx             =   8705
         _cy             =   7646
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":0873
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGDctoCombo 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   79
         Top             =   1680
         Width           =   10095
         _cx             =   17806
         _cy             =   8916
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIncentivos.frx":08A7
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
      Begin NEED2.dtpFecha dtpFechaInicioDctoCombo 
         Height          =   285
         Left            =   -73875
         TabIndex        =   80
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin NEED2.dtpFecha dtpFechaFinDctoCombo 
         Height          =   285
         Left            =   -71595
         TabIndex        =   81
         Top             =   1320
         Width           =   1455
         _extentx        =   2566
         _extenty        =   503
         value           =   41814.4080324074
      End
      Begin MSDataListLib.DataCombo cmbNegocioDctoCombo 
         Height          =   315
         Left            =   -69315
         TabIndex        =   82
         Top             =   960
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
      Begin MSComDlg.CommonDialog cdArchivoDctoCombo 
         Left            =   -65280
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69975
         TabIndex        =   87
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Promo Combo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74940
         TabIndex        =   86
         Top             =   1005
         Width           =   990
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69960
         TabIndex        =   85
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -72405
         TabIndex        =   84
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74805
         TabIndex        =   83
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Entregar"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69375
         TabIndex        =   73
         Top             =   1732
         Width           =   1035
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Min."
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74685
         TabIndex        =   71
         Top             =   1732
         Width           =   705
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo2"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69360
         TabIndex        =   69
         Top             =   2092
         Width           =   660
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74805
         TabIndex        =   65
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -72405
         TabIndex        =   64
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo1"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74610
         TabIndex        =   63
         Top             =   2092
         Width           =   660
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Promo Combo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74940
         TabIndex        =   62
         Top             =   1005
         Width           =   990
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69975
         TabIndex        =   61
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69975
         TabIndex        =   51
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Premio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74430
         TabIndex        =   50
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69960
         TabIndex        =   49
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -72405
         TabIndex        =   48
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74805
         TabIndex        =   47
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74805
         TabIndex        =   36
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -72405
         TabIndex        =   35
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69960
         TabIndex        =   34
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Promo Combo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74940
         TabIndex        =   33
         Top             =   1005
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69975
         TabIndex        =   32
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74805
         TabIndex        =   22
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -72405
         TabIndex        =   21
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69960
         TabIndex        =   20
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Incentivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74595
         TabIndex        =   19
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -69975
         TabIndex        =   18
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6765
         TabIndex        =   8
         Top             =   1380
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6525
         TabIndex        =   7
         Top             =   1020
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmIncentivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
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

Private clsCon_Def As New clsConsulta
Private strSql As String
Dim i_flag As Integer

Private Sub cmdActualizar_Click()
    Dim i As Long
    If optPremio.Value = True Then
        VSFGAplicar.Cols = 11
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 6) = "Cod.Prod."
        VSFGAplicar.TextMatrix(0, 7) = "Producto"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad"
        VSFGAplicar.TextMatrix(0, 9) = "Precio"
        VSFGAplicar.TextMatrix(0, 10) = "Dcto"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pre_loc_nombre,producto.prd_codigo,prd_nombre," & _
                 " pre_loc_cantidad,pre_loc_precio,pre_loc_dcto " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN premio_local ON persona.emp_codigo=premio_local.emp_codigo " & _
                 " AND persona.per_codigo LIKE premio_local.per_codigo " & _
                 " AND persona.tip_ped_codigo=premio_local.tip_ped_codigo " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(premio_local.pre_loc_fecha_desde,10) AND LEFT(premio_local.pre_loc_fecha_hasta,10) " & _
                 " AND premio_local.pre_loc_estado=0 " & _
                 " AND round(pedido.ped_subtotal/1.12,2) BETWEEN premio_local.pre_loc_rango_inferior AND premio_local.pre_loc_rango_superior " & _
                 " INNER JOIN producto ON premio_local.emp_codigo=producto.emp_codigo " & _
                 " AND premio_local.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc,producto.prd_codigo "
    ElseIf optIncentivo.Value = True Then
        VSFGAplicar.Cols = 11
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 6) = "Cod.Prod."
        VSFGAplicar.TextMatrix(0, 7) = "Producto"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad"
        VSFGAplicar.TextMatrix(0, 9) = "Precio"
        VSFGAplicar.TextMatrix(0, 10) = "Dcto"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,inc_loc_nombre,producto.prd_codigo,prd_nombre," & _
                 " inc_loc_cantidad,inc_loc_precio,inc_loc_dcto " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN incentivo_local ON persona.emp_codigo=incentivo_local.emp_codigo " & _
                 " AND persona.per_codigo=incentivo_local.per_codigo " & _
                 " AND persona.tip_ped_codigo=incentivo_local.tip_ped_codigo " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(incentivo_local.inc_loc_fecha_desde,10) AND LEFT(incentivo_local.inc_loc_fecha_hasta,10) " & _
                 " AND incentivo_local.inc_loc_estado=0 " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc,producto.prd_codigo "
    ElseIf optPromoComboPedido.Value = True Then
        VSFGAplicar.Cols = 10
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Cantidad Producto Ingresado"
        VSFGAplicar.TextMatrix(0, 6) = "Cantidad Combo Solicitado"
        VSFGAplicar.TextMatrix(0, 7) = "Cantidad Producto Mínimo"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad Combo a Entregar"
        VSFGAplicar.TextMatrix(0, 9) = "Cantidad Combo a Eliminar"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,COALESCE(productoing,0) as productoing,combosol," & _
                 " pro_com_ped_cantidad_min,pro_com_ped_cantidad_ent,0 as cat_eli " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' "
        strSql = strSql & " INNER JOIN ( SELECT p.emp_codigo,p.ped_codigo,SUM(dp.det_ped_cant_entregada) as combosol, " & _
                 " pro_com_ped_cantidad_min,pro_com_ped_cantidad_ent " & _
                 " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                 " AND p.ped_codigo=dp.ped_codigo " & _
                 " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                 " AND p.ped_fecha BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                 " INNER JOIN det_promo_combo_pedido_ent dpcpe ON pcp.emp_codigo=dpcpe.emp_codigo" & _
                 " AND pcp.tip_ped_codigo=dpcpe.tip_ped_codigo" & _
                 " AND pcp.pro_com_ped_codigo=dpcpe.pro_com_ped_codigo" & _
                 " AND dp.prd_codigo=dpcpe.prd_codigo " & _
                 " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                 " AND p.ped_estado=0 " & _
                 " AND p.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " GROUP BY p.emp_codigo,p.ped_codigo" & _
                 ") pe ON pedido.emp_codigo=pe.emp_codigo " & _
                 " AND pedido.ped_codigo=pe.ped_codigo "
        strSql = strSql & " LEFT JOIN ( SELECT p.emp_codigo,p.ped_codigo, " & _
                 " SUM(dp.det_ped_cant_entregada) as productoing " & _
                 " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                 " AND p.ped_codigo=dp.ped_codigo " & _
                 " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                 " AND p.ped_fecha BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                 " INNER JOIN det_promo_combo_pedido dpcp ON pcp.emp_codigo=dpcp.emp_codigo" & _
                 " AND pcp.tip_ped_codigo=dpcp.tip_ped_codigo" & _
                 " AND pcp.pro_com_ped_codigo=dpcp.pro_com_ped_codigo" & _
                 " AND dp.prd_codigo=dpcp.prd_codigo " & _
                 " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                 " AND p.ped_estado=0 " & _
                 " AND p.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " GROUP BY p.emp_codigo,p.ped_codigo" & _
                 ") pi ON pedido.emp_codigo=pi.emp_codigo " & _
                 " AND pedido.ped_codigo=pi.ped_codigo "
        strSql = strSql & " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc "
    ElseIf optDctoCombo.Value = True Then
        VSFGAplicar.Cols = 12
        VSFGAplicar.Rows = 1
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 6) = "Cod.Prod."
        VSFGAplicar.TextMatrix(0, 7) = "Producto"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad"
        VSFGAplicar.TextMatrix(0, 9) = "Precio"
        VSFGAplicar.TextMatrix(0, 10) = "Dcto"
        VSFGAplicar.TextMatrix(0, 11) = "%Dcto"
        '1 producto
'        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
'                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
'                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dct_dcto/100,pro_com_dct_dcto " & _
'                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
'                 " AND pedido.per_codigo=persona.per_codigo " & _
'                 " AND persona.cat_p_tipo='C' " & _
'                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
'                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
'                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
'                 " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
'                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
'                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
'                 " AND det_pedido.det_ped_cant_pedida>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
'                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
'                 " AND promo_combo_dcto.prd_codigo_2='' AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
'                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
'                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
'                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
'                 " AND pedido.ped_estado=0 " & _
'                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
'        strSql = strSql & " UNION "
        '2 producto 1
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dct_dcto/100,pro_com_dct_dcto " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        '2 producto 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dct_dcto/100,pro_com_dct_dcto " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " ORDER BY ped_codigo,per_codigo "
    ElseIf optPromoCombo.Value = True Then
        VSFGAplicar.Cols = 11
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 6) = "Cod.Prod."
        VSFGAplicar.TextMatrix(0, 7) = "Producto"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad"
        VSFGAplicar.TextMatrix(0, 9) = "Precio"
        VSFGAplicar.TextMatrix(0, 10) = "Dcto"
        'a uno
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_2='' AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b uno
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_2='' AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a dos 1
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a dos 2
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b dos
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a tres 1
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a tres 2
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a tres 3
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_3=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b tres
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 1
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo " & _
                 " INNER JOIN det_pedido dpc ON pedido.emp_codigo=dpc.emp_codigo " & _
                 " AND pedido.ped_codigo=dpc.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND dpc.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 2
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo " & _
                 " INNER JOIN det_pedido dpc ON pedido.emp_codigo=dpc.emp_codigo " & _
                 " AND pedido.ped_codigo=dpc.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND dpc.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 3
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo " & _
                 " INNER JOIN det_pedido dpc ON pedido.emp_codigo=dpc.emp_codigo " & _
                 " AND pedido.ped_codigo=dpc.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND dpc.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_3=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 4
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_pedida,det_pedido.det_ped_precio,det_pedido.det_ped_cant_pedida*det_pedido.det_ped_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo " & _
                 " INNER JOIN det_pedido dpc ON pedido.emp_codigo=dpc.emp_codigo " & _
                 " AND pedido.ped_codigo=dpc.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_4 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpc.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_4=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b cuatro
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100 " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo " & _
                 " INNER JOIN det_pedido dpb ON pedido.emp_codigo=dpb.emp_codigo " & _
                 " AND pedido.ped_codigo=dpb.ped_codigo " & _
                 " INNER JOIN det_pedido dpc ON pedido.emp_codigo=dpc.emp_codigo " & _
                 " AND pedido.ped_codigo=dpc.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE promo_combo.per_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND dpb.det_ped_cant_pedida>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=0 " & _
                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " ORDER BY per_codigo "
    ElseIf optNPrendasAY.Value = True Then
        strSql = " SELECT ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,CONCAT(per_apellido, ' ',per_nombre) as cli " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
                 " AND ped_estado=0 " & _
                 " AND ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
    End If
    clsCon_Def.Ejecutar strSql
    Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    
    If optPromoComboPedido.Value = True Then
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Cantidad Producto Ingresado"
        VSFGAplicar.TextMatrix(0, 6) = "Cantidad Combo Solicitado"
        VSFGAplicar.TextMatrix(0, 7) = "Cantidad Producto Mínimo"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad Combo a Entregar"
        VSFGAplicar.TextMatrix(0, 9) = "Cantidad Combo a Eliminar"
        For i = 1 To VSFGAplicar.Rows - 1
            If Int(VSFGAplicar.TextMatrix(i, 5) / VSFGAplicar.TextMatrix(i, 7)) * VSFGAplicar.TextMatrix(i, 8) < VSFGAplicar.TextMatrix(i, 6) Then
                VSFGAplicar.TextMatrix(i, 9) = VSFGAplicar.TextMatrix(i, 6) - (Int(VSFGAplicar.TextMatrix(i, 5) / VSFGAplicar.TextMatrix(i, 7)) * VSFGAplicar.TextMatrix(i, 8))
            Else
                VSFGAplicar.TextMatrix(i, 9) = 0
            End If
        Next i
    ElseIf optNPrendasAY.Value = True Then
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
    End If
End Sub

Private Sub cmdAplicarDctoCombo_Click()
    Dim i As Long
    Dim j As Long
    
    VSFGDctoCombo.Select 1, VSFGDctoCombo.Cols - 1
    VSFGDctoCombo.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        For i = 1 To VSFGDctoCombo.Rows - 1
            If Val(VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1)) = 4 Then
                strSql = " INSERT INTO promo_combo_dcto (emp_codigo, tip_ped_codigo, pro_com_dct_fecha_desde, pro_com_dct_fecha_hasta," & _
                         " prd_codigo_1, pro_com_dct_cantidad_1, " & _
                         " prd_codigo_2, pro_com_dct_cantidad_2, " & _
                         " prd_codigo_3, pro_com_dct_cantidad_3, " & _
                         " prd_codigo_4, pro_com_dct_cantidad_4, " & _
                         " pro_com_dct_nombre,pro_com_dct_dcto, pro_com_dct_fechamod, pro_com_dct_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioDctoCombo.BoundText & "','" & Format(dtpFechaInicioDctoCombo.Value, "yyyy-mm-dd") & "','" & Format(dtpFechaFinDctoCombo.Value, "yyyy-mm-dd") & "'," & _
                         " '" & VSFGDctoCombo.TextMatrix(i, 0) & "','" & VSFGDctoCombo.TextMatrix(i, 2) & "'," & _
                         " '" & VSFGDctoCombo.TextMatrix(i, 3) & "','" & VSFGDctoCombo.TextMatrix(i, 5) & "'," & _
                         " '" & VSFGDctoCombo.TextMatrix(i, 6) & "','" & VSFGDctoCombo.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGDctoCombo.TextMatrix(i, 9) & "','" & VSFGDctoCombo.TextMatrix(i, 11) & "'," & _
                         " '" & UCase(txtNombreDctoCombo.Text) & "','" & VSFGDctoCombo.TextMatrix(i, 12) & "'," & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "PromoCombos"
    
    Unload Me
End Sub

Private Sub cmdAplicarIncentivo_Click()
    Dim i As Long
    Dim j As Long
    
    VSFGIncentivo.Select 1, VSFGIncentivo.Cols - 1
    VSFGIncentivo.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        For i = 1 To VSFGIncentivo.Rows - 1
            If Val(VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1)) = 2 Then
                If VSFGIncentivo.TextMatrix(i, 1) = "C01477" Then
                VSFGIncentivo.ShowCell i, 1
                End If
                strSql = " INSERT INTO incentivo_local (emp_codigo, tip_ped_codigo, per_codigo, inc_loc_fecha_desde, inc_loc_fecha_hasta, " & _
                         " prd_codigo, inc_loc_cedula, inc_loc_nombre, inc_loc_cantidad, inc_loc_precio, inc_loc_dcto, inc_loc_estado," & _
                         " ped_codigo, inc_loc_fechamod, inc_loc_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioIncentivo.BoundText & "','" & VSFGIncentivo.TextMatrix(i, 1) & "','" & Format(dtpFechaInicioIncentivo.Value, "yyyy-mm-dd") & "','" & Format(dtpFechaFinIncentivo.Value, "yyyy-mm-dd") & "'," & _
                         " '" & VSFGIncentivo.TextMatrix(i, 3) & "','" & VSFGIncentivo.TextMatrix(i, 0) & "','" & txtNombreIncentivo.Text & "','" & VSFGIncentivo.TextMatrix(i, 5) & "','" & VSFGIncentivo.TextMatrix(i, 6) & "','" & VSFGIncentivo.TextMatrix(i, 7) & "',0, " & _
                         " null,CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Incentivos"
    
    Unload Me

End Sub

Private Sub cmdAplicarAplicar_Click()
    Dim i As Long
    Dim j As Long
    Dim dblDcto As Double
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If optPremio.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            For i = 1 To VSFGAplicar.Rows - 1
                strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                         " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                         " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto) " & _
                         " VALUES ('" & strEmpresa & "','PRI',0," & _
                         " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & VSFGAplicar.TextMatrix(i, 10) & "') "
                clsCon_Def.Ejecutar strSql
            Next i
        End If
    ElseIf optIncentivo.Value = True Then
        For i = 1 To VSFGAplicar.Rows - 1
            strSql = " SELECT COALESCE(count(*),0) as n " & _
                     " FROM incentivo_local " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND per_codigo='" & VSFGAplicar.TextMatrix(i, 2) & "' " & _
                     " AND LEFT('" & VSFGAplicar.TextMatrix(i, 1) & "',10) BETWEEN LEFT(inc_loc_fecha_desde,10) AND LEFT(inc_loc_fecha_hasta,10)" & _
                     " AND inc_loc_estado=0 " & _
                     " AND ped_codigo is null " & _
                     " AND prd_codigo='" & VSFGAplicar.TextMatrix(i, 6) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                If clsCon_Def.adorec_Def("n") > 0 Then
                    strSql = " UPDATE incentivo_local " & _
                             " SET inc_loc_estado=1, " & _
                             " ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND per_codigo='" & VSFGAplicar.TextMatrix(i, 2) & "' " & _
                             " AND LEFT('" & VSFGAplicar.TextMatrix(i, 1) & "',10) BETWEEN LEFT(inc_loc_fecha_desde,10) AND LEFT(inc_loc_fecha_hasta,10)" & _
                             " AND inc_loc_estado=0 " & _
                             " AND ped_codigo is null " & _
                             " AND prd_codigo='" & VSFGAplicar.TextMatrix(i, 6) & "' "
                    clsCon_Def.Ejecutar strSql
                    strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                             " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                             " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                             " det_ped_cant_entregada, det_ped_precio,det_ped_dcto) " & _
                             " VALUES ('" & strEmpresa & "','PRI',0," & _
                             " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                             " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                             " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & VSFGAplicar.TextMatrix(i, 10) & "') "
                    clsCon_Def.Ejecutar strSql
                End If
            End If
        Next i
    ElseIf optDctoCombo.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            For i = 1 To VSFGAplicar.Rows - 1
                strSql = " REPLACE INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                         " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                         " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto) " & _
                         " VALUES ('" & strEmpresa & "','PRI',0," & _
                         " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & VSFGAplicar.TextMatrix(i, 10) & "') "
                clsCon_Def.Ejecutar strSql
            Next i
        End If
    
    ElseIf optPromoComboPedido.Value = True Then
        For i = 1 To VSFGAplicar.Rows - 1
            If VSFGAplicar.TextMatrix(i, 9) > 0 Then
                For j = 1 To VSFGAplicar.TextMatrix(i, 9)
                    strSql = " SELECT p.emp_codigo,p.ped_codigo,dp.prd_codigo,dp.det_ped_cant_entregada " & _
                             " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                             " AND p.ped_codigo=dp.ped_codigo " & _
                             " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                             " AND p.ped_fecha BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                             " INNER JOIN det_promo_combo_pedido_ent dpcpe ON pcp.emp_codigo=dpcpe.emp_codigo" & _
                             " AND pcp.tip_ped_codigo=dpcpe.tip_ped_codigo" & _
                             " AND pcp.pro_com_ped_codigo=dpcpe.pro_com_ped_codigo" & _
                             " AND dp.prd_codigo=dpcpe.prd_codigo " & _
                             " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                             " AND p.ped_estado=0 and dp.det_ped_cant_entregada>0 " & _
                             " AND p.ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'"
                    clsAux.Ejecutar strSql
                    If clsAux.adorec_Def("det_ped_cant_entregada") = 1 Then
                        strSql = " DELETE FROM det_pedido " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'" & _
                                 " AND prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
                        clsCon_Def.Ejecutar strSql, "M"
                    ElseIf clsAux.adorec_Def("det_ped_cant_entregada") > 1 Then
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_cant_entregada=det_ped_cant_entregada-1, " & _
                                 " det_ped_cant_pedida=det_ped_cant_pedida-1 " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'" & _
                                 " AND prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
                        clsCon_Def.Ejecutar strSql, "M"
                    End If
                Next j
            End If
        Next i
    ElseIf optNPrendasAY.Value = True Then
        For i = 1 To VSFGAplicar.Rows - 1
            'PROMO PRENDA PRECIO'
            dblDcto = FormatoD2(PromoPrendaPrecio(cmbNegocio.BoundText, VSFGAplicar.TextMatrix(i, 0)))
        Next i
    End If
    MsgBox "Carga Finalizada"
    Unload Me
End Sub

Private Sub cmdAplicarPremio_Click()
    Dim i As Long
    Dim j As Long
    
    VSFGPremio.Select 1, VSFGPremio.Cols - 1
    VSFGPremio.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        For i = 1 To VSFGPremio.Rows - 1
            If Val(VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1)) = 2 Then
                If VSFGPremio.TextMatrix(i, 1) = "C01477" Then
                VSFGPremio.ShowCell i, 1
                End If
                strSql = " INSERT INTO premio_local (emp_codigo, tip_ped_codigo, per_codigo, pre_loc_fecha_desde, pre_loc_fecha_hasta," & _
                         " pre_loc_rango_inferior, pre_loc_rango_superior, prd_codigo, pre_loc_cedula, pre_loc_nombre," & _
                         " pre_loc_cantidad, pre_loc_precio, pre_loc_dcto, pre_loc_estado," & _
                         " pre_loc_fechamod, pre_loc_usumod)" & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioPremio.BoundText & "','" & VSFGPremio.TextMatrix(i, 1) & "','" & Format(dtpFechaInicioPremio.Value, "yyyy-mm-dd") & "','" & Format(dtpFechaFinPremio.Value, "yyyy-mm-dd") & "'," & _
                         " '" & VSFGPremio.TextMatrix(i, 3) & "','" & VSFGPremio.TextMatrix(i, 4) & "','" & VSFGPremio.TextMatrix(i, 5) & "','" & VSFGPremio.TextMatrix(i, 0) & "','" & UCase(txtNombrePremio.Text) & "'," & _
                         " '" & VSFGPremio.TextMatrix(i, 7) & "','" & VSFGPremio.TextMatrix(i, 8) & "','" & VSFGPremio.TextMatrix(i, 9) & "',0, " & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Incentivos"
    
    Unload Me
End Sub

Private Sub cmdAplicarPromoCombo_Click()
    Dim i As Long
    Dim j As Long
    
    VSFGPromoCombo.Select 1, VSFGPromoCombo.Cols - 1
    VSFGPromoCombo.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        For i = 1 To VSFGPromoCombo.Rows - 1
            If Val(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) = 6 Then
                If VSFGPromoCombo.TextMatrix(i, 1) = "C01477" Then
                VSFGPromoCombo.ShowCell i, 1
                End If
                strSql = " INSERT INTO promo_combo (emp_codigo, tip_ped_codigo, per_codigo, pro_com_fecha_desde, pro_com_fecha_hasta," & _
                         " prd_codigo_1, pro_com_cantidad_1, prd_codigo_2, pro_com_cantidad_2," & _
                         " prd_codigo_3, pro_com_cantidad_3, prd_codigo_4, pro_com_cantidad_4,pro_com_cedula, " & _
                         " pro_com_nombre, prd_codigo_t, pro_com_cantidad, pro_com_precio, pro_com_dcto, " & _
                         " pro_com_fechamod, pro_com_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioPromoCombo.BoundText & "','" & VSFGPromoCombo.TextMatrix(i, 1) & "','" & Format(dtpFechaInicioPromoCombo.Value, "yyyy-mm-dd") & "','" & Format(dtpFechaFinPromoCombo.Value, "yyyy-mm-dd") & "'," & _
                         " '" & VSFGPromoCombo.TextMatrix(i, 3) & "','" & VSFGPromoCombo.TextMatrix(i, 5) & "','" & VSFGPromoCombo.TextMatrix(i, 6) & "','" & VSFGPromoCombo.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGPromoCombo.TextMatrix(i, 9) & "','" & VSFGPromoCombo.TextMatrix(i, 11) & "','" & VSFGPromoCombo.TextMatrix(i, 12) & "','" & VSFGPromoCombo.TextMatrix(i, 14) & "','" & VSFGPromoCombo.TextMatrix(i, 0) & "'," & _
                         " '" & UCase(txtNombrePromoCombo.Text) & "','" & VSFGPromoCombo.TextMatrix(i, 15) & "','" & VSFGPromoCombo.TextMatrix(i, 17) & "','" & VSFGPromoCombo.TextMatrix(i, 18) & "','" & VSFGPromoCombo.TextMatrix(i, 19) & "'," & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "PromoCombos"
    
    Unload Me
End Sub

Private Sub cmdAplicarPromoComboPedido_Click()
    Dim i As Long
    Dim j As Long
    num = 3
    VSFGPromoComboPedido1.Select 1, VSFGPromoComboPedido1.Cols - 1
    VSFGPromoComboPedido1.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        strSql = " INSERT INTO promo_combo_pedido (emp_codigo, tip_ped_codigo, pro_com_ped_codigo," & _
                 " pro_com_ped_fecha_desde, pro_com_ped_fecha_hasta, pro_com_ped_cantidad_min," & _
                 " pro_com_ped_cantidad_ent,pro_com_ped_nombre, pro_com_ped_fechamod, pro_com_ped_usumod)" & _
                 " VALUES('" & strEmpresa & "','" & cmbNegocioPromoComboPedido.BoundText & "','" & num & "'," & _
                 " '" & dtpFechaInicioPromoComboPedido.Value & "','" & dtpFechaFinPromoComboPedido.Value & "','" & txtCantMinPromoComboPedido1.Text & "'," & _
                 " '" & txtCantEntPromoComboPedido1.Text & "','" & UCase(txtNombrePromoComboPedido.Text) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
        For i = 1 To VSFGPromoComboPedido1.Rows - 1
            If Val(VSFGPromoComboPedido1.TextMatrix(i, VSFGPromoComboPedido1.Cols - 1)) = 1 Then
                VSFGPromoComboPedido1.ShowCell i, 1
                strSql = " INSERT INTO det_promo_combo_pedido (emp_codigo, tip_ped_codigo, pro_com_ped_codigo, prd_codigo, " & _
                         " det_pro_com_ped_fechamod, det_pro_com_ped_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioPromoComboPedido.BoundText & "','" & num & "','" & VSFGPromoComboPedido1.TextMatrix(i, 0) & "'," & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
    VSFGPromoComboPedido2.Select 1, VSFGPromoComboPedido2.Cols - 1
    VSFGPromoComboPedido2.Sort = flexSortGenericDescending
        For i = 1 To VSFGPromoComboPedido2.Rows - 1
            If Val(VSFGPromoComboPedido2.TextMatrix(i, VSFGPromoComboPedido2.Cols - 1)) = 1 Then
                VSFGPromoComboPedido2.ShowCell i, 1
                strSql = " INSERT INTO det_promo_combo_pedido_ent (emp_codigo, tip_ped_codigo, pro_com_ped_codigo, prd_codigo," & _
                         " det_pro_com_ped_ent_fechamod,det_pro_com_ped_ent_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & cmbNegocioPromoComboPedido.BoundText & "','" & num & "','" & VSFGPromoComboPedido2.TextMatrix(i, 0) & "'," & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "PromoCombosPedidos"
    
    Unload Me
End Sub

Private Sub cmdExplorarDctoCombo_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoDctoCombo.Tag = sDir
    cdArchivoDctoCombo.ShowOpen
    txtArchivoDctoCombo = cdArchivoDctoCombo.FileName
    ChDir sDir
    VSFGDctoCombo.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoDctoCombo.Text <> "") Then
        VSFGDctoCombo.Rows = 1
        VSFGDctoCombo.LoadGrid txtArchivoDctoCombo.Text, flexFileCommaText
        VSFGDctoCombo.Cols = VSFGDctoCombo.Cols + 6
        VSFGDctoCombo.ColPosition(8) = VSFGDctoCombo.Cols - 3
        VSFGDctoCombo.ColPosition(7) = VSFGDctoCombo.Cols - 4
        VSFGDctoCombo.ColPosition(6) = VSFGDctoCombo.Cols - 6
        VSFGDctoCombo.ColPosition(5) = VSFGDctoCombo.Cols - 7
        VSFGDctoCombo.ColPosition(4) = VSFGDctoCombo.Cols - 9
        VSFGDctoCombo.ColPosition(3) = VSFGDctoCombo.Cols - 10
        VSFGDctoCombo.ColPosition(2) = VSFGDctoCombo.Cols - 12
        VSFGDctoCombo.ColPosition(1) = VSFGDctoCombo.Cols - 13
        VSFGDctoCombo.TextMatrix(0, 1) = "Producto1"
        VSFGDctoCombo.TextMatrix(0, 4) = "Producto2"
        VSFGDctoCombo.TextMatrix(0, 7) = "Producto3"
        VSFGDctoCombo.TextMatrix(0, 10) = "Producto4"
        VSFGDctoCombo.TextMatrix(0, 13) = "Observacion"
        VSFGDctoCombo.TextMatrix(0, 14) = "Aplica"
        For i = 1 To VSFGDctoCombo.Rows - 1
            VSFGDctoCombo.ShowCell i, 0
            If VSFGDctoCombo.TextMatrix(i, 0) <> VSFGDctoCombo.TextMatrix(i - 1, 0) Then
                strSql = " SELECT prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND prd_codigo='" & VSFGDctoCombo.TextMatrix(i, 0) & "' "
                
                clsCon_Def.Ejecutar strSql
            End If
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGDctoCombo.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = FormatoD0(VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1)) + 1
            Else
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = 0
                VSFGDctoCombo.Cell(flexcpBackColor, i, 0, i, VSFGDctoCombo.Cols - 1) = vbRed
            End If
            
            If VSFGDctoCombo.TextMatrix(i, 3) = "" Then
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
            Else
                If VSFGDctoCombo.TextMatrix(i, 3) <> VSFGDctoCombo.TextMatrix(i - 1, 3) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGDctoCombo.TextMatrix(i, 3) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGDctoCombo.TextMatrix(i, 4) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
                Else
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = 0
                    VSFGDctoCombo.Cell(flexcpBackColor, i, 0, i, VSFGDctoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGDctoCombo.TextMatrix(i, 6) = "" Then
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
            Else
                If VSFGDctoCombo.TextMatrix(i, 6) <> VSFGDctoCombo.TextMatrix(i - 1, 6) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGDctoCombo.TextMatrix(i, 6) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGDctoCombo.TextMatrix(i, 7) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
                Else
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = 0
                    VSFGDctoCombo.Cell(flexcpBackColor, i, 0, i, VSFGDctoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGDctoCombo.TextMatrix(i, 9) = "" Then
                VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
            Else
                If VSFGDctoCombo.TextMatrix(i, 9) <> VSFGDctoCombo.TextMatrix(i - 1, 9) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGDctoCombo.TextMatrix(i, 9) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGDctoCombo.TextMatrix(i, 10) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) + 1
                Else
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGDctoCombo.TextMatrix(i, VSFGDctoCombo.Cols - 1) = 0
                    VSFGDctoCombo.Cell(flexcpBackColor, i, 0, i, VSFGDctoCombo.Cols - 1) = vbRed
                End If
            End If
            
            
        Next i
    
        VSFGDctoCombo.Select 1, VSFGDctoCombo.Cols - 1
        VSFGDctoCombo.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0

End Sub

Private Sub cmdExplorarIncentivo_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoIncentivo.Tag = sDir
    cdArchivoIncentivo.ShowOpen
    txtArchivoIncentivo = cdArchivoIncentivo.FileName
    ChDir sDir
    VSFGIncentivo.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoIncentivo.Text <> "") Then
        VSFGIncentivo.Rows = 1
        VSFGIncentivo.LoadGrid txtArchivoIncentivo.Text, flexFileCommaText
        VSFGIncentivo.Cols = VSFGIncentivo.Cols + 5
        VSFGIncentivo.ColPosition(4) = VSFGIncentivo.Cols - 3
        VSFGIncentivo.ColPosition(3) = VSFGIncentivo.Cols - 4
        VSFGIncentivo.ColPosition(2) = VSFGIncentivo.Cols - 5
        VSFGIncentivo.ColPosition(1) = 3
        VSFGIncentivo.TextMatrix(0, 1) = "Cod.Cliente"
        VSFGIncentivo.TextMatrix(0, 2) = "Cliente"
        VSFGIncentivo.TextMatrix(0, 4) = "Producto"
        VSFGIncentivo.TextMatrix(0, 8) = "Observacion"
        VSFGIncentivo.TextMatrix(0, 9) = "Aplica"
        For i = 1 To VSFGIncentivo.Rows - 1
            VSFGIncentivo.ShowCell i, 0
            If VSFGIncentivo.TextMatrix(i, 3) <> VSFGIncentivo.TextMatrix(i - 1, 3) Then
                strSql = " SELECT prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND prd_codigo='" & VSFGIncentivo.TextMatrix(i, 3) & "' "
                
                clsCon_Def.Ejecutar strSql
            End If
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGIncentivo.TextMatrix(i, 4) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1) = 1
            Else
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1) = 0
                VSFGIncentivo.Cell(flexcpBackColor, i, 0, i, VSFGIncentivo.Cols - 1) = vbRed
            End If
            If VSFGIncentivo.TextMatrix(i, 0) <> VSFGIncentivo.TextMatrix(i - 1, 0) Then
                strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND per_ruc like '%" & VSFGIncentivo.TextMatrix(i, 0) & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND tip_ped_codigo='" & cmbNegocioIncentivo.BoundText & "' "
                         
                clsCon_DefP.Ejecutar strSql
            End If
            If clsCon_DefP.adorec_Def.RecordCount > 0 Then
                VSFGIncentivo.TextMatrix(i, 1) = clsCon_DefP.adorec_Def("per_codigo")
                VSFGIncentivo.TextMatrix(i, 2) = clsCon_DefP.adorec_Def("cli")
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1) = VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1) + 1
            Else
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 2) = VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                VSFGIncentivo.TextMatrix(i, VSFGIncentivo.Cols - 1) = 0
                VSFGIncentivo.Cell(flexcpBackColor, i, 0, i, VSFGIncentivo.Cols - 1) = vbRed
            End If
            
        Next i
    
        VSFGIncentivo.Select 1, VSFGIncentivo.Cols - 1
        VSFGIncentivo.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0
End Sub

Private Sub cmdExplorarPremio_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoPremio.Tag = sDir
    cdArchivoPremio.ShowOpen
    txtArchivoPremio = cdArchivoPremio.FileName
    ChDir sDir
    VSFGPremio.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoPremio.Text <> "") Then
        VSFGPremio.Rows = 1
        VSFGPremio.LoadGrid txtArchivoPremio.Text, flexFileTabText
        VSFGPremio.Cols = VSFGPremio.Cols + 5
        VSFGPremio.ColPosition(6) = VSFGPremio.Cols - 3
        VSFGPremio.ColPosition(5) = VSFGPremio.Cols - 4
        VSFGPremio.ColPosition(4) = VSFGPremio.Cols - 5
        VSFGPremio.ColPosition(3) = VSFGPremio.Cols - 7
        VSFGPremio.ColPosition(2) = VSFGPremio.Cols - 8
        VSFGPremio.ColPosition(1) = VSFGPremio.Cols - 9
        VSFGPremio.TextMatrix(0, 1) = "Cod.Cliente"
        VSFGPremio.TextMatrix(0, 2) = "Cliente"
        VSFGPremio.TextMatrix(0, 6) = "Producto"
        VSFGPremio.TextMatrix(0, 10) = "Observacion"
        VSFGPremio.TextMatrix(0, 11) = "Aplica"
        For i = 1 To VSFGPremio.Rows - 1
            VSFGPremio.ShowCell i, 0
            If VSFGPremio.TextMatrix(i, 5) <> VSFGPremio.TextMatrix(i - 1, 5) Then
                strSql = " SELECT prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND prd_codigo='" & VSFGPremio.TextMatrix(i, 5) & "' "
                
                clsCon_Def.Ejecutar strSql
            End If
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGPremio.TextMatrix(i, 6) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) = 1
            Else
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) = 0
                VSFGPremio.Cell(flexcpBackColor, i, 0, i, VSFGPremio.Cols - 1) = vbRed
            End If
            If VSFGPremio.TextMatrix(i, 0) <> VSFGPremio.TextMatrix(i - 1, 0) Then
                If VSFGPremio.TextMatrix(i, 0) <> "%" And VSFGPremio.TextMatrix(i, 0) <> "" Then
                strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND per_ruc like '%" & VSFGPremio.TextMatrix(i, 0) & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND tip_ped_codigo='" & cmbNegocioPremio.BoundText & "' "
                         
                clsCon_DefP.Ejecutar strSql
                End If
            End If
            If VSFGPremio.TextMatrix(i, 0) = "%" Then
                VSFGPremio.TextMatrix(i, 1) = "%"
                VSFGPremio.TextMatrix(i, 2) = "Todos los clientes"
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) = VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) + 1
            ElseIf VSFGPremio.TextMatrix(i, 0) <> "" Then
                If clsCon_DefP.adorec_Def.RecordCount > 0 Then
                    VSFGPremio.TextMatrix(i, 1) = clsCon_DefP.adorec_Def("per_codigo")
                    VSFGPremio.TextMatrix(i, 2) = clsCon_DefP.adorec_Def("cli")
                    VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) = VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) + 1
                End If
            Else
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 2) = VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                VSFGPremio.TextMatrix(i, VSFGPremio.Cols - 1) = 0
                VSFGPremio.Cell(flexcpBackColor, i, 0, i, VSFGPremio.Cols - 1) = vbRed
            End If
            
        Next i
    
        VSFGPremio.Select 1, VSFGPremio.Cols - 1
        VSFGPremio.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0

End Sub

Private Sub cmdExplorarPromoCombo_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoPromoCombo.Tag = sDir
    cdArchivoPromoCombo.ShowOpen
    txtArchivoPromoCombo = cdArchivoPromoCombo.FileName
    ChDir sDir
    VSFGPromoCombo.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoPromoCombo.Text <> "") Then
        VSFGPromoCombo.Rows = 1
        VSFGPromoCombo.LoadGrid txtArchivoPromoCombo.Text, flexFileCommaText
        VSFGPromoCombo.Cols = VSFGPromoCombo.Cols + 9
        VSFGPromoCombo.ColPosition(12) = VSFGPromoCombo.Cols - 3
        VSFGPromoCombo.ColPosition(11) = VSFGPromoCombo.Cols - 4
        VSFGPromoCombo.ColPosition(10) = VSFGPromoCombo.Cols - 5
        VSFGPromoCombo.ColPosition(9) = VSFGPromoCombo.Cols - 7
        VSFGPromoCombo.ColPosition(8) = VSFGPromoCombo.Cols - 8
        VSFGPromoCombo.ColPosition(7) = VSFGPromoCombo.Cols - 10
        VSFGPromoCombo.ColPosition(6) = VSFGPromoCombo.Cols - 11
        VSFGPromoCombo.ColPosition(5) = VSFGPromoCombo.Cols - 13
        VSFGPromoCombo.ColPosition(4) = VSFGPromoCombo.Cols - 14
        VSFGPromoCombo.ColPosition(3) = VSFGPromoCombo.Cols - 16
        VSFGPromoCombo.ColPosition(2) = VSFGPromoCombo.Cols - 17
        VSFGPromoCombo.ColPosition(1) = VSFGPromoCombo.Cols - 19
        VSFGPromoCombo.TextMatrix(0, 1) = "Cod.Cliente"
        VSFGPromoCombo.TextMatrix(0, 2) = "Cliente"
        VSFGPromoCombo.TextMatrix(0, 4) = "Producto1"
        VSFGPromoCombo.TextMatrix(0, 7) = "Producto2"
        VSFGPromoCombo.TextMatrix(0, 10) = "Producto3"
        VSFGPromoCombo.TextMatrix(0, 13) = "Producto4"
        VSFGPromoCombo.TextMatrix(0, 16) = "ProductoPromo"
        VSFGPromoCombo.TextMatrix(0, 20) = "Observacion"
        VSFGPromoCombo.TextMatrix(0, 21) = "Aplica"
        For i = 1 To VSFGPromoCombo.Rows - 1
            VSFGPromoCombo.ShowCell i, 0
            If VSFGPromoCombo.TextMatrix(i, 3) <> VSFGPromoCombo.TextMatrix(i - 1, 3) Then
                strSql = " SELECT prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 3) & "' "
                
                clsCon_Def.Ejecutar strSql
            End If
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGPromoCombo.TextMatrix(i, 4) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 5) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 5) <> VSFGPromoCombo.TextMatrix(i - 1, 5) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 5) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 6) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
                Else
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 7) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 7) <> VSFGPromoCombo.TextMatrix(i - 1, 7) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 7) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 8) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
                ElseIf VSFGPromoCombo.TextMatrix(i, 7) <> "" Then
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 9) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 9) <> VSFGPromoCombo.TextMatrix(i - 1, 9) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 9) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 10) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
                Else
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 11) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 11) <> VSFGPromoCombo.TextMatrix(i - 1, 11) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 11) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                End If
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 12) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
                Else
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 0) <> "%" Then
                If VSFGPromoCombo.TextMatrix(i, 0) <> VSFGPromoCombo.TextMatrix(i - 1, 0) Then
                    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli " & _
                             " FROM persona " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND per_ruc like '%" & VSFGPromoCombo.TextMatrix(i, 0) & "' " & _
                             " AND cat_p_tipo='C' " & _
                             " AND tip_ped_codigo='" & cmbNegocioPromoCombo.BoundText & "' "
                             
                    clsCon_DefP.Ejecutar strSql
                End If
                If clsCon_DefP.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 1) = clsCon_DefP.adorec_Def("per_codigo")
                    VSFGPromoCombo.TextMatrix(i, 2) = clsCon_DefP.adorec_Def("cli")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
                Else
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            Else
                VSFGPromoCombo.TextMatrix(i, 1) = "%"
                VSFGPromoCombo.TextMatrix(i, 2) = "-- Todos --"
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            End If
            
        Next i
    
        VSFGPromoCombo.Select 1, VSFGPromoCombo.Cols - 1
        VSFGPromoCombo.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0

End Sub

Private Sub cmdExplorarPromoComboPedido1_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoPromoComboPedido1.Tag = sDir
    cdArchivoPromoComboPedido.ShowOpen
    txtArchivoPromoComboPedido1 = cdArchivoPromoComboPedido.FileName
    ChDir sDir
    VSFGPromoComboPedido1.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoPromoComboPedido1.Text <> "") Then
        VSFGPromoComboPedido1.Rows = 1
        VSFGPromoComboPedido1.LoadGrid txtArchivoPromoComboPedido1.Text, flexFileCommaText
        VSFGPromoComboPedido1.Cols = 3
        VSFGPromoComboPedido1.TextMatrix(0, 1) = "Producto"
        VSFGPromoComboPedido1.TextMatrix(0, 2) = "Aplica"
        For i = 1 To VSFGPromoComboPedido1.Rows - 1
            VSFGPromoComboPedido1.ShowCell i, 0
            strSql = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFGPromoComboPedido1.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGPromoComboPedido1.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGPromoComboPedido1.TextMatrix(i, VSFGPromoComboPedido1.Cols - 1) = 1
            Else
                VSFGPromoComboPedido1.TextMatrix(i, 1) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGPromoComboPedido1.TextMatrix(i, VSFGPromoComboPedido1.Cols - 1) = 0
                VSFGPromoComboPedido1.Cell(flexcpBackColor, i, 0, i, VSFGPromoComboPedido1.Cols - 1) = vbRed
            End If
            
        Next i
    
        VSFGPromoComboPedido1.Select 1, VSFGPromoComboPedido1.Cols - 1
        VSFGPromoComboPedido1.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0

End Sub

Private Sub cmdExplorarPromoComboPedido2_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivoPromoComboPedido2.Tag = sDir
    cdArchivoPromoComboPedido.ShowOpen
    txtArchivoPromoComboPedido2 = cdArchivoPromoComboPedido.FileName
    ChDir sDir
    VSFGPromoComboPedido2.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivoPromoComboPedido2.Text <> "") Then
        VSFGPromoComboPedido2.Rows = 1
        VSFGPromoComboPedido2.LoadGrid txtArchivoPromoComboPedido2.Text, flexFileCommaText
        VSFGPromoComboPedido2.Cols = 3
        VSFGPromoComboPedido2.TextMatrix(0, 1) = "Producto"
        VSFGPromoComboPedido2.TextMatrix(0, 2) = "Aplica"
        For i = 1 To VSFGPromoComboPedido2.Rows - 1
            VSFGPromoComboPedido2.ShowCell i, 0
            strSql = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFGPromoComboPedido2.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFGPromoComboPedido2.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFGPromoComboPedido2.TextMatrix(i, VSFGPromoComboPedido2.Cols - 1) = 1
            Else
                VSFGPromoComboPedido2.TextMatrix(i, 1) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                VSFGPromoComboPedido2.TextMatrix(i, VSFGPromoComboPedido2.Cols - 1) = 0
                VSFGPromoComboPedido2.Cell(flexcpBackColor, i, 0, i, VSFGPromoComboPedido2.Cols - 1) = vbRed
            End If
            
        Next i
    
        VSFGPromoComboPedido2.Select 1, VSFGPromoComboPedido2.Cols - 1
        VSFGPromoComboPedido2.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpFechaInicioAplicar.Value = dtpFechaInicioIncentivo.Value
    dtpFechaFinAplicar.Value = dtpFechaFinIncentivo.Value
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        strSql = " SELECT tip_ped_codigo,tip_ped_nombre " & _
                 " FROM tipo_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tip_ped_nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbNegocioIncentivo.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioIncentivo.ListField = "tip_ped_nombre"
        cmbNegocioIncentivo.BoundColumn = "tip_ped_codigo"
        Set cmbNegocioAplicar.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioAplicar.ListField = "tip_ped_nombre"
        cmbNegocioAplicar.BoundColumn = "tip_ped_codigo"
        Set cmbNegocioPremio.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioPremio.ListField = "tip_ped_nombre"
        cmbNegocioPremio.BoundColumn = "tip_ped_codigo"
        Set cmbNegocioPromoCombo.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioPromoCombo.ListField = "tip_ped_nombre"
        cmbNegocioPromoCombo.BoundColumn = "tip_ped_codigo"
        Set cmbNegocioPromoComboPedido.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioPromoComboPedido.ListField = "tip_ped_nombre"
        cmbNegocioPromoComboPedido.BoundColumn = "tip_ped_codigo"
        Set cmbNegocioDctoCombo.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbNegocioDctoCombo.ListField = "tip_ped_nombre"
        cmbNegocioDctoCombo.BoundColumn = "tip_ped_codigo"
     
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub op_coleccion_Click()
       i_flag = 0
       dcmbColeccion.Enabled = True
       dcmbDescripcion.Enabled = False
       dtpFechaInicioIncentivo.Enabled = False
       dtpFechaFinIncentivo.Enabled = False
End Sub

Private Sub op_disponibleventa_Click()
        i_flag = 2
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicioIncentivo.Enabled = True
       dtpFechaFinIncentivo.Enabled = True
End Sub

Private Sub op_lista_Click()
       i_flag = 1
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = True
       dtpFechaInicioIncentivo.Enabled = False
       dtpFechaFinIncentivo.Enabled = False
End Sub


Private Function PromoPrendaPrecio(TipoPedido As String, Pedido As Double) As Double
    Dim clsPromo As New clsConsulta
    Dim clsConsulta As New clsConsulta
    Dim clsEjecuta As New clsConsulta
    Dim PrendasDePromo As Long
    Dim TotalPrendasDePromo As Long
    Dim NumeroDePromo As Long
    Dim Dcto As Double
    clsConsulta.Inicializar AdoConn, AdoConnMaster
    clsEjecuta.Inicializar AdoConn, AdoConnMaster
    clsPromo.Inicializar AdoConn, AdoConnMaster
    
    PromoPrendaPrecio = 0
    strSql = " SELECT pro_pre_pre_codigo " & _
             " FROM promo_prenda_precio " & _
             " WHERE promo_prenda_precio.emp_codigo='" & strEmpresa & "'" & _
             " AND promo_prenda_precio.tip_ped_codigo='" & TipoPedido & "'" & _
             " AND CURRENT_TIMESTAMP BETWEEN pro_pre_pre_fecha_desde AND pro_pre_pre_fecha_hasta"
    clsPromo.Ejecutar strSql
    While Not clsPromo.adorec_Def.EOF
        strSql = " SELECT det_pedido.dep_codigo,det_pedido.prd_codigo,det_ped_cant_pedida,det_ped_precio,pro_pre_pre_cantidad,det_pro_pre_pre_precio,det_ped_dcto " & _
                 " FROM pedido INNER JOIN det_pedido " & _
                 " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN ( " & _
                 " SELECT promo_prenda_precio.emp_codigo,pro_pre_pre_cantidad,prd_codigo,det_pro_pre_pre_precio" & _
                 " FROM promo_prenda_precio INNER JOIN det_promo_prenda_precio " & _
                 " ON promo_prenda_precio.emp_codigo=det_promo_prenda_precio.emp_codigo " & _
                 " AND promo_prenda_precio.tip_ped_codigo=det_promo_prenda_precio.tip_ped_codigo " & _
                 " AND promo_prenda_precio.pro_pre_pre_codigo=det_promo_prenda_precio.pro_pre_pre_codigo " & _
                 " WHERE promo_prenda_precio.emp_codigo='" & strEmpresa & "'" & _
                 " AND promo_prenda_precio.pro_pre_pre_codigo='" & clsPromo.adorec_Def("pro_pre_pre_codigo") & "'" & _
                 " AND promo_prenda_precio.tip_ped_codigo='" & TipoPedido & "'" & _
                 " AND CURRENT_TIMESTAMP BETWEEN pro_pre_pre_fecha_desde AND pro_pre_pre_fecha_hasta " & _
                 " ) promo " & _
                 " ON det_pedido.emp_codigo=promo.emp_codigo " & _
                 " AND det_pedido.prd_codigo=promo.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
                 " AND pedido.ped_codigo='" & Pedido & "'" & _
                 " ORDER BY det_ped_precio DESC,det_ped_cant_pedida ASC"
        clsConsulta.Ejecutar (strSql), "M"
        If clsConsulta.adorec_Def.RecordCount > 0 Then
            PrendasDePromo = clsConsulta.adorec_Def("pro_pre_pre_cantidad")
            TotalPrendasDePromo = 0
            While Not clsConsulta.adorec_Def.EOF
                TotalPrendasDePromo = TotalPrendasDePromo + clsConsulta.adorec_Def("det_ped_cant_pedida")
                clsConsulta.adorec_Def.MoveNext
            Wend
            If PrendasDePromo <= TotalPrendasDePromo Then
                clsConsulta.adorec_Def.MoveFirst
                NumeroDePromo = Int(TotalPrendasDePromo / PrendasDePromo) * PrendasDePromo
                
                While (Not clsConsulta.adorec_Def.EOF) And NumeroDePromo > 0
                    Dcto = 0
                    If FormatoD0(clsConsulta.adorec_Def("det_ped_cant_pedida")) <= NumeroDePromo Then
                        Dcto = FormatoD0(clsConsulta.adorec_Def("det_ped_cant_pedida")) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                             - FormatoD0(clsConsulta.adorec_Def("det_ped_cant_pedida")) * FormatoD2(clsConsulta.adorec_Def("det_pro_pre_pre_precio"))
                        If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                            Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                        End If
                        NumeroDePromo = NumeroDePromo - FormatoD0(clsConsulta.adorec_Def("det_ped_cant_pedida"))
                    Else
                        Dcto = FormatoD0(NumeroDePromo) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                             - FormatoD0(NumeroDePromo) * FormatoD2(clsConsulta.adorec_Def("det_pro_pre_pre_precio"))
                        If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                            Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                        End If
                        NumeroDePromo = 0
                    End If
                    strSql = " UPDATE det_pedido " & _
                             " SET det_ped_dcto='" & FormatoD4(Dcto) & "'" & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ped_codigo='" & Pedido & "'" & _
                             " AND prd_codigo='" & clsConsulta.adorec_Def("prd_codigo") & "'" & _
                             " AND dep_codigo='" & clsConsulta.adorec_Def("dep_codigo") & "'"
                    clsEjecuta.Ejecutar strSql, "M"
                    PromoPrendaPrecio = PromoPrendaPrecio + FormatoD4(Dcto)
                    clsConsulta.adorec_Def.MoveNext
                Wend
            End If
        End If
        clsPromo.adorec_Def.MoveNext
    Wend
End Function

Private Function ComprobarPedido() As Boolean
    Dim cantTotal As Double
    cantTotal = FormatoD4(txtCantidad.Text)
    strSql = " DROP TABLE IF EXISTS temp1 "
    clsSql.Ejecutar strSql, "M"
    strSql = " CREATE TEMPORARY TABLE temp1(prod VARCHAR(40), cant DECIMAL(14,4)) "
    clsSql.Ejecutar strSql, "M"
    With VSFG
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" And Val(FormatoD4(.TextMatrix(i, 4))) > 0 Then
                strSql = " INSERT INTO temp1 " & _
                         " VALUES ('" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 4) & "') "
                clsSql.Ejecutar strSql, "M"
            End If
        Next i
    End With
    
    strSql = " DROP TABLE IF EXISTS temp2 "
    clsSql.Ejecutar strSql, "M"
    strSql = " CREATE TEMPORARY TABLE temp2 " & _
             " SELECT GROUP_CONCAT(prod,'-',cant ORDER BY prod) as xxx " & _
             " FROM temp1 "
    clsSql.Ejecutar strSql, "M"
    
    strSql = " DROP TABLE IF EXISTS temp3 "
    clsSql.Ejecutar strSql, "M"
    strSql = " CREATE TEMPORARY TABLE temp3(codigo DECIMAL(14,0),xxx TEXT,cant DECIMAl(14,4)) "
    clsSql.Ejecutar strSql, "M"

    strSql = " INSERT INTO temp3 " & _
             " SELECT pedido.ped_codigo as codigo,GROUP_CONCAT(prd_codigo,'-',det_ped_cant_pedida ORDER BY prd_codigo) as xxx,SUM(det_ped_cant_pedida) " & _
             " FROM pedido " & _
             " INNER JOIN det_pedido " & _
             " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
             " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND pedido.per_codigo='" & cmbCliente.BoundText & "' " & _
             " AND pedido.ped_estado!='3' " & _
             " GROUP BY pedido.ped_codigo " & _
             " HAVING SUM(det_ped_cant_pedida)=" & cantTotal
    clsSql.Ejecutar strSql, "M"
        
    strSql = " SELECT temp3.codigo " & _
             " FROM temp2 " & _
             " INNER JOIN temp3 " & _
             " ON temp2.xxx=temp3.xxx "
    clsSql.Ejecutar strSql, "M"
    If clsSql.adorec_Def.RecordCount > 0 Then
        If Trim(clsSql.adorec_Def(0)) <> "" And Not IsNull(clsSql.adorec_Def(0)) Then
            If MsgBox("Este pedido se encuentra duplicado (" & clsSql.adorec_Def(0) & ") para la misma persona" & vbCrLf & "Por favor verifique los datos. Desea generar el pedido?", vbQuestion + vbYesNo, "Comprobar Pedido") = vbNo Then
                ComprobarPedido = True
            Else
                ComprobarPedido = False
            End If
        Else
            ComprobarPedido = False
        End If
    Else
        ComprobarPedido = False
    End If
    strSql = " DROP TABLE IF EXISTS temp1 "
    clsSql.Ejecutar strSql, "M"
    strSql = " DROP TABLE IF EXISTS temp2 "
    clsSql.Ejecutar strSql, "M"
    strSql = " DROP TABLE IF EXISTS temp3 "
    clsSql.Ejecutar strSql, "M"
End Function

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
    strSql = " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbNegocio.RowSource = clsSql.adorec_Def.DataSource
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
    StockMin = 0
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
             " WHERE emp_codigo = '" & strEmpresa & "' AND '" & cmbNegocio.BoundText & "' LIKE CONCAT(tip_ped_codigo) " & _
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
    Set CmbBodega.RowSource = clsBods.adorec_Def.DataSource
    CmbBodega.ListField = "dep_nombre"
    CmbBodega.BoundColumn = "dep_codigo"
    
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
            MsgBox "No se encontró un cliente con CI/RUC " & txtRuc.Text, vbInformation, "CI/RUC"
            If cmbCliente.Text <> "" Then
                LimpiarTodo
            Else
                txtRuc.Text = ""
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
                If VSFG.TextMatrix(Row, 8) = 0 Then
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
        If tipoPed > 0 And tipoPed < 3 Then
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
                 " AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "' " & _
                 " AND CURRENT_DATE BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD2(clsSql.adorec_Def(0)) > FormatoD2(txtDcto.Text) Then
                dctoMax = FormatoD2(clsSql.adorec_Def(0))
            End If
        End If
'''
        VSFG.TextMatrix(Row, 6) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD4(dctoMax) / 100)
        VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
        CalcuTotal
    End If
    
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid en caso de ser necesario
    'If VSFG.Rows - 1 >= OldRow And VSFG.Rows - 1 >= NewRow Then
    If FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 7)) <> 0 Then
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
            Cancel = True
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
                     " producto.prd_nombre, (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio*'" & 1 + dblComision / 100 & "' as lis_pre_p_precio,prd_cambia_precio " & _
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
                     " GROUP BY dep_codigo, prd_codigo " & _
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
                    VSFG.TextMatrix(Row, 6) = 0#
                    'Verifica que el precio de la lista no sea menor al costo del producto y tampoco sea una cotización
                    ''If clsLstPrds.adorec_Def("prd_costo") > clsLstPrds.adorec_Def("lis_pre_p_precio") Then ''''And tipoPed <> 1 Then
                    ''    VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("prd_costo"))
                    ''Else
                        VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("lis_pre_p_precio"))
                    ''End If
                    'Verifica que la existencia del producto sea mayor que cero
                    If clsLstPrds.adorec_Def("exi_cantidad") - IIf(FacDirecto = True, 0, StockMin) > 0 And FormatoD4(VSFG.TextMatrix(Row, 5)) <> 0 Then
                        VSFG.TextMatrix(Row, 4) = 1
                        VSFG.TextMatrix(Row, 11) = 1
                    Else
                        VSFG.TextMatrix(Row, 4) = 0
                        VSFG.TextMatrix(Row, 11) = 0
                    End If
                    
                    
                    dctoMax = 0
                    dctoMax = FormatoD2(txtDcto.Text)
                    strSql = " SELECT prd_pro_porcentaje " & _
                             " FROM producto_promo " & _
                             " WHERE emp_codigo = '" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "' " & _
                             " AND CURRENT_DATE BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                    clsSql.Ejecutar strSql
                    If clsSql.adorec_Def.RecordCount > 0 Then
                        If FormatoD2(clsSql.adorec_Def(0)) > FormatoD2(txtDcto.Text) Then
                            dctoMax = FormatoD2(clsSql.adorec_Def(0))
                        End If
                    End If
                    
                    VSFG.TextMatrix(Row, 6) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD4(dctoMax) / 100)
                    VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
                    VSFG.TextMatrix(Row, 8) = clsLstPrds.adorec_Def("exi_cantidad") - StockMin
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
    If VSFG.Col = 3 And KeyCode = vbKeyF4 And Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col)) <> "" And Len(Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col))) >= 4 Then
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
                tipoPed = 0
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
                tipoPed = 0
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
                tipoPed = 1
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
                tipoPed = 2
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
                tipoPed = 3
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
                tipoPed = 4
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
                         " LEFT JOIN vendedor ON (vendedor.ven_codigo = IF(pedido.ven_codigo='' OR pedido.ven_codigo IS NULL,persona.ven_codigo,pedido.ven_codigo)) AND (vendedor.emp_codigo = pedido.emp_codigo)) " & _
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
    If VSFGTPeds.Col = 1 And VSFGTPeds.Row > 0 And tipoPed > 0 And ban = 0 Then
        If VSFGTPeds.ComboIndex >= 0 Then
            lngCod = VSFGTPeds.ComboItem(VSFGTPeds.ComboIndex)
            VSFG.Editable = flexEDKbdMouse
            VSFG.Rows = 2
            VSFG.Clear 1
            Select Case tipoPed
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
                    strSql = " SELECT if(producto.prd_codigo<>'',producto.prd_codigo,det_prd_com.prd_codigo)as prd_codigo, " & _
                             " if(producto.prd_nombre<>'',producto.prd_nombre,producto_1.prd_nombre) as prd_nombre, " & _
                             " sum(if(isnull(det_prd_com.det_prd_com_cantidad),det_cotizacion.det_cot_cantidad, " & _
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
'''''                    strSqlPrdTemp = " INSERT INTO TempReser SELECT prd_codigo,dep_codigo,sum(if(ped_estado=1,det_ped_cant_entregada,det_ped_cant_pedida)) as cant " & _
'''''                                    " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
'''''                                    " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
'''''                                    " AND ped_estado<=1 GROUP BY prd_codigo,dep_codigo"
'''''                    clsDet.Ejecutar strSqlPrdTemp
                    
                    'Obtiene solo los productos que intervienen en la factura con sus respectivos datos
                    strSql = " SELECT det_egreso.dep_codigo,det_egreso.prd_codigo, prd_nombre, IF(det_egr_cantidad<SUM(exi_cantidad),det_egr_cantidad,SUM(exi_cantidad)),det_egr_precio,0,IF(det_egr_cantidad<SUM(exi_cantidad),det_egr_cantidad,SUM(exi_cantidad))*det_egr_precio as tot,COALESCE(SUM(exi_cantidad),0) as exist,(producto.prd_costo/(1 - 0.1)) as prd_costo,prd_cambia_precio " & _
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
''''                    strSqlPrdTemp = " INSERT INTO TempReser SELECT prd_codigo,dep_codigo,sum(if(ped_estado=1,det_ped_cant_entregada,det_ped_cant_pedida)) as cant " & _
''''                                    " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
''''                                    " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
''''                                    " AND ped_estado<=1 GROUP BY prd_codigo,dep_codigo"
''''                    clsDet.Ejecutar strSqlPrdTemp
                    'Obtiene solo los productos que intervienen en la factura con sus respectivos datos
                    strSql = " SELECT det_pedido.dep_codigo,det_pedido.prd_codigo, prd_nombre, det_ped_cant_pedida,det_ped_precio,det_ped_dcto as dcto,det_ped_cant_pedida*det_ped_precio as tot,COALESCE(SUM(exi_cantidad),0)+det_ped_cant_pedida as exist,(producto.prd_costo/(1 - 0.1)) as prd_costo,prd_cambia_precio " & _
                             " FROM (det_pedido INNER JOIN producto ON producto.emp_codigo = det_pedido.emp_codigo  AND producto.prd_codigo = det_pedido.prd_codigo) " & _
                             " INNER JOIN existencia ON existencia.emp_codigo = det_pedido.emp_codigo AND existencia.prd_codigo = det_pedido.prd_codigo AND existencia.dep_codigo = det_pedido.dep_codigo " & _
                             " Where det_pedido.emp_codigo='" & strEmpresa & "' And ped_codigo=" & lngCod & _
                             " GROUP BY dep_codigo,prd_codigo, prd_nombre, det_ped_cant_pedida,det_ped_precio " & _
                             " ORDER BY producto.prd_codigo "
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


