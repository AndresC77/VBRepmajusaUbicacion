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
   ClientLeft      =   3735
   ClientTop       =   465
   ClientWidth     =   14280
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
   ScaleWidth      =   14280
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
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
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
      Tab(0).Control(3)=   "Label32"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpFechaFinAplicar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpFechaInicioAplicar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "VSFGAplicar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbNegocioAplicar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAplicarAplicar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdActualizar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optIncentivo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optPromoCombo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optPremio"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optPromoComboPedido"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optDctoCombo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "optNPrendasAY"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "optPromoPremioPorMontoMarca"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPedido"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "optDctoFecha"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "optDctoMonto"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Cargar Incentivos"
      TabPicture(1)   =   "frmIncentivos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "cdArchivoIncentivo"
      Tab(1).Control(6)=   "cmbNegocioIncentivo"
      Tab(1).Control(7)=   "dtpFechaFinIncentivo"
      Tab(1).Control(8)=   "dtpFechaInicioIncentivo"
      Tab(1).Control(9)=   "VSFGIncentivo"
      Tab(1).Control(10)=   "txtNombreIncentivo"
      Tab(1).Control(11)=   "cmdExplorarIncentivo"
      Tab(1).Control(12)=   "txtArchivoIncentivo"
      Tab(1).Control(13)=   "cmdAplicarIncentivo"
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
      Tab(3).Control(0)=   "txtNombrePremio"
      Tab(3).Control(1)=   "cmdExplorarPremio"
      Tab(3).Control(2)=   "txtArchivoPremio"
      Tab(3).Control(3)=   "cmdAplicarPremio"
      Tab(3).Control(4)=   "VSFGPremio"
      Tab(3).Control(5)=   "dtpFechaInicioPremio"
      Tab(3).Control(6)=   "dtpFechaFinPremio"
      Tab(3).Control(7)=   "cmbNegocioPremio"
      Tab(3).Control(8)=   "cdArchivoPremio"
      Tab(3).Control(9)=   "Label18"
      Tab(3).Control(10)=   "Label17"
      Tab(3).Control(11)=   "Label16"
      Tab(3).Control(12)=   "Label15"
      Tab(3).Control(13)=   "Label14"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Cargar Promo Combo Pedido"
      TabPicture(4)   =   "frmIncentivos.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label19"
      Tab(4).Control(1)=   "Label20"
      Tab(4).Control(2)=   "Label21"
      Tab(4).Control(3)=   "Label22"
      Tab(4).Control(4)=   "Label23"
      Tab(4).Control(5)=   "Label24"
      Tab(4).Control(6)=   "Label25"
      Tab(4).Control(7)=   "Label26"
      Tab(4).Control(8)=   "VSFGPromoComboPedido2"
      Tab(4).Control(9)=   "cdArchivoPromoComboPedido"
      Tab(4).Control(10)=   "cmbNegocioPromoComboPedido"
      Tab(4).Control(11)=   "dtpFechaFinPromoComboPedido"
      Tab(4).Control(12)=   "dtpFechaInicioPromoComboPedido"
      Tab(4).Control(13)=   "VSFGPromoComboPedido1"
      Tab(4).Control(14)=   "txtNombrePromoComboPedido"
      Tab(4).Control(15)=   "cmdExplorarPromoComboPedido1"
      Tab(4).Control(16)=   "txtArchivoPromoComboPedido1"
      Tab(4).Control(17)=   "cmdAplicarPromoComboPedido"
      Tab(4).Control(18)=   "cmdExplorarPromoComboPedido2"
      Tab(4).Control(19)=   "txtArchivoPromoComboPedido2"
      Tab(4).Control(20)=   "txtCantMinPromoComboPedido1"
      Tab(4).Control(21)=   "txtCantEntPromoComboPedido1"
      Tab(4).ControlCount=   22
      TabCaption(5)   =   "Cargar Dcto x Combo"
      TabPicture(5)   =   "frmIncentivos.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label27"
      Tab(5).Control(1)=   "Label28"
      Tab(5).Control(2)=   "Label29"
      Tab(5).Control(3)=   "Label30"
      Tab(5).Control(4)=   "Label31"
      Tab(5).Control(5)=   "cdArchivoDctoCombo"
      Tab(5).Control(6)=   "cmbNegocioDctoCombo"
      Tab(5).Control(7)=   "dtpFechaFinDctoCombo"
      Tab(5).Control(8)=   "dtpFechaInicioDctoCombo"
      Tab(5).Control(9)=   "VSFGDctoCombo"
      Tab(5).Control(10)=   "cmdAplicarDctoCombo"
      Tab(5).Control(11)=   "txtArchivoDctoCombo"
      Tab(5).Control(12)=   "cmdExplorarDctoCombo"
      Tab(5).Control(13)=   "txtNombreDctoCombo"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Carga Promo Premio por monto Marca"
      TabPicture(6)   =   "frmIncentivos.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.OptionButton optDctoMonto 
         Caption         =   "Dcto x Monto"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3840
         TabIndex        =   94
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton optDctoFecha 
         Caption         =   "Descuento por Fecha"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3840
         TabIndex        =   93
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPedido 
         Height          =   330
         Left            =   8160
         TabIndex        =   91
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optPromoPremioPorMontoMarca 
         Caption         =   "Premio por monto Marca"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3840
         TabIndex        =   90
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optNPrendasAY 
         Caption         =   "n Prendas a $x.xx"
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
         Left            =   -66105
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
         Left            =   -66465
         TabIndex        =   68
         Top             =   2040
         Width           =   3360
      End
      Begin VB.CommandButton cmdExplorarPromoComboPedido2 
         Caption         =   "..."
         Height          =   315
         Left            =   -63105
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
         Left            =   10680
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
         Width           =   13575
         _cx             =   23945
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
         FormatString    =   $"frmIncentivos.frx":03CE
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
         Left            =   11685
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
         Format          =   69206019
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
         Left            =   11685
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
         Format          =   69206019
         CurrentDate     =   37463
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGIncentivo 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   14
         Top             =   1680
         Width           =   13575
         _cx             =   23945
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
         FormatString    =   $"frmIncentivos.frx":0544
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
         Width           =   13575
         _cx             =   23945
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
         FormatString    =   $"frmIncentivos.frx":05E4
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
         Width           =   13575
         _cx             =   23945
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
         FormatString    =   $"frmIncentivos.frx":0775
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
         Width           =   6735
         _cx             =   11880
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
         FormatString    =   $"frmIncentivos.frx":085B
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
      Begin MSComDlg.CommonDialog cdArchivoPromoComboPedido 
         Left            =   -67800
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPromoComboPedido2 
         Height          =   4335
         Left            =   -68040
         TabIndex        =   66
         Top             =   2400
         Width           =   6735
         _cx             =   11880
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
         FormatString    =   $"frmIncentivos.frx":088F
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
         Width           =   13575
         _cx             =   23945
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
         FormatString    =   $"frmIncentivos.frx":08C3
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
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7440
         TabIndex        =   92
         Top             =   960
         Width           =   525
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
         Left            =   -67215
         TabIndex        =   73
         Top             =   1725
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
         Left            =   -67200
         TabIndex        =   69
         Top             =   2085
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
         Left            =   10965
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
         Left            =   10725
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
Option Explicit
Private clsCon_Def As New clsConsulta
Private strSql As String
Dim i_flag As Integer
Dim dblIVA As Double

Public Sub cmdActualizar_Click()
    Dim i As Long
    Dim clsAux As New clsConsulta
    Dim clsAux2 As New clsConsulta
    Dim FiltroPedido As String
    If txtPedido.Text <> "" Then
        FiltroPedido = " AND pedido.ped_codigo='" & txtPedido.Text & "' "
    Else
        FiltroPedido = ""
    End If
    clsAux.Inicializar AdoConn, AdoConnMaster
    clsAux2.Inicializar AdoConn, AdoConnMaster
    If optDctoFecha.Value = True Then
        VSFGAplicar.Cols = 6
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 4) = "DCTO"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,'10' AS pDCTO " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " AND pedido.ped_fechamod BETWEEN '2018-05-29 00:00:00' and '2018-06-04 23:59:59' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_subtotal >= 60 AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '2018-05-29 00:00:00' and '2018-06-04 23:59:59' " & _
                 " ORDER BY per_ruc "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
        
    ElseIf optPremio.Value = True Then
        VSFGAplicar.Cols = 12
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
        VSFGAplicar.TextMatrix(0, 11) = "ComoIncentivo"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pre_loc_nombre,producto.prd_codigo,prd_nombre," & _
                 " pre_loc_cantidad,pre_loc_precio,pre_loc_cantidad*pre_loc_dcto,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN premio_local ON persona.emp_codigo=premio_local.emp_codigo " & _
                 " AND persona.per_codigo LIKE premio_local.per_codigo " & _
                 " AND persona.tip_ped_codigo=premio_local.tip_ped_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(premio_local.pre_loc_fecha_desde,10) AND LEFT(premio_local.pre_loc_fecha_hasta,10) " & _
                 " AND premio_local.pre_loc_total_cantidad='T' " & _
                 " AND premio_local.pre_loc_estado=0 " & _
                 " AND round(pedido.ped_subtotal/1." & dblIVA & ",2) BETWEEN premio_local.pre_loc_rango_inferior AND premio_local.pre_loc_rango_superior " & _
                 " INNER JOIN producto ON premio_local.emp_codigo=producto.emp_codigo " & _
                 " AND premio_local.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION " & _
                 " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pre_loc_nombre,producto.prd_codigo,prd_nombre," & _
                 " pre_loc_cantidad,pre_loc_precio,pre_loc_cantidad*pre_loc_dcto,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN ("
        strSql = strSql & _
                 " SELECT pedido.emp_codigo,pedido.ped_codigo,SUM(det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada) as cantTotal " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " AND producto.prd_incentivo=0 AND producto.prd_codigo NOT LIKE 'PR-%' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " GROUP BY pedido.emp_codigo,pedido.ped_codigo" & _
                 " HAVING SUM(det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>0" & _
                 ") cant ON pedido.emp_codigo=cant.emp_codigo " & _
                 " AND pedido.ped_codigo=cant.ped_codigo "
        strSql = strSql & " INNER JOIN premio_local ON persona.emp_codigo=premio_local.emp_codigo " & _
                 " AND persona.per_codigo LIKE premio_local.per_codigo " & _
                 " AND persona.tip_ped_codigo=premio_local.tip_ped_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(premio_local.pre_loc_fecha_desde,10) AND LEFT(premio_local.pre_loc_fecha_hasta,10) " & _
                 " AND premio_local.pre_loc_total_cantidad='C' " & _
                 " AND premio_local.pre_loc_estado=0 " & _
                 " AND round(cant.cantTotal,2) BETWEEN premio_local.pre_loc_rango_inferior AND premio_local.pre_loc_rango_superior " & _
                 " INNER JOIN producto ON premio_local.emp_codigo=producto.emp_codigo " & _
                 " AND premio_local.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc,producto.prd_codigo "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    ElseIf optPromoPremioPorMontoMarca.Value = True Then
        VSFGAplicar.Cols = 12
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
        VSFGAplicar.TextMatrix(0, 11) = "ComoIncentivo"
        strSql = " SELECT ped_codigo,ped_fechamod,tip_ped_codigo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        
        clsAux.Ejecutar strSql
        
        While Not clsAux.adorec_Def.EOF
            strSql = " SELECT pro_mar_nombre,pro_mar_filtro_codigos,pro_mar_total,producto.prd_codigo,prd_nombre," & _
                     " pro_mar_cantidad,pro_mar_precio,pro_mar_dcto " & _
                     " FROM promo_marcas " & _
                     " INNER JOIN producto ON promo_marcas.emp_codigo=producto.emp_codigo " & _
                     " AND promo_marcas.prd_codigo=producto.prd_codigo " & _
                     " WHERE promo_marcas.emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ped_codigo='" & clsAux.adorec_Def("tip_ped_codigo") & "' " & _
                     " AND '" & clsAux.adorec_Def("ped_fechamod") & "' between pro_mar_fecha_desde and pro_mar_fecha_hasta "
            clsAux2.Ejecutar strSql
            While Not clsAux2.adorec_Def.EOF
                strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                         " CONCAT(per_apellido, ' ',per_nombre) as cli,'" & clsAux2.adorec_Def("pro_mar_nombre") & "' as np,'" & clsAux2.adorec_Def("prd_codigo") & "' as prd_codigo,'" & clsAux2.adorec_Def("prd_nombre") & "' as prd_nombre," & _
                         " '" & clsAux2.adorec_Def("pro_mar_cantidad") & "' as pre_loc_cantidad,'" & clsAux2.adorec_Def("pro_mar_precio") & "' as pre_loc_precio,'" & clsAux2.adorec_Def("pro_mar_dcto") & "' as pre_loc_dcto," & _
                         " ROUND((SUM(det_ped_cant_entregada*det_ped_precio) - SUM(ROUND(det_ped_cant_entregada*det_ped_precio*if(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2))),2) ,ROUND((SUM(det_ped_cant_entregada*det_ped_precio) - SUM(ROUND(det_ped_cant_entregada*det_ped_precio*if(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2))),2) as t,prd_incentivo" & _
                         " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                         " AND pedido.per_codigo=persona.per_codigo " & _
                         " AND persona.cat_p_tipo='C' " & _
                         " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                         " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo  " & _
                         " AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                         " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                         " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                         " AND producto.mar_codigo " & clsAux2.adorec_Def("pro_mar_filtro_codigos") & _
                         " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                         " AND producto_promo.prd_pro_fechaini<=LEFT(pedido.ped_fechamod,10) AND producto_promo.prd_pro_fechafin>=LEFT(pedido.ped_fechamod,10) AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                         " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                         " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                         " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                         " AND pedido.ped_codigo='" & clsAux.adorec_Def("ped_codigo") & "'" & _
                         " GROUP BY ped_codigo " & _
                         " HAVING t>='" & clsAux2.adorec_Def("pro_mar_total") & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGAplicar.AddItem clsCon_Def.adorec_Def(0) & vbTab & clsCon_Def.adorec_Def(1) & vbTab & clsCon_Def.adorec_Def(2) & vbTab & _
                                        clsCon_Def.adorec_Def(3) & vbTab & clsCon_Def.adorec_Def(4) & vbTab & clsCon_Def.adorec_Def(5) & vbTab & _
                                        clsCon_Def.adorec_Def(6) & vbTab & clsCon_Def.adorec_Def(7) & vbTab & clsCon_Def.adorec_Def(8) & vbTab & _
                                        clsCon_Def.adorec_Def(9) & vbTab & clsCon_Def.adorec_Def(10) & vbTab & clsCon_Def.adorec_Def(11)
                End If
                
                clsAux2.adorec_Def.MoveNext
            Wend
            clsAux.adorec_Def.MoveNext
        Wend
        'MsgBox "AAA"
    ElseIf optIncentivo.Value = True Then
        VSFGAplicar.Cols = 13
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
        VSFGAplicar.TextMatrix(0, 11) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 12) = "FAAC"
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,inc_loc_nombre,producto.prd_codigo,prd_nombre," & _
                 " inc_loc_cantidad,inc_loc_precio,inc_loc_cantidad*inc_loc_dcto,prd_incentivo,ped_egr_codigo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN incentivo_local ON persona.emp_codigo=incentivo_local.emp_codigo " & _
                 " AND persona.per_codigo=incentivo_local.per_codigo " & _
                 " AND persona.tip_ped_codigo=incentivo_local.tip_ped_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(incentivo_local.inc_loc_fecha_desde,10) AND LEFT(incentivo_local.inc_loc_fecha_hasta,10) " & _
                 " AND incentivo_local.inc_loc_estado=0 " & _
                 " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
                 " AND incentivo_local.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc,producto.prd_codigo "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
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
        strSql = strSql & " INNER JOIN ( SELECT p.emp_codigo,p.ped_codigo,SUM(dp.det_ped_cant_entregada+dp.det_ped_cant_programada) as combosol, " & _
                 " pro_com_ped_cantidad_min,pro_com_ped_cantidad_ent " & _
                 " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                 " AND p.ped_codigo=dp.ped_codigo " & _
                 " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                 " AND p.ped_fechamod BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                 " INNER JOIN det_promo_combo_pedido_ent dpcpe ON pcp.emp_codigo=dpcpe.emp_codigo" & _
                 " AND pcp.tip_ped_codigo=dpcpe.tip_ped_codigo" & _
                 " AND pcp.pro_com_ped_codigo=dpcpe.pro_com_ped_codigo" & _
                 " AND dp.prd_codigo=dpcpe.prd_codigo " & _
                 " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                 " AND p.ped_estado in (0,1) " & Replace(FiltroPedido, "pedido", "p") & _
                 " AND p.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " GROUP BY p.emp_codigo,p.ped_codigo, pro_com_ped_cantidad_min,pro_com_ped_cantidad_ent" & _
                 ") pe ON pedido.emp_codigo=pe.emp_codigo " & _
                 " AND pedido.ped_codigo=pe.ped_codigo "
        strSql = strSql & " LEFT JOIN ( SELECT p.emp_codigo,p.ped_codigo, " & _
                 " SUM(dp.det_ped_cant_entregada+dp.det_ped_cant_programada) as productoing " & _
                 " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                 " AND p.ped_codigo=dp.ped_codigo " & _
                 " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                 " AND p.ped_fechamod BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                 " INNER JOIN det_promo_combo_pedido dpcp ON pcp.emp_codigo=dpcp.emp_codigo" & _
                 " AND pcp.tip_ped_codigo=dpcp.tip_ped_codigo" & _
                 " AND pcp.pro_com_ped_codigo=dpcp.pro_com_ped_codigo" & _
                 " AND dp.prd_codigo=dpcp.prd_codigo " & _
                 " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                 " AND p.ped_estado in (0,1) " & Replace(FiltroPedido, "pedido", "p") & _
                 " AND p.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " GROUP BY p.emp_codigo,p.ped_codigo" & _
                 ") pi ON pedido.emp_codigo=pi.emp_codigo " & _
                 " AND pedido.ped_codigo=pi.ped_codigo "
        strSql = strSql & " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
                 " ORDER BY per_ruc "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    ElseIf optDctoCombo.Value = True Then
        VSFGAplicar.Cols = 14
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
        VSFGAplicar.TextMatrix(0, 12) = "MinimoPromo"
        VSFGAplicar.TextMatrix(0, 13) = "CantidadPromo"
        '1 producto
'        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
'                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
'                 " det_pedido.det_ped_cant_confirmada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_confirmada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto " & _
'                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
'                 " AND pedido.per_codigo=persona.per_codigo " & _
'                 " AND persona.cat_p_tipo='C' " & _
'                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
'                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
'                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
'                 " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
'                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
'                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
'                 " AND det_pedido.det_ped_cant_confirmada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
'                 " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
'                 " AND promo_combo_dcto.prd_codigo_2='' AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
'                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
'                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
'                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
'                 " AND pedido.ped_estado in (0,1) " & _
'                 " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
'        strSql = strSql & " UNION "
        '2 producto 1
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto, " & _
                 " round(det_pedido.det_ped_cant_entregada/promo_combo_dcto.pro_com_dct_cantidad_1,0,1) as minimopromo,promo_combo_dcto.pro_com_dct_cantidad_1 as cantidadpromo " & _
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
                 " AND det_pedido.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        '2 producto 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto, " & _
                 " round(det_pedido.det_ped_cant_entregada/promo_combo_dcto.pro_com_dct_cantidad_2,0,1) as minimopromo,promo_combo_dcto.pro_com_dct_cantidad_2 as cantidadpromo " & _
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
                 " AND det_pedido.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_3='' AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        
        '3 producto 1
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto, " & _
                 " round(det_pedido.det_ped_cant_entregada/promo_combo_dcto.pro_com_dct_cantidad_1,0,1) as minimopromo,promo_combo_dcto.pro_com_dct_cantidad_1 as cantidadpromo " & _
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
        strSql = strSql & " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND det_pedido.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND dpa.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo_dcto.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        '3 producto 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto, " & _
                 " round(det_pedido.det_ped_cant_entregada/promo_combo_dcto.pro_com_dct_cantidad_2,0,1) as minimopromo,promo_combo_dcto.pro_com_dct_cantidad_2 as cantidadpromo " & _
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
        strSql = strSql & " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND det_pedido.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo_dcto.prd_codigo_3 " & _
                 " AND dpb.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        '3 producto 3
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_dct_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dct_dcto/100.00,pro_com_dct_dcto, " & _
                 " round(det_pedido.det_ped_cant_entregada/promo_combo_dcto.pro_com_dct_cantidad_3,0,1) as minimopromo,promo_combo_dcto.pro_com_dct_cantidad_3 as cantidadpromo " & _
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
        strSql = strSql & " INNER JOIN promo_combo_dcto ON persona.emp_codigo=promo_combo_dcto.emp_codigo " & _
                 " AND persona.tip_ped_codigo=promo_combo_dcto.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo_dcto.prd_codigo_3 " & _
                 " AND det_pedido.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_3 " & _
                 " AND dpa.prd_codigo=promo_combo_dcto.prd_codigo_1 " & _
                 " AND dpa.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo_dcto.prd_codigo_2 " & _
                 " AND dpb.det_ped_cant_entregada>=promo_combo_dcto.pro_com_dct_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo_dcto.pro_com_dct_fecha_desde,10) AND LEFT(promo_combo_dcto.pro_com_dct_fecha_hasta,10) " & _
                 " AND promo_combo_dcto.prd_codigo_4='' " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                 " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " ORDER BY ped_codigo,per_codigo,pro_com_dct_nombre,minimopromo "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    ElseIf optPromoCombo.Value = True Then
        VSFGAplicar.Cols = 13
        VSFGAplicar.TextMatrix(0, 0) = "Pedido"
        VSFGAplicar.TextMatrix(0, 1) = "Fecha"
        VSFGAplicar.TextMatrix(0, 2) = "Cod.Cliente"
        VSFGAplicar.TextMatrix(0, 3) = "CI/RUC"
        VSFGAplicar.TextMatrix(0, 4) = "Cliente"
        VSFGAplicar.TextMatrix(0, 5) = "Incentivo"
        VSFGAplicar.TextMatrix(0, 6) = "Cod.Prod."
        VSFGAplicar.TextMatrix(0, 7) = "Producto"
        VSFGAplicar.TextMatrix(0, 8) = "Cantidad"
        VSFGAplicar.TextMatrix(0, 9) = "CantidadProg"
        VSFGAplicar.TextMatrix(0, 10) = "Precio"
        VSFGAplicar.TextMatrix(0, 11) = "Dcto"
        VSFGAplicar.TextMatrix(0, 12) = "ComoIncentivo"
        'a uno
        strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_cant_programada,det_pedido.det_ped_precio," & _
                 " (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_2='' AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b uno
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,pro_com_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*pro_com_precio*pro_com_dcto/100.00,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_2='' AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo "
        strSql = strSql & " UNION "
        'a dos 1
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_cant_programada,det_pedido.det_ped_precio," & _
                 " (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a dos 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,det_pedido.det_ped_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " UNION "
        'b dos
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,0,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100.00,prd_incentivo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN det_pedido dpa ON pedido.emp_codigo=dpa.emp_codigo " & _
                 " AND pedido.ped_codigo=dpa.ped_codigo "
        strSql = strSql & " INNER JOIN promo_combo ON persona.emp_codigo=promo_combo.emp_codigo " & _
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_3='' AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a tres 1
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,det_pedido.det_ped_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " UNION "
        'a tres 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_cant_programada,det_pedido.det_ped_precio," & _
                 " (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a tres 3
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,det_pedido.det_ped_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_3=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " UNION "
        'b tres
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " pro_com_cantidad,0,pro_com_precio,pro_com_cantidad*pro_com_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_4='' " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 1
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,det_pedido.det_ped_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND (dpc.det_ped_cant_entregada+dpc.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_1=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " UNION "
        'a cuatro 2
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_cant_programada,det_pedido.det_ped_precio,det_pedido.det_ped_cant_entregada*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND (dpc.det_ped_cant_entregada+dpc.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_2=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'a cuatro 3
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,det_pedido.det_ped_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND (dpc.det_ped_cant_entregada+dpc.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_3=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " UNION "
        'a cuatro 4
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " det_pedido.det_ped_cant_entregada,det_pedido.det_ped_cant_programada,det_pedido.det_ped_precio," & _
                 " (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)*det_pedido.det_ped_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_4 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpc.det_ped_cant_entregada+dpc.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_4=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
        strSql = strSql & " UNION "
        'b cuatro
        strSql = strSql & " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, " & _
                 " CONCAT(per_apellido, ' ',per_nombre) as cli,pro_com_nombre,producto.prd_codigo,prd_nombre," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada))),0,pro_com_precio," & _
                 " FLOOR(pro_com_cantidad/promo_combo.pro_com_cantidad_1*SUM((det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)))*pro_com_precio*pro_com_dcto/100.00,prd_incentivo " & _
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
                 " AND persona.per_codigo LIKE CONCAT('',promo_combo.per_codigo) " & _
                 " AND persona.tip_ped_codigo=promo_combo.tip_ped_codigo " & _
                 " AND det_pedido.prd_codigo=promo_combo.prd_codigo_1 " & _
                 " AND (det_pedido.det_ped_cant_entregada+det_pedido.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_1 " & _
                 " AND dpa.prd_codigo=promo_combo.prd_codigo_2 " & _
                 " AND (dpa.det_ped_cant_entregada+dpa.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_2 " & _
                 " AND dpb.prd_codigo=promo_combo.prd_codigo_3 " & _
                 " AND (dpb.det_ped_cant_entregada+dpb.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_3 " & _
                 " AND dpc.prd_codigo=promo_combo.prd_codigo_4 " & _
                 " AND (dpc.det_ped_cant_entregada+dpc.det_ped_cant_programada)>=promo_combo.pro_com_cantidad_4 " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN LEFT(promo_combo.pro_com_fecha_desde,10) AND LEFT(promo_combo.pro_com_fecha_hasta,10) " & _
                 " AND promo_combo.prd_codigo_t!='' " & _
                 " INNER JOIN producto ON promo_combo.emp_codigo=producto.emp_codigo " & _
                 " AND promo_combo.prd_codigo_t=producto.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND pedido.ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59' GROUP BY pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,per_apellido,per_nombre,pro_com_nombre,producto.prd_codigo, prd_nombre,pro_com_cantidad,promo_combo.pro_com_cantidad_1,pro_com_precio,pro_com_dcto,prd_incentivo,det_pedido.det_ped_precio"
        strSql = strSql & " ORDER BY per_codigo,ped_codigo,prd_codigo "
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    ElseIf optNPrendasAY.Value = True Then
        strSql = " SELECT ped_codigo,ped_fecha,pedido.per_codigo,per_ruc,CONCAT(per_apellido, ' ',per_nombre) as cli " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocioAplicar.BoundText & "' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
                 " AND ped_estado in (0,1) AND pedido.ped_reprogramado=0 " & FiltroPedido & _
                 " AND ped_fechamod BETWEEN '" & Format(dtpFechaInicioAplicar.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFinAplicar.Value, "yyyy-mm-dd hh:mm") & ":59'"
    
        clsCon_Def.Ejecutar strSql
        Set VSFGAplicar.DataSource = clsCon_Def.adorec_Def.DataSource
    End If
    
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

Public Sub cmdAplicarAplicar_Click()
    Dim i As Long
    Dim j As Long
    Dim dblDcto As Double
    Dim EsRegalo As Integer
    Dim NuevoProductoPedido As Integer
    Dim clsAux As New clsConsulta
    Dim clsPedido As New clsPedidos
    
    Dim Promo As String
    Dim maxPromo As Integer
    
    clsPedido.Inicializar AdoConn, AdoConnMaster
    clsAux.Inicializar AdoConn, AdoConnMaster
    If optDctoFecha.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            For i = 1 To VSFGAplicar.Rows - 1
                strSql = " UPDATE det_pedido" & _
                         " SET det_ped_dcto=ROUND(det_ped_cant_pedida*det_ped_precio*.10,2) " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "' AND det_ped_dcto=0" & _
                         " AND prd_codigo NOT LIKE 'PR-%'"
                clsCon_Def.Ejecutar strSql, "M", False
                clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
            Next i
        End If
    ElseIf optPremio.Value = True Or Me.optPromoPremioPorMontoMarca.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            For i = 1 To VSFGAplicar.Rows - 1
                strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                         " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                         " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto," & _
                         " det_ped_incentivo) " & _
                         " VALUES ('" & strEmpresa & "','PRI','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                         " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & VSFGAplicar.TextMatrix(i, 10) & "', " & _
                         " '" & VSFGAplicar.TextMatrix(i, 11) & "') "
                clsCon_Def.Ejecutar strSql, "M", False
                clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
                VSFGAplicar.ShowCell i, 0
            Next i
        End If
    ElseIf optPromoCombo.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            For i = 1 To VSFGAplicar.Rows - 1
                If VSFGAplicar.TextMatrix(i, 0) <> VSFGAplicar.TextMatrix(i - 1, 0) Or VSFGAplicar.TextMatrix(i, 6) <> VSFGAplicar.TextMatrix(i - 1, 6) Then
                    NuevoProductoPedido = 1
                Else
                    NuevoProductoPedido = 0
                End If
                strSql = " EXEC SP_det_pedido_Mantenimiento " & _
                         " '" & strEmpresa & "','" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','PRI'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 10) & "','" & VSFGAplicar.TextMatrix(i, 11) & "','" & VSFGAplicar.TextMatrix(i, 5) & "'," & _
                         " '" & strUsuario & "', " & _
                         " 0," & NuevoProductoPedido & " "
                clsCon_Def.Ejecutar strSql, "M"
                clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
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
                    clsCon_Def.Ejecutar strSql, "M"
                    EsRegalo = 0
                    If FormatoD2(VSFGAplicar.TextMatrix(i, 8) * VSFGAplicar.TextMatrix(i, 9)) <= FormatoD2(VSFGAplicar.TextMatrix(i, 10)) Then
                        EsRegalo = 1
                    End If
                    strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                             " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                             " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                             " det_ped_cant_entregada, det_ped_precio,det_ped_dcto," & _
                             " det_ped_incentivo) " & _
                             " VALUES ('" & strEmpresa & "','PRI','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                             " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                             " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) & "'," & _
                             " '" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & VSFGAplicar.TextMatrix(i, 10) & "'," & _
                             " '" & VSFGAplicar.TextMatrix(i, 11) & "') "
                    clsCon_Def.Ejecutar strSql, "M", False
'                    strSQL = " INSERT INTO det_egreso(emp_codigo,egr_codigo,tip_egr_codigo,prd_codigo,dep_codigo,det_egr_cantidad,det_egr_precio,det_egr_costo,det_egr_dcto,det_egr_pdcto,det_egr_fechamod,det_egr_usumod) VALUES('RYB','" & VSFGAplicar.TextMatrix(i, 11) & "','FAC','" & VSFGAplicar.TextMatrix(i, 6) & "','PRI','" & VSFGAplicar.TextMatrix(i, 8) & "','" & VSFGAplicar.TextMatrix(i, 9) & "',0,0,0,CURRENT_TIMESTAMP,'ADMIN')"
'
'                    clsCon_Def.Ejecutar strSQL, "M"
'                    If VSFGAplicar.TextMatrix(i, 6) = "NIVEL1" Then
'                    strSQL = " INSERT INTO det_contenedor_mercaderia(emp_codigo,con_mer_codigo,prd_codigo,con_mer_codigo_origen,con_mer_codigo_destino,det_con_mer_fecha,tip_mov_codigo,mov_codigo,det_con_mer_cantidad,det_con_mer_fechamod,det_con_mer_usumod) VALUES('RYB','15120001784','" & VSFGAplicar.TextMatrix(i, 6) & "','15120001784',0,CURRENT_TIMESTAMP,'FAC','" & VSFGAplicar.TextMatrix(i, 11) & "','" & VSFGAplicar.TextMatrix(i, 8) & "',CURRENT_TIMESTAMP,'ADMIN')"
'                    ElseIf VSFGAplicar.TextMatrix(i, 6) = "NIVEL2" Then
'                    strSQL = " INSERT INTO det_contenedor_mercaderia(emp_codigo,con_mer_codigo,prd_codigo,con_mer_codigo_origen,con_mer_codigo_destino,det_con_mer_fecha,tip_mov_codigo,mov_codigo,det_con_mer_cantidad,det_con_mer_fechamod,det_con_mer_usumod) VALUES('RYB','16030000243','" & VSFGAplicar.TextMatrix(i, 6) & "','16030000243',0,CURRENT_TIMESTAMP,'FAC','" & VSFGAplicar.TextMatrix(i, 11) & "','" & VSFGAplicar.TextMatrix(i, 8) & "',CURRENT_TIMESTAMP,'ADMIN')"
'                    Else
'                    strSQL = " INSERT INTO det_contenedor_mercaderia(emp_codigo,con_mer_codigo,prd_codigo,con_mer_codigo_origen,con_mer_codigo_destino,det_con_mer_fecha,tip_mov_codigo,mov_codigo,det_con_mer_cantidad,det_con_mer_fechamod,det_con_mer_usumod) VALUES('RYB','16030000243','" & VSFGAplicar.TextMatrix(i, 6) & "','16030000243',0,CURRENT_TIMESTAMP,'FAC','" & VSFGAplicar.TextMatrix(i, 11) & "','" & VSFGAplicar.TextMatrix(i, 8) & "',CURRENT_TIMESTAMP,'ADMIN')"
'                    End If
'                    clsCon_Def.Ejecutar strSQL, "M"
                    clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
                End If
            End If
        Next i
    ElseIf optDctoCombo.Value = True Then
        If VSFGAplicar.Rows > 1 Then
            Promo = 0
            maxPromo = 0
            For i = 1 To VSFGAplicar.Rows - 1
                If Promo <> VSFGAplicar.TextMatrix(i, 5) Then
                VSFGAplicar.ShowCell i, 13
                    maxPromo = VSFGAplicar.TextMatrix(i, 12)
                End If
                Promo = VSFGAplicar.TextMatrix(i, 5)
                EsRegalo = 0
                If FormatoD2(VSFGAplicar.TextMatrix(i, 8) * VSFGAplicar.TextMatrix(i, 9)) <= FormatoD2(VSFGAplicar.TextMatrix(i, 10)) Then
                    EsRegalo = 1
                End If
                
                strSql = " SELECT prd_pro_porcentaje " & _
                         " FROM producto_promo " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND prd_codigo='" & VSFGAplicar.TextMatrix(i, 6) & "' " & _
                         " AND CURRENT_TIMESTAMP BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                clsCon_Def.Ejecutar strSql
                dblDcto = VSFGAplicar.TextMatrix(i, 11)
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGAplicar.TextMatrix(i, 10) = FormatoD2(VSFGAplicar.TextMatrix(i, 8) * VSFGAplicar.TextMatrix(i, 9) * clsCon_Def.adorec_Def("prd_pro_porcentaje") / 100)
                    VSFGAplicar.TextMatrix(i, 11) = FormatoD2(clsCon_Def.adorec_Def("prd_pro_porcentaje"))
                Else
                    VSFGAplicar.TextMatrix(i, 10) = FormatoD2(0)
                    VSFGAplicar.TextMatrix(i, 11) = FormatoD2(0)
                End If
                
                strSql = " DELETE FROM det_pedido WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND dep_codigo='PRI' AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "' " & _
                         " AND prd_codigo='" & VSFGAplicar.TextMatrix(i, 6) & "'"
                clsCon_Def.Ejecutar strSql, "M"
                
                strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                         " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                         " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto" & _
                         " ) " & _
                         " VALUES ('" & strEmpresa & "','PRI',0," & _
                         " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 0) & "','" & VSFGAplicar.TextMatrix(i, 6) & "','" & VSFGAplicar.TextMatrix(i, 8) - maxPromo & "'," & _
                         " '" & VSFGAplicar.TextMatrix(i, 8) - maxPromo & "','" & VSFGAplicar.TextMatrix(i, 9) & "','" & FormatoD4(VSFGAplicar.TextMatrix(i, 10)) & "'" & _
                         " ) "
                clsCon_Def.Ejecutar strSql, "M"
                strSql = " UPDATE det_pedido " & _
                         " SET det_ped_cant_pedida=det_ped_cant_pedida+" & maxPromo & "," & _
                         " det_ped_cant_entregada=det_ped_cant_entregada+" & maxPromo & "," & _
                         " det_ped_precio=" & VSFGAplicar.TextMatrix(i, 9) & "," & _
                         " det_ped_dcto=det_ped_dcto+" & FormatoD2(maxPromo * VSFGAplicar.TextMatrix(i, 9) * dblDcto / 100) & " " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_codigo=" & VSFGAplicar.TextMatrix(i, 0) & " " & _
                         " AND prd_codigo='" & VSFGAplicar.TextMatrix(i, 6) & "' " & _
                         " AND dep_codigo='PRI' "
                clsCon_Def.Ejecutar (strSql), "M"
                clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
            Next i
        End If
    
    ElseIf optPromoComboPedido.Value = True Then
        For i = 1 To VSFGAplicar.Rows - 1
            If VSFGAplicar.TextMatrix(i, 9) > 0 Then
                For j = 1 To VSFGAplicar.TextMatrix(i, 9)
                    strSql = " SELECT p.emp_codigo,p.ped_codigo,dp.prd_codigo,dp.det_ped_cant_entregada,dp.det_ped_cant_programada " & _
                             " FROM pedido p INNER JOIN det_pedido dp ON p.emp_codigo=dp.emp_codigo " & _
                             " AND p.ped_codigo=dp.ped_codigo " & _
                             " INNER JOIN promo_combo_pedido pcp ON p.emp_codigo=pcp.emp_codigo " & _
                             " AND p.ped_fechamod BETWEEN pcp.pro_com_ped_fecha_desde AND pcp.pro_com_ped_fecha_hasta " & _
                             " INNER JOIN det_promo_combo_pedido_ent dpcpe ON pcp.emp_codigo=dpcpe.emp_codigo" & _
                             " AND pcp.tip_ped_codigo=dpcpe.tip_ped_codigo" & _
                             " AND pcp.pro_com_ped_codigo=dpcpe.pro_com_ped_codigo" & _
                             " AND dp.prd_codigo=dpcpe.prd_codigo " & _
                             " WHERE p.emp_codigo='" & strEmpresa & "' " & _
                             " AND p.ped_estado in (0,1) and dp.det_ped_cant_entregada>0 " & _
                             " AND p.ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'"
                    clsAux.Ejecutar strSql
                    If clsAux.adorec_Def("det_ped_cant_entregada") + clsAux.adorec_Def("det_ped_cant_programada") = 1 Then
                        strSql = " DELETE FROM det_pedido " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'" & _
                                 " AND prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
                        clsCon_Def.Ejecutar strSql, "M"
                        clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
                    ElseIf clsAux.adorec_Def("det_ped_cant_entregada") > 1 Then
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_cant_entregada=det_ped_cant_entregada-1, " & _
                                 " det_ped_cant_pedida=det_ped_cant_pedida-1 " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'" & _
                                 " AND prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
                        clsCon_Def.Ejecutar strSql, "M"
                        clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
                    ElseIf clsAux.adorec_Def("det_ped_cant_programada") > 1 Then
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_cant_programada=det_ped_cant_programada-1, " & _
                                 " det_ped_cant_pedida=det_ped_cant_pedida-1 " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo='" & VSFGAplicar.TextMatrix(i, 0) & "'" & _
                                 " AND prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
                        clsCon_Def.Ejecutar strSql, "M"
                        clsPedido.RecalculoTotal VSFGAplicar.TextMatrix(i, 0)
                    End If
                Next j
            End If
        Next i
    ElseIf optNPrendasAY.Value = True Then
        For i = 1 To VSFGAplicar.Rows - 1
            'PROMO PRENDA PRECIO'
            VSFGAplicar.ShowCell i, 0
            dblDcto = FormatoD2(PromoPrendaPrecio(cmbNegocioAplicar.BoundText, VSFGAplicar.TextMatrix(i, 0), "'" & Format(VSFGAplicar.TextMatrix(i, 1), "yyyyMMdd") & "'", False))
            If dblDcto > 0 Then
                'MsgBox dblDcto & " " & VSFGAplicar.TextMatrix(i, 0)
            End If
            dblDcto = 0
        Next i
    End If
    'MsgBox "Carga Finalizada"
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
        VSFGDctoCombo.LoadGrid txtArchivoDctoCombo.Text, flexFileTabText
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
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFGPromoCombo.TextMatrix(i, 4) = clsCon_Def.adorec_Def("prd_nombre")
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                Else
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                    VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                End If
            Else
                VSFGPromoCombo.TextMatrix(i, 4) = VSFGPromoCombo.TextMatrix(i - 1, 4)
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 6) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 6) <> VSFGPromoCombo.TextMatrix(i - 1, 6) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 6) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                        VSFGPromoCombo.TextMatrix(i, 7) = clsCon_Def.adorec_Def("prd_nombre")
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                    Else
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                        VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                    End If
                Else
                    VSFGPromoCombo.TextMatrix(i, 7) = VSFGPromoCombo.TextMatrix(i - 1, 7)
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
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
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                        VSFGPromoCombo.TextMatrix(i, 10) = clsCon_Def.adorec_Def("prd_nombre")
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                    Else
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                        VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                    End If
                Else
                    VSFGPromoCombo.TextMatrix(i, 10) = VSFGPromoCombo.TextMatrix(i - 1, 10)
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                End If
            End If
            
            If VSFGPromoCombo.TextMatrix(i, 12) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 12) <> VSFGPromoCombo.TextMatrix(i - 1, 12) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 12) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                        VSFGPromoCombo.TextMatrix(i, 13) = clsCon_Def.adorec_Def("prd_nombre")
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                    Else
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                        VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                    End If
                Else
                    VSFGPromoCombo.TextMatrix(i, 13) = VSFGPromoCombo.TextMatrix(i - 1, 13)
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                End If
            End If
            VSFGPromoCombo.ShowCell i, 13
            If VSFGPromoCombo.TextMatrix(i, 15) = "" Then
                VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) + 1
            Else
                If VSFGPromoCombo.TextMatrix(i, 15) <> VSFGPromoCombo.TextMatrix(i - 1, 15) Then
                    strSql = " SELECT prd_nombre " & _
                             " FROM producto " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND prd_codigo='" & VSFGPromoCombo.TextMatrix(i, 15) & "' "
                    
                    clsCon_Def.Ejecutar strSql
                    VSFGPromoCombo.ShowCell i, 16
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                        VSFGPromoCombo.TextMatrix(i, 16) = clsCon_Def.adorec_Def("prd_nombre")
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                    Else
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = "PRODUCTO NO ENCONTRADO - No se aplicara"
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                        VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                    End If
                Else
                    VSFGPromoCombo.TextMatrix(i, 16) = VSFGPromoCombo.TextMatrix(i - 1, 16)
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
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
                    If clsCon_DefP.adorec_Def.RecordCount > 0 Then
                        VSFGPromoCombo.TextMatrix(i, 1) = clsCon_DefP.adorec_Def("per_codigo")
                        VSFGPromoCombo.TextMatrix(i, 2) = clsCon_DefP.adorec_Def("cli")
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
                    Else
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) = VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                        VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = 0
                        VSFGPromoCombo.Cell(flexcpBackColor, i, 0, i, VSFGPromoCombo.Cols - 1) = vbRed
                    End If
                Else
                    VSFGPromoCombo.TextMatrix(i, 1) = VSFGPromoCombo.TextMatrix(i - 1, 1)
                    VSFGPromoCombo.TextMatrix(i, 2) = VSFGPromoCombo.TextMatrix(i - 1, 2)
                    VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1) = FormatoD0(VSFGPromoCombo.TextMatrix(i, VSFGPromoCombo.Cols - 1)) + 1
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

Public Sub CmdSalir_Click()
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
    
        strSql = " SELECT COALESCE(par_numero,0) as par_numero,COALESCE(par_texto,'') as par_texto,COALESCE(cta_nombre,'') as cta_nombre " & _
                 " FROM parametro INNER JOIN ctaconta ON parametro.emp_codigo=ctaconta.emp_codigo AND parametro.par_texto=ctaconta.cta_codigo " & _
                 " WHERE parametro.emp_codigo='" & strEmpresa & "' " & _
                 " AND par_codigo='IVAC'"
        clsCon_Def.Ejecutar strSql
        dblIVA = clsCon_Def.adorec_Def("par_numero")
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
