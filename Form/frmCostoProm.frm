VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCostoProm 
   BackColor       =   &H00DDDDDD&
   Caption         =   "Chequeo de Liquidación"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   Icon            =   "frmCostoProm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   12885
   Begin VB.TextBox txtTFob 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1995
      TabIndex        =   33
      Text            =   "0.0000"
      Top             =   6840
      Width           =   1260
   End
   Begin VB.TextBox txtFS 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1995
      TabIndex        =   32
      Text            =   "0.0000"
      Top             =   7200
      Width           =   1260
   End
   Begin VB.TextBox txtOG 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1995
      TabIndex        =   31
      Text            =   "0.0000"
      Top             =   7560
      Width           =   1260
   End
   Begin VB.TextBox txtImp 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1995
      TabIndex        =   30
      Text            =   "0.0000"
      Top             =   7920
      Width           =   1260
   End
   Begin VB.TextBox txtTotGastos 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10200
      TabIndex        =   29
      Top             =   840
      Width           =   1020
   End
   Begin VB.TextBox txtTotPed 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10200
      TabIndex        =   28
      Top             =   360
      Width           =   1020
   End
   Begin VB.TextBox txtAgeAfi 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   25
      Top             =   840
      Width           =   1770
   End
   Begin VB.TextBox txtEmb 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   24
      Top             =   360
      Width           =   1770
   End
   Begin VB.TextBox txtObser 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   6960
      Width           =   6255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox txtFechaIng 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   18
      Top             =   840
      Width           =   1170
   End
   Begin VB.TextBox txtCodCli 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3765
      TabIndex        =   10
      Top             =   2055
      Width           =   2220
   End
   Begin VB.TextBox txtFaxProveedor 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7815
      TabIndex        =   9
      Top             =   2715
      Width           =   2130
   End
   Begin VB.TextBox txtTelProveedor 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   8
      Top             =   2400
      Width           =   2130
   End
   Begin VB.TextBox txtRucProveedor 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7815
      TabIndex        =   7
      Top             =   2040
      Width           =   2130
   End
   Begin VB.TextBox txtDirProveedor 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3765
      TabIndex        =   6
      Top             =   2715
      Width           =   2205
   End
   Begin VB.TextBox txtNomP 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3765
      TabIndex        =   5
      Top             =   2385
      Width           =   2220
   End
   Begin MSDataListLib.DataCombo dcmbNumPed 
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfgTablaGC 
      Height          =   3255
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   11655
      _cx             =   20558
      _cy             =   5741
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCostoProm.frx":030A
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL FOB:"
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
      Left            =   1020
      TabIndex        =   37
      Top             =   6840
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL Flete/Seguro:"
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
      TabIndex        =   36
      Top             =   7200
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL OtrosGastos:"
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
      TabIndex        =   35
      Top             =   7560
      Width           =   1530
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL Imp:"
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
      Left            =   1095
      TabIndex        =   34
      Top             =   7920
      Width           =   840
   End
   Begin VB.Label lblTotGastos 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTALGASTOS:"
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
      Left            =   8760
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblTotPed 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PEDIDO:"
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
      Left            =   8760
      TabIndex        =   26
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblAgeAfi 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Ag. Afianzado:"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblEmbarcador 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Embarcador:"
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
      Left            =   5040
      TabIndex        =   22
      Top             =   480
      Width           =   915
   End
   Begin VB.Label lblObser 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
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
      Left            =   3480
      TabIndex        =   21
      Top             =   7080
      Width           =   1155
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de pedido:"
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
      TabIndex        =   19
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label lblProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del Proveedor"
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
      Left            =   5760
      TabIndex        =   17
      Top             =   1680
      Width           =   1680
   End
   Begin VB.Label lblFaxProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
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
      Left            =   6990
      TabIndex        =   16
      Top             =   2775
      Width           =   315
   End
   Begin VB.Label lblTelProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   6945
      TabIndex        =   15
      Top             =   2430
      Width           =   675
   End
   Begin VB.Label lblRucProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
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
      Left            =   6960
      TabIndex        =   14
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label lblDirProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
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
      Left            =   2760
      TabIndex        =   13
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lblNomProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   2445
      Width           =   600
   End
   Begin VB.Label lblCodProveedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
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
      Left            =   2760
      TabIndex        =   11
      Top             =   2115
      Width           =   540
   End
   Begin VB.Label lblNumPed 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Pedido:"
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
      TabIndex        =   0
      Top             =   480
      Width           =   1350
   End
End
Attribute VB_Name = "frmCostoProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para calcular el costo promedio de los productos recién ingresados    #
'#  a la empresa por Importaciones, visualizando la existencia y totales de     #
'#  ingreso, existencia y total final                                           #
'#                                                                              #
'#  frmCostoProm V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para calcular el costo promedio de los productos ingresados a la    #
'#  empresa por concepto de Importaciones.                                      #
'#  En esta ventana solo podemos escoger o ingresar valores en el combo         #
'#  de Numero de Pedido. Automaticamente se calculan y recuperan el resto       #
'#  de valores en la tabla o grid.                                              #
'#                                                                              #
'#  Se puede escoger el número de pedido o ingresar dicho número en el combo    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    persona    : En esta tabla se consulta los datos del proveedor al que se  #
'#                 le adquirió la mercadería y se importó.                      #
'#    det_ingreso_imp      : En esta tabla se consulta los detalles del ingreso.#
'#    producto   : En esta tabla se consulta el nombre del producto y en DONDE  #
'#                 se actulizará el nuevo costo del producto.                   #
'#    pedido_importacion   : En esta tabla se consulta las fechas y estado del  #
'#                           pedido de importacion.                             #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#    limpiarFxGD() : Permite borrar el flexgrid utilizado para cuando se       #
'#                    realiza un cambio de documento.                           #
'#                    Calcula automaticamente el contenido del resto de celdas  #
'#                    a partir del %peo y fob ingresados.                       #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'
Private clsCadaPrd As New clsConsulta
Private clsDivCant As New clsConsulta
Private clsIngImp_det As New clsConsulta
Private clsImp_det As New clsConsulta
Private clsConsu As New clsConsulta
Private clsExis As New clsConsulta

Private Sub cmdImprimir_Click()
    drptLiquidacionImp.Tag = dcmbNumPed.BoundText
    drptLiquidacionImp.Orientation = rptOrientLandscape
    drptLiquidacionImp.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCadaPrd = Nothing
    Set clsDivCant = Nothing
    Set clsIngImp_det = Nothing
    Set clsImp_det = Nothing
    Set clsConsu = Nothing
    Set clsExis = Nothing
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub dcmbNumPed_Change()

'Despliego los datos segú el dato ingresado o seleccioado en el data combo
'clsConsu.adorec_Def.MoveFirst
If dcmbNumPed.Text = "" Then
Exit Sub
End If

 If (CSng(dcmbNumPed.BoundText) > 99999) Then
        Call limpiarFxGD
        txtCodCli.Text = ""
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtFechaIng.Text = ""
        TxtObser.Text = ""
        txtTotGastos.Text = ""
        txtAgeAfi.Text = ""
        txtEmb.Text = ""
        txtTotPed.Text = ""
        Exit Sub
    End If
    
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "ped_imp_codigo = '" & dcmbNumPed.BoundText & "'", , adSearchForward
    
    
    If clsConsu.adorec_Def.EOF = False Then
        'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
        txtCodCli.Text = clsConsu.adorec_Def("per_codigo")
        txtNomP.Text = clsConsu.adorec_Def("per_apellido") & " " & clsConsu.adorec_Def("per_nombre")
        txtFechaIng.Text = Format(clsConsu.adorec_Def("ped_imp_fecha_ped"), "yyyy-mmm-dd")
        txtRucProveedor.Text = clsConsu.adorec_Def("per_ruc")
        txtDirProveedor.Text = clsConsu.adorec_Def("per_direccion")
        txtTelProveedor.Text = clsConsu.adorec_Def("per_telf")
        txtFaxProveedor.Text = clsConsu.adorec_Def("per_fax")
        txtEmb.Text = clsConsu.adorec_Def("emb_nombre")
        txtAgeAfi.Text = clsConsu.adorec_Def("age_afi_nombre")
        txtTotPed.Text = clsConsu.adorec_Def("ped_imp_total_pedido")
        txtTotGastos.Text = clsConsu.adorec_Def("ped_imp_total_gastos")
        TxtObser.Text = clsConsu.adorec_Def("ped_imp_observacion")
        Call limpiarFxGD
        
        'consulta para llenar flexgrid
        strSQL = " select b.prd_codigo, a.prd_nombre,SUM(b.det_ing_imp_cantidad) ," & _
                " AVG(b.det_ing_imp_fob),SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_fob) " & _
                " , SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_cif) - SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_fob)" & _
                " , AVG(b.det_ing_imp_cif),SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_cif) " & _
                " , SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_costofinal) - SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_cif)" & _
                " , AVG(b.det_ing_imp_costofinal),SUM(b.det_ing_imp_cantidad)*AVG(b.det_ing_imp_costofinal) " & _
                " , (AVG(b.det_ing_imp_costofinal)-AVG(b.det_ing_imp_fob))/AVG(b.det_ing_imp_fob) * 100.00" & _
                " from det_ingreso_imp b INNER JOIN producto a ON a.prd_codigo = b.prd_codigo and a.emp_codigo=b.emp_codigo " & _
                " where b.emp_codigo = '" & strEmpresa & "'" & _
                " and b.tip_ing_codigo = 'IIM' " & _
                " and b.ped_imp_codigo = " & dcmbNumPed.BoundText & _
                " GROUP BY prd_codigo" & _
                " ORDER BY prd_codigo"
        clsImp_det.Ejecutar (strSQL)
        ' llena el flex grid
        If (clsImp_det.adorec_Def.RecordCount > 0) Then
            Set vsfgTablaGC.DataSource = clsImp_det.adorec_Def.DataSource
        End If
        Dim i As Long
        txtTFob = Val(0#)
        txtFS = Val(0#)
        txtOG = Val(0#)
        txtImp = Val(0#)
        For i = 1 To vsfgTablaGC.Rows - 1
            txtTFob = Val(txtTFob) + Val(vsfgTablaGC.TextMatrix(i, 5))
            txtFS = Val(txtFS) + Val(vsfgTablaGC.TextMatrix(i, 6))
            txtOG = Val(txtOG) + Val(vsfgTablaGC.TextMatrix(i, 9))
            txtImp = Val(txtImp) + Val(vsfgTablaGC.TextMatrix(i, 11))
        Next i
        
    Else
        Call limpiarFxGD
        txtCodCli.Text = ""
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtFechaIng.Text = ""
        txtEmb.Text = ""
        txtAgeAfi.Text = ""
        txtTotPed.Text = ""
        txtTotGastos.Text = ""
        TxtObser.Text = ""
               
    End If
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub



Private Sub dcmbNumPed_KeyPress(KeyAscii As Integer)
'controla que no se ingrese caracteres no permitidos como letras,etc.
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
 'Centra esta forma dentro de la forma MDI
    Me.Width = 13005
    Me.Height = 9255
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsImp_det.Inicializar AdoConn, AdoConnMaster
    clsExis.Inicializar AdoConn, AdoConnMaster
    
    
    'Ejecuta un SQL contra la base de datos
    strSQL = "select CONCAT(a.ped_imp_codigo) AS ped_imp_codigo, c.ped_imp_fecha_ped," & _
                " c.ped_imp_total_pedido,c.ped_imp_total_gastos," & _
                " c.ped_imp_observacion," & _
                " e.emb_nombre," & _
                " d.age_afi_nombre," & _
                " b.per_codigo, b.per_nombre," & _
                " b.per_apellido, b.per_ruc, b.per_direccion," & _
                " b.per_telf, b.per_fax,CONCAT(c.ped_imp_numero,' (',c.ped_imp_codigo,')') as ped_imp_numero from det_ingreso_imp a, persona b, pedido_importacion c," & _
                " embarcador e, agente_afianzado d" & _
                " where  c.per_codigo = b.per_codigo and c.ped_imp_estado = 'LI'" & _
                " and a.emp_codigo = '" & strEmpresa & "' " & _
                " and a.ped_imp_codigo = c.ped_imp_codigo" & _
                " and  a.emp_codigo = c.emp_codigo " & _
                " and d.age_afi_codigo=c.age_afi_codigo " & _
                " and e.emb_codigo=c.emb_codigo " & _
                " group by a.ped_imp_codigo " & _
                " order by a.ped_imp_codigo "

    clsConsu.Ejecutar (strSQL)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Importaciones para calcular Costo Promedio en el sistema.", vbInformation, "SisAdmi"
        'cmdAceptar.Enabled = False
        dcmbNumPed.Enabled = False
        
        Exit Sub
    Else
        Set dcmbNumPed.RowSource = clsConsu.adorec_Def.DataSource
        dcmbNumPed.BoundColumn = "ped_imp_codigo"
        dcmbNumPed.ListField = "ped_imp_numero"
    End If
    
     ' initializa el flexgrid
    'vsfgTablaGC.Editable = flexEDKbdMouse
    'vsfgTablaGC.AllowUserResizing = flexResizeBoth

errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
       
        End Select
End Sub

Private Sub Form_Activate()
clsConsu.Actualizar
  If (clsConsu.adorec_Def.RecordCount <> 0) Then
        Set dcmbNumPed.RowSource = clsConsu.adorec_Def.DataSource
        dcmbNumPed.BoundColumn = "ped_imp_codigo"
        dcmbNumPed.ListField = "ped_imp_numero"
    End If

End Sub


Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Double
    vsfgTablaGC.Tag = "N"
    'vsfgDetalleImp.Rows = 2
    
    
'    For X = 1 To vsfgDetalleImp.Rows - 1
'       For Y = 1 To vsfgDetalleImp.Cols - 1
'           vsfgDetalleImp.TextMatrix(X, Y) = ""
'        Next Y
'    Next X
    vsfgTablaGC.Rows = 1
    vsfgTablaGC.Clear 1
    vsfgTablaGC.Tag = "T"
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub





