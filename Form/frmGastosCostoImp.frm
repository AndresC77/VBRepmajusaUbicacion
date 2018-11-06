VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGastosCostoImp 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Costos a productos y Distribución de Gastos"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12255
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
      Left            =   1680
      TabIndex        =   37
      Top             =   7080
      Width           =   1260
   End
   Begin VB.TextBox txtAran 
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
      Left            =   1680
      TabIndex        =   35
      Top             =   6720
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
      Left            =   1680
      TabIndex        =   33
      Top             =   6360
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
      Left            =   1680
      TabIndex        =   31
      Top             =   6000
      Width           =   1260
   End
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
      Left            =   1680
      TabIndex        =   29
      Top             =   5640
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
      Left            =   10361
      TabIndex        =   17
      Top             =   555
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
      Left            =   10361
      TabIndex        =   28
      Top             =   188
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
      Left            =   6641
      TabIndex        =   25
      Top             =   555
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
      Left            =   6641
      TabIndex        =   24
      Top             =   188
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
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   6120
      Width           =   6255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6180
      TabIndex        =   4
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      Top             =   7080
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
      Left            =   2314
      TabIndex        =   18
      Top             =   555
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
      Left            =   3540
      TabIndex        =   10
      Top             =   1215
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
      Left            =   7590
      TabIndex        =   9
      Top             =   1875
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
      Left            =   7590
      TabIndex        =   8
      Top             =   1545
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
      Left            =   7590
      TabIndex        =   7
      Top             =   1200
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
      Left            =   3540
      TabIndex        =   6
      Top             =   1875
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
      Left            =   3540
      TabIndex        =   5
      Top             =   1545
      Width           =   2220
   End
   Begin MSDataListLib.DataCombo dcmbNumPed 
      Height          =   330
      Left            =   2314
      TabIndex        =   1
      Top             =   195
      Width           =   2655
      _ExtentX        =   4683
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
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   12015
      _cx             =   1991463625
      _cy             =   1991448173
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGastosCostoImp.frx":0000
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
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
      Left            =   840
      TabIndex        =   38
      Top             =   7080
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL Aran:"
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
      Left            =   720
      TabIndex        =   36
      Top             =   6720
      Width           =   960
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
      Left            =   150
      TabIndex        =   34
      Top             =   6360
      Width           =   1530
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
      Left            =   165
      TabIndex        =   32
      Top             =   6000
      Width           =   1515
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
      Left            =   765
      TabIndex        =   30
      Top             =   5640
      Width           =   915
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
      Left            =   8914
      TabIndex        =   27
      Top             =   600
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
      Left            =   8914
      TabIndex        =   26
      Top             =   240
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
      Left            =   5194
      TabIndex        =   23
      Top             =   600
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
      Left            =   5194
      TabIndex        =   22
      Top             =   240
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
      Left            =   3120
      TabIndex        =   21
      Top             =   6120
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
      Left            =   874
      TabIndex        =   19
      Top             =   600
      Width           =   1245
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
      Left            =   6765
      TabIndex        =   16
      Top             =   1935
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
      Left            =   6720
      TabIndex        =   15
      Top             =   1590
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
      Left            =   6735
      TabIndex        =   14
      Top             =   1320
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
      Left            =   2535
      TabIndex        =   13
      Top             =   1920
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
      Left            =   2535
      TabIndex        =   12
      Top             =   1605
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
      Left            =   2535
      TabIndex        =   11
      Top             =   1275
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
      Left            =   874
      TabIndex        =   0
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmGastosCostoImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para visualizar los ingresos de mercadería realizados por concepto de #
'#  Importaciones,  y que permite asignar los costos correspondientes           #
'#  y la distribución de los gastos.                                            #
'#                                                                              #
'#  frmGastosCostoImp V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los ingresos de mercadería a una determinada  emp-   #
'#  presa por concepto de Importaciones.                                        #
'#  En esta ventana solo podemos ingresar valores en los campos de %peso y      #
'#  fob. Automaticamente se calculan el resto de valores en la tabla o grid     #
'#  como total de fob y cif, total de flete y seguro, y otros gastos; un costo  #
'#  unitario final y costs finales totales, etc.                                #
'#                                                                              #
'#  Se puede escoger el número de documento o ingresar dicho número en el combo #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    persona    : En esta tabla se consulta los datos del proveedor al que se  #
'#                 le adquirió la mercadería y se importó.                      #
'#    det_ingreso_imp      : En esta tabla se consulta los detalles del ingreso.#
'#    producto   : En esta tabla se consulta el nombre del producto.            #
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
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'
Private clsCadaPrd As New clsConsulta
Private clsDivCant As New clsConsulta
Private clsIngImp_det As New clsConsulta
Private clsImp_det As New clsConsulta
Private clsConsu As New clsConsulta
Private clsCon_det As New clsConsulta
Private clscon As New clsConsulta
Private clsIng As New clsConsulta
Private clsEgr As New clsConsulta
Private PesoT As Double
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
    Set clsCon_det = Nothing
    Set clscon = Nothing
End Sub

Private Sub cmdAceptar_Click()
    clsIngImp_det.Inicializar AdoConn, AdoConnMaster
    clsDivCant.Inicializar AdoConn, AdoConnMaster
    clsCadaPrd.Inicializar AdoConn, AdoConnMaster
    clsIng.Inicializar AdoConn, AdoConnMaster
    clsEgr.Inicializar AdoConn, AdoConnMaster
    Dim PPP As Double
    Dim saldo As Long
    Dim SAux As Long
    Dim strSQL As String
    
    'Dim prd As String, str As String, str2 As String, strSql As String
    
    'consulta para ver los productos iguales pero con distinto codigo de ingreso
    strSQL = " select b.prd_codigo, b.ing_codigo," & _
            " count(b.ing_codigo) as c" & _
            " from det_ingreso_imp b where" & _
            " b.emp_codigo = '" & strEmpresa & "'" & _
            " and b.tip_ing_codigo = 'IIM' and b.det_ing_imp_cantidad!=0" & _
            " and b.ped_imp_codigo ='" & dcmbNumPed.BoundText & _
            "' GROUP BY b.prd_codigo, b.ing_codigo" & _
            " ORDER BY prd_codigo"
    
    'InputBox "-", "-", strSql
    
    clsDivCant.Ejecutar strSQL
    
    
    'If clsDivCant.adorec_Def.EOF = True Then
    ' MsgBox "Consulta vacía"
    'Else
    ' MsgBox "Consulta NO vacía"
    'End If
    
    clsDivCant.adorec_Def.MoveFirst
    
    For i = 1 To vsfgTablaGC.Rows - 1
            prd = clsDivCant.adorec_Def("prd_codigo")
            ' para separar la cantidad correspondiente a cada uno
            If clsDivCant.adorec_Def("c") > 1 Then
                strSQL = " select a.prd_codigo,a.ing_codigo, a.det_ing_imp_cantidad" & _
                        " from det_ingreso_imp a" & _
                        " where a.emp_codigo = '" & strEmpresa & "'" & _
                        " and a.tip_ing_codigo = 'IIM'" & _
                        " and a.ped_imp_codigo = " & dcmbNumPed.BoundText & _
                        " and a.prd_codigo='" & prd & "'"
                        
                clsCadaPrd.Ejecutar strSQL
                clsCadaPrd.adorec_Def.MoveFirst
                For c = 1 To clsDivCant.adorec_Def("c")
                     strSQL = " UPDATE det_ingreso_imp SET det_ing_imp_fob=" & _
                           Val(Format(vsfgTablaGC.TextMatrix(i, 4))) & ", " & _
                            " det_ing_imp_cif=" & _
                           Val(Format(vsfgTablaGC.TextMatrix(i, 7), "###0.0000")) & ", " & _
                            " det_ing_imp_arancel= " & _
                           Val(Format(vsfgTablaGC.TextMatrix(i, 10), "###0.0000")) & ", " & _
                            " det_ing_imp_costofinal= " & _
                           Val(Format(vsfgTablaGC.TextMatrix(i, 12), "###0.0000")) & _
                            " WHERE emp_codigo='" & strEmpresa & "' AND ped_imp_codigo=" & dcmbNumPed.BoundText & " AND " & _
                            " prd_codigo='" & prd & "' AND ing_codigo=" & clsCadaPrd.adorec_Def("ing_codigo") & " AND tip_ing_codigo='IIM'"
                    clsIngImp_det.Ejecutar strSQL, "M"
                    
                    strSQL = " UPDATE det_ingreso SET det_ing_precio=" & _
                            Val(Format(vsfgTablaGC.TextMatrix(i, 12), "###0.0000")) & _
                            ", det_ing_costo=" & PPP & _
                            " WHERE emp_codigo='" & strEmpresa & "' AND " & _
                            " prd_codigo='" & prd & "' AND ing_codigo=" & clsCadaPrd.adorec_Def("ing_codigo") & " AND tip_ing_codigo='IIM'"

                    clsIngImp_det.Ejecutar strSQL, "M"
                    
                    clsCadaPrd.adorec_Def.MoveNext
                Next c
            Else ' si no esta el producto repetido
                strSQL = " UPDATE det_ingreso_imp SET det_ing_imp_fob=" & _
                      Val(Format(vsfgTablaGC.TextMatrix(i, 4), "###0.0000")) & ", " & _
                       " det_ing_imp_cif=" & _
                      Val(Format(vsfgTablaGC.TextMatrix(i, 7), "####0.0000")) & ", " & _
                       " det_ing_imp_arancel= " & _
                      Val(Format(vsfgTablaGC.TextMatrix(i, 10), "####0.0000")) & ", " & _
                       " det_ing_imp_costofinal= " & _
                      Val(Format(vsfgTablaGC.TextMatrix(i, 12), "####0.0000")) & _
                       " WHERE emp_codigo='" & strEmpresa & "' AND ped_imp_codigo=" & dcmbNumPed.BoundText & " AND " & _
                       " prd_codigo='" & prd & "' AND ing_codigo=" & clsDivCant.adorec_Def("ing_codigo") & " AND tip_ing_codigo='IIM'"
                       
                 clsIngImp_det.Ejecutar strSQL, "M"
                 
                 strSQL = " UPDATE det_ingreso SET det_ing_precio=" & _
                      Val(Format(vsfgTablaGC.TextMatrix(i, 12), "####0.0000")) & _
                       ",det_ing_costo=" & PPP & _
                       " WHERE emp_codigo='" & strEmpresa & "' AND " & _
                       " prd_codigo='" & prd & "' AND ing_codigo=" & clsDivCant.adorec_Def("ing_codigo") & " AND tip_ing_codigo='IIM'"

                 clsIngImp_det.Ejecutar strSQL, "M"
             
             
             End If
             'Actualiza lista de precio del proveedor
             strSQL = " DELETE FROM persona_producto " & _
                      " WHERE emp_codigo='" & strEmpresa & "' " & _
                      " AND per_codigo='" & txtCodCli.Text & "' " & _
                      " AND prd_codigo='" & prd & "'"
             'clsIngImp_det.Ejecutar strSql, "M"
             strSQL = " INSERT INTO persona_producto " & _
                      " SELECT '" & txtCodCli.Text & "','" & strEmpresa & "',prd_codigo,det_ing_precio,det_ing_fechamod,det_ing_usumod " & _
                      " FROM ingreso INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND det_ingreso.prd_codigo='" & prd & "'" & _
                      " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                      " AND per_codigo='" & txtCodCli.Text & "' " & _
                      " AND ingreso.tip_ing_codigo='IIM' " & _
                      " ORDER BY ing_fecha DESC LIMIT 0,1"
            'clsIngImp_det.Ejecutar strSql, "M"
            clsDivCant.adorec_Def.MoveNext
            'clsImp_det.adorec_Def.MoveNext
    Next i
'**************************************************************
    strSQL = " select prd_codigo, ingreso.ing_codigo, ingreso.tip_ing_codigo,ingreso.ing_fecha " & _
            " from det_ingreso_imp INNER JOIN ingreso ON det_ingreso_imp.emp_codigo=ingreso.emp_codigo AND det_ingreso_imp.ing_codigo=ingreso.ing_codigo AND det_ingreso_imp.tip_ing_codigo=ingreso.tip_ing_codigo " & _
            " where det_ingreso_imp.emp_codigo = '" & strEmpresa & "'" & _
            " and det_ingreso_imp.tip_ing_codigo = 'IIM'" & _
            " and det_ingreso_imp.ped_imp_codigo ='" & dcmbNumPed.BoundText & _
            "' " & _
            " ORDER BY ingreso.ing_fecha,prd_codigo "
    clsDivCant.Ejecutar strSQL
    Dim clsCos As New clsCostear
    clsCos.Inicializar AdoConn, AdoConnMaster
    While Not clsDivCant.adorec_Def.EOF
        clsCos.Recostear Format(clsDivCant.adorec_Def("ing_fecha"), "yyyy-mm") & "-01", Format(clsDivCant.adorec_Def("ing_fecha"), "yyyy") & "-12-31", clsDivCant.adorec_Def("prd_codigo")
        clsDivCant.adorec_Def.MoveNext
    Wend
'**************************************************************
    str2 = " UPDATE pedido_importacion SET ped_imp_estado='LI' " & _
           " WHERE emp_codigo='" & strEmpresa & "' AND ped_imp_codigo=" & dcmbNumPed.BoundText
           
    clsIngImp_det.Ejecutar str2, "M"
    
    MsgBox "Ingresado"
    Unload Me
    
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub dcmbNumPed_Change()
'On Error GoTo errhandler
'Despliego los datos segú el dato ingresado o seleccioado en el data combo
'clsConsu.adorec_Def.MoveFirst
clscon.Inicializar AdoConn, AdoConnMaster
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
        If Val(txtTotGastos.Text) <= 0 Then
            MsgBox "Esta importacion no tiene asignado gastos", , "Importaciones"
        End If
        TxtObser.Text = clsConsu.adorec_Def("ped_imp_observacion")
        Call limpiarFxGD
        
        'llenar flexgrid
        strSQL = " select a.prd_codigo, a.prd_nombre," & _
                " sum(b.det_ing_imp_cantidad) As totalCant, b.ing_codigo" & _
                " from det_ingreso_imp b INNER JOIN producto a " & _
                " ON a.prd_codigo = b.prd_codigo" & _
                " AND a.emp_codigo=b.emp_codigo " & _
                " where a.emp_codigo = '" & strEmpresa & "'" & _
                " AND b.tip_ing_codigo = 'IIM' " & _
                " AND b.ped_imp_codigo = " & dcmbNumPed.BoundText & _
                " GROUP BY a.prd_codigo,a.prd_nombre, b.ing_codigo" & _
                " HAVING sum(b.det_ing_imp_cantidad)<>0 " & _
                " ORDER BY a.prd_codigo"
        
        clsImp_det.Ejecutar strSQL
        If (clsImp_det.adorec_Def.RecordCount > 0) Then
            Dim i As Long
            clsImp_det.adorec_Def.MoveFirst
            For i = 1 To clsImp_det.adorec_Def.RecordCount
                strSQL = " select COALESCE(det_ped_imp_precio,0) AS det_ped_imp_precio " & _
                " from det_pedido_imp where" & _
                " prd_codigo='" & clsImp_det.adorec_Def("prd_codigo") & "'" & _
                " and emp_codigo = '" & strEmpresa & "'" & _
                " and ped_imp_codigo = " & dcmbNumPed.BoundText
                clscon.Ejecutar strSQL
                vsfgTablaGC.AddItem ""
                vsfgTablaGC.TextMatrix(i, 1) = clsImp_det.adorec_Def("prd_codigo")
                vsfgTablaGC.TextMatrix(i, 2) = 0
                vsfgTablaGC.TextMatrix(i, 3) = clsImp_det.adorec_Def("totalCant")
                If clscon.adorec_Def.RecordCount > 0 Then
                    vsfgTablaGC.TextMatrix(i, 4) = clscon.adorec_Def("det_ped_imp_precio")
                Else
                    vsfgTablaGC.TextMatrix(i, 4) = 0
                End If
                vsfgTablaGC.TextMatrix(i, 10) = 0
                clsImp_det.adorec_Def.MoveNext
            Next i
        End If
        
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
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
 'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 200
    
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    clsImp_det.Inicializar AdoConn, AdoConnMaster
    
    'Ejecuta un SQL contra la base de datos
    strSQL = " SELECT c.ped_imp_codigo, c.ped_imp_fecha_ped, c.ped_imp_total_pedido," & _
                " COALESCE(c.ped_imp_total_gastos,0) as ped_imp_total_gastos, c.ped_imp_observacion, COALESCE(e.emb_nombre,'-') AS emb_nombre, COALESCE(d.age_afi_nombre,'-') AS age_afi_nombre, b.per_codigo, b.per_nombre," & _
                " b.per_apellido, b.per_ruc, b.per_direccion, b.per_telf, b.per_fax,CONCAT(c.ped_imp_numero,' (',c.ped_imp_codigo,')') as ped_imp_numero" & _
                " FROM pedido_importacion c " & _
                " INNER JOIN persona b ON c.per_codigo = b.per_codigo AND c.emp_codigo = b.emp_codigo " & _
                " LEFT JOIN embarcador e ON e.emb_codigo=c.emb_codigo " & _
                " LEFT JOIN agente_afianzado d ON d.age_afi_codigo=c.age_afi_codigo " & _
                " where  c.ped_imp_estado = 'BO'" & _
                " and c.emp_codigo = '" & strEmpresa & "' " & _
                " order by c.ped_imp_codigo "

    clsConsu.Ejecutar strSQL
    
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Importaciones para calcular costos y distribuir gastos en el sistema.", vbInformation, "SisAdmi"
        cmdAceptar.Enabled = False
        dcmbNumPed.Enabled = False
        Exit Sub
    Else
        Set dcmbNumPed.RowSource = clsConsu.adorec_Def.DataSource
        dcmbNumPed.ListField = "ped_imp_numero"
        dcmbNumPed.BoundColumn = "ped_imp_codigo"
    End If
    
     ' initializa el flexgrid
    vsfgTablaGC.Editable = flexEDKbdMouse
    vsfgTablaGC.AllowUserResizing = flexResizeBoth

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
        dcmbNumPed.ListField = "ped_imp_numero"
        dcmbNumPed.BoundColumn = "ped_imp_codigo"
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

Private Sub vsfgTablaGC_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'para que no se pueda escribir en las columnas que se indica
If NewCol = 1 Then
                SendKeys vbKeyTab
End If

If NewCol = 3 Or NewCol = 5 Or NewCol = 6 Or NewCol = 7 Or NewCol = 8 Or NewCol = 9 Or NewCol = 11 Or NewCol = 12 Or NewCol = 13 Then
         If Abs(NewCol - OldCol) = 1 Then
            If NewCol > OldCol Then
                SendKeys vbKeyTab
            Else
                SendKeys vbKeyLeft
            End If
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub vsfgTablaGC_CellChanged(ByVal Row As Long, ByVal Col As Long)

If (Col = 4 And vsfgTablaGC.TextMatrix(Row, Col) <> "") Then
    vsfgTablaGC.TextMatrix(Row, Col + 1) = CStr(Val(vsfgTablaGC.TextMatrix(Row, Col)) * Val(vsfgTablaGC.TextMatrix(Row, Col - 1)))
End If
'MsgBox Row                 3
'MsgBox vsfgTablaGC.Rows    4
If Col = 10 Then
    vsfgTablaGC.TextMatrix(Row, 11) = Val(Format(vsfgTablaGC.TextMatrix(Row, 10), "###0.0000")) * vsfgTablaGC.TextMatrix(Row, 8) / 100#
End If
'para ver si estan todas las celdas de PESO y FOB llenas
Dim i As Long, llenado As Long, PesoT As Double, Tfob As Double, TPCIF As Double, TOG As Double, TAran As Double
llenado = 1: PesoT = 0: Tfob = 0: TPCIF = 0: TOG = 0: TAran = 0

For i = 1 To vsfgTablaGC.Rows - 1
    If (vsfgTablaGC.TextMatrix(i, 4) <> "" And vsfgTablaGC.TextMatrix(i, 2) <> "") Then
        llenado = llenado + 1
        PesoT = PesoT + vsfgTablaGC.TextMatrix(i, 2)
    
    End If
    Tfob = Tfob + Val(vsfgTablaGC.TextMatrix(i, 5))
    TPCIF = TPCIF + Val(vsfgTablaGC.TextMatrix(i, 6))
    TOG = TOG + Val(vsfgTablaGC.TextMatrix(i, 9))
    TAran = TAran + Val(vsfgTablaGC.TextMatrix(i, 11))
Next i
txtTFob.Text = Format(Tfob, "###0.00")
txtFS.Text = Format(TPCIF, "###0.00")
txtOG.Text = Format(TOG, "###0.00")
txtAran.Text = Format(TAran, "###0.00")
txtImp.Text = Format(Tfob + TPCIF + TOG + TAran, "###0.00")

If llenado = i Then ' si todos los datos en peso y fob han sido INGRESADOS

'verifica que la suma de los porcentajes de PESO sea 100.00
'If PesoT <> 100 Then
'MsgBox "La suma de los porcentajes de peso no es 100 (" & PesoT & ")"
'cmdAceptar.Enabled = False
'Exit Sub
'Else
cmdAceptar.Enabled = True
'End If

'MsgBox PesoT & " " & Tfob

    strSQL = " SELECT d.gas_imp_codigo,d.det_gas_imp_valor," & _
            " g.gas_imp_calcula_a,g.gas_imp_prorra_a," & _
            " g.gas_imp_tipo " & _
            " FROM det_gasto_imp d INNER JOIN gasto_importacion g" & _
            " ON ( d.gas_imp_codigo=g.gas_imp_codigo" & _
            " AND d.emp_codigo='" & strEmpresa & "')" & _
            " WHERE d.ped_imp_codigo =" & clsConsu.adorec_Def("ped_imp_codigo") & _
            " ORDER BY g.gas_imp_tipo DESC "
    clsCon_det.Ejecutar strSQL
    
For fila = 1 To vsfgTablaGC.Rows - 1  'para llenar fila por fila
    Dim TotalParaCIF As Double ' para contabilizar los gastos de flete y seguro
    TotalParaCIF = 0:
    
    'obtenemos datos para calcular los gastos y costos
    If clsCon_det.adorec_Def.RecordCount > 0 Then
    clsCon_det.adorec_Def.MoveFirst
    End If
    For i = 1 To clsCon_det.adorec_Def.RecordCount
            'si es FLETE o SEGURO
            If clsCon_det.adorec_Def("gas_imp_tipo") = "S" Or clsCon_det.adorec_Def("gas_imp_tipo") = "F" Then
                ' si prorra al Peso
                If clsCon_det.adorec_Def("gas_imp_prorra_a") = "P" Then
                    TotalParaCIF = TotalParaCIF + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 2)) / PesoT
                    'MsgBox TotalParaCIF & " " & Val(clsCon_det.adorec_Def("det_gas_imp_valor")) & " P"
                Else ' entonces prorra al COSTO "C"
                If Tfob > 0 Then
                    TotalParaCIF = TotalParaCIF + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 5)) / Tfob
                End If
                    'MsgBox TotalParaCIF & " " & Val(clsCon_det.adorec_Def("det_gas_imp_valor")) & " C"
                End If
            End If
        clsCon_det.adorec_Def.MoveNext
    Next i

vsfgTablaGC.TextMatrix(fila, 6) = CStr(TotalParaCIF)
vsfgTablaGC.TextMatrix(fila, 7) = CStr(Val(vsfgTablaGC.TextMatrix(fila, 4)) + TotalParaCIF / vsfgTablaGC.TextMatrix(fila, 3))
vsfgTablaGC.TextMatrix(fila, 8) = CStr(Val(vsfgTablaGC.TextMatrix(fila, 7)) * vsfgTablaGC.TextMatrix(fila, 3))

Next fila

Dim TotalCIF As Double
TotalCIF = 0
For fila = 1 To vsfgTablaGC.Rows - 1
TotalCIF = TotalCIF + vsfgTablaGC.TextMatrix(fila, 8)
Next fila

For fila = 1 To vsfgTablaGC.Rows - 1  'para llenar fila por fila
    Dim TotalOtros As Double ' para contabilizar los Otros gastos
    TotalOtros = 0
    
    'obtenemos datos para calcular los gastos y costos
    If clsCon_det.adorec_Def.RecordCount > 0 Then
    clsCon_det.adorec_Def.MoveFirst
    End If
    
    For i = 1 To clsCon_det.adorec_Def.RecordCount
            ' si es OTRO gasto
            If clsCon_det.adorec_Def("gas_imp_tipo") = "-" Then
                ' si prorra al Peso
                If clsCon_det.adorec_Def("gas_imp_prorra_a") = "P" Then
                    'si calcula al cif
                    If clsCon_det.adorec_Def("gas_imp_calcula_a") = "C" Then
                        TotalOtros = TotalOtros + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 2)) / PesoT
                    Else ' al fob
                        TotalOtros = TotalOtros + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 2)) / PesoT
                    End If
                ' entonces prorra al COSTO "C"
                Else
                    
                    'si calcula al cif
                    If clsCon_det.adorec_Def("gas_imp_calcula_a") = "C" Then
                        If TotalCIF > 0 Then
                        TotalOtros = TotalOtros + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 8)) / TotalCIF
                        End If
                    Else ' al fob
                        If Tfob > 0 Then
                        TotalOtros = TotalOtros + Val(clsCon_det.adorec_Def("det_gas_imp_valor")) * Val(vsfgTablaGC.TextMatrix(fila, 5)) / Tfob
                        End If
                    End If
                    
                End If
            End If

        clsCon_det.adorec_Def.MoveNext
    Next i

vsfgTablaGC.TextMatrix(fila, 9) = CStr(TotalOtros)
vsfgTablaGC.TextMatrix(fila, 12) = CStr(Val(vsfgTablaGC.TextMatrix(fila, 7)) + (TotalOtros + Val(Format(vsfgTablaGC.TextMatrix(fila, 11), "###0.0000"))) / vsfgTablaGC.TextMatrix(fila, 3))
vsfgTablaGC.TextMatrix(fila, 13) = CStr(Val(vsfgTablaGC.TextMatrix(fila, 12)) * vsfgTablaGC.TextMatrix(fila, 3))
If Val(vsfgTablaGC.TextMatrix(fila, 4)) <> 0 Then
    vsfgTablaGC.TextMatrix(fila, 14) = (Val(vsfgTablaGC.TextMatrix(fila, 12)) - Val(vsfgTablaGC.TextMatrix(fila, 4))) / Val(vsfgTablaGC.TextMatrix(fila, 4))
End If

Next fila


End If ' end if llenado = i

Exit Sub

errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
       
        End Select

End Sub



Private Sub vsfgTablaGC_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
   

   'Valido que solo se pueda dar enter en los campos que se indica
 If (Col = 1) Or (Col = 3) Or (Col = 5) Or Col = 6 Or Col = 7 Or Col = 8 Or Col = 9 Or Col = 11 Or Col = 12 Or (Col = 13) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
   End If
    
    
  'Valido que solo se pueda ingresar números  en el campo Peso
 If Col = 2 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> 46) Then
            KeyAscii = 0
        End If
    End If
    
'Valido que solo se pueda ingresar números  en el campo FOB
 If Col = 4 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> 46) Then
            KeyAscii = 0
        End If
End If
End Sub





