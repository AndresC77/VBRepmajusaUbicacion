VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRealizaDevolucionProducto 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar Devolucion de Productos"
   ClientHeight    =   8460
   ClientLeft      =   3330
   ClientTop       =   2010
   ClientWidth     =   12600
   Icon            =   "frmRealizaDevolucionProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdReImp 
      Caption         =   "&Imprimir Formulario"
      Height          =   375
      Left            =   7200
      TabIndex        =   47
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Extras del Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   9975
      Begin VB.TextBox txtAutorizacion 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   990
         Width           =   1815
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNumAux 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Top             =   615
         Width           =   1815
      End
      Begin VB.TextBox txtCaduca 
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1060
      End
      Begin VB.TextBox txtDcto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpCaduca 
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
         Left            =   8400
         TabIndex        =   11
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "MM/yyy"
         Format          =   66387971
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   7935
         TabIndex        =   9
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
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
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   1695
         _extentx        =   2990
         _extenty        =   450
         value           =   42009.6059722222
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
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
         Left            =   720
         TabIndex        =   45
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Auxiliar"
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
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
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
         Left            =   3840
         TabIndex        =   30
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizacion"
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
         Left            =   3840
         TabIndex        =   29
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
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
         Left            =   3840
         TabIndex        =   28
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblFpago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Caducidad del Doc"
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
         TabIndex        =   26
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
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
         TabIndex        =   25
         Top             =   630
         Width           =   780
      End
      Begin VB.Label lblV 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   12375
      Begin VB.TextBox TxtObserv 
         Height          =   555
         Left            =   1200
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3360
         Width           =   4695
      End
      Begin VB.TextBox TxtRecargo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox TxtIva 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox TxtDesc 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   10740
         TabIndex        =   35
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox TxtSubTotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtCantIng 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
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
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2895
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   12075
         _cx             =   1970557747
         _cy             =   1970541554
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRealizaDevolucionProducto.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
      Begin MSDataListLib.DataCombo dcmbIVA 
         Height          =   315
         Left            =   8040
         TabIndex        =   49
         Top             =   3720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "% IVA:"
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
         Left            =   7440
         TabIndex        =   50
         Top             =   3750
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observ:"
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
         TabIndex        =   43
         Top             =   3435
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   9840
         TabIndex        =   42
         Top             =   3270
         Width           =   630
      End
      Begin VB.Label LblIva 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA X%"
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
         Left            =   9840
         TabIndex        =   41
         Top             =   3750
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Var SubT:"
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
         Left            =   9840
         TabIndex        =   40
         Top             =   3510
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   9840
         TabIndex        =   39
         Top             =   3990
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   9840
         TabIndex        =   38
         Top             =   4350
         Width           =   450
      End
      Begin VB.Label lblCantINg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         TabIndex        =   20
         Top             =   3270
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9975
      Begin VB.TextBox txtRuc 
         Height          =   285
         Left            =   7080
         TabIndex        =   1
         Top             =   255
         Width           =   2415
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   255
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo cmbFormulario 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label14 
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
         Left            =   6480
         TabIndex        =   48
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   46
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
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
         Left            =   240
         TabIndex        =   32
         Top             =   645
         Width           =   525
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
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4950
      TabIndex        =   15
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   600
      Picture         =   "frmRealizaDevolucionProducto.frx":04CF
      Top             =   7800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   360
      Picture         =   "frmRealizaDevolucionProducto.frx":05FB
      Top             =   7800
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmRealizaDevolucionProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso de mercadería a los depòsitos por concepto de         #
'#  importaciones se permite crear estos ingresos                               #
'#  frmIngImportacion  V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite ingresar los productos a los diferentes depòsitos       #
'#  de la compañía por concepto de importaciones , solo se permite el ingreso   #
'#  de tales datos para posteriormente actualizar las existencias.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se almacenan los nuevos ingresos de mercadería #
'#    det_ingreso: En estatabla se almacena los detalles de cada ingreso        #
'#    persona    : Se consulta los proveedores de la empresa                    #
'#    deposito   : Se consulta los depositos o bodegas de la empresa            #
'#    producto   : Se consulta los productos de la empresa                      #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#               limpiarFxGD()   Permite borrar los datos que se encuentran     #
'#                               en el flexGrid para realizar un nuevo ingreso  #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Public Neg As String
Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private clsCon_Prd2 As New clsConsulta
Private clsFPago As New clsConsulta
Private strSql As String
Private IngAsi As Boolean

Private Sub cmbCliente_Validate(Cancel As Boolean)
    Dim Gerente As String, Director As String, entra As Boolean, tipNeg As String
    entra = False: Gerente = "": Director = ""
    Dim Vend As String, ForPag As String
        
    If cmbCliente.BoundText <> "" Then
        strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,tip_ped_codigo,per_codigo_ref,per_codigo_ref2,for_pag_codigo,ven_codigo " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            tipNeg = clsCon_Def.adorec_Def(1)
            Gerente = clsCon_Def.adorec_Def(2)
            Director = clsCon_Def.adorec_Def(3)
            Vend = clsCon_Def.adorec_Def(5)
            ForPag = clsCon_Def.adorec_Def(4)
            If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                MsgBox "El Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                'cmdAceptar.Enabled = False
                'entra = True
            End If
        End If
        If entra = False Then
            strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_codigo='" & Gerente & "' " & _
                     " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                    MsgBox "El Gerente de Zona del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                    'cmdAceptar.Enabled = False
                    'entra = True
                End If
            End If
        End If
        
        If entra = False Then
            strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_codigo='" & Director & "' " & _
                     " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                    MsgBox "El Director del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                    'cmdAceptar.Enabled = False
                    'entra = True
                End If
            End If
        End If
        
        cmbVendedor.BoundText = Vend
        CmbFpago.BoundText = ForPag
        
    End If
    cmbFormulario = ""
    strSql = " SELECT DISTINCT cambio.cam_codigo,cam_factura " & _
             " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
             " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
             " AND det_cambio.prd_codigo_ped IS NOT NULL " & _
             " AND det_cambio.tip_ing_codigo='' AND det_cambio.ing_codigo='0' " & _
             " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
             " AND cambio.per_codigo='" & cmbCliente.BoundText & "' " & _
             " ORDER BY cambio.cam_codigo "
    clsCon_Def.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    Set cmbFormulario.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbFormulario.ListField = "cam_codigo"
        cmbFormulario.BoundColumn = "cam_factura"
    Else
        cmbFormulario = "No hay formularios del cliente "
    End If
    
End Sub


Private Sub cmbFactura_Validate(Cancel As Boolean)
    CargaProductos
End Sub

Private Sub cmbFormulario_Validate(Cancel As Boolean)
    txtNumAux.Text = cmbFormulario.BoundText
    dcmbIVA.BoundText = RevisivaCodigoIVAFactura(txtNumAux.Text)
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If strPtoFactura <> clsCon_Def.adorec_Def(0) Then
                LimpiarTodo
            End If
            strPtoFactura = clsCon_Def.adorec_Def(0)
        End If
    Else
        Exit Sub
    End If
    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
End Sub

Private Sub cmdAceptar_Click()
    Dim clsIngreso As New clsInventario
    Dim clsAsiento As New clsContable
    Dim clsCambio As New clsCambio
    Dim clsCta As New clsCtaXx
    Dim i As Long
    Dim strObserv As String
    Dim booGuardar As Boolean
    Dim TotalRet As Double
    Dim cue_p_c_codigo As Double
    Dim strTipCompAsiento As String
    Dim booSinIva As Boolean
    Dim NumeroAsiento As String
    booSinIva = False
    'dtpFecha.value = "2018-09-03"
    If txtSerie.Locked = False And txtDocumento.Locked = False And txtAutorizacion.Locked = False Then
        If Trim(txtSerie.Text) = "" Or Trim(txtDocumento.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
            MsgBox "Llene los campos del Documento", vbInformation, "Documento"
            Exit Sub
        End If
    End If
    If FormatoD2(TxtTotal.Text) = 0 Then
        MsgBox "No tiene seleccionado ningun producto", vbInformation, "Documento"
        Exit Sub
    End If
    If CmbFpago.Text = "" Then
        MsgBox "Llene Forma de Pago", vbInformation, "Documento"
        Exit Sub
    End If
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    clsCambio.Inicializar AdoConn, AdoConnMaster
    clsCambio.strDoc = cmbFormulario.Text
    'NOTA DE CREDITO
    
    strSql = " SELECT tip_ped_ptofac " & _
             " FROM tipo_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    clsCon_Def.Ejecutar strSql
    
    txtSerie.Text = clsCon_Def.adorec_Def("tip_ped_ptofac") & strSucursal
    
    strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
             " AND tip_ing_codigo='DCL' AND ing_codigo like '" & FormatoD0(txtSerie.Text) & "%' " & _
             " AND LEN(ing_codigo)>10 AND ing_codigo NOT IN ( " & _
                    " SELECT i_s.ing_codigo FROM ingreso_salto i_s " & _
                    " WHERE i_s.emp_codigo='" & strEmpresa & "' AND i_s.tip_ing_codigo='DCL' " & _
                    " AND i_s.ing_codigo like '" & FormatoD0(txtSerie.Text) & "%')" & _
             " ORDER BY ing_numero DESC,ing_fecha DESC,ing_codigo DESC "
    clsCon_Def.Ejecutar strSql
    
    
    
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        txtSerie.Text = clsCon_Def.adorec_Def("ing_serie")
        txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
        txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
    Else
        txtDocumento.Text = 1
        txtAutorizacion.Text = "0"
    End If
    
    booGuardar = clsIngreso.NuevoIng(True, "DCL", False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, Me.cmbCliente.BoundText, dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), 0, 0, 0, dcmbIVA.BoundText)
    If booGuardar = True Then
        strTipCompAsiento = "A"
        strObserv = UCase("NOTA DE CREDITO" & clsIngreso.strDoc & vbNewLine & "PERSONA: " & cmbCliente.Text & vbNewLine & "DOCUMENTO: " & txtSerie.Text & Format(txtDocumento.Text, "0000000") & vbNewLine & TxtObserv.Text)
        If IngAsi = True Then
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            clsAsiento.NuevoAsiento strTipCompAsiento, dtpFecha.value, 0, 0, TxtTotal.Text, strObserv
            clsIngreso.ModificaIng , , , , , , clsAsiento.NumAsiento
            NumeroAsiento = clsAsiento.NumAsiento
        End If
        
        With vsfgDetalle
            For i = 1 To VSFG.Rows - 1
                If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
                    clsIngreso.NuevoDetIng VSFG.TextMatrix(i, 3), "PRI", FormatoD4(VSFG.TextMatrix(i, 5)), FormatoD8(VSFG.TextMatrix(i, 6)), FormatoD4(VSFG.TextMatrix(i, 9)), FormatoD4(VSFG.TextMatrix(i, 7)), Abs(FormatoD0(VSFG.TextMatrix(i, 10)))
                    clsCambio.AsignarIngreso cmbFormulario.Text, VSFG.TextMatrix(i, 2), VSFG.TextMatrix(i, 3), VSFG.TextMatrix(i, 11), "DCL", clsIngreso.strDoc
                End If
            Next i
            InicializarContenedorRecurrente
        End With
        clsIngreso.DetRetenciones
        If IngAsi = True And CmbFpago.Visible = True Then
            clsFPago.adorec_Def.MoveFirst
            strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
            clsFPago.adorec_Def.Find strComparar
            'Inserta un nuevo registro de la cuenta por cobrar*/
            clsCta.Inicializar AdoConn, AdoConnMaster
            
            clsCta.IngAsientoIng clsAsiento, clsIngreso
            
            'Aplica la Nota de Credito
            clsCta.strPersona = clsIngreso.strPersona
            clsCta.strTipoCta = "C"
            clsCta.AplicaNC clsIngreso.strDoc, "2011-01-01", dtpFecha.value
            DocElectronico "04", (clsIngreso.strDoc)
            MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
            Set clsCta = Nothing
            Set clsAsiento = Nothing
        End If
        
        
        Dim rpTNC As New frmReporte
        rpTNC.strNumero = clsIngreso.strDoc
        rpTNC.strReporte = "rptNotaCredito"
        rpTNC.Show
        
        Dim rpTNC2 As New frmReporte
        rpTNC2.strNumero = clsIngreso.strDoc
        rpTNC2.strReporte = "rptNotaCreditoUbicacion"
        rpTNC2.Show
'
'        Dim rpMov1 As New frmReporte
'        rpMov1.strNumero = clsIngreso.strDoc
'        rpMov1.strTipo = clsIngreso.strTipo
'        rpMov1.strReporte = "rptDetalleAdjunto"
'        rpMov1.Show

        Unload Me
    End If

End Sub

Public Sub cmdBuscar_Click()
    Dim clsCon_Cam As New clsConsulta
    Dim i As Long
    Dim j As Long
    clsCon_Cam.Inicializar AdoConn, AdoConnMaster
    If Left(cmbFormulario.BoundText, 1) = "R" Then
        strSql = " SELECT '0' as sel, cambio.cam_codigo,mot_aju_codigo," & _
                 " prd_codigo_ing,ping.prd_nombre as prd_nombre_ping,det_cam_cantidad,det_factura_ryb.det_egr_precio," & _
                 " ROUND(det_factura_ryb.det_egr_dcto/det_factura_ryb.det_egr_cantidad,4) as dct,ROUND(det_cam_cantidad*det_factura_ryb.det_egr_precio-ROUND(det_factura_ryb.det_egr_dcto/det_factura_ryb.det_egr_cantidad*det_cam_cantidad,4),2) as tot, " & _
                 " ping.prd_costo as prd_costo_ping,ping.prd_iva, " & _
                 " COALESCE(prd_codigo_ped,'') as prd_codigo_ped ,COALESCE(pped.prd_nombre,'') as prd_nombre_pped ,COALESCE(pped.prd_costo,'0') as prd_costo_pped " & _
                 " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
                 " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
                 " AND det_cambio.prd_codigo_ped IS NOT NULL " & _
                 " AND (det_cambio.tip_ing_codigo='' or det_cambio.tip_ing_codigo is null) AND (det_cambio.ing_codigo=0 or det_cambio.ing_codigo is null) " & _
                 " INNER JOIN det_factura_ryb ON cambio.emp_codigo=det_factura_ryb.emp_codigo " & _
                 " AND cambio.cam_factura=CONCAT('R',det_factura_ryb.egr_codigo) AND det_cambio.prd_codigo_ing=det_factura_ryb.prd_codigo" & _
                 " AND det_factura_ryb.tip_egr_codigo='FAC' " & _
                 " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
                 " LEFT JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
                 " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
                 " AND cambio.cam_codigo='" & cmbFormulario.Text & "' " & _
                 " AND cambio.per_codigo='" & cmbCliente.BoundText & "' " & _
                 " ORDER BY cambio.cam_codigo,mot_aju_codigo,ping.prd_nombre "
    Else
        strSql = " SELECT DISTINCT '0' as sel, cambio.cam_codigo,det_cambio.mot_aju_codigo," & _
                 " det_cambio.prd_codigo_ing,ping.prd_nombre as prd_nombre_ping,det_cambio.det_cam_cantidad," & _
                 " COALESCE(det_egreso.det_egr_precio,de2.det_egr_precio) as det_egr_precio," & _
                 " ROUND(COALESCE(det_egreso.det_egr_dcto/det_egreso.det_egr_cantidad,de2.det_egr_dcto/de2.det_egr_cantidad)*det_cambio.det_cam_cantidad,4) as dct," & _
                 " ROUND(det_cambio.det_cam_cantidad*COALESCE(det_egreso.det_egr_precio,de2.det_egr_precio)-ROUND(COALESCE(det_egreso.det_egr_dcto/det_egreso.det_egr_cantidad*det_cambio.det_cam_cantidad,de2.det_egr_dcto/de2.det_egr_cantidad*det_cambio.det_cam_cantidad),4),2) as tot, " & _
                 " ping.prd_costo as prd_costo_ping,ping.prd_iva, " & _
                 " COALESCE(det_cambio.prd_codigo_ped,'') as prd_codigo_ped ,COALESCE(pped.prd_nombre,'') as prd_nombre_pped ," & _
                 " COALESCE(pped.prd_costo,'0') as prd_costo_pped "
        strSql = strSql & " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
                 " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
                 " AND det_cambio.prd_codigo_ped IS NOT NULL " & _
                 " AND (det_cambio.tip_ing_codigo='' or det_cambio.tip_ing_codigo is null) AND (det_cambio.ing_codigo=0 or det_cambio.ing_codigo is null) " & _
                 " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
                 " LEFT JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
                 " LEFT JOIN det_egreso ON cambio.emp_codigo=det_egreso.emp_codigo " & _
                 " AND cambio.cam_factura=det_egreso.egr_codigo AND det_cambio.prd_codigo_ing=det_egreso.prd_codigo" & _
                 " AND det_egreso.tip_egr_codigo='FAC' "
        strSql = strSql & " LEFT JOIN cambio c2 ON cambio.emp_codigo=c2.emp_codigo" & _
                 " AND cambio.cam_codigo!=c2.cam_codigo" & _
                 " AND cambio.per_codigo=c2.per_codigo" & _
                 " AND cambio.cam_factura=c2.cam_factura" & _
                 " LEFT JOIN det_cambio dc2 ON c2.emp_codigo=dc2.emp_codigo " & _
                 " AND c2.cam_codigo=dc2.cam_codigo " & _
                 " AND det_cambio.prd_codigo_ing=dc2.prd_codigo_ped " & _
                 " LEFT JOIN det_egreso de2 ON cambio.emp_codigo=de2.emp_codigo " & _
                 " AND cambio.cam_factura=de2.egr_codigo AND dc2.prd_codigo_ing=de2.prd_codigo" & _
                 " AND de2.tip_egr_codigo='FAC' "
        strSql = strSql & " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
                 " AND cambio.cam_codigo='" & cmbFormulario.Text & "' " & _
                 " AND cambio.per_codigo='" & cmbCliente.BoundText & "' AND COALESCE(det_egreso.det_egr_precio,de2.det_egr_precio) IS NOT NULL" & _
                 "  "
        strSql = strSql & " UNION "
        strSql = strSql & " SELECT DISTINCT '0' as sel, cambio.cam_codigo,det_cambio.mot_aju_codigo," & _
                 " det_cambio.prd_codigo_ing,ping.prd_nombre as prd_nombre_ping,det_cambio.det_cam_cantidad," & _
                 " COALESCE(det_factura_ryb.det_egr_precio,de3.det_egr_precio) as det_egr_precio," & _
                 " ROUND(COALESCE(det_factura_ryb.det_egr_dcto/det_factura_ryb.det_egr_cantidad,de3.det_egr_dcto/de3.det_egr_cantidad)*det_cambio.det_cam_cantidad,4) as dct," & _
                 " ROUND(det_cambio.det_cam_cantidad*COALESCE(det_factura_ryb.det_egr_precio,de3.det_egr_precio)-ROUND(COALESCE(det_factura_ryb.det_egr_dcto/det_factura_ryb.det_egr_cantidad*det_cambio.det_cam_cantidad,de3.det_egr_dcto/de3.det_egr_cantidad*det_cambio.det_cam_cantidad),4),2) as tot, " & _
                 " ping.prd_costo as prd_costo_ping,ping.prd_iva, " & _
                 " COALESCE(det_cambio.prd_codigo_ped,'') as prd_codigo_ped ,COALESCE(pped.prd_nombre,'') as prd_nombre_pped ," & _
                 " COALESCE(pped.prd_costo,'0') as prd_costo_pped "
        strSql = strSql & " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
                 " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
                 " AND det_cambio.prd_codigo_ped IS NOT NULL " & _
                 " AND (det_cambio.tip_ing_codigo='' or det_cambio.tip_ing_codigo is null) AND (det_cambio.ing_codigo=0 or det_cambio.ing_codigo is null) " & _
                 " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
                 " LEFT JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
                 " LEFT JOIN det_factura_ryb ON cambio.emp_codigo=det_factura_ryb.emp_codigo " & _
                 " AND cambio.cam_factura=det_factura_ryb.egr_codigo AND det_cambio.prd_codigo_ing=det_factura_ryb.prd_codigo" & _
                 " AND det_factura_ryb.tip_egr_codigo='FAC' "
        strSql = strSql & " LEFT JOIN cambio c2 ON cambio.emp_codigo=c2.emp_codigo" & _
                 " AND cambio.cam_codigo!=c2.cam_codigo" & _
                 " AND cambio.per_codigo=c2.per_codigo" & _
                 " AND cambio.cam_factura=c2.cam_factura" & _
                 " LEFT JOIN det_cambio dc2 ON c2.emp_codigo=dc2.emp_codigo " & _
                 " AND c2.cam_codigo=dc2.cam_codigo " & _
                 " AND det_cambio.prd_codigo_ing=dc2.prd_codigo_ped " & _
                 " LEFT JOIN det_factura_ryb de3 ON cambio.emp_codigo=de3.emp_codigo " & _
                 " AND cambio.cam_factura=de3.egr_codigo AND dc2.prd_codigo_ing=de3.prd_codigo" & _
                 " AND de3.tip_egr_codigo='FAC' "
        strSql = strSql & " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
                 " AND cambio.cam_codigo='" & cmbFormulario.Text & "' " & _
                 " AND cambio.per_codigo='" & cmbCliente.BoundText & "' AND COALESCE(det_factura_ryb.det_egr_precio,de3.det_egr_precio) IS NOT NULL" & _
                 "  "
    End If
    clsCon_Cam.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Cam.adorec_Def.DataSource
    i = 1
    j = 0
    While Not clsCon_Cam.adorec_Def.EOF
        For j = 0 To VSFG.Cols - 1
            VSFG.TextMatrix(i, j) = clsCon_Cam.adorec_Def(j)
        Next j
        i = i + 1
        clsCon_Cam.adorec_Def.MoveNext
    Wend
    RevisarDatos
End Sub

Private Sub RevisarDatos()
    Dim clsCon_TipDoc As New clsConsulta
    clsCon_TipDoc.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT tip_ing_codigo, tip_ing_nombre,tip_ing_impuesto,tip_ing_persona,tip_ing_cx_p_c,tip_ing_recargo,tip_ing_numsri,tip_ing_cos_pre,tip_ing_retencion " & _
             " FROM tipo_ingreso WHERE emp_codigo = '" & strEmpresa & "' AND tip_ing_codigo='DCL'"
    clsCon_TipDoc.Ejecutar strSql
    
    If Right(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) = "S" Then
        IngAsi = True
    Else
        IngAsi = False
    End If
    If Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) <> "N" Then
        lblFpago.Visible = True
        CmbFpago.Visible = True
    Else
        lblFpago.Visible = False
        CmbFpago.Visible = False
    End If
    CmbFpago.Tag = Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1)
    If clsCon_TipDoc.adorec_Def("tip_ing_numsri") = "F" Then
        txtSerie.Text = strSucursal & strPtoFactura
        txtDocumento.Text = ""
        txtAutorizacion.Text = strAutorFactura
        txtCaduca.Text = ""
        txtSerie.Locked = True
        txtDocumento.Locked = True
        txtAutorizacion.Locked = True
        dtpCaduca.Enabled = False
    ElseIf clsCon_TipDoc.adorec_Def("tip_ing_numsri") = "P" Then
        
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        
        txtSerie.Text = clsCon_Def.adorec_Def("tip_ped_ptofac") & strSucursal
        
        
        strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
                 " FROM ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
                 " AND tip_ing_codigo='DCL' AND ing_codigo like '" & FormatoD0(txtSerie.Text) & "%' AND LEN(ing_codigo)>10 " & _
                 " ORDER BY ing_numero DESC,ing_fecha DESC,ing_codigo DESC "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtSerie.Text = clsCon_Def.adorec_Def("ing_serie")
            txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
            txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
            If clsCon_Def.adorec_Def("ing_caduca") <> "00/0000" Then
                If clsCon_Def.adorec_Def("ing_caduca") <> "" Then
                    dtpCaduca.value = clsCon_Def.adorec_Def("ing_caduca")
                End If
                txtCaduca.Text = clsCon_Def.adorec_Def("ing_caduca")
            Else
                dtpCaduca.value = Format(HoyDia, "mm\/yyyy")
                txtCaduca.Text = ""
            End If
        Else
            txtSerie.Text = "002" & PtoEmiDocEle & ""
            txtCaduca.Text = "01/2015"
            txtDocumento.Text = "1"
            txtAutorizacion.Text = "0"
            dtpCaduca.value = Format(HoyDia, "mm\/yyyy")
        End If
        txtSerie.Locked = False
        txtDocumento.Locked = False
        txtAutorizacion.Locked = False
        dtpCaduca.Enabled = True
    End If
    cmbVendedor.Visible = True
    lblV.Visible = True

End Sub

Private Sub cmdReImp_Click()
    Dim rpMov1 As New frmReporte
    rpMov1.strNumero = FormatoD0(cmbFormulario.Text)
    rpMov1.strReporte = "rptTckAjuste"
    rpMov1.Show
    Dim rpMov2 As New frmReporte
    rpMov2.strNumero = FormatoD0(cmbFormulario.Text)
    rpMov2.strReporte = "rptAjuste"
    rpMov2.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dcmbIVA_Validate(Cancel As Boolean)
    TxtIva.Tag = dcmbIVA.Text
    CalculaTotal
End Sub

Private Sub dtpCaduca_Change()
    txtCaduca.Text = Format(dtpCaduca.value, "mm\/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSql = ""
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_Prd2 = Nothing
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd2.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    dtpFecha.value = HoyDia
    'dtpFecha.value = "2018-09-03"
    
    dtpFecha.Enabled = False
    
    Dim au As Integer
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " WHERE cod_iva_enuso=1 "
     clsCon_Def.Ejecutar (strSql)
     au = clsCon_Def.adorec_Def("cod_iva_codigo")
    TxtIva.Tag = clsCon_Def.adorec_Def("cod_iva_porcentaje")
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " ORDER BY cod_iva_porcentaje"
     clsCon_Def.Ejecutar (strSql)
     Set dcmbIVA.RowSource = clsCon_Def.adorec_Def.DataSource
     dcmbIVA.ListField = "cod_iva_porcentaje"
     dcmbIVA.BoundColumn = "cod_iva_codigo"
     dcmbIVA.BoundText = au
    
    
    
    cargarTipoPedido
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar (strSql)
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
    
    '****** VENDEDORES
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as nombV " & _
             " FROM vendedor " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY nombV "
    clsCon_Def.Ejecutar (strSql)
    Set cmbVendedor.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbVendedor.ListField = "nombV"
    cmbVendedor.BoundColumn = "ven_codigo"
    
    'PonerBotones
End Sub

Private Sub VsfgDetalleIng_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleIng.MouseRow
    c = vsfgDetalleIng.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = (vsfgDetalleIng.Rows - 1)) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleIng.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleIng.Cell(flexcpLeft, r, c) + vsfgDetalleIng.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         vsfgDetalleIng.RemoveItem (r)
         'PonerBotones
         CalculaTotal
    Else
        vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub cargarTipoPedido()
    strSql = " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo like '" & Neg & "'" & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbNegocio.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo like '" & Neg & "'" & _
             " AND tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
End Sub


Private Sub CargaProductos()

    'Carga los motivos de ajuste
    strSql = " SELECT mot_aju_codigo,mot_aju_nombre " & _
             " FROM motivo_ajuste " & _
             " Where emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY mot_aju_nombre "
    clsCon_Def.Ejecutar strSql
    vsfgDetalleIng.ColComboList(1) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "*mot_aju_codigo, mot_aju_nombre", "mot_aju_codigo")
    
    'Consulto los productos de la factura
    
    strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
             " AND det_egreso.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_egreso.tip_egr_codigo = 'FAC' " & _
             " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
             " AND prd_baja=0 ORDER BY prd_codigo "
    clsCon_Prd.Ejecutar strSql
    vsfgDetalleIng.ColComboList(2) = vsfgDetalleIng.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    'Consulto los productos de la factura
    strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
             " AND det_egreso.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_egreso.tip_egr_codigo = 'FAC' " & _
             " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
             " AND prd_baja=0 ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSql
    vsfgDetalleIng.ColComboList(3) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    
    'Consulto los productos para el cambio
    
    strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM producto INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
             " AND lis_pre_p_precio!=0 " & _
             " AND prd_baja=0 ORDER BY prd_codigo "
    clsCon_Prd2.Ejecutar strSql
    vsfgDetalleIng.ColComboList(6) = vsfgDetalleIng.BuildComboList(clsCon_Prd2.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    'Consulto los productos de la factura
    strSql = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM producto INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
             " AND lis_pre_p_precio!=0 " & _
             " AND prd_baja=0 ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSql
    vsfgDetalleIng.ColComboList(7) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")

End Sub

Private Sub vsfgDetalleIng_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If vsfgDetalleIng.TextMatrix(Row, Col) <> "" And Not IsNumeric(vsfgDetalleIng.TextMatrix(Row, Col)) Then
            MsgBox "Ingrese valores numéricos en Cantidad", vbInformation, "Detalle"
            vsfgDetalleIng.TextMatrix(Row, Col) = 0
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Or Col = 3 Then
        If vsfgDetalleIng.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero un motivo", vbInformation, "Motivos"
            vsfgDetalleIng.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
    End If
    If Col = 2 Then
        vsfgDetalleIng.TextMatrix(Row, 3) = vsfgDetalleIng.TextMatrix(Row, 2)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalleIng.TextMatrix(Row, 6) = 0
    ElseIf Col = 3 Then
        vsfgDetalleIng.TextMatrix(Row, 2) = vsfgDetalleIng.TextMatrix(Row, 3)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalleIng.TextMatrix(Row, 6) = 0
    ElseIf Col = 4 Then
        vsfgDetalleIng.TextMatrix(Row, 6) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalleIng.TextMatrix(Row, 5)))
    End If
    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 2) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 3) <> "" And Val(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 4)) <> 0 Then
        vsfgDetalleIng.AddItem ""
        vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 0) = vsfgDetalleIng.Rows - 1
        vsfgDetalleIng.Cell(flexcpPicture, vsfgDetalleIng.Rows - 1, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, vsfgDetalleIng.Rows - 1, 0) = flexAlignRightCenter
        If vsfgDetalleIng.Rows > 2 Then
             vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) = vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 2, 1)
        End If
    End If
    CalculaTotal
End Sub

Private Sub CalculaTotal()
    Dim i As Long
    Dim totalIng As Double
    Dim SubtotalIng As Double
    Dim DctoIng As Double
    Dim IVAIng As Double
    Dim CantIng As Double
    totalIng = 0
    CantIng = 0
    
    For i = 1 To VSFG.Rows - 1
        If Abs(FormatoD0(VSFG.TextMatrix(i, 0))) = 1 Then
            CantIng = CantIng + FormatoD4(VSFG.TextMatrix(i, 5))
            totalIng = totalIng + FormatoD4(VSFG.TextMatrix(i, 8))
            DctoIng = DctoIng + FormatoD4(VSFG.TextMatrix(i, 7))
            If Abs(FormatoD0(VSFG.TextMatrix(i, 10))) = 1 Then
                IVAIng = IVAIng + FormatoD4(VSFG.TextMatrix(i, 8))
            End If
        End If
    Next i
    txtCantIng.Text = FormatoD4(CantIng)
    TxtSubTotal.Text = FormatoD2(totalIng + DctoIng)
    TxtDesc.Text = FormatoD2(DctoIng)
    TxtIva.Text = FormatoD2(IVAIng * FormatoD2(TxtIva.Tag) / 100#)
    TxtTotal.Text = FormatoD2(TxtSubTotal.Text - TxtDesc.Text + TxtIva.Text)
End Sub

Private Sub txtNumAux_Change()
    dcmbIVA.BoundText = RevisivaCodigoIVAFactura(txtNumAux.Text)
    TxtIva.Tag = dcmbIVA.Text
    CalculaTotal
End Sub

Private Sub txtRuc_Validate(Cancel As Boolean)
    
    If Trim(txtRuc.Text) <> "" Then
        strSql = " SELECT per_codigo " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND per_ruc='" & txtRuc.Text & "'"
        clsCon_Def.Ejecutar strSql
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If cmbCliente.BoundText <> clsCon_Def.adorec_Def(0) Then
                cmbCliente.BoundText = clsCon_Def.adorec_Def(0)
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

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 And Row > 0 Then
        If Abs(VSFG.TextMatrix(Row, 0)) = 0 Then
            VSFG.Select Row, 1, Row, VSFG.Cols - 1
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
        Else
            VSFG.Select Row, 1, Row, VSFG.Cols - 1
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HC0FFFF
        End If
    End If
    CalculaTotal
End Sub
