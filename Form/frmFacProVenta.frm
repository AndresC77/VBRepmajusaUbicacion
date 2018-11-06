VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmV_FacProVenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación de Proyecto de Trabajo"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmFacProVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9705
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle:"
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
      Height          =   5175
      Left            =   945
      TabIndex        =   6
      Top             =   2640
      Width           =   7815
      Begin VB.TextBox TxtObserv 
         Height          =   285
         Left            =   360
         MaxLength       =   250
         TabIndex        =   15
         Top             =   4680
         Width           =   7095
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
         Left            =   6240
         TabIndex        =   14
         Top             =   3728
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4080
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3480
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFG 
         Height          =   2055
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   7200
         _cx             =   12700
         _cy             =   3625
         _ConvInfo       =   1
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFacProVenta.frx":030A
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
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFGReca 
         Height          =   855
         Left            =   390
         TabIndex        =   16
         Top             =   3240
         Width           =   4305
         _cx             =   7594
         _cy             =   1508
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         FormatString    =   $"frmFacProVenta.frx":0416
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
      End
      Begin VB.Label Label7 
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
         TabIndex        =   23
         Top             =   4440
         Width           =   1185
      End
      Begin VB.Label Label6 
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
         TabIndex        =   22
         Top             =   3000
         Width           =   765
      End
      Begin VB.Label Label1 
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
         Left            =   4980
         TabIndex        =   21
         Top             =   3037
         Width           =   1155
      End
      Begin VB.Label LblIva 
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
         Left            =   5040
         TabIndex        =   20
         Top             =   3270
         Width           =   615
      End
      Begin VB.Label Label8 
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
         Left            =   5040
         TabIndex        =   19
         Top             =   3765
         Width           =   825
      End
      Begin VB.Label Label4 
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
         Left            =   5040
         TabIndex        =   18
         Top             =   3480
         Width           =   750
      End
      Begin VB.Label Label5 
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
         Left            =   5040
         TabIndex        =   17
         Top             =   4110
         Width           =   915
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
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   7
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Proyectos de Trabajo:"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9495
      Begin VSFlex7Ctl.VSFlexGrid VSFGProT 
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   9045
         _cx             =   15954
         _cy             =   2566
         _ConvInfo       =   1
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFacProVenta.frx":0496
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
      End
      Begin MSDataListLib.DataCombo CmbTipoFac 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   368
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   368
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
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
         Left            =   7440
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59506691
         CurrentDate     =   37463
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Factura:"
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
         Left            =   285
         TabIndex        =   29
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Left            =   3480
         TabIndex        =   28
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   6840
         TabIndex        =   27
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdDeBaja 
      Caption         =   "Dar de Baja"
      Height          =   375
      Left            =   1770
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Facturar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Timer TmrAct 
      Interval        =   20000
      Left            =   600
      Top             =   5880
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   615
      Picture         =   "frmFacProVenta.frx":05A5
      Top             =   5520
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   375
      Picture         =   "frmFacProVenta.frx":06D1
      ToolTipText     =   "Elimina una Fila"
      Top             =   5520
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmV_FacProVenta"
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

Private clsPedidos As New clsConsulta
Private clsSql As New clsConsulta
Private clsTFac As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsFPago As New clsConsulta
Private clsRet As New clsConsulta
Private clsVer As New clsConsulta
Private clsCantEnt As New clsConsulta
Private clsProm As New clsPrePromCot
Private iva As Double, codPer As String, fila As Long, CabGrid As String

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

Private Sub CalcuReca()
    'Calcula el total del pedido
    Dim Suma As Double
    For i = 1 To VSFGReca.Rows - 1
        Suma = Suma + Val(Replace(VSFGReca.TextMatrix(i, 3), ",", "."))
    Next i
    TxtRecargo = Format(Suma, "##,##0.00")
    txtTotal = Format(Suma + Val(txtIva) + Val(txtSubTotal), "##,##0.00")
End Sub

Private Sub CalcuTotal()
    'Calcula es total del pedido
    Dim Suma As Double, Columna As Long
    'Busca cual es la columna del total
    For i = 0 To VSFG.Cols - 1
        If VSFG.TextMatrix(0, i) = "Total" Then
            Columna = i
            Exit For
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        Suma = Suma + Val(Replace(VSFG.TextMatrix(i, Columna), ",", "."))
    Next i
    'Coloca los totales parciales de la factura
    txtSubTotal = Format(Suma, "##,##0.00")
    txtIva = Format(Suma * iva / 100, "##,##0.00")
    txtTotal = Format(Suma + Val(txtIva) + Val(TxtRecargo) - Val(TxtDesc), "##,##0.00")
End Sub

Private Sub CmbFpago_Change()
    CmdLimpiar = True
End Sub

Private Sub CmbTipoFac_Change()
    CmdLimpiar = True
End Sub
Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
     'Detiene la actualización automática de los proyectos de trabajo a mostrar
    TmrAct.Enabled = False
'****** INGRESO
    Dim num As Long, codEgr As Long
    'Realiza un ingreso al proyecto de trabajo de los productos a facturar para poder cuadrar con bodega
    'Obtiene el código con el que se debe insertar el nuevo ingreso
    strSql = " SELECT if(max(ing_codigo),max(ing_codigo),0) as num " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='DPR' "
    clsSql.Ejecutar (strSql)
    num = clsSql.adorec_Def("num") + 1
    'Inserta la cabecera del ingreso
    strSql = " INSERT INTO ingreso (ing_codigo, tip_ing_codigo, emp_codigo, " & _
             " per_codigo, ing_fecha, ing_subtotal, ing_subtotal_o, ing_dcto, ing_impuesto, " & _
             " ing_total, ing_observacion, ing_numasiento, ing_fechamod, ing_usumod) " & _
             " VALUES (" & num & ",'DPR','" & strEmpresa & "','" & VSFGProT.TextMatrix(fila, 7) & "','" & Format(dtpFecha, "yyyy-MM-dd") & "', " & _
             " 0, 0, 0, 0, 0, 'Ingreso para Facturar proyecto de Trabajo', '', CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
    clsSql.Ejecutar (strSql)
    'Inserta un detalle de proyecto de trabajo
    strSql = " INSERT INTO det_pro_tra (pro_tra_codigo, emp_codigo, det_pro_tra_tipo, " & _
             " det_pro_tra_codigo, det_pro_tra_ie, det_pro_tra_fechamod, det_pro_tra_usumod) " & _
             " VALUES (" & VSFGProT.TextMatrix(fila, 0) & ",'" & strEmpresa & "', 'DPR'," & num & ",'I',CURRENT_TIMESTAMP,substring_index(USER(),'@',1)) "
    clsSql.Ejecutar (strSql)
    'Inserta los detalles de ingreso al proyecto
    strSql = " INSERT INTO det_ingreso (emp_codigo,ing_codigo,tip_ing_codigo,prd_codigo,dep_codigo," & _
             " det_ing_cantidad,det_ing_precio,det_ing_fechamod,det_ing_usumod) " & _
             " SELECT '" & strEmpresa & "'," & num & ",'DPR',AuxI_E.prd_codigo,'" & VSFGProT.TextMatrix(fila, 6) & "', " & _
             " egr-ing,PromPre,CURRENT_TIMESTAMP,substring_index(USER(),'@',1) " & _
             " FROM AuxI_E INNER JOIN PrePromCot ON AuxI_E.prd_codigo = PrePromCot.prd_codigo " & _
             " Order By AuxI_E.prd_codigo "
    clsSql.Ejecutar (strSql)
'****** EGRESO
    'Realiza el egreso definitivo de la factura
    'Obtiene el código con el que se debe insertar el nuevo egreso
    strSql = " SELECT if(max(egr_codigo),max(egr_codigo),0) as num " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='FAC' "
    clsSql.Ejecutar (strSql)
    codEgr = clsSql.adorec_Def("num") + 1
    'Inserta la cabecera del egreso
    strSql = " INSERT INTO egreso (egr_codigo, tip_egr_codigo, emp_codigo, for_pag_codigo, " & _
             " per_codigo, egr_fecha, egr_subtotal, egr_subtotal_o, egr_dcto, egr_impuesto, " & _
             " egr_total, egr_observacion, egr_numasiento, egr_fechamod, egr_usumod) " & _
             " VALUES (" & codEgr & ",'FAC','" & strEmpresa & "','" & CmbFpago.BoundText & "','" & VSFGProT.TextMatrix(fila, 7) & "', " & _
             " '" & Format(dtpFecha, "yyyy-MM-dd") & "'," & txtSubTotal & ",'" & TxtRecargo & "','" & TxtDesc & "'" & _
             " ," & txtIva & "," & txtTotal & ",'" & TxtObserv & "','',CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
    clsSql.Ejecutar (strSql)
    'Inserta un detalle de proyecto de trabajo
    strSql = " INSERT INTO det_pro_tra (pro_tra_codigo, emp_codigo, det_pro_tra_tipo, " & _
             " det_pro_tra_codigo, det_pro_tra_ie, det_pro_tra_fechamod, det_pro_tra_usumod) " & _
             " VALUES (" & VSFGProT.TextMatrix(fila, 0) & ",'" & strEmpresa & "','FAC'," & num & ", " & _
             " 'E', CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
    clsSql.Ejecutar (strSql)
    'Inserta los detalles de ingreso al proyecto
    strSql = " INSERT INTO det_egreso (emp_codigo,egr_codigo,tip_egr_codigo,prd_codigo,dep_codigo," & _
             " det_egr_cantidad,det_egr_precio,det_egr_fechamod,det_egr_usumod) " & _
             " SELECT '" & strEmpresa & "'," & num & ",'FAC',AuxI_E.prd_codigo,'" & VSFGProT.TextMatrix(fila, 6) & "', " & _
             " egr-ing,PromPre,CURRENT_TIMESTAMP,substring_index(USER(),'@',1) " & _
             " FROM AuxI_E INNER JOIN PrePromCot ON AuxI_E.prd_codigo = PrePromCot.prd_codigo " & _
             " Order By AuxI_E.prd_codigo "
    clsSql.Ejecutar (strSql)
    clsProm.elimTabla
'****** ACTUALIZA PROYECTO DE VENTA Y TRABAJO
    'Actualiza el proyecto de trabajo
    strSql = " UPDATE pro_trabajo " & _
             " SET pro_tra_estado = 2, " & _
             " pro_egr_codigo = " & num & ", " & _
             " pro_tra_tipo_factura = " & CmbTipoFac.BoundText & ", " & _
             " pro_tra_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_tra_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_tra_codigo=" & VSFGProT.TextMatrix(fila, 0) & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
    'Actualiza el estado del proyecto de venta a vendido
    strSql = " UPDATE proyecto_venta SET " & _
             " pro_ven_estado = 3, " & _
             " pro_ven_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_ven_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_ven_codigo=" & VSFGProT.TextMatrix(fila, 8) & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
'****** COTIZACION
    'Actualiza la cotizacion a vendida
    strSql = " UPDATE cotizacion SET cot_estado=2, " & _
             " cot_fechamod=CURRENT_TIMESTAMP," & _
             " cot_usumod=substring_index(USER(),'@',1) " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo=" & VSFGProT.TextMatrix(fila, 1)
    clsSql.Ejecutar (strSql)
'****** RETENCIONES
    'Realiza las retenciones al cliente
    Dim retencion As Double
    'Obtiene todas las retenciones de un cliente
    strSql = " SELECT retencion.ret_codigo, ret_porcentaje, ret_gravara " & _
             " FROM retencion INNER JOIN persona_ret ON (retencion.ret_codigo = persona_ret.ret_codigo) " & _
             " AND (retencion.emp_codigo = persona_ret.emp_codigo) " & _
             " WHERE persona_ret.per_codigo='" & VSFGProT.TextMatrix(fila, 7) & "' AND persona_ret.emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
    'Genera e inserta las retenciones de una persona
    While Not clsSql.adorec_Def.EOF
        'Verifica a que item de la factura afecta la retención
        Select Case clsSql.adorec_Def("ret_gravara")
            Case "SubTotal"
                retencion = Val(txtSubTotal) * clsSql.adorec_Def("ret_porcentaje") / 100
            Case "IVA0%"
                retencion = Val(TxtRecargo) * clsSql.adorec_Def("ret_porcentaje") / 100
            Case "IVA"
                retencion = Val(txtIva) * clsSql.adorec_Def("ret_porcentaje") / 100
            Case "Total"
                retencion = Val(txtTotal) * clsSql.adorec_Def("ret_porcentaje") / 100
        End Select
        'Inserta un detalle de ingreso de retención
        strSql = " INSERT INTO det_egreso_ret " & _
                 " (emp_codigo, ret_codigo, egr_codigo, tip_egr_codigo, " & _
                 " det_egr_ret_valor, det_egr_ret_fechamod, det_egr_ret_usumod) " & _
                 " VALUES ('" & strEmpresa & "','" & clsSql.adorec_Def("ret_codigo") & "'," & _
                 codEgr & ", 'FAC'," & Replace(retencion, ",", ".") & ", CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
        clsRet.Ejecutar (strSql)
        clsSql.adorec_Def.MoveNext
    Wend
'****** RECARGOS
    'Genera los posibles recargos que podujo esta factura
    For i = 1 To VSFGReca.Rows - 1
        If VSFGReca.TextMatrix(i, 1) <> "" Then
            strSql = " INSERT INTO det_egreso_c (emp_codigo,egr_codigo,tip_egr_codigo,oca_codigo, " & _
                     " det_egr_c_cantidad,det_egr_c_precio,det_egr_c_fechamod,det_egr_c_usumod) " & _
                     " VALUES ('" & strEmpresa & "'," & codEgr & ",'FAC','" & VSFGReca.TextMatrix(i, 1) & "'" & _
                     " ,1," & Replace(VSFGReca.TextMatrix(i, 3), ",", ".") & " ,CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
            clsSql.Ejecutar (strSql)
        End If
    Next i
'****** MENSAJE
    'Actualiza el grid que muestra los proyectos de venta actuales
    clsPedidos.Actualizar
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
    MsgBox "Proyecto No. " & LblPedido & " facturado.", vbInformation, "Proyecto"
    CmdConfirmar.Enabled = False
    CmdLimpiar = True
    'Reactiva el control timer
    TmrAct.Enabled = True
End Sub

Private Sub CmdDeBaja_Click()
    'Verifica si no se completa el pedido
    Dim Resp As Integer
    'Verifica que el usuario esté seguro de dar de baja al pedido
    Resp = MsgBox("Está seguro de dar de baja al Proyecto Nº. " & LblPedido, vbInformation + vbYesNo, "De Baja")
    If Resp = vbNo Then
        Exit Sub
    End If
    'Da de baja al proyecto de trabajo
    strSql = " UPDATE pro_trabajo SET pro_tra_estado=1, " & _
             " pro_tra_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_tra_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_tra_codigo=" & LblPedido & " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
    'Da de baja al proyecto de venta
    strSql = " UPDATE proyecto_venta SET pro_ven_estado=1, " & _
             " pro_ven_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_ven_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_ven_codigo=" & VSFGProT.TextMatrix(fila, 8) & " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
    'Da de baja la cotización
    strSql = " UPDATE cotizacion SET cot_estado=3, " & _
             " cot_fechamod=CURRENT_TIMESTAMP," & _
             " cot_usumod=substring_index(USER(),'@',1) " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo=" & VSFGProT.TextMatrix(fila, 1)
    clsSql.Ejecutar (strSql)
    'Limpia los grids que mostraban datos del pedido
    MsgBox "Proyecto No. " & LblPedido & " dado de Baja.", vbInformation, "De Baja"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
End Sub

Private Sub CmdLimpiar_Click()
    'Muestra el formulario como si se hubiera cargado por primera vez
    VSFGProT.Select 0, 0, 0, 0
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGReca.Clear 1
    VSFGReca.Rows = 2
    VSFGReca.Enabled = False
    CmdConfirmar.Enabled = False
    CmdDeBaja.Enabled = False
    txtSubTotal = ""
    txtTotal = ""
    TxtRecargo = ""
    txtIva = ""
    TxtDesc = ""
    LblPedido = "-"
    fila = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - mdiPrincipal.Height / 40
    '- (Me.Height / 6)
    'Inicializa los objetos de conexión con la base de datos
    clsPedidos.Inicializar AdoConn
    clsTFac.Inicializar AdoConn
    clsRecargos.Inicializar AdoConn
    clsSql.Inicializar AdoConn
    clsFPago.Inicializar AdoConn
    clsRet.Inicializar AdoConn
    clsVer.Inicializar AdoConn
    'Consulta todos los proyectos de trabajo que se pueden facturar
    strSql = " SELECT pro_tra_codigo, cot_codigo, SUBSTRING(cot_fecha,1,10) as fechaCot, CONCAT(SUBSTRING(pro_ven_descricion,1,25),' ...') as descPro, " & _
             " CONCAT(per_apellido,' ',per_nombre) as nombC, CONCAT(ven_apellido,' ',ven_nombre) as nombV, deposito.dep_codigo, persona.per_codigo, " & _
             " proyecto_venta.pro_ven_codigo, cot_subtotal " & _
             " FROM ((((persona INNER JOIN proyecto_venta ON (persona.per_codigo = proyecto_venta.per_codigo) AND (persona.emp_codigo = proyecto_venta.emp_codigo)) " & _
             " INNER JOIN vendedor ON (vendedor.emp_codigo = proyecto_venta.emp_codigo) AND (vendedor.ven_codigo = proyecto_venta.ven_codigo)) " & _
             " INNER JOIN cotizacion ON (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo) AND (proyecto_venta.emp_codigo = cotizacion.emp_codigo)) " & _
             " INNER JOIN pro_trabajo ON (proyecto_venta.emp_codigo = pro_trabajo.emp_codigo) AND (proyecto_venta.pro_ven_codigo = pro_trabajo.pro_ven_codigo)) " & _
             " INNER JOIN deposito ON (pro_trabajo.pro_dep_codigo = deposito.dep_codigo) AND (pro_trabajo.emp_codigo = deposito.emp_codigo) " & _
             " WHERE cot_estado=1 AND proyecto_venta.emp_codigo='" & strEmpresa & "' " & _
             " AND pro_tra_estado=0 " & _
             " ORDER BY pro_trabajo.pro_tra_codigo "
    clsPedidos.Ejecutar (strSql)
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
    'Consulta los recargos que puede manejar una empresa
    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Obtiene el IVA vigente para realizar la factura
    strSql = " SELECT par_numero " & _
             " FROM parametro " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='IVAV'"
    clsSql.Ejecutar (strSql)
    iva = clsSql.adorec_Def("par_numero")
    lblIva = "IVA " & iva & " %:"
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSql = " SELECT for_pag_codigo, for_pag_nombre " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar (strSql)
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
    CmbFpago = "CONTADO"
    'Consulta todos los tipos de factura y los muestra en un combo
    strSql = " SELECT * FROM tipo_factura "
    clsTFac.Ejecutar (strSql)
    Set CmbTipoFac.RowSource = clsTFac.adorec_Def.DataSource
    CmbTipoFac.ListField = "tipo_fac_descripcion"
    CmbTipoFac.BoundColumn = "tipo_fac_codigo"
    CmbTipoFac = "Lo Entregado"
    'Coloca los botones de eliminar fila en el grid de recargos
    PonerBotones
    'Coloca la fecha actual
    dtpFecha.Value = Date
    CabGrid = VSFG.FormatString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Elimina la tabla temporal de precios promedio de productos de una cotización
    clsProm.elimTabla
End Sub

Private Sub LblPedido_Click()

End Sub

'Verifica cada 10 segundos si existe un nuevo pedido a revisar
Private Sub TmrAct_Timer()
    clsPedidos.Actualizar
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
    With VSFGProT
        If fila > 0 Then
            'Coloca en blanco todas las celdas del cuerpo del grid
            .Select 1, 0, .Rows - 1, .Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = &H80000005
            'Coloca la fila seleccionada en color azul
            .Row = fila
            .Select fila, 0, fila, .Cols - 1
            .FillStyle = flexFillRepeat
        End If
    End With
End Sub

Private Sub TxtDesc_Change()
    TxtDesc = Replace(TxtDesc, ",", ".")
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
    txtIva = Replace(txtIva, ",", ".")
End Sub

Private Sub TxtRecargo_Change()
    TxtRecargo = Replace(TxtRecargo, ",", ".")
End Sub

Private Sub TxtSubTotal_Change()
    txtSubTotal = Replace(txtSubTotal, ",", ".")
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
    txtTotal = Replace(txtTotal, ",", ".")
End Sub

Private Sub VSFGProT_Click()
    'Selecciona toda una fila del grid cuando se ha dado un click sobre ella
    With VSFGProT
        If .Row > 0 Then
            fila = .Row
            LblPedido = .TextMatrix(fila, 0)
            VSFGReca.Enabled = True
            'Coloca en blanco todas las celdas del cuerpo del grid
            .Select 1, 0, .Rows - 1, .Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = &HFFFFFF
            'Coloca la fila seleccionada en color azul
            .Select fila, 0, fila, .Cols - 1
            .FillStyle = flexFillRepeat
            CmdConfirmar.Enabled = True
        '******* EGRESO E INGRESO AL PROYECTO
            strSql = " DROP TABLE IF EXISTS AuxI_E "
            clsSql.Ejecutar (strSql)
            'Obtiene la cantidad de productos entregados al proyecto de trabajo suma(egr - ing)
            'Crea una tabla temporal con el total de las devoluviones hechas proyecto
            strSql = " CREATE TEMPORARY TABLE AuxI_E " & _
                     " SELECT det_ingreso.prd_codigo, 0 As egr, Sum(det_ingreso.det_ing_cantidad) AS ing " & _
                     " FROM (det_pro_tra INNER JOIN ingreso ON (det_pro_tra.det_pro_tra_codigo = ingreso.ing_codigo) " & _
                     " AND (det_pro_tra.emp_codigo = ingreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = ingreso.tip_ing_codigo)) " & _
                     " INNER JOIN det_ingreso ON (ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo) AND (ingreso.emp_codigo = det_ingreso.emp_codigo) " & _
                     " AND (ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                     " WHERE det_pro_tra.det_pro_tra_ie='I' AND det_pro_tra.pro_tra_codigo=" & .TextMatrix(fila, 0) & _
                     " AND det_pro_tra.det_pro_tra_tipo='DPR' AND det_ingreso.emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY det_ingreso.prd_codigo "
            clsSql.Ejecutar (strSql)
            'Crea una tabla temporal con el total de las notas de remision hechas al proyecto
            strSql = " INSERT INTO AuxI_E " & _
                     " SELECT det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) AS egr, 0 as ing " & _
                     " FROM (det_pro_tra INNER JOIN egreso ON (det_pro_tra.det_pro_tra_codigo = egreso.egr_codigo) " & _
                     " AND (det_pro_tra.emp_codigo = egreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = egreso.tip_egr_codigo)) " & _
                     " INNER JOIN det_egreso ON (egreso.tip_egr_codigo = det_egreso.tip_egr_codigo) AND (egreso.emp_codigo = det_egreso.emp_codigo) " & _
                     " AND (egreso.egr_codigo = det_egreso.egr_codigo) " & _
                     " WHERE det_pro_tra.det_pro_tra_ie='E' AND det_pro_tra.pro_tra_codigo=" & .TextMatrix(fila, 0) & _
                     " AND det_pro_tra.det_pro_tra_tipo='NRP' AND det_egreso.emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY det_egreso.prd_codigo "
            clsSql.Ejecutar (strSql)
            'Crea la tabla temporal de precios promedio de productos de la cotización seleccionada
            clsProm.crearTabla .TextMatrix(fila, 1), .TextMatrix(fila, 7), strEmpresa
            'Coloca las cabeceras de grid de detalles
            VSFG.FormatString = CabGrid
        '******* TIPO FACTURA
            TxtObserv = ""
            TxtObserv.Locked = False
            txtSubTotal.Locked = True
            'Verifica que tipo de factura se va a realizar
            Select Case CmbTipoFac.BoundText
                Case 0 'Lo cotizado
                    'Consulta los datos de la cotización seleccionada en el grid de proyectos de trabajo
                    strSql = " SELECT '" & .TextMatrix(fila, 6) & "' as Bod, det_cotizacion.prd_codigo, IF(producto.prd_nombre<>'',producto.prd_nombre,producto_compuesto.prd_com_nombre) as nombPrd, " & _
                             " det_cot_cantidad, det_cot_precio, (det_cot_cantidad*det_cot_precio) as total, cot_dcto, cot_observacion, cot_subtotal " & _
                             " FROM (((proyecto_venta INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo)) " & _
                             " INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
                             " LEFT JOIN producto ON (det_cotizacion.emp_codigo = producto.emp_codigo) AND (det_cotizacion.prd_codigo = producto.prd_codigo)) " & _
                             " LEFT JOIN producto_compuesto ON (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo) AND (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo) " & _
                             " WHERE proyecto_venta.emp_codigo='" & strEmpresa & "' AND cotizacion.cot_codigo=" & .TextMatrix(fila, 1) & " " & _
                             " ORDER BY proyecto_venta.pro_ven_codigo, cotizacion.cot_codigo "
                    clsVer.Ejecutar (strSql)
                    'Muestra el descuento y las observaciones de la cotización seleccionada
                    TxtDesc = clsVer.adorec_Def("cot_dcto")
                    TxtObserv = clsVer.adorec_Def("cot_observacion")
                    TxtObserv.Locked = True
                    txtSubTotal = clsVer.adorec_Def("cot_subtotal")
                Case 1 'Lo entregado
                    'Cruza las tablas de Egreso, Ingreso a proyecto con la tabla de precios promedios de productos de la cotización
                    'para obtener finalmente lo que se ha entregado al proyecto de trabajo
                    strSql = " SELECT '" & .TextMatrix(fila, 6) & "' as Bod, AuxI_E.prd_codigo, prd_nombre, " & _
                             " Sum(egr)-Sum(ing) as cant, PromPre, (Sum(egr)-Sum(ing))*PromPre as Total " & _
                             " FROM AuxI_E INNER JOIN PrePromCot ON (AuxI_E.prd_codigo = PrePromCot.prd_codigo) " & _
                             " Group By AuxI_E.prd_codigo " & _
                             " Order By AuxI_E.prd_codigo "
                    clsVer.Ejecutar (strSql)
                Case 2 'El servicio
                    'Coloca en el detalle de la factura que se está facturando el servicio
                    VSFG.Clear
                    VSFG.Rows = 2
                    VSFG.Cols = 3
                    VSFG.FormatString = "^Bodega|^Descripción|^Total"
                    VSFG.TextMatrix(1, 0) = .TextMatrix(fila, 6)
                    VSFG.TextMatrix(1, 1) = "Facturación del proyecto: " & .TextMatrix(fila, 3) & " por servicio"
                    VSFG.TextMatrix(1, 2) = Format(.TextMatrix(fila, 9), "###0.00")
                    VSFG.ColAlignment(2) = flexAlignRightCenter
                    VSFG.AutoSize 0, 2
                    txtSubTotal.Locked = False
                    txtSubTotal = VSFG.TextMatrix(1, 2)
            End Select
            'Muestra los datos en el grid de detalle
            If Val(CmbTipoFac.BoundText) < 2 Then
                Set VSFG.DataSource = clsVer.adorec_Def.DataSource
            End If
        Else
            CmdLimpiar = True
        End If
        'Calcula el total de la factura si hay filas mostradas
        If VSFG.Rows > 1 Then
            CalcuTotal
            CmdDeBaja.Enabled = True
        Else
            CmdLimpiar = True
        End If
    End With
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

Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
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
        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - X
        If d > imgBtnDn.Width Then Exit Sub
        
        ' click was on a button: do the work
         .Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
        Respuesta = MsgBox(Mensaje, Estilo, Título)
            
        'Recorro el FlexGrid para poner números a las filas
            
        If Respuesta = vbYes Then
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
