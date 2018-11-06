VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmV_FacProVenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación de Proyecto de Trabajo"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmV_FacProVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   9735
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
      Height          =   5415
      Left            =   960
      TabIndex        =   16
      Top             =   3000
      Width           =   7815
      Begin VB.TextBox TxtObserv 
         Height          =   285
         Left            =   360
         MaxLength       =   250
         TabIndex        =   10
         Top             =   4800
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2055
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   7200
         _cx             =   12700
         _cy             =   3625
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
         FormatString    =   $"frmV_FacProVenta.frx":030A
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
         Editable        =   1
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
         Height          =   855
         Left            =   390
         TabIndex        =   4
         Top             =   3240
         Width           =   4305
         _cx             =   7594
         _cy             =   1508
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
         FormatString    =   $"frmV_FacProVenta.frx":0418
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
         TabIndex        =   25
         Top             =   4560
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   3277
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   3517
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5040
         TabIndex        =   19
         Top             =   4117
         Width           =   1065
      End
      Begin VB.Label LblDetalle 
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
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   1200
         TabIndex        =   17
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
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   9495
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   7560
         TabIndex        =   30
         Top             =   488
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "ACT"
         Height          =   1455
         Left            =   9120
         TabIndex        =   29
         Top             =   1080
         Width           =   255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGProT 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   9045
         _cx             =   15954
         _cy             =   2566
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
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_FacProVenta.frx":0498
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
      Begin MSDataListLib.DataCombo CmbTipoFac 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   488
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   4560
         TabIndex        =   1
         Top             =   488
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         TabIndex        =   28
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   3600
         TabIndex        =   27
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Left            =   6870
         TabIndex        =   26
         Top             =   540
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdDeBaja 
      Caption         =   "Dar de Baja"
      Height          =   375
      Left            =   1785
      TabIndex        =   11
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6495
      TabIndex        =   14
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Facturar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3375
      TabIndex        =   12
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   4935
      TabIndex        =   13
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   495
      Picture         =   "frmV_FacProVenta.frx":05A7
      Top             =   6600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   255
      Picture         =   "frmV_FacProVenta.frx":06D3
      ToolTipText     =   "Elimina una Fila"
      Top             =   6600
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
Private clsExis As New clsConsulta
Private clsProm As New clsPrePromCot
Private IVA As Double, CodPer As String, fila As Long, CabGrid As String

Private Sub cmdActualizar_Click()
    Actualizar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    'Elimina la tabla temporal de precios promedio de productos de una cotización
    clsProm.elimTabla
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidos = Nothing
    Set clsSql = Nothing
    Set clsTFac = Nothing
    Set clsRecargos = Nothing
    Set clsFPago = Nothing
    Set clsRet = Nothing
    Set clsVer = Nothing
    Set clsCantEnt = Nothing
    Set clsProm = Nothing
    Set clsExis = Nothing
End Sub

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
        Suma = Suma + Val(Format(VSFGReca.TextMatrix(i, 3), "###0.00"))
    Next i
    TxtRecargo = Format(Suma, "####0.00")
    TxtTotal = Format(Suma + Val(Format(TxtIva, "###0.00")) + Val(Format(TxtSubTotal, "###0.00")), "####0.00")
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
        Suma = Suma + Val(Format(VSFG.TextMatrix(i, Columna), "###0.00"))
    Next i
    'Coloca los totales parciales de la factura
    TxtSubTotal = Format(Suma, "####0.00")
    TxtIva = Format(Suma * IVA / 100, "####0.00")
    TxtTotal = Format(Suma + Val(Format(TxtIva, "###0.00")) + Val(Format(TxtRecargo, "###0.00")) - Val(Format(TxtDesc, "###0.00")), "####0.00")
End Sub

Private Sub CmbFpago_Change()
    CmdLimpiar = True
End Sub

Private Sub CmbTipoFac_Change()
    CmdLimpiar = True
End Sub
Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
     'Detiene la actualización automática de los proyectos de trabajo a mostrar
    'TmrAct.Enabled = False
'****** INGRESO
    Dim num As Double, codEgr As Double
    Dim booPasar As Boolean
    Dim booGuardar As Boolean
    Dim booCerrar As Boolean
    Dim CodPer As String
    Dim codVen As String
    Dim clsEgreso As New clsInventario
    Dim clsMovE As New clsInventario
    Dim clsMovI As New clsInventario
    Dim clsAsiento As New clsContable
    Dim clsCta As New clsCtaXx
    
    Dim GuiaAutomatica As Boolean
    
    'Realiza un ingreso al proyecto de trabajo de los productos a facturar para poder cuadrar con bodega
    If CmbFpago.Text = "" Then Exit Sub
    'Obtiene el código con el que se debe insertar el nuevo egreso
    CodPer = VSFGProT.TextMatrix(fila, 7)
    codVen = CmdConfirmar.Tag
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsMovE.Inicializar AdoConn, AdoConnMaster
    clsMovI.Inicializar AdoConn, AdoConnMaster
    booGuardar = clsEgreso.NuevoEgr(True, "FAC", True, strSucursal, strPtoFactura, , CmbFpago.BoundText, CodPer, Format(dtpFecha.Value, "yyyy-MM-dd"), , codVen, TxtObserv, , strAutorFactura, strCaducaFactura, FormatoD2(TxtSubTotal), FormatoD2(TxtRecargo), FormatoD2(TxtDesc), FormatoD2(TxtIva), FormatoD2(TxtTotal), 0, , , CodigoIVA)
    If booGuardar = True Then
        codEgr = clsEgreso.strDoc
        clsAsiento.Inicializar AdoConn, AdoConnMaster
        clsAsiento.NuevoAsiento "F", dtpFecha.Value, 0, 0, TxtTotal.Text, "FACTURA " & codEgr
        'Inserta la cabecera del egreso
        clsEgreso.ModificaEgr , , , , , , clsAsiento.NumAsiento
        
        If MsgBox("Quiere Cerrar el Proyecto?", vbYesNo + vbQuestion, "Proyecto de Trabajo") = vbYes Then
            booCerrar = True
        Else
            booCerrar = False
        End If
        If booCerrar = True Then
    '************************DEVOLUCION A PRINCIPAL
    '        //Crea una tabla temporal con el total de las trasnferencias hechas al proyecto
            strSQL = " CREATE TEMPORARY TABLE AuxI " & _
                      " SELECT COALESCE(det_ingreso.prd_codigo,det_egreso.prd_codigo) as prd_codigo, Sum(COALESCE(det_ingreso.det_ing_cantidad,0)-COALESCE(det_egreso.det_egr_cantidad,0)) AS cantidad " & _
                      " FROM (det_pro_tra LEFT JOIN ingreso ON (det_pro_tra.det_pro_tra_codigo = ingreso.ing_codigo) " & _
                      " AND (det_pro_tra.emp_codigo = ingreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = ingreso.tip_ing_codigo)) " & _
                      " LEFT JOIN det_ingreso ON (ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo) AND (ingreso.emp_codigo = det_ingreso.emp_codigo) " & _
                      " AND (ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                      " LEFT JOIN egreso ON (det_pro_tra.det_pro_tra_codigo = egreso.egr_codigo) AND (det_pro_tra.emp_codigo = egreso.emp_codigo) " & _
                      " AND (det_pro_tra.det_pro_tra_tipo = egreso.tip_egr_codigo) AND egreso.tip_egr_codigo='ETR' " & _
                      " LEFT JOIN det_egreso ON (egreso.tip_egr_codigo = det_egreso.tip_egr_codigo) AND (egreso.emp_codigo = det_egreso.emp_codigo) AND (egreso.egr_codigo = det_egreso.egr_codigo) AND det_egreso.tip_egr_codigo='ETR' " & _
                      " WHERE if(det_pro_tra.det_pro_tra_ie='I',det_pro_tra.det_pro_tra_tipo='ITR',if(det_pro_tra.det_pro_tra_ie='E',det_pro_tra.det_pro_tra_tipo='ETR',1=0))" & _
                      " AND det_pro_tra.pro_tra_codigo='" & VSFGProT.TextMatrix(fila, 0) & "' " & _
                      " AND det_pro_tra.emp_codigo='" & strEmpresa & "' " & _
                      " GROUP BY prd_codigo "
            clsSql.Ejecutar strSQL, "M"
    
    '        //Crea una tabla temporal con el total de los productos pedidos sumando los componentes de productos compuestos
            strSQL = " CREATE TEMPORARY TABLE AuxP " & _
                      " SELECT if(producto.prd_codigo<>'',producto.prd_codigo,det_prd_com.prd_codigo)as prd_codigo, " & _
                      " sum(if(isnull(det_prd_com.det_prd_com_cantidad),det_cotizacion.det_cot_cantidad,det_prd_com.det_prd_com_cantidad * det_cotizacion.det_cot_cantidad))as cantidad,det_cotizacion.det_cot_precio as prec " & _
                      " FROM (((det_cotizacion LEFT JOIN producto_compuesto ON (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo)" & _
                      " AND (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo)) " & _
                      " LEFT JOIN producto ON (det_cotizacion.prd_codigo = producto.prd_codigo) AND (det_cotizacion.emp_codigo = producto.emp_codigo)) " & _
                      " LEFT JOIN det_prd_com ON (producto_compuesto.emp_codigo = det_prd_com.emp_codigo) AND (producto_compuesto.prd_com_codigo = det_prd_com.prd_com_codigo)) " & _
                      " LEFT JOIN producto AS producto_1 ON (det_prd_com.emp_codigo = producto_1.emp_codigo) AND (det_prd_com.prd_codigo = producto_1.prd_codigo) " & _
                      " WHERE det_cotizacion.cot_codigo='" & VSFGProT.TextMatrix(fila, 1) & "' AND det_cotizacion.emp_codigo='" & strEmpresa & "' " & _
                      " GROUP BY prd_codigo "
            clsSql.Ejecutar strSQL, "M"
    '        //echo $strSql;
    '        //Crea una tabla temporal con el total de las devoluviones del proyecto
            strSQL = " CREATE TEMPORARY TABLE AuxI2 " & _
                      " SELECT det_ingreso.prd_codigo, Sum(det_ingreso.det_ing_cantidad) AS cantidad_i " & _
                      " FROM (det_pro_tra INNER JOIN ingreso ON (det_pro_tra.det_pro_tra_codigo = ingreso.ing_codigo) " & _
                      " AND (det_pro_tra.emp_codigo = ingreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = ingreso.tip_ing_codigo)) " & _
                      " INNER JOIN det_ingreso ON (ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo) AND (ingreso.emp_codigo = det_ingreso.emp_codigo) " & _
                      " AND (ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                      " WHERE det_pro_tra.det_pro_tra_ie='I' AND det_pro_tra.pro_tra_codigo='" & VSFGProT.TextMatrix(fila, 0) & "' " & _
                      " AND (det_pro_tra.det_pro_tra_tipo='DPR' OR det_pro_tra.det_pro_tra_tipo='FAC') AND det_ingreso.emp_codigo='" & strEmpresa & "' " & _
                      " GROUP BY det_ingreso.prd_codigo "
            clsSql.Ejecutar strSQL, "M"
    '        //Crea una tabla temporal con el total de las notas de remision hechas al proyecto
            strSQL = " CREATE TEMPORARY TABLE AuxE2 " & _
                      " SELECT det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) AS cantidad_e " & _
                      " FROM (det_pro_tra INNER JOIN egreso ON (det_pro_tra.det_pro_tra_codigo = egreso.egr_codigo) " & _
                      " AND (det_pro_tra.emp_codigo = egreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = egreso.tip_egr_codigo)) " & _
                      " INNER JOIN det_egreso ON (egreso.tip_egr_codigo = det_egreso.tip_egr_codigo) AND (egreso.emp_codigo = det_egreso.emp_codigo) " & _
                      " AND (egreso.egr_codigo = det_egreso.egr_codigo) " & _
                      " WHERE det_pro_tra.det_pro_tra_ie='E' AND det_pro_tra.pro_tra_codigo='" & VSFGProT.TextMatrix(fila, 0) & "' " & _
                      " AND (det_pro_tra.det_pro_tra_tipo='NRP' OR det_pro_tra.det_pro_tra_tipo='FAC') AND det_egreso.emp_codigo='" & strEmpresa & "' " & _
                      " GROUP BY det_egreso.prd_codigo "
            clsSql.Ejecutar strSQL, "M"
            
    '            //******* CODIGO EGRESO E INGRESO
    '        //Obtiene el código con el que se debe insertar el nuevo egreso de productos
            clsMovE.NuevoEgr False, "ETR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , "DEVOLUCION PROYECTO " & VSFGProT.TextMatrix(fila, 3)
            clsMovI.NuevoIng False, "ITR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , "DEVOLUCION PROYECTO " & VSFGProT.TextMatrix(fila, 3)
            Dim Ntrasf As Double
            Ntrasf = clsMovE.strDoc
    '    //******* DETALLE PROYECTO TRABAJO
    '        //Inserta un detalle de proyecto de trabajo
            strSQL = " INSERT INTO det_pro_tra (pro_tra_codigo, emp_codigo, det_pro_tra_tipo, det_pro_tra_codigo, det_pro_tra_ie, " & _
                      " det_pro_tra_fechamod, det_pro_tra_usumod) " & _
                      " VALUES ('" & VSFGProT.TextMatrix(fila, 0) & "', '" & strEmpresa & "', 'ITR', '" & Ntrasf & "', 'I', CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
            clsSql.Ejecutar strSQL, "M"
    '       strSql = " INSERT INTO det_egreso (emp_codigo,egr_codigo,tip_egr_codigo,prd_codigo,dep_codigo," .
    '                          " det_egr_cantidad,det_egr_precio,det_egr_costo,det_egr_fechamod,det_egr_usumod) " .
    '                          " VALUES ('" . $empresa . "'," . $num . ",'ETR','" . $matAuxi2[$i][3] . "'" .
    '                          " ,'" . $matAuxi2[$i][1] . "'," . $matAuxi2[$i][5] . "," . $matAuxi2[$i][6] .
    '                          " ,CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) ";
    '        //Obtiene un resumen de la candidad de prodiuctos necesaria a transferir a la bodega del proyecto
            strSQL = " SELECT AuxE2.prd_codigo,'" & VSFGProT.TextMatrix(fila, 6) & "' as dep_codigo, " & _
                     " (COALESCE(AuxI.cantidad,0) + COALESCE(AuxI2.cantidad_i,0) - COALESCE(AuxE2.cantidad_e,0)) as cantidad,COALESCE(AuxP.prec,0) as precio,prd_costo as costo " & _
                     " FROM AuxE2 INNER JOIN producto ON AuxE2.prd_codigo=producto.prd_codigo AND producto.emp_codigo='" & strEmpresa & "' " & _
                     " LEFT JOIN AuxI ON AuxE2.prd_codigo = AuxI.prd_codigo " & _
                     " LEFT JOIN AuxI2 ON AuxE2.prd_codigo=AuxI2.prd_codigo " & _
                     " LEFT JOIN AuxP ON AuxE2.prd_codigo=AuxP.prd_codigo " & _
                     " WHERE (COALESCE(AuxI.cantidad,0) + COALESCE(AuxI2.cantidad_i,0) - COALESCE(AuxE2.cantidad_e,0)) <> 0"
            clsSql.Ejecutar strSQL, "M"
            While Not clsSql.adorec_Def.EOF
                clsMovE.NuevoDetEgr clsSql.adorec_Def("prd_codigo"), clsSql.adorec_Def("dep_codigo"), clsSql.adorec_Def("cantidad"), clsSql.adorec_Def("precio"), clsSql.adorec_Def("costo"), 0, 1
                clsSql.adorec_Def.MoveNext
            Wend
            strSQL = " SELECT AuxE2.prd_codigo,'" & strBodega & "' as dep_codigo, " & _
                     " (COALESCE(AuxI.cantidad,0) + COALESCE(AuxI2.cantidad_i,0) - COALESCE(AuxE2.cantidad_e,0)) as cantidad,COALESCE(AuxP.prec,0) as precio,prd_costo as costo " & _
                     " FROM AuxE2 INNER JOIN producto ON AuxE2.prd_codigo=producto.prd_codigo AND producto.emp_codigo='" & strEmpresa & "' " & _
                     " LEFT JOIN AuxI ON AuxE2.prd_codigo = AuxI.prd_codigo " & _
                     " LEFT JOIN AuxI2 ON AuxE2.prd_codigo=AuxI2.prd_codigo " & _
                     " LEFT JOIN AuxP ON AuxE2.prd_codigo=AuxP.prd_codigo " & _
                     " WHERE (COALESCE(AuxI.cantidad,0) + COALESCE(AuxI2.cantidad_i,0) - COALESCE(AuxE2.cantidad_e,0)) <> 0"
            clsSql.Ejecutar strSQL, "M"
            While Not clsSql.adorec_Def.EOF
                clsMovI.NuevoDetIng clsSql.adorec_Def("prd_codigo"), clsSql.adorec_Def("dep_codigo"), clsSql.adorec_Def("cantidad"), clsSql.adorec_Def("precio"), clsSql.adorec_Def("costo"), 0, 1
                clsSql.adorec_Def.MoveNext
            Wend
            InicializarContenedorRecurrente
            strSQL = " DROP TABLE AuxP "
            clsSql.Ejecutar strSQL, "M"
            strSQL = " DROP TABLE AuxI "
            clsSql.Ejecutar strSQL, "M"
            strSQL = " DROP TABLE AuxI2 "
            clsSql.Ejecutar strSQL, "M"
            strSQL = " DROP TABLE AuxE2 "
            clsSql.Ejecutar strSQL, "M"
    '*******************************************
        End If
        
        'Obtiene el código con el que se debe insertar el nuevo ingreso
        clsMovI.NuevoIng True, "DPR", False, strSucursal, strPtoFactura, , , CodPer, Format(dtpFecha.Value, "yyyy-MM-dd"), , , "Ingreso para facturar proyecto de trabajo " & VSFGProT.TextMatrix(fila, 3)
        num = clsMovI.strDoc
        'Inserta un detalle de proyecto de trabajo
        strSQL = " INSERT INTO det_pro_tra (pro_tra_codigo, emp_codigo, det_pro_tra_tipo, " & _
                 " det_pro_tra_codigo, det_pro_tra_ie, det_pro_tra_fechamod, det_pro_tra_usumod) " & _
                 " VALUES (" & VSFGProT.TextMatrix(fila, 0) & ",'" & strEmpresa & "', 'DPR'," & num & ",'I',CURRENT_TIMESTAMP,substring_index(USER(),'@',1)) "
        clsSql.Ejecutar strSQL, "M"
        'Inserta los detalles de ingreso al proyecto
            For i = 1 To VSFG.Rows - 1
                If VSFG.TextMatrix(i, 3) > 0 Then
                    clsMovI.NuevoDetIng VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 0), VSFG.TextMatrix(i, 3), VSFG.TextMatrix(i, 4), VSFG.TextMatrix(i, 6), 0, 1
                End If
            Next i
            InicializarContenedorRecurrente
    '****** EGRESO
        'Realiza el egreso definitivo de la factura
        'Inserta un detalle de proyecto de trabajo
        strSQL = " INSERT INTO det_pro_tra (pro_tra_codigo, emp_codigo, det_pro_tra_tipo, " & _
                 " det_pro_tra_codigo, det_pro_tra_ie, det_pro_tra_fechamod, det_pro_tra_usumod) " & _
                 " VALUES (" & VSFGProT.TextMatrix(fila, 0) & ",'" & strEmpresa & "','FAC'," & codEgr & ", " & _
                 " 'E', CURRENT_TIMESTAMP, substring_index(USER(),'@',1)) "
        clsSql.Ejecutar strSQL, "M"
        'Inserta los detalles de ingreso al proyecto
'        While Not clsSql.adorec_Def.EOF
'
'            clsSql.adorec_Def.MoveNext
'        Wend
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 3) > 0 Then
                strSQL = " SELECT prd_costo " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                clsSql.Ejecutar strSQL
                clsEgreso.NuevoDetEgr VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 0), VSFG.TextMatrix(i, 3), VSFG.TextMatrix(i, 4), clsSql.adorec_Def("prd_costo"), 0, 1
            End If
        Next i
        clsProm.elimTabla
        If booCerrar = True Then
        '****** ACTUALIZA PROYECTO DE VENTA Y TRABAJO
            'Actualiza el proyecto de trabajo
            strSQL = " UPDATE pro_trabajo " & _
                     " SET pro_tra_estado = 2, " & _
                     " pro_egr_codigo = " & codEgr & ", " & _
                     " pro_tra_tipo_factura = " & CmbTipoFac.BoundText & ", " & _
                     " pro_tra_fechamod=CURRENT_TIMESTAMP, " & _
                     " pro_tra_usumod=substring_index(USER(),'@',1) " & _
                     " WHERE pro_tra_codigo=" & VSFGProT.TextMatrix(fila, 0) & _
                     " AND emp_codigo='" & strEmpresa & "' "
            clsSql.Ejecutar strSQL, "M"
            'Actualiza el estado del proyecto de venta a vendido
            strSQL = " UPDATE proyecto_venta SET " & _
                     " pro_ven_estado = 3, " & _
                     " pro_ven_fechamod=CURRENT_TIMESTAMP, " & _
                     " pro_ven_usumod=substring_index(USER(),'@',1) " & _
                     " WHERE pro_ven_codigo=" & VSFGProT.TextMatrix(fila, 8) & _
                     " AND emp_codigo='" & strEmpresa & "' "
            clsSql.Ejecutar strSQL, "M"
        '****** COTIZACION
            'Actualiza la cotizacion a vendida
            strSQL = " UPDATE cotizacion SET cot_estado=2, " & _
                     " cot_fechamod=CURRENT_TIMESTAMP," & _
                     " cot_usumod=substring_index(USER(),'@',1) " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo=" & VSFGProT.TextMatrix(fila, 1)
            clsSql.Ejecutar strSQL, "M"
        End If
    '****** RETENCIONES
        clsEgreso.DetRetenciones
    '****** RECARGOS
        'Genera los posibles recargos que podujo esta factura
        For i = 1 To VSFGReca.Rows - 1
            If VSFGReca.TextMatrix(i, 1) <> "" Then
                clsEgreso.NuevoDetEgrRecargo VSFGReca.TextMatrix(i, 1), FormatoD2(VSFGReca.TextMatrix(i, 3))
            End If
        Next i
    '****** MENSAJE
    
    '*****************************
            clsFPago.adorec_Def.MoveFirst
        strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
        'Inserta un nuevo registro de la cuenta por cobrar*/
        clsCta.Inicializar AdoConn, AdoConnMaster
        clsFPago.adorec_Def.Find strComparar
        clsCta.NuevaCta "C", 1, "00", Format(dtpFecha.Value, "yyyy-MM-dd"), Format(DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha.Value), "yyyy-MM-dd"), CodPer, "Factura # " & codEgr & " - " & TxtObserv, strSucursal & strPtoFactura, Right(codEgr, 7), strAutorFactura, strCaducaFactura, clsEgreso.dblTotalProd, clsEgreso.dblTotalServ, clsEgreso.dblTotalProdIVA, clsEgreso.dblTotalServIVA, 2, clsEgreso.dblIVA, clsEgreso.dblSubTotal0, 0, 0, 0, clsEgreso.dblTotal, clsAsiento.NumAsiento

        clsCta.IngAsientoEgr clsAsiento, clsEgreso
        Set clsCta = Nothing
        Set clsAsiento = Nothing

        'Actualiza el grid que muestra los proyectos de venta actuales
        clsPedidos.Actualizar
        Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
        MsgBox "Proyecto No. " & LblPedido & " facturado.", vbInformation, "Proyecto"
        CmdConfirmar.Enabled = False
    End If
    CmdLimpiar = True
    If booGuardar = True Then
        Dim RepFactura As New frmReporte
        RepFactura.strNumero = codEgr
        'no se usa
        GuiaAutomatica = False
        RepFactura.strReporte = IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola")
        RepFactura.Show
    End If
    'Reactiva el control timer
    'TmrAct.Enabled = True
    
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
    strSQL = " UPDATE pro_trabajo SET pro_tra_estado=1, " & _
             " pro_tra_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_tra_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_tra_codigo=" & LblPedido & " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSQL, "M"
    'Da de baja al proyecto de venta
    strSQL = " UPDATE proyecto_venta SET pro_ven_estado=1, " & _
             " pro_ven_fechamod=CURRENT_TIMESTAMP, " & _
             " pro_ven_usumod=substring_index(USER(),'@',1) " & _
             " WHERE pro_ven_codigo=" & VSFGProT.TextMatrix(fila, 8) & " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSQL, "M"
    'Da de baja la cotización
    strSQL = " UPDATE cotizacion SET cot_estado=3, " & _
             " cot_fechamod=CURRENT_TIMESTAMP," & _
             " cot_usumod=substring_index(USER(),'@',1) " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo=" & VSFGProT.TextMatrix(fila, 1)
    clsSql.Ejecutar strSQL, "M"
    'Limpia los grids que mostraban datos del pedido
    MsgBox "Proyecto No. " & LblPedido & " dado de Baja.", vbInformation, "De Baja"
    CmdLimpiar = True
    'Actualiza el grid que muestra los pedidos actuales
    clsPedidos.Actualizar
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
End Sub

Private Sub cmdLimpiar_Click()
    'Muestra el formulario como si se hubiera cargado por primera vez
    VSFGProT.Select 0, 0, 0, 0
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGReca.Clear 1
    VSFGReca.Rows = 2
    VSFGReca.Enabled = False
    CmdConfirmar.Enabled = False
    CmdDeBaja.Enabled = False
    TxtSubTotal = ""
    TxtTotal = ""
    TxtRecargo = ""
    TxtIva = ""
    TxtDesc = ""
    LblPedido = "-"
    fila = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    '- (Me.Height / 6)
    'Inicializa los objetos de conexión con la base de datos
    clsPedidos.Inicializar AdoConn, AdoConnMaster
    clsTFac.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    clsRet.Inicializar AdoConn, AdoConnMaster
    clsVer.Inicializar AdoConn, AdoConnMaster
    'Consulta todos los proyectos de trabajo que se pueden facturar
    strSQL = " SELECT pro_tra_codigo, cot_codigo, SUBSTRING(cot_fecha,1,10) as fechaCot, CONCAT(SUBSTRING(pro_ven_descricion,1,25),' ...') as descPro, " & _
             " CONCAT(per_apellido,' ',per_nombre) as nombC, proyecto_venta.ven_codigo as nombV, deposito.dep_codigo, persona.per_codigo, " & _
             " proyecto_venta.pro_ven_codigo, cot_subtotal " & _
             " FROM ((((persona INNER JOIN proyecto_venta ON (persona.per_codigo = proyecto_venta.per_codigo) AND (persona.emp_codigo = proyecto_venta.emp_codigo)) " & _
             " INNER JOIN vendedor ON (vendedor.emp_codigo = proyecto_venta.emp_codigo) AND (vendedor.ven_codigo = proyecto_venta.ven_codigo)) " & _
             " INNER JOIN cotizacion ON (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo) AND (proyecto_venta.emp_codigo = cotizacion.emp_codigo)) " & _
             " INNER JOIN pro_trabajo ON (proyecto_venta.emp_codigo = pro_trabajo.emp_codigo) AND (proyecto_venta.pro_ven_codigo = pro_trabajo.pro_ven_codigo)) " & _
             " INNER JOIN deposito ON (pro_trabajo.pro_dep_codigo = deposito.dep_codigo) AND (pro_trabajo.emp_codigo = deposito.emp_codigo) " & _
             " WHERE cot_estado=1 AND proyecto_venta.emp_codigo='" & strEmpresa & "' " & _
             " AND pro_tra_estado=0 " & _
             " ORDER BY pro_trabajo.pro_tra_codigo "
    clsPedidos.Ejecutar strSQL
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGProT.DataSource = clsPedidos.adorec_Def.DataSource
    'Consulta los recargos que puede manejar una empresa
    strSQL = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar strSQL
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Obtiene el IVA vigente para realizar la factura
    IVA = PorIVA
    LblIva = "IVA " & IVA & " %:"
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSQL = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar strSQL
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
    'Consulta todos los tipos de factura y los muestra en un combo
    strSQL = " SELECT * FROM tipo_factura "
    clsTFac.Ejecutar strSQL
    Set CmbTipoFac.RowSource = clsTFac.adorec_Def.DataSource
    CmbTipoFac.ListField = "tipo_fac_descripcion"
    CmbTipoFac.BoundColumn = "tipo_fac_codigo"
    CmbTipoFac = "Lo Entregado"
    'Coloca los botones de eliminar fila en el grid de recargos
    PonerBotones
    'Coloca la fecha actual
    dtpFecha.Value = HoyDia
    CabGrid = VSFG.FormatString
End Sub

'Verifica cada 10 segundos si existe un nuevo pedido a revisar
Private Sub Actualizar()
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
    TxtDesc = Format(TxtDesc, "###0.00")
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
    TxtIva = Format(TxtIva, "###0.00")
End Sub

Private Sub TxtRecargo_Change()
    TxtRecargo = Format(TxtRecargo, "###0.00")
End Sub

Private Sub TxtSubTotal_Change()
    TxtSubTotal = Format(TxtSubTotal, "###0.00")
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
    TxtTotal = Format(TxtTotal, "###0.00")
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    VSFG.TextMatrix(Row, 5) = Val(Format(VSFG.TextMatrix(Row, 4), "###0.0000")) * Val(Format(VSFG.TextMatrix(Row, 3), "###0.00"))
    CalcuTotal
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Row > 0 And (Col = 4 Or Col = 3)) Then
        Cancel = True
    End If
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
            strSQL = " DROP TABLE IF EXISTS AuxI_E "
            clsSql.Ejecutar strSQL
            CmdConfirmar.Tag = .TextMatrix(fila, 5)
            'Obtiene la cantidad de productos entregados al proyecto de trabajo suma(egr - ing)
            'Crea una tabla temporal con el total de las devoluviones hechas proyecto
            strSQL = " CREATE TEMPORARY TABLE AuxI_E " & _
                     " SELECT det_ingreso.prd_codigo, 00000000 As egr, Sum(det_ingreso.det_ing_cantidad) AS ing,0000000.0000 as precio,Sum(det_ingreso.det_ing_cantidad*det_ing_costo)/Sum(det_ingreso.det_ing_cantidad) as costo " & _
                     " FROM (det_pro_tra INNER JOIN ingreso ON (det_pro_tra.det_pro_tra_codigo = ingreso.ing_codigo) " & _
                     " AND (det_pro_tra.emp_codigo = ingreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = ingreso.tip_ing_codigo)) " & _
                     " INNER JOIN det_ingreso ON (ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo) AND (ingreso.emp_codigo = det_ingreso.emp_codigo) " & _
                     " AND (ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                     " WHERE det_pro_tra.det_pro_tra_ie='I' AND det_pro_tra.pro_tra_codigo=" & .TextMatrix(fila, 0) & _
                     " AND det_pro_tra.det_pro_tra_tipo='DPR' AND det_ingreso.emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY det_ingreso.prd_codigo "
            clsSql.Ejecutar strSQL
            'Crea una tabla temporal con el total de las notas de remision hechas al proyecto
            strSQL = " INSERT INTO AuxI_E " & _
                     " SELECT det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) AS egr, 00000000 as ing,Sum(det_egreso.det_egr_cantidad*det_egr_precio)/Sum(det_egreso.det_egr_cantidad) as precio,Sum(det_egreso.det_egr_cantidad*det_egr_costo)/Sum(det_egreso.det_egr_cantidad) as costo " & _
                     " FROM (det_pro_tra INNER JOIN egreso ON (det_pro_tra.det_pro_tra_codigo = egreso.egr_codigo) " & _
                     " AND (det_pro_tra.emp_codigo = egreso.emp_codigo) AND (det_pro_tra.det_pro_tra_tipo = egreso.tip_egr_codigo)) " & _
                     " INNER JOIN det_egreso ON (egreso.tip_egr_codigo = det_egreso.tip_egr_codigo) AND (egreso.emp_codigo = det_egreso.emp_codigo) " & _
                     " AND (egreso.egr_codigo = det_egreso.egr_codigo) " & _
                     " WHERE det_pro_tra.det_pro_tra_ie='E' AND det_pro_tra.pro_tra_codigo=" & .TextMatrix(fila, 0) & _
                     " AND det_pro_tra.det_pro_tra_tipo='NRP' AND det_egreso.emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY det_egreso.prd_codigo "
            clsSql.Ejecutar strSQL
            'Crea la tabla temporal de precios promedio de productos de la cotización seleccionada
            clsProm.crearTabla .TextMatrix(fila, 1), .TextMatrix(fila, 7), strEmpresa
            'Coloca las cabeceras de grid de detalles
            VSFG.FormatString = CabGrid
        '******* TIPO FACTURA
            TxtObserv = ""
            TxtObserv.Locked = False
            TxtSubTotal.Locked = True
            'Verifica que tipo de factura se va a realizar
            Select Case CmbTipoFac.BoundText
                Case 0 'Lo cotizado
                    'Consulta los datos de la cotización seleccionada en el grid de proyectos de trabajo
                    strSQL = " SELECT '" & .TextMatrix(fila, 6) & "' as Bod, det_cotizacion.prd_codigo, IF(producto.prd_nombre<>'',producto.prd_nombre,producto_compuesto.prd_com_nombre) as nombPrd, " & _
                             " det_cot_cantidad, det_cot_precio, (det_cot_cantidad*det_cot_precio) as total, cot_dcto, cot_observacion, cot_subtotal " & _
                             " FROM (((proyecto_venta INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo)) " & _
                             " INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
                             " LEFT JOIN producto ON (det_cotizacion.emp_codigo = producto.emp_codigo) AND (det_cotizacion.prd_codigo = producto.prd_codigo)) " & _
                             " LEFT JOIN producto_compuesto ON (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo) AND (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo) " & _
                             " WHERE proyecto_venta.emp_codigo='" & strEmpresa & "' AND cotizacion.cot_codigo=" & .TextMatrix(fila, 1) & " " & _
                             " ORDER BY proyecto_venta.pro_ven_codigo, cotizacion.cot_codigo "
                    clsVer.Ejecutar strSQL
                    'Muestra el descuento y las observaciones de la cotización seleccionada
                    TxtDesc = clsVer.adorec_Def("cot_dcto")
                    TxtObserv = clsVer.adorec_Def("cot_observacion")
                    TxtObserv.Locked = True
                    TxtSubTotal = clsVer.adorec_Def("cot_subtotal")
                Case 1 'Lo entregado
                    'Cruza las tablas de Egreso, Ingreso a proyecto con la tabla de precios promedios de productos de la cotización
                    'para obtener finalmente lo que se ha entregado al proyecto de trabajo
                    strSQL = " SELECT '" & .TextMatrix(fila, 6) & "' as Bod, AuxI_E.prd_codigo, prd_nombre, " & _
                             " Sum(egr)-Sum(ing) as cant, COALESCE(Sum(egr*precio)/Sum(egr),0) as prec, (Sum(egr)-Sum(ing))*COALESCE(Sum(egr*precio)/Sum(egr),0) as Total,COALESCE(Sum(egr*costo)/Sum(egr),0) as costo " & _
                             " FROM AuxI_E INNER JOIN producto ON AuxI_E.prd_codigo=producto.prd_codigo AND producto.emp_codigo='" & strEmpresa & "' LEFT JOIN PrePromCot ON (AuxI_E.prd_codigo = PrePromCot.prd_codigo) " & _
                             " Group By AuxI_E.prd_codigo HAVING cant!=0 " & _
                             " Order By AuxI_E.prd_codigo "
                    clsVer.Ejecutar strSQL
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
                    TxtSubTotal.Locked = False
                    TxtSubTotal = VSFG.TextMatrix(1, 2)
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

Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
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
        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        
        ' click was on a button: do the work
         .Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
            
        'Recorro el FlexGrid para poner números a las filas
            
        If respuesta = vbYes Then
             Dim i As Long
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
