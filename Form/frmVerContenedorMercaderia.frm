VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerContenedorMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contenedores de Mercaderia"
   ClientHeight    =   7440
   ClientLeft      =   1290
   ClientTop       =   1380
   ClientWidth     =   19470
   Icon            =   "frmVerContenedorMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   19470
   Begin VB.CommandButton cmdContenedorVacios 
      Caption         =   "Anular &Vacios"
      Height          =   360
      Left            =   10920
      TabIndex        =   25
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdCrearContenedorDescontinuados 
      Caption         =   "&Crear Cont. Desco."
      Height          =   360
      Left            =   9120
      TabIndex        =   24
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimirFull 
      Caption         =   "&Imprimir 1ra Vez"
      Height          =   360
      Left            =   11760
      TabIndex        =   23
      Top             =   960
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimirDet 
      Caption         =   "Imprimir Detalle"
      Height          =   360
      Left            =   11760
      TabIndex        =   19
      Top             =   480
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimirSTK 
      Caption         =   "Imprimir Sticker"
      Height          =   360
      Left            =   11760
      TabIndex        =   18
      Top             =   120
      Width           =   1700
   End
   Begin VB.CommandButton cmdReubicarContenedor 
      Caption         =   "&Reubicar Contenedor"
      Height          =   360
      Left            =   3720
      TabIndex        =   16
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdTransferirPrendas 
      Caption         =   "&Transferir Prendas"
      Height          =   360
      Left            =   5520
      TabIndex        =   15
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdAnularContenedor 
      Caption         =   "&Anular Contenedor"
      Height          =   360
      Left            =   1920
      TabIndex        =   14
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdAuditarContenedo 
      Caption         =   "&Auditar Contenedor"
      Height          =   360
      Left            =   7320
      TabIndex        =   13
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdCrearContenedor 
      Caption         =   "&Crear Contenedor"
      Height          =   360
      Left            =   105
      TabIndex        =   7
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   12930
      TabIndex        =   6
      Top             =   6960
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      Begin VB.CheckBox chkVacios 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Vacios"
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
         Height          =   255
         Left            =   9120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkFiltroEstado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Estado"
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
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Contenedor"
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkFiltroBodega 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Bodega"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo cmbBodega 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   5880
         TabIndex        =   21
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5880
         TabIndex        =   22
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contenedor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   495
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bodega"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   495
         Width           =   3255
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   1800
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   6540
      _cx             =   11536
      _cy             =   3175
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
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerContenedorMercaderia.frx":030A
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG1 
      Height          =   3480
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   6540
      _cx             =   11536
      _cy             =   6138
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
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerContenedorMercaderia.frx":042A
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   3480
      Left            =   6720
      TabIndex        =   12
      Top             =   3360
      Width           =   7020
      _cx             =   12382
      _cy             =   6138
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
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerContenedorMercaderia.frx":04AF
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG3 
      Height          =   1800
      Left            =   6720
      TabIndex        =   17
      Top             =   1440
      Width           =   7020
      _cx             =   12382
      _cy             =   3175
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerContenedorMercaderia.frx":05CF
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VSFlex8LCtl.VSFlexGrid vsfgCaracteristica 
      Height          =   4080
      Left            =   13800
      TabIndex        =   27
      Top             =   120
      Width           =   5580
      _cx             =   9842
      _cy             =   7197
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   6000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerContenedorMercaderia.frx":06B8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
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
   Begin VB.Image imgPic 
      BorderStyle     =   1  'Fixed Single
      Height          =   2505
      Left            =   13800
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   4005
   End
End
Attribute VB_Name = "frmVerContenedorMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Private strContenedor As String
Private clsContenedor As New clsContenedor
Private strProducto As String
Private Sub IniDato()
    Tipo = "Contenedor de Mercaderia"
    Tipo2 = "Contenedor de Mercaderia"
    Me.Caption = Tipo
End Sub

Private Sub chkFiltroBodega_Click()
    If chkFiltroBodega.Value = 1 Then
        cmbBodega.Enabled = True
    Else
        cmbBodega.Enabled = False
    End If

End Sub

Private Sub chkFiltroEstado_Click()
    If chkFiltroEstado.Value = 1 Then
        cmbEstado.Enabled = True
    Else
        cmbEstado.Enabled = False
    End If
End Sub

Private Sub cmbBodega_Validate(Cancel As Boolean)
    Carga
End Sub

Private Sub cmdAnularContenedor_Click()
    If strContenedor <> "" Then
        clsContenedor.AnularContenedor
    End If
End Sub

Private Sub cmdAuditarContenedo_Click()
    If strContenedor <> "" Then
        frmAuditoriaContenedor.Show
        frmAuditoriaContenedor.txtCodigoOrigen.Text = strContenedor
    End If
End Sub

Private Sub cmdContenedorVacios_Click()
    frmContenedorMercaderiaVacio.Show
End Sub

Private Sub cmdCrearContenedor_Click()
    frmContenedorMercaderia.Show
End Sub

Private Sub cmdCrearContenedorDescontinuados_Click()
    frmContenedorMercaderiaDesc.Show
End Sub

Private Sub cmdImprimirDet_Click()
    If strContenedor <> "" Then
        clsContenedor.ImprimirLista 1, False
    End If
End Sub

Private Sub cmdImprimirFull_Click()
    If strContenedor <> "" Then
        clsContenedor.ImprimirLista 2
        clsContenedor.ImprimirSTK
        clsContenedor.CambiaEstado "0"
    End If
End Sub

Private Sub cmdImprimirSTK_Click()
    If strContenedor <> "" Then
        clsContenedor.ImprimirSTK False
    End If
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSql = " SELECT contenedor_mercaderia.con_mer_codigo,con_mer_fecha,est_con_mer_codigo,dep_codigo,ubi_bod_codigo,con_mer_observacion,con_mer_fechamod,con_mer_usumod, '0' as modi " & _
             " FROM contenedor_mercaderia " & _
             " WHERE emp_codigo = '" & strEmpresa & "' "
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  con_mer_codigo LIKE  '" & txtCodigo.Text & "%'"
    End If
    If chkFiltroBodega.Value = 1 Then
        strSql = strSql & " AND  dep_codigo LIKE '%" & cmbBodega.BoundText & "%' "
    End If
    If chkFiltroEstado.Value = 1 Then
        strSql = strSql & " AND  est_con_mer_codigo LIKE '%" & cmbEstado.BoundText & "%' "
    End If
    
    If chkFiltroEstado.Value = 1 Then
        strSql = strSql & " AND  est_con_mer_codigo LIKE '%" & cmbEstado.BoundText & "%' "
    End If
    If chkVacios.Value = 1 Then
        strSql = strSql & " AND est_con_mer_codigo!=-1 AND NOT EXISTS " & _
                 "( " & _
                 " SELECT DISTINCT con_mer_codigo " & _
                 " FROM ( " & _
                 " SELECT CM.con_mer_codigo,det_contenedor_mercaderia.prd_codigo," & _
                 " SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                 " FROM contenedor_mercaderia CM INNER JOIN det_contenedor_mercaderia ON CM.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND CM.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE CM.emp_codigo = '" & strEmpresa & "' " & _
                 " AND est_con_mer_codigo!=-1 " & _
                 " GROUP BY CM.con_mer_codigo,det_contenedor_mercaderia.prd_codigo " & _
                 " HAVING SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)!=0 " & _
                 " ) lleno WHERE lleno.con_mer_codigo=contenedor_mercaderia.con_mer_codigo" & _
                 " ) "
    End If
    strSql = strSql & " ORDER BY contenedor_mercaderia.con_mer_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    strSql = " SELECT dep_codigo as cod, dep_nombre as nomb" & _
             " FROM deposito " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By nomb "
    clsCon_Def.Ejecutar (strSql)
    
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "cod, *nomb", "cod")
    
    strSql = " SELECT est_con_mer_codigo as cod, est_con_mer_descripcion as nomb" & _
             " FROM est_contenedor_mercaderia " & _
             " Order By nomb "
    clsCon_Def.Ejecutar (strSql)
    
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "cod, *nomb", "cod")
    
    'ucrtVSFG.PonerNum
    If VSFG.Rows > 1 Then
        VSFG_AfterRowColChange 0, 0, 1, 1
    Else
        strContenedor = ""
    End If
End Sub

Private Sub cmdReubicarContenedor_Click()
    If VSFG.Row > 0 And VSFG.TextMatrix(VSFG.Row, 0) <> "" Then
        frmReubicarContenedorMercaderia.txtCodigo.Text = VSFG.TextMatrix(VSFG.Row, 0)
        frmReubicarContenedorMercaderia.dtpFecha.Value = VSFG.TextMatrix(VSFG.Row, 1)
        frmReubicarContenedorMercaderia.txtBodega.Text = VSFG.Cell(flexcpTextDisplay, VSFG.Row, 3)
        frmReubicarContenedorMercaderia.txtUbicacion.Text = VSFG.TextMatrix(VSFG.Row, 4)
        frmReubicarContenedorMercaderia.TxtObser.Text = VSFG.TextMatrix(VSFG.Row, 5)
        frmReubicarContenedorMercaderia.Show
    End If
End Sub

Private Sub cmdTransferirPrendas_Click()
    If strContenedor <> "" Then
        frmTransferirPrendasContenedorMercaderia.Show
        frmTransferirPrendasContenedorMercaderia.txtCodigoOrigen.Text = strContenedor
    End If
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or strContenedor = "" Then
        strContenedor = VSFG.TextMatrix(NewRow, 0)
        clsContenedor.SetContenedor strContenedor
        CargaHistorial
        CargaDetalle
        CargaCaracteristicas
    End If
    
End Sub

Private Sub borrarImagenes()
    Dim i As Long
    On Error Resume Next
    For i = 1 To (vsfgCaracteristica.Rows - 1) * 10
        Kill Trim(App.Path) & "\v" & i & ".jpeg"
    Next i
End Sub

Private Sub CargaCaracteristicas()
    Dim numarchivos As Integer
    Dim numarchivosJ As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Arch As String
    Dim ArchJ As String
    Dim rs As New ADODB.Recordset
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    
    mystream.Type = adTypeBinary
    borrarImagenes
    vsfgCaracteristica.Rows = 1
    rs.Open " SELECT con_mer_car_tipo,con_mer_car_numero,con_mer_car_observacion,con_mer_car_foto " & _
             " FROM contenedor_mercaderia_caracteristica " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND con_mer_codigo LIKE  '" & strContenedor & "%'" & _
             " ORDER BY CASE con_mer_car_tipo WHEN 'A' THEN 0 " & _
             " WHEN 'M' THEN 1 " & _
             " WHEN 'C' THEN 2 " & _
             " WHEN 'I' THEN 3 END ", AdoConn
    
    If rs.RecordCount > 0 Then
        Width = 19560
        mystream.Open
        
        i = 0
        While Not rs.EOF
            i = i + 1
            vsfgCaracteristica.AddItem rs!con_mer_car_tipo & vbTab & rs!con_mer_car_numero & vbTab & rs!con_mer_car_observacion
            If IsNull(rs!con_mer_car_foto) = False Then
                mystream.Write rs!con_mer_car_foto
                Arch = App.Path & "\v" & i & ".jpeg"
                mystream.SaveToFile Arch, adSaveCreateOverWrite
                vsfgCaracteristica.TextMatrix(i, 3) = Arch
            End If
            rs.MoveNext
        Wend
        mystream.Close
    Else
        Width = 13890
    End If
    rs.Close
End Sub

Private Sub CargaHistorial()
    
    strSql = " SELECT con_mer_his_fecha,dep_codigo,ubi_bod_codigo,con_mer_his_observacion,con_mer_his_fechamod,con_mer_his_usumod " & _
             " FROM contenedor_mercaderia_historia " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND con_mer_codigo LIKE  '" & strContenedor & "%'" & _
             " ORDER BY con_mer_his_fecha "
    clsCon_Def.Ejecutar strSql
    Set VSFG3.DataSource = clsCon_Def.adorec_Def.DataSource
    
    strSql = " SELECT dep_codigo as cod, dep_nombre as nomb" & _
             " FROM deposito " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By nomb "
    clsCon_Def.Ejecutar (strSql)
    
    VSFG3.ColComboList(1) = VSFG3.BuildComboList(clsCon_Def.adorec_Def, "cod, *nomb", "cod")
    
End Sub

Private Sub CargaDetalle()
    Set VSFG1.DataSource = clsContenedor.adorec_DetalleContenedor
    If VSFG1.Rows > 1 Then
        VSFG1_AfterRowColChange 0, 0, 1, 1
    Else
        strProducto = ""
    End If
    
End Sub

Private Sub CargaKardex()
    
    strSql = " SELECT det_con_mer_fecha,tip_mov_codigo,mov_codigo,con_mer_codigo_origen,con_mer_codigo_destino,IIF(con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad as tot,det_con_mer_fechamod,det_con_mer_usumod " & _
             " FROM det_contenedor_mercaderia " & _
             " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_contenedor_mercaderia.con_mer_codigo LIKE  '" & strContenedor & "%'" & _
             " AND det_contenedor_mercaderia.prd_codigo LIKE  '" & strProducto & "%'" & _
             " ORDER BY det_con_mer_fecha,tot "
    clsCon_Def.Ejecutar strSql
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    
    
    strSql = " SELECT tip_ing_codigo as cod, tip_ing_nombre as nomb" & _
             " FROM tipo_ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " UNION " & _
             " SELECT tip_egr_codigo, tip_egr_nombre " & _
             " FROM tipo_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By nomb "
    clsCon_Def.Ejecutar (strSql)
    
    VSFG2.ColComboList(1) = VSFG2.BuildComboList(clsCon_Def.adorec_Def, "cod, *nomb", "cod")
    
End Sub

Private Sub VSFG_DblClick()
    Dim i As Long
    Set DAT = New frmDatos
    If VSFG.Row >= 1 Then
        DAT.Show
        DAT.VSFG.Rows = VSFG.Cols
        For i = 1 To VSFG.Cols - 1
            DAT.VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(0, i)
            DAT.VSFG.Cell(flexcpText, i, 1) = VSFG.Cell(flexcpTextDisplay, VSFG.Row, i)
            If VSFG.ColComboList(i) <> "" Then
                DAT.VSFG.TextMatrix(i, 2) = VSFG.ColComboList(i)
                DAT.VSFG.Cell(flexcpText, i, 3) = VSFG.Cell(flexcpText, VSFG.Row, i)
            End If
        Next i
        DAT.VSFG.Cell(flexcpBackColor, 1, 1, DAT.VSFG.Rows - 1, 1) = VSFG.Cell(flexcpBackColor, VSFG.Row, VSFG.Col)
        DAT.VSFG.RowHidden(DAT.VSFG.Rows - 1) = True
        Set DAT.VSFGOrigen = VSFG
        DAT.VSFGOrigen.Tag = VSFG.Row
        DAT.Caption = Tipo
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.Value = 1 Then
        txtCodigo.Enabled = True
    Else
        txtCodigo.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    borrarImagenes
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Width = 13890
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsContenedor.Inicializar AdoConn, AdoConnMaster
    IniDato
    If ContenedorInventario = False Then
        txtCodigo.Text = Format(HoyDia, "YYmm")
    Else
        txtCodigo.Text = "0000"
    End If
    
    Carga
    
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbBodega.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbBodega.ListField = "dep_nombre"
    cmbBodega.BoundColumn = "dep_codigo"
    
    
    strSql = " SELECT est_con_mer_codigo, est_con_mer_descripcion " & _
             " FROM est_contenedor_mercaderia " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbEstado.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbEstado.ListField = "est_con_mer_descripcion"
    cmbEstado.BoundColumn = "est_con_mer_codigo"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or strProducto = "" Then
        strProducto = VSFG1.TextMatrix(NewRow, 0)
        CargaKardex
    End If
End Sub

Private Sub vsfgCaracteristica_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If OldRow <> NewRow And NewRow > 0 Then
        Set imgPic.Picture = Nothing
        
        If NewRow >= 3 Then
            
            If fso.FileExists(Trim(App.Path) & "\v" & NewRow & ".jpeg") = True Then
                imgPic.Picture = LoadPicture(Trim(App.Path) & "\v" & NewRow & ".jpeg")
            End If
            
        End If
    End If
End Sub


