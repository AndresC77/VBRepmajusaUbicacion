VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmContenedorMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contenedores de Mercaderia"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19980
   Icon            =   "frmContenedorMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   19980
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4305
      Left            =   16920
      Picture         =   "frmContenedorMercaderia.frx":030A
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.CommandButton cmdSeleccionarImagen 
      Caption         =   "Seleccionar Imagen"
      Height          =   360
      Left            =   17760
      TabIndex        =   38
      Top             =   2760
      Width           =   1700
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "&Explorar..."
      Height          =   360
      Left            =   15960
      TabIndex        =   37
      Top             =   2760
      Width           =   1700
   End
   Begin VSFlex8LCtl.VSFlexGrid vsfgCaracteristica 
      Height          =   6960
      Left            =   10200
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   5580
      _cx             =   9842
      _cy             =   12277
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   6000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmContenedorMercaderia.frx":59A3
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
      Editable        =   2
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
   Begin VB.CommandButton cmdCaracteristicas 
      Caption         =   "Caracteristicas"
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      Top             =   2355
      Width           =   1455
   End
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   5160
      TabIndex        =   32
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chkResta 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Resta"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   2415
      Width           =   1335
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   8160
      TabIndex        =   30
      Top             =   2355
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Orden de Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   9945
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   2400
      End
      Begin VB.TextBox txtObservacionOrdenCompra 
         Height          =   645
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   4800
      End
      Begin MSDataListLib.DataCombo cmbOrdenCompra 
         Height          =   330
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ord. Compra:"
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
         TabIndex        =   29
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Left            =   660
         TabIndex        =   28
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGCarga 
      Height          =   975
      Left            =   7320
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
      _cx             =   5106
      _cy             =   1720
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
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
   Begin VB.TextBox txtCopias 
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Text            =   "1"
      Top             =   7200
      Width           =   495
   End
   Begin VB.CommandButton cmdAnularContenedor 
      Caption         =   "&Anular Contenedor"
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.CommandButton cmdCrearContenedor 
      Caption         =   "&Ingresar Contenedor"
      Height          =   360
      Left            =   105
      TabIndex        =   2
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   11400
      TabIndex        =   1
      Top             =   7200
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Contenedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.CheckBox chkInventario 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Inventario"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8400
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   248
         Width           =   1815
      End
      Begin VB.TextBox TxtObser 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   8400
      End
      Begin MSDataListLib.DataCombo cmbBodega 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo cmbUbicacion 
         Height          =   315
         Left            =   4755
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
         Left            =   4755
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   66584579
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   7200
         TabIndex        =   14
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   6780
         TabIndex        =   16
         Top             =   645
         Width           =   345
      End
      Begin VB.Label LblObser 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         Left            =   270
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contenedor:"
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
         TabIndex        =   13
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Creación:"
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
         Left            =   3495
         TabIndex        =   12
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   3930
         TabIndex        =   9
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega:"
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
         Left            =   645
         TabIndex        =   7
         Top             =   645
         Width           =   600
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4320
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   9945
      _cx             =   17542
      _cy             =   7620
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmContenedorMercaderia.frx":5A63
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
      FrozenRows      =   1
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGRes 
      Height          =   6960
      Left            =   10200
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   3045
      _cx             =   5371
      _cy             =   12277
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmContenedorMercaderia.frx":5B06
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
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   9360
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG1 
      Height          =   6960
      Left            =   10200
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   5580
      _cx             =   9842
      _cy             =   12277
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
      SelectionMode   =   1
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
      FormatString    =   $"frmContenedorMercaderia.frx":5B6C
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16920
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgPic 
      BorderStyle     =   1  'Fixed Single
      Height          =   2500
      Left            =   15840
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4000
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   4440
      TabIndex        =   33
      Top             =   2430
      Width           =   555
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   7200
      Width           =   3435
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#Copias:"
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
      Left            =   3675
      TabIndex        =   18
      Top             =   7275
      Width           =   645
   End
End
Attribute VB_Name = "frmContenedorMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private anchoAnteriorCarac As Long
Private anchoAnteriorInvent As Long
Private anchoAnteriorOC As Long

Private Sub chkInventario_Click()
    If chkInventario.Value = 0 Then
        ContenedorInventario = False
        lblEstado.Caption = "--"
        VSFGRes.Visible = False
        cmbOrdenCompra.BoundText = ""
        txtEstado.Text = ""
        txtObservacionOrdenCompra.Text = ""
        cmbOrdenCompra.Locked = False
        'Me.Width = 10275
        Width = anchoAnteriorInvent
    Else
        ContenedorInventario = True
        lblEstado.Caption = "INVENTARIO"
        VSFGRes.Visible = True
        cmbOrdenCompra.BoundText = ""
        txtEstado.Text = ""
        txtObservacionOrdenCompra.Text = ""
        cmbOrdenCompra.Locked = True
        VSFG1.Visible = False
        anchoAnteriorInvent = Me.Width
        Me.Width = 13440
    End If
    
End Sub

Private Sub cmbBodega_Validate(Cancel As Boolean)
    CargaUbica
End Sub

Private Sub CargaUbica()
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & cmbBodega.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsCon_Def.Ejecutar strSql
    Set cmbUbicacion.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbUbicacion.ListField = "ubi_bod_codigo"
    cmbUbicacion.BoundColumn = "ubi_bod_codigo"
End Sub

Private Sub CargaOrdenCompra()
    Dim clsAux As New clsConsulta
    Dim codigoEstado As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    If cmbOrdenCompra.BoundText <> "" Then
        VSFG1.Visible = True
        anchoAnteriorOC = Width
        Me.Width = 15945
        strSql = " SELECT ord_com_codigo,est_ord_com_descripcion,ord_com_observacion " & _
                 " FROM orden_compra INNER JOIN est_orden_compra ON orden_compra.est_ord_com_codigo=est_orden_compra.est_ord_com_codigo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & cmbOrdenCompra.BoundText & "'"
        clsAux.Ejecutar strSql
        If clsAux.adorec_Def.RecordCount > 0 Then
            txtEstado.Text = clsAux.adorec_Def("est_ord_com_descripcion")
            txtObservacionOrdenCompra.Text = clsAux.adorec_Def("ord_com_observacion")
        Else
            txtEstado.Text = ""
            txtObservacionOrdenCompra.Text = ""
        End If
    Else
        VSFG1.Visible = False
'        If chkInventario.Value = 0 Then
'            Me.Width = 10275
'        Else
'            Me.Width = 13440
'        End If
        Width = anchoAnteriorOC
    End If
    
    
End Sub


Private Sub cmbOrdenCompra_Validate(Cancel As Boolean)
    If cmbOrdenCompra.BoundText <> "" Then
        CargaOrdenCompra
        CargarDetalle
    End If
End Sub

Private Sub CargarDetalle()
    Dim strContenedores As String
    Dim i As Long
    VSFG1.Clear 1
    VSFG1.Rows = 3
    If cmbOrdenCompra.BoundText <> "" Then
        strSql = " SELECT producto.prd_codigo,prd_nombre,COALESCE(cantContenedores,0),det_ord_com_cantidad,COALESCE(cantRecibida,0),det_ord_com_cantidad-(COALESCE(cantContenedores,0)+COALESCE(cantRecibida,0)) " & _
                 " FROM det_orden_compra INNER JOIN preproducto_producto ON det_orden_compra.emp_codigo=preproducto_producto.emp_codigo " & _
                 " AND det_orden_compra.pre_codigo=preproducto_producto.pre_codigo " & _
                 " AND det_orden_compra.col_codigo=preproducto_producto.col_codigo " & _
                 " AND det_orden_compra.tal_codigo=preproducto_producto.tal_codigo " & _
                 " INNER JOIN producto " & _
                 " ON preproducto_producto.emp_codigo=producto.emp_codigo " & _
                 " AND preproducto_producto.prd_codigo=producto.prd_codigo "
        strSql = strSql & " LEFT JOIN (" & _
                 " SELECT contenedor_mercaderia.emp_codigo,prd_codigo,SUM(det_con_mer_cantidad) as cantContenedores " & _
                 " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                 " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND contenedor_mercaderia.ord_com_codigo='" & cmbOrdenCompra.BoundText & "'" & _
                 " GROUP BY contenedor_mercaderia.emp_codigo,prd_codigo " & _
                 " ) cm ON producto.emp_codigo=cm.emp_codigo AND producto.prd_codigo=cm.prd_codigo"
        strSql = strSql & " LEFT JOIN (" & _
                 " SELECT recepcion_mercaderia.emp_codigo,prd_codigo,SUM(det_con_mer_cantidad) as cantRecibida " & _
                 " FROM recepcion_mercaderia INNER JOIN det_recepcion_mercaderia " & _
                 " ON recepcion_mercaderia.emp_codigo=det_recepcion_mercaderia.emp_codigo " & _
                 " AND recepcion_mercaderia.rec_mer_codigo=det_recepcion_mercaderia.rec_mer_codigo " & _
                 " INNER JOIN contenedor_mercaderia ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                 " WHERE recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND recepcion_mercaderia.ord_com_codigo='" & cmbOrdenCompra.BoundText & "' " & _
                 " GROUP BY recepcion_mercaderia.emp_codigo,prd_codigo " & _
                 " ) rm ON producto.emp_codigo=rm.emp_codigo AND producto.prd_codigo=rm.prd_codigo"
        strSql = strSql & " WHERE det_orden_compra.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_orden_compra.ord_com_codigo='" & cmbOrdenCompra.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
    End If
    VSFG1.MergeCells = flexMergeRestrictRows
    VSFG1.MergeCol(0) = True
    VSFG1.Subtotal flexSTSum, -1, 2, , vbYellow, , True, "TOTAL"
    VSFG1.Subtotal flexSTSum, -1, 3, , vbYellow, , True, "TOTAL"
    VSFG1.Subtotal flexSTSum, -1, 4, , vbYellow, , True, "TOTAL"
    'VSFG1.Subtotal flexSTSum, -1, 7, , vbYellow, , True, "TOTAL"
End Sub

Private Sub cmbUbicacion_Change()
    cmbUbicacion.Text = Replace(cmbUbicacion.Text, "'", "")
End Sub

Private Sub cmdCaracteristicas_Click()
    If cmdCaracteristicas.Caption = "Quitar Caract." Then
        Width = anchoAnteriorCarac
        vsfgCaracteristica.Visible = False
        cmdCaracteristicas.Caption = "Caracteristicas"
        
    Else
        anchoAnteriorCarac = Width
        Width = 20070
        vsfgCaracteristica.Visible = True
        cmdCaracteristicas.Caption = "Quitar Caract."
        vsfgCaracteristica.Rows = 1
        vsfgCaracteristica.AddItem "A"
        vsfgCaracteristica.AddItem "M"
        vsfgCaracteristica.AddItem "C"
    End If
End Sub

Private Sub cmdCargar_Click()
    Dim sDir As String
    Dim strCodigo As String
    Dim dblCantidad As Double
    Dim i As Long
    sDir = CurDir
    cdArchivo.ShowOpen
    If cdArchivo.FileName <> "" Then
         VSFGCarga.LoadGrid cdArchivo.FileName, flexFileTabText
         If VSFGCarga.Rows > 1 Then
            For i = 1 To VSFGCarga.Rows - 1
                If Trim(VSFGCarga.TextMatrix(i, 0)) <> "" And FormatoD4(VSFGCarga.TextMatrix(i, 0)) <> 0 Then
                    AgregarProd Trim(VSFGCarga.TextMatrix(i, 0)), , VSFGCarga.TextMatrix(i, 1)
                End If
            Next i
         End If
    End If
    ChDir sDir
End Sub

Private Sub cmdCrearContenedor_Click()
    Dim i As Long
    Dim strCont As String
    Dim clsCont As New clsContenedor
    clsCont.Inicializar AdoConn, AdoConnMaster
    If cmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        If FormatoD0(VSFG.TextMatrix(1, 3)) > 0 Then
            clsCont.NuevoContenedor strSucursal, strPtoFactura, dtpFecha.Value, 0, cmbBodega.BoundText, cmbUbicacion.BoundText, TxtObser.Text, 0, cmbTipo.BoundText, ContenedorInventario, cmbOrdenCompra.BoundText
            For i = 2 To VSFG.Rows - 1
                clsCont.AgregarDetalle VSFG.TextMatrix(i, 0), "", clsCont.strNumContenedor, dtpFecha.Value, VSFG.TextMatrix(i, 3)
            Next i
        End If
        For i = 1 To vsfgCaracteristica.Rows - 1
            clsCont.AgregarDetalleCaracteristica vsfgCaracteristica.TextMatrix(i, 0), FormatoD4(vsfgCaracteristica.TextMatrix(i, 1)), vsfgCaracteristica.TextMatrix(i, 2), vsfgCaracteristica.TextMatrix(i, 3)
'            If i >= 3 And vsfgCaracteristica.TextMatrix(i, 3) <> "" Then
'                GuardaFoto i, clsCont
'            End If
        Next i
        clsCont.ImprimirSTK
        clsCont.ImprimirLista FormatoD0(txtCopias.Text)
        MsgBox "Contenedor " & clsCont.strNumContenedor & " ingresado"
        Unload Me
    Else
        MsgBox "No tiene registrada la ubicacion"
    End If
End Sub

Private Sub cmdExplorar_Click()
    Dim factor As Double
    Dim anchoPic As Long
    Dim altoPic As Long
    Dim anchoImg As Long
    Dim altoImg As Long
    cdArchivo.ShowOpen
    If cdArchivo.FileName <> "" Then
        pic.Picture = LoadPicture(cdArchivo.FileName)
        anchoPic = pic.Width
        altoPic = pic.Height
        anchoImg = 4000
        altoImg = 2500
        If anchoImg / anchoPic > altoImg / altoPic Then
            factor = altoImg / altoPic
        Else
            factor = anchoImg / anchoPic
        End If
        pic.PaintPicture pic.Picture, 0, 0, FormatoD0(anchoPic / factor), FormatoD0(altoPic / factor)
        imgPic.Width = FormatoD0(anchoPic * factor)
        imgPic.Height = FormatoD0(altoPic * factor)
        imgPic.Picture = pic.Picture
    End If
End Sub

Private Sub cmdSeleccionarImagen_Click()
    If imgPic.Picture <> 0 Then
        SavePicture imgPic.Picture, Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg"
        'frmContenedorMercaderia.vsfgCaracteristica.TextMatrix(frmContenedorMercaderia.vsfgCaracteristica.Row, frmContenedorMercaderia.vsfgCaracteristica.Col) = Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg"
        'vsfgCaracteristica.CellPicture = imgPic.Picture
        vsfgCaracteristica.Cell(flexcpPicture, vsfgCaracteristica.Row, 3) = LoadPicture(Trim(App.Path) & "\" & vsfgCaracteristica.Row & ".jpeg")
        vsfgCaracteristica.TextMatrix(vsfgCaracteristica.Row, 3) = Trim(App.Path) & "\" & vsfgCaracteristica.Row & ".jpeg"
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Long
    If chkInventario.Value = 1 Then
        strSql = " SELECT det_con_mer_usumod as Usuario, sum(det_con_mer_cantidad)  as Cantidad " & _
                 " FROM contenedor_mercaderia inner join det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE contenedor_mercaderia.con_mer_codigo>='18000000000' " & _
                 " AND FORMAT(det_contenedor_mercaderia.det_con_mer_fecha,'yyyy-MM-dd')=FORMAT(CURRENT_TIMESTAMP,'yyyy-MM-dd') " & _
                 " GROUP BY det_con_mer_usumod " & _
                 " ORDER BY Cantidad DESC "
        clsCon_Def.Ejecutar strSql
        Set VSFGRes.DataSource = clsCon_Def.adorec_Def.DataSource
        For i = 1 To VSFGRes.Rows - 1
            VSFGRes.TextMatrix(i, 0) = i
        Next i
        VSFGRes.Subtotal flexSTSum, -1, 2, "###0", vbBlue, vbWhite
    End If
    If ContenedorInventario = False Then
        chkInventario.Value = 0
        lblEstado.Caption = "--"
    Else
        chkInventario.Value = 1
        lblEstado.Caption = "INVENTARIO"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 1 To (vsfgCaracteristica.Rows - 1) * 10
        Kill Trim(App.Path) & "\" & i & ".jpeg"
    Next i
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
    Me.Width = 10275
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT ord_com_codigo " & _
             " FROM orden_compra " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND est_ord_com_codigo BETWEEN 0 and 19"
    clsCon_Def.Ejecutar strSql
    cmbOrdenCompra.ListField = "ord_com_codigo"
    cmbOrdenCompra.BoundColumn = "ord_com_codigo"
    
    Set cmbOrdenCompra.RowSource = clsCon_Def.adorec_Def.DataSource
    
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbBodega.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbBodega.ListField = "dep_nombre"
    cmbBodega.BoundColumn = "dep_codigo"
    cmbBodega.BoundText = "PRI"
    cmbBodega_Validate False
    cmbUbicacion.BoundText = "0000"
    
    strSql = " SELECT tip_mer_con_codigo, tip_mer_con_nombre " & _
             " FROM tipo_mercaderia_contenedor " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbTipo.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbTipo.ListField = "tip_mer_con_nombre"
    cmbTipo.BoundColumn = "tip_mer_con_codigo"
    
    cmbTipo.BoundText = "0"
    
    dtpFecha.Value = Ahora
    VSFG.SubtotalPosition = flexSTAbove
    VSFG.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFG.Cell(flexcpFontSize, 1, 0, 1, VSFG.Cols - 1) = VSFG.Cell(flexcpFontSize, 1, 0, 1, VSFG.Cols - 1) + 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProd UCase(txtLector.Text), chkResta.Value
        txtLector.Text = ""
        chkResta.Value = False
    ElseIf KeyCode = vbKeyF6 Then
        cmdCargar.Visible = Not cmdCargar.Visible
    End If
End Sub

Private Sub AgregarProd(codigo As String, Optional Resta As Boolean = False, Optional Cantidad As Double = 1)
    Dim i As Long
    Dim pas As Boolean
    pas = False

    With VSFG1
        If cmbOrdenCompra.BoundText <> "" Then
            For i = 1 To .Rows - 1
                If codigo = .TextMatrix(i, 0) Then
                    .ShowCell i, 2
                    .Select i, 0
                    If Resta = False Then
                        .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) + Cantidad
                    Else
                        .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) - Cantidad
                    End If
                    pas = True
                    Exit For
                End If
            Next i
            If pas = False Then
                MsgBox "El producto no corresponde a la orden de compra seleccionada", vbCritical, "Orden de Compra"
                Exit Sub
            Else
                pas = False
            End If
        End If
    End With
InicioFor:
    With VSFG
        For i = 1 To .Rows - 1
            If codigo = .TextMatrix(i, 0) Then
                .ShowCell i, 0
                .Select i, 0
                If Resta = False Then
                    .TextMatrix(i, 3) = Val(Format(.TextMatrix(i, 3), "###0")) + Cantidad
                Else
                    .TextMatrix(i, 3) = Val(Format(.TextMatrix(i, 3), "###0")) - Cantidad
                End If
                pas = True
                Exit For
            ElseIf Trim(.TextMatrix(i, 0)) = "" Then
                .RemoveItem i
                GoTo InicioFor
            End If
        Next i
        If pas = False Then
            strSql = " SELECT prd_codigo, prd_nombre,prd_peso,COALESCE(clc_nombre,'') as colecc " & _
                     " FROM producto LEFT JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo AND  producto.clc_codigo=coleccion.clc_codigo" & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & codigo & "'"
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                If FormatoD4(clsCon_Def.adorec_Def("prd_peso")) <= 0 And chkInventario.Value = 0 Then
'                    frmDefinePeso.txtCodigo = clsCon_Def.adorec_Def("prd_codigo")
'                    frmDefinePeso.txtNombre = clsCon_Def.adorec_Def("prd_nombre")
'                    frmDefinePeso.Show vbModal
                End If
                If Resta = False Then
                    .AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & clsCon_Def.adorec_Def("colecc") & vbTab & Cantidad
                Else
                    .AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & clsCon_Def.adorec_Def("colecc") & vbTab & -1 * Cantidad
                End If
                .ShowCell .Rows - 1, 0
            Else
                MsgBox "El producto no existe en la base de datos." & vbNewLine & _
                       "No se ingresara.", vbInformation, "Productos"
            End If
        End If
        
        .Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
        '.ShowCell 1, 2
    End With
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub VSFG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow <> 0 And NewRow < VSFG1.Rows - 1 Then
        VSFG1.BackColorSel = IIf(VSFG1.TextMatrix(NewRow, 5) < 0, vbMagenta, vbBlue)
    ElseIf NewRow = VSFG1.Rows - 1 Then
        VSFG1.BackColorSel = vbGreen
    End If
End Sub

Private Sub VSFG1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 And VSFG1.TextMatrix(Row, 3) <> "" And VSFG1.TextMatrix(Row, 4) <> "" Then
        VSFG1.TextMatrix(Row, 5) = FormatoD2(VSFG1.TextMatrix(Row, 3)) - (FormatoD2(VSFG1.TextMatrix(Row, 2)) + FormatoD2(VSFG1.TextMatrix(Row, 4)))
        
        VSFG1.Subtotal flexSTSum, -1, 2, , vbYellow, , True, "TOTAL"
        VSFG1.Subtotal flexSTSum, -1, 3, , vbYellow, , True, "TOTAL"
        VSFG1.Subtotal flexSTSum, -1, 4, , vbYellow, , True, "TOTAL"
        'VSFG1.Subtotal flexSTSum, -1, 7, , vbYellow, , True, "TOTAL"
    End If
    If Col = 5 And Row > 0 Then
        VSFG1.Cell(flexcpBackColor, Row, 0, Row, VSFG1.Cols - 1) = IIf(VSFG1.TextMatrix(Row, 5) < 0, vbRed, vbWhite)
        VSFG1.Cell(flexcpForeColor, Row, 0, Row, VSFG1.Cols - 1) = IIf(VSFG1.TextMatrix(Row, 5) < 0, vbWhite, vbBlack)
        VSFG1.Cell(flexcpForeColor, Row, 0, Row, VSFG1.Cols - 1) = IIf(VSFG1.TextMatrix(Row, 5) < 0, vbWhite, vbBlack)
    End If
End Sub

Private Sub vsfgCaracteristica_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim fSO As Object
    
    Set fSO = CreateObject("Scripting.FileSystemObject")
    If OldRow <> NewRow Then
        Set imgPic.Picture = Nothing
        
        If NewRow >= 3 Then
            cmdExplorar.Enabled = True
            cmdSeleccionarImagen.Enabled = True
            
            If fSO.FileExists(Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg") = True Then
                imgPic.Picture = LoadPicture(Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg")
            End If
            
        Else
            cmdExplorar.Enabled = False
            cmdSeleccionarImagen.Enabled = False
        End If
    End If
End Sub


Private Sub vsfgCaracteristica_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgCaracteristica_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim i As Long
    Dim Car As String
End Sub

Private Sub VSFGCaracteristica_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 Then
        vsfgCaracteristica.TextMatrix(Row, Col) = UCase(vsfgCaracteristica.TextMatrix(Row, Col))
    End If
    If vsfgCaracteristica.TextMatrix(vsfgCaracteristica.Rows - 1, 0) <> "" And (vsfgCaracteristica.TextMatrix(vsfgCaracteristica.Rows - 1, 1) <> "" Or vsfgCaracteristica.TextMatrix(vsfgCaracteristica.Rows - 1, 2) <> "") Then
        vsfgCaracteristica.AddItem "I"
    End If
End Sub
