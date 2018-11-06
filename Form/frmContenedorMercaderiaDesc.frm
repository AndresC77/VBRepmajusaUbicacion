VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContenedorMercaderiaDesc 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contenedores de Mercaderia"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   Icon            =   "frmContenedorMercaderiaDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12960
   Begin VB.CommandButton cmdCrearContenedorVerde 
      Caption         =   "&Crear Contenedor Verde"
      Height          =   360
      Left            =   6600
      TabIndex        =   22
      Top             =   6480
      Width           =   3495
   End
   Begin VB.CommandButton cmdCrearContenedorAmarillo 
      Caption         =   "&Crear Contenedor Amarillo"
      Height          =   360
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Width           =   3495
   End
   Begin VB.CommandButton cmdCrearContenedorRojo 
      Caption         =   "&Crear Contenedor Rojo"
      Height          =   360
      Left            =   6600
      TabIndex        =   20
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CheckBox chkResta 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Resta"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   10320
      TabIndex        =   14
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCrearContenedorBlanco 
      Caption         =   "&Crear Contenedor Banco"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   11040
      TabIndex        =   1
      Top             =   6480
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
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
         TabIndex        =   5
         Top             =   930
         Width           =   6000
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
         Width           =   2535
         _ExtentX        =   4471
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
         Format          =   66125827
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1320
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
         Left            =   900
         TabIndex        =   25
         Top             =   1365
         Width           =   345
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGBlanco 
      Height          =   2040
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   6225
      _cx             =   10980
      _cy             =   3598
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   14737632
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
      FormatString    =   $"frmContenedorMercaderiaDesc.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGRojo 
      Height          =   2040
      Left            =   6600
      TabIndex        =   17
      Top             =   1920
      Width           =   6225
      _cx             =   10980
      _cy             =   3598
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   255
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   255
      BackColorAlternate=   14737632
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
      FormatString    =   $"frmContenedorMercaderiaDesc.frx":038F
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGAmarillo 
      Height          =   2040
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   6225
      _cx             =   10980
      _cy             =   3598
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   65535
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   65535
      BackColorAlternate=   14737632
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
      FormatString    =   $"frmContenedorMercaderiaDesc.frx":0414
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGVerde 
      Height          =   2040
      Left            =   6600
      TabIndex        =   19
      Top             =   4440
      Width           =   6225
      _cx             =   10980
      _cy             =   3598
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   65280
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   65280
      BackColorAlternate=   14737632
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
      FormatString    =   $"frmContenedorMercaderiaDesc.frx":0499
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
   Begin VB.Label lblTipoContenedor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7920
      TabIndex        =   23
      Top             =   120
      Width           =   4815
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
      Left            =   9600
      TabIndex        =   15
      Top             =   1155
      Width           =   555
   End
End
Attribute VB_Name = "frmContenedorMercaderiaDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String

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

Private Sub cmbUbicacion_Change()
    cmbUbicacion.Text = Replace(cmbUbicacion.Text, "'", "")
End Sub

Private Sub cmdCrearContenedorAmarillo_Click()
    Dim i As Long
    Dim strCont As String
    Dim clsCont As New clsContenedor
    clsCont.Inicializar AdoConn, AdoConnMaster
    If cmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        If FormatoD0(VSFGAmarillo.TextMatrix(1, 2)) > 0 Then
            clsCont.NuevoContenedor strSucursal, strPtoFactura, dtpFecha.Value, 0, cmbBodega.BoundText, cmbUbicacion.BoundText, TxtObser.Text, 2, cmbTipo.BoundText
            For i = 2 To VSFGAmarillo.Rows - 1
                clsCont.AgregarDetalle VSFGAmarillo.TextMatrix(i, 0), "", clsCont.strNumContenedor, dtpFecha.Value, VSFGAmarillo.TextMatrix(i, 2)
            Next i
        End If
        clsCont.ImprimirSTK
        clsCont.ImprimirLista 2
        Set clsCont = Nothing
        VSFGAmarillo.Clear 1
        VSFGAmarillo.Rows = 2
        VSFGAmarillo.Subtotal flexSTSum, -1, 2, , vbYellow, vbBlack, True, "TOTAL Amarillo"
    Else
        MsgBox "No tiene registrada la ubicacion"
    End If
End Sub

Private Sub cmdCrearContenedorBlanco_Click()
    Dim i As Long
    Dim strCont As String
    Dim clsCont As New clsContenedor
    clsCont.Inicializar AdoConn, AdoConnMaster
    If cmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        If FormatoD0(VSFGBlanco.TextMatrix(1, 2)) > 0 Then
            clsCont.NuevoContenedor strSucursal, strPtoFactura, dtpFecha.Value, 0, cmbBodega.BoundText, cmbUbicacion.BoundText, TxtObser.Text, 0, cmbTipo.BoundText
            For i = 2 To VSFGBlanco.Rows - 1
                clsCont.AgregarDetalle VSFGBlanco.TextMatrix(i, 0), "", clsCont.strNumContenedor, dtpFecha.Value, VSFGBlanco.TextMatrix(i, 2)
            Next i
        End If
        clsCont.ImprimirSTK
        clsCont.ImprimirLista 2
        Set clsCont = Nothing
        VSFGBlanco.Clear 1
        VSFGBlanco.Rows = 2
        VSFGBlanco.Subtotal flexSTSum, -1, 2, , vbWhite, vbBlack, True, "TOTAL Blanco"
    Else
        MsgBox "No tiene registrada la ubicacion"
    End If
End Sub

Private Sub cmdCrearContenedorRojo_Click()
    Dim i As Long
    Dim strCont As String
    Dim clsCont As New clsContenedor
    clsCont.Inicializar AdoConn, AdoConnMaster
    If cmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        If FormatoD0(VSFGRojo.TextMatrix(1, 2)) > 0 Then
            clsCont.NuevoContenedor strSucursal, strPtoFactura, dtpFecha.Value, 0, cmbBodega.BoundText, cmbUbicacion.BoundText, TxtObser.Text, 3, cmbTipo.BoundText
            For i = 2 To VSFGRojo.Rows - 1
                clsCont.AgregarDetalle VSFGRojo.TextMatrix(i, 0), "", clsCont.strNumContenedor, dtpFecha.Value, VSFGRojo.TextMatrix(i, 2)
            Next i
        End If
        clsCont.ImprimirSTK
        clsCont.ImprimirLista 2
        Set clsCont = Nothing
        VSFGRojo.Clear 1
        VSFGRojo.Rows = 2
        VSFGRojo.Subtotal flexSTSum, -1, 2, , vbRed, vbWhite, True, "TOTAL Rojo"
    Else
        MsgBox "No tiene registrada la ubicacion"
    End If
End Sub

Private Sub cmdCrearContenedorVerde_Click()
    Dim i As Long
    Dim strCont As String
    Dim clsCont As New clsContenedor
    clsCont.Inicializar AdoConn, AdoConnMaster
    If cmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        If FormatoD0(VSFGVerde.TextMatrix(1, 2)) > 0 Then
            clsCont.NuevoContenedor strSucursal, strPtoFactura, dtpFecha.Value, 0, cmbBodega.BoundText, cmbUbicacion.BoundText, TxtObser.Text, 1, cmbTipo.BoundText
            For i = 2 To VSFGVerde.Rows - 1
                clsCont.AgregarDetalle VSFGVerde.TextMatrix(i, 0), "", clsCont.strNumContenedor, dtpFecha.Value, VSFGVerde.TextMatrix(i, 2)
            Next i
        End If
        clsCont.ImprimirSTK
        clsCont.ImprimirLista 2
        Set clsCont = Nothing
        VSFGVerde.Clear 1
        VSFGVerde.Rows = 2
        VSFGVerde.Subtotal flexSTSum, -1, 2, , vbGreen, vbBlack, True, "TOTAL Verde"
    Else
        MsgBox "No tiene registrada la ubicacion"
    End If
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

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
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
    VSFGBlanco.SubtotalPosition = flexSTAbove
    VSFGBlanco.Subtotal flexSTSum, -1, 2, , vbWhite, vbBlack, True, "TOTAL"
    VSFGBlanco.Cell(flexcpFontSize, 1, 0, 1, VSFGBlanco.Cols - 1) = VSFGBlanco.Cell(flexcpFontSize, 1, 0, 1, VSFGBlanco.Cols - 1) + 2
    VSFGVerde.SubtotalPosition = flexSTAbove
    VSFGVerde.Subtotal flexSTSum, -1, 2, , vbGreen, vbBlack, True, "TOTAL"
    VSFGVerde.Cell(flexcpFontSize, 1, 0, 1, VSFGVerde.Cols - 1) = VSFGVerde.Cell(flexcpFontSize, 1, 0, 1, VSFGVerde.Cols - 1) + 2
    VSFGAmarillo.SubtotalPosition = flexSTAbove
    VSFGAmarillo.Subtotal flexSTSum, -1, 2, , vbYellow, vbBlack, True, "TOTAL"
    VSFGAmarillo.Cell(flexcpFontSize, 1, 0, 1, VSFGAmarillo.Cols - 1) = VSFGAmarillo.Cell(flexcpFontSize, 1, 0, 1, VSFGAmarillo.Cols - 1) + 2
    VSFGRojo.SubtotalPosition = flexSTAbove
    VSFGRojo.Subtotal flexSTSum, -1, 2, , vbRed, vbWhite, True, "TOTAL"
    VSFGRojo.Cell(flexcpFontSize, 1, 0, 1, VSFGRojo.Cols - 1) = VSFGRojo.Cell(flexcpFontSize, 1, 0, 1, VSFGRojo.Cols - 1) + 2
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
    End If
End Sub

Private Sub AgregarProd(codigo As String, Optional Resta As Boolean = False)
    Dim i As Long
    Dim pas As Boolean
    Dim valResta As Integer
InicioFor:
    pas = False
    If Resta = False Then
        valResta = 1
    Else
        valResta = -1
    End If
    If pas = False Then
        With VSFGBlanco
            For i = 1 To .Rows - 1
                If codigo = .TextMatrix(i, 0) Then
                    .ShowCell i, 0
                    .Select i, 0
                    .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) + valResta
                    pas = True
                    .Subtotal flexSTSum, -1, 2, , vbWhite, vbBlack, True, "TOTAL Blanco"
                    lblTipoContenedor.BackColor = vbWhite
                    lblTipoContenedor.Caption = .TextMatrix(i, 1)
                    Exit For
                ElseIf Trim(.TextMatrix(i, 0)) = "" Then
                    .RemoveItem i
                    GoTo InicioFor
                End If
            Next i
        End With
    End If
    If pas = False Then
        With VSFGRojo
            For i = 1 To .Rows - 1
                If codigo = .TextMatrix(i, 0) Then
                    .ShowCell i, 0
                    .Select i, 0
                    .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) + valResta
                    .Subtotal flexSTSum, -1, 2, , vbRed, vbWhite, True, "TOTAL Rojo"
                    lblTipoContenedor.BackColor = vbRed
                    lblTipoContenedor.Caption = .TextMatrix(i, 1)
                    pas = True
                    Exit For
                ElseIf Trim(.TextMatrix(i, 0)) = "" Then
                    .RemoveItem i
                    GoTo InicioFor
                End If
            Next i
        End With
    End If
    If pas = False Then
        With VSFGAmarillo
            For i = 1 To .Rows - 1
                If codigo = .TextMatrix(i, 0) Then
                    .ShowCell i, 0
                    .Select i, 0
                    .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) + valResta
                    .Subtotal flexSTSum, -1, 2, , vbYellow, vbBlack, True, "TOTAL Amarillo"
                    lblTipoContenedor.BackColor = vbYellow
                    lblTipoContenedor.Caption = .TextMatrix(i, 1)
                    pas = True
                    Exit For
                ElseIf Trim(.TextMatrix(i, 0)) = "" Then
                    .RemoveItem i
                    GoTo InicioFor
                End If
            Next i
        End With
    End If
    If pas = False Then
        With VSFGVerde
            For i = 1 To .Rows - 1
                If codigo = .TextMatrix(i, 0) Then
                    .ShowCell i, 0
                    .Select i, 0
                    .TextMatrix(i, 2) = Val(Format(.TextMatrix(i, 2), "###0")) + valResta
                    .Subtotal flexSTSum, -1, 2, , vbGreen, vbBlack, True, "TOTAL Verde"
                    lblTipoContenedor.BackColor = vbGreen
                    lblTipoContenedor.Caption = .TextMatrix(i, 1)
                    pas = True
                    Exit For
                ElseIf Trim(.TextMatrix(i, 0)) = "" Then
                    .RemoveItem i
                    GoTo InicioFor
                End If
            Next i
        End With
    End If
    
    If pas = False Then
        strSql = " SELECT producto.prd_codigo, producto.prd_nombre,LEFT(clc_nombre,5) as colec,DATEDIFF(CURRENT_TIMESTAMP,max(ing_fecha)) as dias " & _
                 " FROM producto INNER JOIN det_ingreso " & _
                 " ON producto.emp_codigo=det_ingreso.emp_codigo " & _
                 " AND producto.prd_codigo=det_ingreso.prd_codigo " & _
                 " AND det_ingreso.tip_ing_codigo IN('COM','IIM','AAU','ITN') " & _
                 " INNER JOIN ingreso " & _
                 " ON det_ingreso.emp_codigo=ingreso.emp_codigo " & _
                 " AND det_ingreso.tip_ing_codigo=ingreso.tip_ing_codigo " & _
                 " AND det_ingreso.ing_codigo=ingreso.ing_codigo " & _
                 " AND ingreso.ing_anulado=0 " & _
                 " INNER JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo " & _
                 " AND producto.clc_codigo=coleccion.clc_codigo " & _
                 " WHERE producto.emp_codigo='" & strEmpresa & "'" & _
                 " AND producto.prd_codigo='" & codigo & "'" & _
                 " GROUP BY producto.prd_codigo, producto.prd_nombre,clc_nombre"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If UCase(clsCon_Def.adorec_Def("colec")) <> "SALDO" Then
                VSFGBlanco.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & valResta
                VSFGBlanco.Subtotal flexSTSum, -1, 2, , vbWhite, vbBlack, True, "TOTAL Blanco"
                lblTipoContenedor.BackColor = vbWhite
                lblTipoContenedor.Caption = clsCon_Def.adorec_Def("prd_nombre")
            ElseIf FormatoD0(clsCon_Def.adorec_Def("dias")) < 500 Then
                VSFGVerde.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & valResta
                VSFGVerde.Subtotal flexSTSum, -1, 2, , vbGreen, vbBlack, True, "TOTAL Verde"
                lblTipoContenedor.BackColor = vbGreen
                lblTipoContenedor.Caption = clsCon_Def.adorec_Def("prd_nombre")
            ElseIf 500 <= FormatoD0(clsCon_Def.adorec_Def("dias")) And FormatoD0(clsCon_Def.adorec_Def("dias")) < 1000 Then
                VSFGAmarillo.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & valResta
                VSFGAmarillo.Subtotal flexSTSum, -1, 2, , vbYellow, vbBlack, True, "TOTAL Amarillo"
                lblTipoContenedor.BackColor = vbYellow
                lblTipoContenedor.Caption = clsCon_Def.adorec_Def("prd_nombre")
            Else
                VSFGRojo.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & valResta
                VSFGRojo.Subtotal flexSTSum, -1, 2, , vbRed, vbWhite, True, "TOTAL Rojo"
                lblTipoContenedor.BackColor = vbRed
                lblTipoContenedor.Caption = clsCon_Def.adorec_Def("prd_nombre")
            End If
        Else
            MsgBox "El producto no existe en la base de datos." & vbNewLine & _
                   "No se ingresara.", vbInformation, "Productos"
        End If
    End If
    
    
End Sub
