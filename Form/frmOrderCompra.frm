VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderCompra 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Order de Compra"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   Icon            =   "frmOrderCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11835
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   2760
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   11625
      _cx             =   20505
      _cy             =   4868
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
      FormatString    =   $"frmOrderCompra.frx":030A
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
   Begin VB.CommandButton cmdCrearOrden 
      Caption         =   "&Crear Orden"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   9960
      TabIndex        =   1
      Top             =   5160
      Width           =   1700
   End
   Begin VB.Frame Frame1 
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11625
      Begin VB.TextBox txtNoAux 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   2760
      End
      Begin VB.TextBox TxtObser 
         Height          =   525
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1410
         Width           =   5760
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
         Left            =   8955
         TabIndex        =   5
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   69206019
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbProveedor 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   5760
         _ExtentX        =   10160
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
      Begin MSComCtl2.DTPicker dtpFechaEnvio 
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
         Left            =   8940
         TabIndex        =   11
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
         Format          =   69206019
         CurrentDate     =   37463
      End
      Begin MSComCtl2.DTPicker dtpFechaRecepcion 
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
         Left            =   8940
         TabIndex        =   13
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
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
         Format          =   69206019
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo cmbTipoTalla 
         Height          =   330
         Left            =   8940
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   330
         Left            =   1320
         TabIndex        =   20
         Top             =   960
         Width           =   5760
         _ExtentX        =   10160
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
      Begin VB.Label Label6 
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
         Left            =   315
         TabIndex        =   21
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Talla:"
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
         Left            =   8190
         TabIndex        =   17
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Est.Recepcion:"
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
         Left            =   7335
         TabIndex        =   14
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Envio:"
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
         Left            =   7980
         TabIndex        =   12
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Aux:"
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
         Left            =   585
         TabIndex        =   10
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
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
         Left            =   8415
         TabIndex        =   6
         Top             =   300
         Width           =   495
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
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2760
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   11625
      _cx             =   20505
      _cy             =   4868
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
      FormatString    =   $"frmOrderCompra.frx":0392
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
      Height          =   2760
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   11625
      _cx             =   20505
      _cy             =   4868
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOrderCompra.frx":0467
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
End
Attribute VB_Name = "frmOrderCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private intTallas As Integer
Private strTallas() As String
Public strTipoOrdenCompra As String

Private Sub cmbTipoTalla_Validate(Cancel As Boolean)
    Dim clsAux As New clsConsulta
    Dim i As Integer
    clsAux.Inicializar AdoConn, AdoConnMaster
    If cmbTipoTalla.BoundText = "%" Then
        VSFG.Visible = True
        VSFG1.Visible = False
        intTallas = 0
        VSFG.Clear 1
        VSFG1.Clear 1
    Else
        VSFG.Visible = False
        VSFG1.Visible = True
        VSFG.Clear 1
        VSFG1.Clear 1
        strSql = " SELECT tipo_talla_talla.tal_codigo,tal_nombre " & _
                 " FROM tipo_talla_talla INNER JOIN talla " & _
                 " ON tipo_talla_talla.emp_codigo=talla.emp_codigo " & _
                 " AND tipo_talla_talla.tal_codigo=talla.tal_codigo" & _
                 " WHERE tipo_talla_talla.emp_codigo='" & strEmpresa & "' " & _
                 " AND tipo_talla_talla.tip_tal_codigo='" & cmbTipoTalla.BoundText & "' " & _
                 " ORDER BY tip_tal_tal_orden"
        clsAux.Ejecutar strSql
        intTallas = clsAux.adorec_Def.RecordCount
        ReDim strTallas(intTallas) As String
        VSFG1.Cols = 6 + intTallas
        VSFG1.TextMatrix(0, 5) = "Total"
        VSFG1.ColPosition(5) = VSFG1.Cols - 1
        VSFG1.TextMatrix(0, 4) = "Precio"
        VSFG1.ColPosition(4) = VSFG1.Cols - 2
        i = 0
        While Not clsAux.adorec_Def.EOF
            VSFG1.TextMatrix(0, 3 + i) = clsAux.adorec_Def("tal_nombre")
            strTallas(i) = clsAux.adorec_Def("tal_codigo")
            VSFG1.ColWidth(3 + i) = 800
            clsAux.adorec_Def.MoveNext
            i = i + 1
        Wend
        VSFG1.TextMatrix(0, 3 + i) = "T.Unidades"
        VSFG1.ColWidth(3 + i) = 1000
        VSFG1.Cell(flexcpAlignment, 0, 3, 0, VSFG1.Cols - 1) = 4
        VSFG1.ColWidth(VSFG1.Cols - 2) = 1000
        VSFG1.ColWidth(VSFG1.Cols - 1) = 1000
    End If
End Sub

Private Sub cmdCrearOrden_Click()
    Dim num As String
    Dim i As Long
    Dim j As Integer
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If CmbFpago.MatchedWithList = False Or cmbProveedor.MatchedWithList = False Then
        MsgBox "No tienen seleccionado proveedor o Forma de Pago", vbInformation
        Exit Sub
    End If
    strSql = " BEGIN TRAN "
    clsAux.Ejecutar strSql, "M"
    strSql = " SELECT COALESCE(MAX(ord_com_codigo),0)+1 as n " & _
             " FROM orden_compra WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
    clsAux.Ejecutar strSql, "M"
    num = 1
    If clsAux.adorec_Def.RecordCount > 0 Then
        num = clsAux.adorec_Def("n")
    End If

    strSql = " INSERT INTO orden_compra(emp_codigo,ord_com_tipo,ord_com_codigo,per_codigo,est_ord_com_codigo," & _
             " ord_com_numaux,ord_com_fecha,ord_com_fecha_envio,ord_com_fecha_entrega," & _
             " ord_com_observacion,for_pag_codigo,ord_com_fechamod,ord_com_usumod)" & _
             " VALUES('" & strEmpresa & "','" & strTipoOrdenCompra & "','" & num & "','" & cmbProveedor.BoundText & "','0'," & _
             " '" & UCase(txtNoAux.Text) & "','" & dtpFecha.Value & "','" & dtpFechaEnvio.Value & "','" & dtpFechaRecepcion.Value & "'," & _
             " '" & UCase(TxtObser.Text) & "','" & CmbFpago.BoundText & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar strSql, "M"
    strSql = " COMMIT TRAN "
    clsAux.Ejecutar strSql, "M"
    If strTipoOrdenCompra = "P" Then
        If cmbTipoTalla.BoundText = "%" Then
            With VSFG
                For i = 1 To .Rows - 1
                    If FormatoD2(.TextMatrix(i, 4)) <> 0 Then
                        strSql = " SELECT det_ord_com_cantidad,det_ord_com_precio " & _
                                 " FROM det_orden_compra " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_codigo='" & num & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' " & _
                                 " AND pre_codigo='" & .TextMatrix(i, 0) & "' AND col_codigo='" & .TextMatrix(i, 2) & "' AND tal_codigo='" & .TextMatrix(i, 3) & "'"
                        clsAux.Ejecutar strSql, "M"
                        If clsAux.adorec_Def.RecordCount = 0 Then
                            strSql = " INSERT INTO det_orden_compra(emp_codigo,ord_com_tipo,ord_com_codigo,pre_codigo," & _
                                     " col_codigo,tal_codigo,det_ord_com_cantidad,det_ord_com_precio," & _
                                     " det_ord_com_fechamod,det_ord_com_usumod) " & _
                                     " VALUES('" & strEmpresa & "','" & strTipoOrdenCompra & "','" & num & "','" & .TextMatrix(i, 0) & "', " & _
                                     " '" & .TextMatrix(i, 2) & "', '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', '" & .TextMatrix(i, 5) & "', " & _
                                     " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                        Else
                            strSql = " UPDATE det_orden_compra " & _
                                     " SET det_ord_com_cantidad='" & clsAux.adorec_Def("det_ord_com_cantidad") + .TextMatrix(i, 4) & "'," & _
                                     " det_ord_com_precio='" & (clsAux.adorec_Def("det_ord_com_cantidad") * clsAux.adorec_Def("det_ord_com_precio") + .TextMatrix(i, 4) * .TextMatrix(i, 5)) / (clsAux.adorec_Def("det_ord_com_cantidad") + .TextMatrix(i, 4)) & "'" & _
                                     " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' AND ord_com_codigo='" & num & "' " & _
                                     " AND pre_codigo='" & .TextMatrix(i, 0) & "' AND col_codigo='" & .TextMatrix(i, 2) & "' AND tal_codigo='" & .TextMatrix(i, 3) & "'"
                        End If
                        clsAux.Ejecutar strSql, "M"
                    End If
                Next i
            End With
        Else
            With VSFG1
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Then
                        For j = 3 To intTallas + 2
                            If FormatoD2(.TextMatrix(i, j)) <> 0 Then
                                strSql = " SELECT det_ord_com_cantidad,det_ord_com_precio " & _
                                         " FROM det_orden_compra " & _
                                         " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_codigo='" & num & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' " & _
                                         " AND pre_codigo='" & .TextMatrix(i, 0) & "' AND col_codigo='" & .TextMatrix(i, 2) & "' AND tal_codigo='" & strTallas(j - 3) & "'"
                                clsAux.Ejecutar strSql, "M"
                                If clsAux.adorec_Def.RecordCount = 0 Then
                                    strSql = " INSERT INTO det_orden_compra(emp_codigo,ord_com_tipo,ord_com_codigo,pre_codigo," & _
                                             " col_codigo,tal_codigo,det_ord_com_cantidad,det_ord_com_precio," & _
                                             " det_ord_com_fechamod,det_ord_com_usumod) " & _
                                             " VALUES('" & strEmpresa & "','" & strTipoOrdenCompra & "','" & num & "','" & .TextMatrix(i, 0) & "', " & _
                                             " '" & .TextMatrix(i, 2) & "', '" & strTallas(j - 3) & "', '" & .TextMatrix(i, j) & "', '" & .TextMatrix(i, .Cols - 2) & "', " & _
                                             " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                                Else
                                    strSql = " UPDATE det_orden_compra " & _
                                             " SET det_ord_com_cantidad='" & clsAux.adorec_Def("det_ord_com_cantidad") + .TextMatrix(i, j) & "'," & _
                                             " det_ord_com_precio='" & (clsAux.adorec_Def("det_ord_com_cantidad") * clsAux.adorec_Def("det_ord_com_precio") + .TextMatrix(i, j) * .TextMatrix(i, .Cols - 2)) / (clsAux.adorec_Def("det_ord_com_cantidad") + .TextMatrix(i, j)) & "'" & _
                                             " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' AND ord_com_codigo='" & num & "' " & _
                                             " AND pre_codigo='" & .TextMatrix(i, 0) & "' AND col_codigo='" & .TextMatrix(i, 2) & "' AND tal_codigo='" & strTallas(j - 3) & "'"
                                End If
                                clsAux.Ejecutar strSql, "M"
                            End If
                        Next j
                    End If
                Next i
            End With
        End If
    ElseIf strTipoOrdenCompra = "S" Then
        
        With VSFG2
            For i = 1 To .Rows - 1
                If FormatoD2(.TextMatrix(i, 3)) <> 0 Then
                    strSql = " SELECT det_ord_com_s_cantidad,det_ord_com_s_precio " & _
                             " FROM det_orden_compra_s " & _
                             " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_codigo='" & num & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' " & _
                             " AND det_ord_com_s_descripcion='" & UCase(.TextMatrix(i, 0)) & "'"
                    clsAux.Ejecutar strSql, "M"
                    If clsAux.adorec_Def.RecordCount = 0 Then
                        strSql = " INSERT INTO det_orden_compra_s(emp_codigo,ord_com_tipo,ord_com_codigo,det_ord_com_s_descripcion," & _
                                 " det_ord_com_s_cantidad,det_ord_com_s_precio," & _
                                 " det_ord_com_s_fechamod,det_ord_com_s_usumod) " & _
                                 " VALUES('" & strEmpresa & "','" & strTipoOrdenCompra & "','" & num & "','" & UCase(.TextMatrix(i, 0)) & "', " & _
                                 " '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', " & _
                                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    Else
                        strSql = " UPDATE det_orden_compra_s " & _
                                 " SET det_ord_com_s_cantidad='" & clsAux.adorec_Def("det_ord_com_s_cantidad") + .TextMatrix(i, 1) & "'," & _
                                 " det_ord_com_s_precio='" & (clsAux.adorec_Def("det_ord_com_s_cantidad") * clsAux.adorec_Def("det_ord_com_s_precio") + .TextMatrix(i, 1) * .TextMatrix(i, 2)) / (clsAux.adorec_Def("det_ord_com_s_cantidad") + .TextMatrix(i, 1)) & "'" & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND ord_com_tipo='" & strTipoOrdenCompra & "' AND ord_com_codigo='" & num & "' " & _
                                 " AND det_ord_com_s_descripcion='" & UCase(.TextMatrix(i, 0)) & "'"
                    End If
                    clsAux.Ejecutar strSql, "M"
                End If
            Next i
        End With
    End If
    Unload Me
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
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsCon_Def.Ejecutar (strSql)
    Set CmbFpago.RowSource = clsCon_Def.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"

    
    If strTipoOrdenCompra = "P" Then
        strSql = " SELECT col_codigo,col_nombre " & _
                 " FROM color " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY col_nombre"
        clsCon_Def.Ejecutar strSql
        
        VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
        VSFG1.ColComboList(2) = VSFG1.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
        strSql = " SELECT tal_codigo,tal_nombre " & _
                 " FROM talla " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tal_nombre"
        clsCon_Def.Ejecutar strSql
        
        VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "tal_codigo, *tal_nombre", "tal_codigo")
        
        strSql = " SELECT '%' as tip_tal_codigo,' -Todas-' as tip_tal_nombre UNION SELECT tip_tal_codigo, tip_tal_nombre " & _
                 " FROM tipo_talla WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tip_tal_nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbTipoTalla.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbTipoTalla.ListField = "tip_tal_nombre"
        cmbTipoTalla.BoundColumn = "tip_tal_codigo"
        cmbTipoTalla.BoundText = "%"
        cmbTipoTalla.Visible = True
        VSFG.Visible = True
        VSFG1.Visible = False
        VSFG2.Visible = False
    Else
        cmbTipoTalla.Visible = False
        VSFG.Visible = False
        VSFG1.Visible = False
        VSFG2.Visible = True
    End If
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ', per_nombre) as nomb " & _
             " FROM persona WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='P' " & _
             " ORDER BY nomb "
    clsCon_Def.Ejecutar strSql
    Set cmbProveedor.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbProveedor.ListField = "nomb"
    cmbProveedor.BoundColumn = "per_codigo"
    
    dtpFecha.Value = HoyDia
    dtpFechaEnvio.Value = HoyDia
    dtpFechaRecepcion.Value = HoyDia
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Or Col = VSFG.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Or Col = VSFG1.Cols - 3 Or Col = VSFG1.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = VSFG2.Cols - 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strRef As String
    Dim i As Long
    If Col = 0 Then
        strSql = " SELECT pre_nombre " & _
                 " FROM preproducto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG.TextMatrix(Row, 0) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            strRef = ""
            For i = 1 To VSFG.Rows - 1
                strRef = strRef & "'" & VSFG.TextMatrix(i, 0) & "',"
            Next i
            strRef = Left(strRef, Len(strRef) - 1)
            VSFG.TextMatrix(Row, 1) = clsCon_Def.adorec_Def("pre_nombre")
            
            strSql = " SELECT DISTINCT color.col_codigo,col_nombre " & _
                     " FROM color inner join preproducto_producto ON color.emp_codigo=preproducto_producto.emp_codigo" & _
                     " AND color.col_codigo=preproducto_producto.col_codigo" & _
                     " WHERE preproducto_producto.emp_codigo='" & strEmpresa & "' and preproducto_producto.pre_codigo in (" & strRef & ")" & _
                     " ORDER BY col_nombre"
            clsCon_Def.Ejecutar strSql
            
            VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
            VSFG1.ColComboList(2) = VSFG1.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
            
        Else
            MsgBox "Esta referencia no existe", vbInformation, "PreProducto"
            VSFG.TextMatrix(Row, 0) = ""
            VSFG.TextMatrix(Row, 1) = ""
            VSFG.TextMatrix(Row, 2) = ""
            VSFG.TextMatrix(Row, 3) = ""
        End If
    ElseIf Col = 2 Then
        strSql = " SELECT pre_codigo " & _
                 " FROM preproducto_producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG.TextMatrix(Row, 0) & "'" & _
                 " AND col_codigo='" & VSFG.TextMatrix(Row, 2) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            MsgBox "Este color no esta registrado para la referencia", vbInformation, "PreProducto"
            For i = 2 To VSFG.Cols - 1
                VSFG.TextMatrix(Row, i) = ""
            Next i
        End If
    ElseIf Col = 3 Then
        strSql = " SELECT pre_codigo " & _
                 " FROM preproducto_producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG.TextMatrix(Row, 0) & "'" & _
                 " AND col_codigo='" & VSFG.TextMatrix(Row, 2) & "'" & _
                 " AND tal_codigo='" & VSFG.TextMatrix(Row, 3) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            MsgBox "Esta talla no esta registrada para la referencia y color", vbInformation, "PreProducto"
            VSFG.TextMatrix(Row, 3) = ""
        End If
    ElseIf Col = 4 Or Col = 5 Then
        VSFG.TextMatrix(Row, 6) = FormatoD2(FormatoD2(VSFG.TextMatrix(Row, 4)) * FormatoD4(VSFG.TextMatrix(Row, 5)))
        If FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 6)) <> 0 Then
            VSFG.AddItem ""
        End If
    ElseIf Col = 6 Then
        If FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 6)) <> 0 Then
            VSFG.AddItem ""
        End If
    End If
    
End Sub

Private Sub VSFG1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strRef As String
    Dim i As Integer
    If Row = 0 Then Exit Sub
    If Col = 0 Then
        strSql = " SELECT pre_nombre " & _
                 " FROM preproducto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG1.TextMatrix(Row, 0) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            strRef = ""
            For i = 1 To VSFG1.Rows - 1
                strRef = strRef & "'" & VSFG1.TextMatrix(i, 0) & "',"
            Next i
            strRef = Left(strRef, Len(strRef) - 1)
            VSFG1.TextMatrix(Row, 1) = clsCon_Def.adorec_Def("pre_nombre")
            strSql = " SELECT DISTINCT color.col_codigo,col_nombre " & _
                     " FROM color inner join preproducto_producto ON color.emp_codigo=preproducto_producto.emp_codigo" & _
                     " AND color.col_codigo=preproducto_producto.col_codigo" & _
                     " WHERE preproducto_producto.emp_codigo='" & strEmpresa & "' and preproducto_producto.pre_codigo in (" & strRef & ")" & _
                     " ORDER BY col_nombre"
            clsCon_Def.Ejecutar strSql
            
            VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
            VSFG1.ColComboList(2) = VSFG1.BuildComboList(clsCon_Def.adorec_Def, "col_codigo, *col_nombre", "col_codigo")
        Else
            MsgBox "Esta referencia no existe", vbInformation, "PreProducto"
            For i = 0 To VSFG1.Cols - 1
                VSFG1.TextMatrix(Row, i) = ""
            Next i
        End If
    ElseIf Col = 2 Then
        strSql = " SELECT pre_codigo " & _
                 " FROM preproducto_producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG1.TextMatrix(Row, 0) & "'" & _
                 " AND col_codigo='" & VSFG1.TextMatrix(Row, 2) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            MsgBox "Este color no esta registrado para la referencia", vbInformation, "PreProducto"
            VSFG1.TextMatrix(Row, 2) = ""
        End If
        For i = 3 To VSFG1.Cols - 1
            VSFG1.TextMatrix(Row, i) = ""
        Next i
    ElseIf Col >= 3 And Col <= VSFG1.Cols - 4 Then
        strSql = " SELECT pre_codigo " & _
                 " FROM preproducto_producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND pre_codigo='" & VSFG1.TextMatrix(Row, 0) & "'" & _
                 " AND col_codigo='" & VSFG1.TextMatrix(Row, 2) & "'" & _
                 " AND tal_codigo='" & strTallas(Col - 3) & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            MsgBox "Esta talla no esta registrada para la referencia y color", vbInformation, "PreProducto"
            VSFG1.TextMatrix(Row, Col) = ""
        End If
        VSFG1.TextMatrix(Row, VSFG1.Cols - 3) = 0
        For i = 3 To 2 + intTallas
            VSFG1.TextMatrix(Row, VSFG1.Cols - 3) = FormatoD4(FormatoD4(VSFG1.TextMatrix(Row, VSFG1.Cols - 3)) + FormatoD4(VSFG1.TextMatrix(Row, i)))
        Next i
        VSFG1.TextMatrix(Row, VSFG1.Cols - 1) = FormatoD2(FormatoD2(VSFG1.TextMatrix(Row, VSFG1.Cols - 3)) * FormatoD4(VSFG1.TextMatrix(Row, VSFG1.Cols - 2)))
    ElseIf Col = VSFG1.Cols - 3 Or Col = VSFG1.Cols - 2 Then
        VSFG1.TextMatrix(Row, VSFG1.Cols - 1) = FormatoD2(FormatoD2(VSFG1.TextMatrix(Row, VSFG1.Cols - 3)) * FormatoD4(VSFG1.TextMatrix(Row, VSFG1.Cols - 2)))
        If FormatoD2(VSFG1.TextMatrix(VSFG1.Rows - 1, VSFG1.Cols - 3)) <> 0 Then
            VSFG1.AddItem ""
        End If
    ElseIf Col = VSFG1.TextMatrix(Row, VSFG1.Cols - 1) Then
        If FormatoD2(VSFG1.TextMatrix(VSFG.Rows - 1, VSFG1.Cols - 1)) <> 0 Then
            VSFG1.AddItem ""
        End If
    End If
    
End Sub

Private Sub VSFG2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Or Col = 1 Or Col = 2 Then
        VSFG2.TextMatrix(Row, 3) = FormatoD2(FormatoD2(VSFG2.TextMatrix(Row, 1)) * FormatoD4(VSFG2.TextMatrix(Row, 2)))
        If FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 3)) <> 0 And VSFG2.TextMatrix(VSFG2.Rows - 1, 0) <> "" Then
            VSFG2.AddItem ""
        End If
    ElseIf Col = 3 Then
        If FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 3)) <> 0 And VSFG2.TextMatrix(VSFG2.Rows - 1, 0) <> "" Then
            VSFG2.AddItem ""
        End If
    End If
    
End Sub
