VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransferirPrendasContenedorMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contenedores de Mercaderia"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   Icon            =   "frmTransferirPrendasContenedorMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   13980
   Begin VB.CommandButton cmdTransferirCompleto 
      Caption         =   "&Transferir Completo"
      Height          =   360
      Left            =   120
      TabIndex        =   28
      Top             =   5760
      Width           =   1700
   End
   Begin VB.Frame Frame2 
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
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6585
      Begin VB.TextBox TxtObserOrigen 
         Height          =   645
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   923
         Width           =   5280
      End
      Begin VB.TextBox txtCodigoOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   248
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpFechaOrigen 
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
         Left            =   4275
         TabIndex        =   19
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataListLib.DataCombo cmbBodegaOrigen 
         Height          =   315
         Left            =   1080
         TabIndex        =   26
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo cmbUbicacionOrigen 
         Height          =   315
         Left            =   4275
         TabIndex        =   27
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
         Left            =   30
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
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
         Left            =   405
         TabIndex        =   23
         Top             =   645
         Width           =   600
      End
      Begin VB.Label Label7 
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
         Left            =   3450
         TabIndex        =   22
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label6 
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
         Left            =   3015
         TabIndex        =   21
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   11520
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdAplicarTransferencia 
      Caption         =   "&Aplicar Transferencia"
      Height          =   360
      Left            =   6840
      TabIndex        =   2
      Top             =   5880
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   12240
      TabIndex        =   1
      Top             =   5880
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Contenedor Destino"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   7185
      Begin VB.TextBox txtCodigoDestino 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   248
         Width           =   1815
      End
      Begin VB.TextBox TxtObserDestino 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   930
         Width           =   6000
      End
      Begin MSDataListLib.DataCombo cmbBodegaDestino 
         Height          =   315
         Left            =   1080
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
      Begin MSDataListLib.DataCombo cmbUbicacionDestino 
         Height          =   315
         Left            =   4515
         TabIndex        =   8
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
      Begin MSComCtl2.DTPicker dtpFechaDestino 
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
         Left            =   4515
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
         Left            =   120
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
         Left            =   3255
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
         Left            =   3690
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
         Left            =   405
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
         Left            =   30
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGDestino 
      Height          =   3840
      Left            =   6840
      TabIndex        =   4
      Top             =   1920
      Width           =   7065
      _cx             =   12462
      _cy             =   6773
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransferirPrendasContenedorMercaderia.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGOrigen 
      Height          =   3840
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   6585
      _cx             =   11615
      _cy             =   6773
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransferirPrendasContenedorMercaderia.frx":03AC
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
      Left            =   10920
      TabIndex        =   15
      Top             =   1635
      Width           =   555
   End
End
Attribute VB_Name = "frmTransferirPrendasContenedorMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Public clsContenedorOrigen As New clsContenedor
Private clsContenedorDestino As New clsContenedor

Private Sub cmbBodegaOrigen_Validate(Cancel As Boolean)
    CargaUbicaOrigen
End Sub

Private Sub CargaUbicaOrigen()
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & cmbBodegaOrigen.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsCon_Def.Ejecutar strSql
    Set cmbUbicacionOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbUbicacionOrigen.ListField = "ubi_bod_codigo"
    cmbUbicacionOrigen.BoundColumn = "ubi_bod_codigo"
End Sub

Private Sub cmbBodegaDestino_Validate(Cancel As Boolean)
    CargaUbicaDestino
End Sub

Private Sub CargaUbicaDestino()
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & cmbBodegaDestino.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsAux.Ejecutar strSql
    Set cmbUbicacionDestino.RowSource = clsAux.adorec_Def.DataSource
    cmbUbicacionDestino.ListField = "ubi_bod_codigo"
    cmbUbicacionDestino.BoundColumn = "ubi_bod_codigo"
End Sub

Private Sub cmdAplicarTransferencia_Click()
    Dim i As Long
    Dim j As Long
    Dim strProd() As String
    j = 1
    If txtCodigoDestino.Text = "" Then
        clsContenedorDestino.NuevoContenedor strSucursal, strPtoFactura, Ahora, 0, cmbBodegaDestino.BoundText, cmbUbicacionDestino.BoundText, TxtObserDestino.Text
    End If
    For i = 2 To VSFGDestino.Rows - 1
        If FormatoD0(VSFGDestino.TextMatrix(i, 3)) > 0 Then
            ReDim Preserve strProd(j)
            strProd(j) = VSFGDestino.TextMatrix(i, 0) & vbTab & VSFGDestino.TextMatrix(i, 3)
            j = j + 1
        End If
    Next i
    clsContenedorOrigen.TransferirPrendasA strProd, clsContenedorDestino
    MsgBox "Se creo el contenedor No.: " & clsContenedorDestino.strNumContenedor, , "Contenedor"
    Unload Me
End Sub

Private Sub cmdTransferirCompleto_Click()
    Dim i As Long
    Dim j As Long
    VSFGDestino.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFGOrigen.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    For i = 2 To VSFGOrigen.Rows - 1
        For j = 1 To VSFGOrigen.TextMatrix(i, 2)
            txtLector.Text = VSFGOrigen.TextMatrix(i, 0)
            txtLector_KeyDown vbKeyReturn, 1
        Next j
    Next i
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
    clsContenedorDestino.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbBodegaDestino.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbBodegaDestino.ListField = "dep_nombre"
    cmbBodegaDestino.BoundColumn = "dep_codigo"
    Set cmbBodegaOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbBodegaOrigen.ListField = "dep_nombre"
    cmbBodegaOrigen.BoundColumn = "dep_codigo"
    VSFGDestino.SubtotalPosition = flexSTAbove
    VSFGOrigen.SubtotalPosition = flexSTAbove
    VSFGDestino.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFGOrigen.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFGDestino.Cell(flexcpFontSize, 1, 0, 1, VSFGDestino.Cols - 1) = VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) + 2
    VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) = VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) + 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtCodigoDestino_Validate(Cancel As Boolean)
    Dim i As Long
    If clsContenedorDestino.SetContenedor(txtCodigoDestino.Text) = True Then
        dtpFechaDestino.Value = clsContenedorDestino.strFecha
        cmbBodegaDestino.BoundText = clsContenedorDestino.strBodega
        cmbBodegaDestino.Locked = True
        cmbBodegaDestino_Validate False
        cmbUbicacionDestino.BoundText = clsContenedorDestino.strUbicacion
        cmbUbicacionDestino.Locked = True
        TxtObserDestino.Text = ""
        Set VSFGDestino.DataSource = clsContenedorDestino.adorec_DetalleContenedor
        VSFGDestino.Cols = VSFGDestino.Cols + 2
        VSFGDestino.TextMatrix(0, VSFGDestino.Cols - 2) = "Incrementar"
        VSFGDestino.TextMatrix(0, VSFGDestino.Cols - 1) = "Modi"
        txtCodigoOrigen_Change
    Else
        cmbUbicacionDestino.Locked = False
        cmbBodegaDestino.Locked = False
        MsgBox "Ese contenedor no esta registrado ", vbCritical, "Contenedor"
        clsContenedorDestino.setVacio
    End If
End Sub

Private Sub txtCodigoOrigen_Change()
    clsContenedorOrigen.SetContenedor txtCodigoOrigen.Text
    dtpFechaOrigen.Value = clsContenedorOrigen.strFecha
    cmbBodegaOrigen.BoundText = clsContenedorOrigen.strBodega
    cmbBodegaOrigen_Validate False
    cmbUbicacionOrigen.BoundText = clsContenedorOrigen.strUbicacion
    TxtObserOrigen.Text = clsContenedorOrigen.strObservacion
    Set VSFGOrigen.DataSource = clsContenedorOrigen.adorec_DetalleContenedor
    VSFGOrigen.Cols = VSFGOrigen.Cols + 2
    VSFGOrigen.TextMatrix(0, VSFGOrigen.Cols - 2) = "Descargar"
    VSFGOrigen.TextMatrix(0, VSFGOrigen.Cols - 1) = "Modi"
    
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProd UCase(txtLector.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub AgregarProd(codigo As String, Optional EsAux As Boolean = True)
    Dim i As Long
    Dim j As Long
    Dim pas As Boolean
    pas = False
InicioFor:
    For i = 1 To VSFGOrigen.Rows - 1
        If codigo = VSFGOrigen.TextMatrix(i, 0) Then
            If FormatoD2(VSFGOrigen.TextMatrix(i, 2)) >= FormatoD2(VSFGOrigen.TextMatrix(i, 3)) + 1 Then
                VSFGOrigen.ShowCell i, 0
                VSFGOrigen.Select i, 0
                VSFGOrigen.TextMatrix(i, 3) = Val(Format(VSFGOrigen.TextMatrix(i, 3), "###0")) + 1
                pas = True
                Exit For
            Else
                'MsgBox "No tiene suficiente mercaderia", vbCritical, "Contenedor"
                Exit For
            End If
        End If
    Next i
    
    If pas = True Then
        pas = False
        For j = 1 To VSFGDestino.Rows - 1
            If codigo = VSFGDestino.TextMatrix(j, 0) Then
                VSFGDestino.ShowCell j, 0
                VSFGDestino.Select j, 0
                VSFGDestino.TextMatrix(j, 3) = Val(Format(VSFGDestino.TextMatrix(j, 3), "###0")) + 1
                pas = True
                Exit For
            End If
        Next j
        If pas = False Then
            VSFGDestino.AddItem VSFGOrigen.TextMatrix(i, 0) & vbTab & VSFGOrigen.TextMatrix(i, 1) & vbTab & "0" & vbTab & "1"
        End If
        VSFGDestino.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
        VSFGOrigen.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    Else
        MsgBox "Ese producto no esta registrado en este contenedor o ya no tiene existencia", vbCritical, "Contenedor"
    End If
End Sub
