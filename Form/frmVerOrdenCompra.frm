VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerOrdenCompra 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Ordenes de Crompra Producto"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   Icon            =   "frmVerOrdenCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13425
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   2400
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   13185
      _cx             =   23257
      _cy             =   4233
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
      FormatString    =   $"frmVerOrdenCompra.frx":030A
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
   Begin VB.CommandButton cmdCambiarFechaProgramada 
      Caption         =   "Cambiar Fecha Prog."
      Height          =   360
      Left            =   7320
      TabIndex        =   27
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimirOrden 
      Caption         =   "Imprimir Orden"
      Height          =   360
      Left            =   9840
      TabIndex        =   26
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdCambiarFechaEntrega 
      Caption         =   "Cambiar Fecha Ent."
      Height          =   360
      Left            =   5520
      TabIndex        =   25
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdAnularOrden 
      Caption         =   "Anular Orden"
      Height          =   360
      Left            =   1920
      TabIndex        =   24
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdEnviarOrden 
      Caption         =   "Enviar Orden"
      Height          =   360
      Left            =   3720
      TabIndex        =   23
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdNuevaOrden 
      Caption         =   "Nueva Orden"
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   11640
      TabIndex        =   6
      Top             =   6240
      Width           =   1700
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13185
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   255
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmVerOrdenCompra.frx":03D0
            Left            =   1320
            List            =   "frmVerOrdenCompra.frx":03FB
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   705
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
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
            Left            =   480
            TabIndex        =   16
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66256899
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
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
            Left            =   1920
            TabIndex        =   17
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66256899
            CurrentDate     =   37463
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   9480
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Orden Compra"
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
         Left            =   3720
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroNombre 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Proveedor"
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
         Left            =   3720
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo cmbProveedor 
         Height          =   330
         Left            =   3720
         TabIndex        =   10
         Top             =   1560
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. Recepción Merca"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   1335
         Width           =   5625
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   1440
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   13185
      _cx             =   23257
      _cy             =   2540
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerOrdenCompra.frx":0464
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
      Height          =   2400
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   13185
      _cx             =   23257
      _cy             =   4233
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
      FormatString    =   $"frmVerOrdenCompra.frx":0634
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
End
Attribute VB_Name = "frmVerOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Private FechaI As Variant
Private FechaF As Variant
Private strOrdenCompra As String
Public strTipoOrdenCompra As String

Private Sub IniDato()
    Tipo = "Orden de Compra"
    Tipo2 = "la Orden de Compra"
    Me.Caption = Tipo
End Sub

Private Sub chkFechas_Click()
    If chkFechas.Value = 1 Then
        Label1.Caption = "Fecha Inicial"
        Label2.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Label1.Caption = "Fecha"
        Label2.Enabled = False
        Fecha2.Enabled = False
    End If
End Sub

Private Sub cmbMesI_Change()
    CambiarFecha
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1"
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
End Sub

Private Sub cmdAnularOrden_Click()
    Dim strMotivo As String
    Dim clsAnula As New clsConsulta
    Dim booAnula As Boolean
    clsAnula.Inicializar AdoConn, AdoConnMaster
    
    booAnula = True
    strSql = " SELECT COUNT(*) as n FROM recepcion_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ord_com_codigo='" & strOrdenCompra & "'"
    clsAnula.Ejecutar strSql
    If clsAnula.adorec_Def("n") > 0 Then
        booAnula = False
    End If
    strSql = " SELECT COUNT(*) as n FROM contenedor_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ord_com_codigo='" & strOrdenCompra & "'"
    clsAnula.Ejecutar strSql
    If clsAnula.adorec_Def("n") > 0 Then
        booAnula = False
    End If
    If booAnula = True Then
        strMotivo = Trim(InputBox("Ingrese el motivo para la anuacion", "Orden de Compra"))
        If strMotivo <> "" Then
            strSql = " UPDATE orden_compra " & _
                     " SET est_ord_com_codigo=-1," & _
                     " ord_com_observacion=CONCAT('" & UCase(strMotivo) & vbNewLine & " - ',ord_com_observacion)," & _
                     " ord_com_fechamod=CURRENT_TIMESTAMP, ord_com_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ord_com_codigo='" & VSFG.TextMatrix(VSFG.Row, 0) & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
            clsCon_Def.Ejecutar strSql
            MsgBox "Orden de Compra ANULADA", vbInformation, "Orden de Compra"
            Carga
        End If
    Else
        MsgBox "Orden de Compra ya tiene ingresado recepción y/o contenedor", vbInformation, "Orden de Compra"
    End If
End Sub

Private Sub cmdCambiarFechaEntrega_Click()
    Dim strFecha As String
    strFecha = "nuevo"
    While strFecha <> "" And IsDate(strFecha) = False
        strFecha = Trim(InputBox("Ingrese la nueva fecha de entrega aaaa-mm-dd", "Orden de Compra", Format(VSFG.TextMatrix(VSFG.Row, 6), "yyyy-mm-dd")))
    Wend
    If Trim(strFecha) <> "" Then
        strSql = " UPDATE orden_compra " & _
                 " SET ord_com_observacion=CONCAT('SE CAMBIA FECHA DE ENTREGA: " & Format(VSFG.TextMatrix(VSFG.Row, 6), "yyyy-mm-dd") & vbNewLine & "',ord_com_observacion)," & _
                 " ord_com_fecha_entrega='" & strFecha & "'," & _
                 " ord_com_fechamod=CURRENT_TIMESTAMP, ord_com_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & VSFG.TextMatrix(VSFG.Row, 0) & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
        clsCon_Def.Ejecutar strSql
        MsgBox "Fecha de Entrega Modificada", , "Orden de Compra"
        Carga
    End If
End Sub

Private Sub cmdCambiarFechaProgramada_Click()
    Dim strFecha As String
    strFecha = "nuevo"
    While strFecha <> "" And IsDate(strFecha) = False
        strFecha = Trim(InputBox("Ingrese la nueva fecha y hora PROGRAMADA de entrega aaaa-mm-dd hh:mm", "Orden de Compra", Format(VSFG.TextMatrix(VSFG.Row, 7), "yyyy-mm-dd hh:mm")))
    Wend
    If IsDate(Trim(strFecha)) = True Then
        strSql = " UPDATE orden_compra " & _
                 " SET ord_com_observacion=CONCAT('SE CAMBIA FECHA PROGRAMADA DE ENTREGA: " & Format(VSFG.TextMatrix(VSFG.Row, 7), "yyyy-mm-dd hh:mm") & vbNewLine & "',ord_com_observacion)," & _
                 " ord_com_fecha_programada='" & Trim(strFecha) & "'," & _
                 " ord_com_fechamod=CURRENT_TIMESTAMP, ord_com_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & VSFG.TextMatrix(VSFG.Row, 0) & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
        clsCon_Def.Ejecutar strSql
        MsgBox "Fecha de Entrega Modificada", , "Orden de Compra"
        Carga
    End If
End Sub

Private Sub cmdEnviarOrden_Click()
    Dim strCco As String
    frmReporte.strNumero = VSFG.TextMatrix(VSFG.Row, 0)
    If strTipoOrdenCompra = "P" Then
        frmReporte.strReporte = "rptOrdenCompraTalla"
        strCco = "coordinaciondeinventarios@rbimportadores.com; asistentejsn@rbimportadores.com; auxiliarjsn@rbimportadores.com"
    ElseIf strTipoOrdenCompra = "S" Then
        frmReporte.strReporte = "rptOrdenCompraSumi"
        strCco = "jefeadministrativo@rbimportadores.com"
    End If
    frmReporte.Show
    frmReporte.Form_Activate
    frmReporte.VSRpt.RenderToFile "OrdenCompra" & VSFG.TextMatrix(VSFG.Row, 0) & ".pdf", vsrPDF
    
    EnviarMail NombreComercial & " Compras", CorreoCompras, VSFG.TextMatrix(VSFG.Row, 1), VSFG.TextMatrix(VSFG.Row, 11), strCco, "Orden de compra " & VSFG.TextMatrix(VSFG.Row, 0), _
                "Estimad@" & vbNewLine & _
                VSFG.TextMatrix(VSFG.Row, 1) & vbNewLine & vbNewLine & _
                "Adjunto encontrarás la orden de compra segun hemos acordado." & vbNewLine & vbNewLine & _
                "Recuerda revisar la fecha de entrega que es el: " & Format(VSFG.TextMatrix(VSFG.Row, 6), "yyyy-MM-dd") & "." & vbNewLine & vbNewLine & _
                "Si tienes alguna novedad por favor no dudes en comunicarte nosotros." & vbNewLine & vbNewLine & _
                "Saludos Cordiales" & vbNewLine & _
                "Compras" & vbNewLine & _
                NombreComercial, "OrdenCompra" & VSFG.TextMatrix(VSFG.Row, 0) & ".pdf"
    Kill "OrdenCompra" & VSFG.TextMatrix(VSFG.Row, 0) & ".pdf"
    Unload frmReporte
    strSql = " UPDATE orden_compra " & _
             " SET est_ord_com_codigo=1," & _
             " ord_com_fecha_envio=CURRENT_TIMESTAMP," & _
             " ord_com_fechamod=CURRENT_TIMESTAMP, ord_com_usumod='" & strUsuario & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ord_com_codigo='" & VSFG.TextMatrix(VSFG.Row, 0) & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
    clsCon_Def.Ejecutar strSql
    MsgBox "Orden de Compra ENVIADA", vbInformation, "Orden de Compra"
    Carga
End Sub

Private Sub cmdImprimirOrden_Click()
    frmReporte.strNumero = VSFG.TextMatrix(VSFG.Row, 0)
    If strTipoOrdenCompra = "P" Then
        frmReporte.strReporte = "rptOrdenCompraTalla"
    ElseIf strTipoOrdenCompra = "S" Then
        frmReporte.strReporte = "rptOrdenCompraSumi"
    End If
    frmReporte.Show
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub

Private Sub Carga()
    strSql = " SELECT ord_com_codigo,concat(per_apellido,' ',per_nombre) as provee,orden_compra.est_ord_com_codigo,est_ord_com_descripcion," & _
             " ord_com_fecha, ord_com_fecha_envio, ord_com_fecha_entrega, ord_com_fecha_programada, ord_com_fecha_recepcion,for_pag_nombre, ord_com_observacion,per_email," & _
             " ord_com_fechamod, ord_com_usumod " & _
             " FROM orden_compra INNER JOIN persona " & _
             " ON orden_compra.emp_codigo=persona.emp_codigo " & _
             " AND orden_compra.per_codigo=persona.per_codigo " & _
             " INNER JOIN est_orden_compra " & _
             " ON orden_compra.est_ord_com_codigo=est_orden_compra.est_ord_com_codigo" & _
             " INNER JOIN forma_pago " & _
             " ON orden_compra.emp_codigo=forma_pago.emp_codigo" & _
             " AND orden_compra.for_pag_codigo=forma_pago.for_pag_codigo" & _
             " WHERE orden_compra.emp_codigo LIKE '" & strEmpresa & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'"
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  ord_com_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSql = strSql & " AND  orden_compra.per_codigo LIKE '" & cmbProveedor.BoundText & "' "
    End If
    
    If chkFiltroFecha.Value = 1 Then
        If Option1.Value = True Then
            strSql = strSql & " AND ord_com_fecha BETWEEN '" & FechaI & " 00:00:00' AND '" & FechaF & " 23:59:59' "
        ElseIf Option2.Value = True Then
            strSql = strSql & " AND ord_com_fecha BETWEEN '" & Fecha1 & " 00:00:00' AND '" & Fecha2 & " 23:59:59' "
        End If
    End If
    
    strSql = strSql & " ORDER BY ord_com_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    If VSFG.Rows > 1 Then
        VSFG_AfterRowColChange 0, 0, 1, 1
    Else
        strRecepcion = ""
    End If
End Sub

Private Sub cmdNuevaOrden_Click()
    frmOrderCompra.strTipoOrdenCompra = strTipoOrdenCompra
    frmOrderCompra.Show
End Sub

Private Sub Fecha1_Change()
    If chkFechas.Value = 0 Then
        Fecha2 = Fecha1
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Label1.Enabled = False
        Fecha1.Enabled = False
        Label2.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        Label1.Enabled = True
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.Value = 1 Then
            Label2.Enabled = True
            Fecha2.Enabled = True
        End If
    End If
End Sub

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.Value = 1 Then
        cmbProveedor.Enabled = True
    Else
        cmbProveedor.Enabled = False
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
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    If strTipoOrdenCompra = "P" Then
        VSFG1.Visible = True
        VSFG2.Visible = False
    ElseIf strTipoOrdenCompra = "S" Then
        VSFG1.Visible = False
        VSFG2.Visible = True
    End If
    Fecha1 = HoyDia
    Fecha2 = HoyDia
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    CambiarFecha
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.EOF = False Then
        Set cmbProveedor.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbProveedor.ListField = "nombre"
        cmbProveedor.BoundColumn = "per_codigo"
    End If
    
    IniDato
    Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim clsAnula As New clsConsulta
    Dim booAnula As Boolean
    clsAnula.Inicializar AdoConn, AdoConnMaster
    If (OldRow <> NewRow Or strRecepcion = "") And NewRow > 0 Then
        strOrdenCompra = VSFG.TextMatrix(NewRow, 0)
        CargaDetalle
        booAnula = True
        strSql = " SELECT COUNT(*) as n FROM recepcion_mercaderia " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & strOrdenCompra & "'"
        clsAnula.Ejecutar strSql
        If clsAnula.adorec_Def("n") > 0 Then
            booAnula = False
        End If
        strSql = " SELECT COUNT(*) as n FROM contenedor_mercaderia " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ord_com_codigo='" & strOrdenCompra & "'"
        clsAnula.Ejecutar strSql
        If clsAnula.adorec_Def("n") > 0 Then
            booAnula = False
        End If
        
        If VSFG.TextMatrix(NewRow, 2) = 0 Or VSFG.TextMatrix(NewRow, 2) = 1 Then
            cmdAnularOrden.Enabled = booAnula
            cmdCambiarFechaEntrega.Enabled = True
            cmdCambiarFechaProgramada.Enabled = False
            cmdEnviarOrden.Enabled = True
        ElseIf VSFG.TextMatrix(NewRow, 2) = 2 Then
            cmdAnularOrden.Enabled = booAnula
            cmdCambiarFechaEntrega.Enabled = True
            cmdCambiarFechaProgramada.Enabled = True
            cmdEnviarOrden.Enabled = True
        Else
            cmdAnularOrden.Enabled = False
            cmdCambiarFechaEntrega.Enabled = False
            cmdCambiarFechaProgramada.Enabled = False
            cmdEnviarOrden.Enabled = False
        End If
    End If
End Sub

Private Sub CargaDetalle()
    If strTipoOrdenCompra = "P" Then
        strSql = " SELECT det_orden_compra.pre_codigo,pre_nombre,col_nombre,tal_nombre,det_ord_com_cantidad,det_ord_com_precio,ROUND(det_ord_com_cantidad*det_ord_com_precio,4),det_ord_com_fechamod,det_ord_com_usumod " & _
                 " FROM det_orden_compra INNER JOIN preproducto_producto " & _
                 " ON det_orden_compra.emp_codigo=preproducto_producto.emp_codigo " & _
                 " AND det_orden_compra.pre_codigo=preproducto_producto.pre_codigo " & _
                 " AND det_orden_compra.col_codigo=preproducto_producto.col_codigo " & _
                 " AND det_orden_compra.tal_codigo=preproducto_producto.tal_codigo " & _
                 " INNER JOIN preproducto " & _
                 " ON det_orden_compra.emp_codigo=preproducto.emp_codigo " & _
                 " AND det_orden_compra.pre_codigo=preproducto.pre_codigo " & _
                 " INNER JOIN color " & _
                 " ON preproducto_producto.emp_codigo=color.emp_codigo " & _
                 " AND preproducto_producto.col_codigo=color.col_codigo " & _
                 " INNER JOIN talla " & _
                 " ON preproducto_producto.emp_codigo=talla.emp_codigo " & _
                 " AND preproducto_producto.tal_codigo=talla.tal_codigo " & _
                 " WHERE det_orden_compra.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_orden_compra.ord_com_codigo LIKE  '" & strOrdenCompra & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'" & _
                 " ORDER BY pre_nombre,col_nombre,tal_nombre "
        clsCon_Def.Ejecutar strSql
        Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            VSFG1.Subtotal flexSTSum, -1, 4, , vbBlue, vbWhite, True, "TOTAL"
            VSFG1.TextMatrix(VSFG1.Rows - 1, 6) = FormatoD2(VSFG1.Aggregate(flexSTSum, 1, 6, VSFG1.Rows - 2, 6))
        End If
    ElseIf strTipoOrdenCompra = "S" Then
        strSql = " SELECT det_orden_compra_s.det_ord_com_s_descripcion,det_ord_com_s_cantidad,det_ord_com_s_precio,ROUND(det_ord_com_s_cantidad*det_ord_com_s_precio,4),det_ord_com_s_fechamod,det_ord_com_s_usumod " & _
                 " FROM det_orden_compra_s " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND ord_com_codigo LIKE  '" & strOrdenCompra & "' AND ord_com_tipo='" & strTipoOrdenCompra & "'" & _
                 " ORDER BY det_ord_com_s_descripcion "
        clsCon_Def.Ejecutar strSql
        Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            VSFG2.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
            VSFG2.TextMatrix(VSFG1.Rows - 1, 3) = FormatoD2(VSFG2.Aggregate(flexSTSum, 1, 3, VSFG1.Rows - 2, 3))
        End If
    End If
End Sub
