VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerRecepcionMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepciones de Mercaderia"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   Icon            =   "frmVerRecepcionMercaderia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13470
   Begin VB.CommandButton cmdMoverRecepcionNueva 
      Caption         =   "Mover Recepción Nueva"
      Height          =   360
      Left            =   2040
      TabIndex        =   24
      Top             =   6240
      Width           =   2895
   End
   Begin VB.CommandButton cmdNuevaRecepcion 
      Caption         =   "Nueva Recepción"
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
            ItemData        =   "frmVerRecepcionMercaderia.frx":030A
            Left            =   1320
            List            =   "frmVerRecepcionMercaderia.frx":0335
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
            Format          =   65404931
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
            Format          =   65404931
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
         Caption         =   "Filtrar Recepcion de Merc."
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
      Height          =   1920
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   7980
      _cx             =   14076
      _cy             =   3387
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerRecepcionMercaderia.frx":039E
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
      Height          =   1920
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   7980
      _cx             =   14076
      _cy             =   3387
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerRecepcionMercaderia.frx":04A7
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
      Height          =   3840
      Left            =   8160
      TabIndex        =   23
      Top             =   2280
      Width           =   5220
      _cx             =   9208
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerRecepcionMercaderia.frx":05B0
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
Attribute VB_Name = "frmVerRecepcionMercaderia"
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
Private strRecepcion As String
Private strBodega As String
Private strUbicacion As String
Private Sub IniDato()
    Tipo = "Recepción de Mercadería"
    Tipo2 = "la Recepción de Mercadería"
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

Private Sub cmdMostrar_Click()
    Carga
End Sub

Private Sub Carga()
    strSql = " SELECT rec_mer_codigo,concat(per_apellido,' ',per_nombre) as provee,est_rec_mer_descripcion," & _
             " rec_mer_fecha,rec_mer_factura,rec_mer_observacion,rec_mer_fechamod, rec_mer_usumod " & _
             " FROM recepcion_mercaderia INNER JOIN persona " & _
             " ON recepcion_mercaderia.emp_codigo=persona.emp_codigo " & _
             " AND recepcion_mercaderia.per_codigo=persona.per_codigo " & _
             " INNER JOIN est_recepcion_mercaderia" & _
             " ON recepcion_mercaderia.est_rec_mer_codigo=est_recepcion_mercaderia.est_rec_mer_codigo" & _
             " WHERE recepcion_mercaderia.emp_codigo LIKE '" & strEmpresa & "'"
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  rec_mer_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSql = strSql & " AND  recepcion_mercaderia.per_codigo LIKE '" & cmbProveedor.BoundText & "' "
    End If
    
    If chkFiltroFecha.Value = 1 Then
        If Option1.Value = True Then
            strSql = strSql & " AND rec_mer_fecha BETWEEN '" & FechaI & " 00:00:00' AND '" & FechaF & " 23:59:59' "
        ElseIf Option2.Value = True Then
            strSql = strSql & " AND rec_mer_fecha BETWEEN '" & Fecha1 & " 00:00:00' AND '" & Fecha2 & " 23:59:59' "
        End If
    End If
    
    strSql = strSql & " ORDER BY rec_mer_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    If VSFG.Rows > 1 Then
        VSFG_AfterRowColChange 0, 0, 1, 1
    Else
        strRecepcion = ""
    End If
End Sub

Private Sub cmdMoverRecepcionNueva_Click()
    If VSFG.Row > 0 And VSFG.TextMatrix(VSFG.Row, 0) <> "" Then
        frmReubicarRecepcionContenedorMercaderia.txtCodigo.Text = VSFG.TextMatrix(VSFG.Row, 0)
        frmReubicarRecepcionContenedorMercaderia.dtpFecha.Value = VSFG.TextMatrix(VSFG.Row, 3)
        frmReubicarRecepcionContenedorMercaderia.txtBodega.Text = strBodega
        frmReubicarRecepcionContenedorMercaderia.txtUbicacion.Text = strUbicacion
        frmReubicarRecepcionContenedorMercaderia.TxtObser.Text = VSFG.TextMatrix(VSFG.Row, 5)
        frmReubicarRecepcionContenedorMercaderia.Show
    End If
End Sub

Private Sub cmdNuevaRecepcion_Click()
    frmRecepcionMercaderia.Show
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

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
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
    Dim strFac As String
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    ElseIf KeyCode = vbKeyF6 And Shift = 1 Then
        strFac = InputBox("Cambio numero de Factura de la recepcion: " & VSFG.TextMatrix(VSFG.Row, 0), "Recepcion", VSFG.TextMatrix(VSFG.Row, 4))
        If strFac <> "" Then
            strSql = " UPDATE recepcion_mercaderia " & _
                     " SET rec_mer_factura='" & strFac & "'," & _
                     " rec_mer_observacion=CONCAT('FACTURA ANTERIOR: " & VSFG.TextMatrix(VSFG.Row, 4) & " - ',rec_mer_observacion), rec_mer_usumod='" & strUsuario & "', rec_mer_fechamod=CURRENT_TIMESTAMP " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND rec_mer_codigo='" & VSFG.TextMatrix(VSFG.Row, 0) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        End If
    End If
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (OldRow <> NewRow Or strRecepcion = "") And NewRow > 0 Then
        strRecepcion = VSFG.TextMatrix(NewRow, 0)
        CargaContenedores
        CargaDetalle
    End If
End Sub

Private Sub CargaDetalle()
    
    strSql = " SELECT det_contenedor_mercaderia.prd_codigo,prd_nombre,SUM(det_con_mer_cantidad) " & _
             " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
             " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
             " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
             " INNER JOIN det_contenedor_mercaderia " & _
             " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
             " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
             " INNER JOIN producto " & _
             " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
             " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
             " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & strRecepcion & "%'" & _
             " GROUP BY det_contenedor_mercaderia.prd_codigo,prd_nombre " & _
             " ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    VSFG2.Subtotal flexSTSum, -1, 2, , vbBlue, vbWhite, True, "TOTAL"
End Sub

Private Sub CargaContenedores()
    Dim Pasa As Boolean
    strSql = " SELECT det_recepcion_mercaderia.con_mer_codigo,con_mer_fecha,est_con_mer_descripcion,dep_nombre,ubi_bod_codigo,con_mer_observacion,con_mer_fechamod,con_mer_usumod " & _
             " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
             " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
             " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
             " INNER JOIN est_contenedor_mercaderia " & _
             " ON contenedor_mercaderia.est_con_mer_codigo=est_contenedor_mercaderia.est_con_mer_codigo" & _
             " INNER JOIN deposito " & _
             " ON contenedor_mercaderia.emp_codigo=deposito.emp_codigo " & _
             " AND contenedor_mercaderia.dep_codigo=deposito.dep_codigo " & _
             " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & strRecepcion & "'" & _
             " ORDER BY con_mer_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
    strSql = " SELECT DISTINCT dep_codigo,ubi_bod_codigo " & _
             " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
             " ON det_recepcion_mercaderia.emp_codigo = contenedor_mercaderia.emp_codigo " & _
             " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
             " WHERE det_recepcion_mercaderia.emp_codigo ='" & strEmpresa & "' " & _
             " AND rec_mer_codigo='" & strRecepcion & "'"
    clsCon_Def.Ejecutar strSql
    Pasa = True
    If clsCon_Def.adorec_Def.RecordCount = 1 Then
        Pasa = True
        strBodega = clsCon_Def.adorec_Def("dep_codigo")
        strUbicacion = clsCon_Def.adorec_Def("ubi_bod_codigo")
    Else
        Pasa = False
    End If
    
    If Pasa = True Then
        strSql = " SELECT DISTINCT count(*) as n " & _
                 " FROM det_recepcion_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                 " ON det_recepcion_mercaderia.emp_codigo = det_contenedor_mercaderia.emp_codigo " & _
                 " AND det_recepcion_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE det_recepcion_mercaderia.emp_codigo ='" & strEmpresa & "' " & _
                 " AND rec_mer_codigo='" & strRecepcion & "' " & _
                 " AND det_contenedor_mercaderia.tip_mov_codigo!=''" & _
                 " AND det_contenedor_mercaderia.mov_codigo!=0"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def("n") > 0 Then
            Pasa = False
        End If
    End If
    
    If Pasa = True Then
        cmdMoverRecepcionNueva.Enabled = True
    Else
        cmdMoverRecepcionNueva.Enabled = False
    End If
End Sub

