VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPreProducto 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre Productos"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14280
   Icon            =   "frmPreProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14280
   Begin VB.CommandButton cmdFichaTecnica 
      Caption         =   "&Ver Ficha Técnica"
      Height          =   360
      Left            =   120
      TabIndex        =   29
      Top             =   7560
      Width           =   1700
   End
   Begin VB.CommandButton cmdGenerarEAN 
      Caption         =   "&Generar EAN"
      Height          =   360
      Left            =   11160
      TabIndex        =   25
      Top             =   7560
      Width           =   1700
   End
   Begin VB.CommandButton cmdConsultarDetalle 
      Caption         =   "&Consultar Detalle"
      Height          =   360
      Left            =   12000
      TabIndex        =   24
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdAceptarDetalle 
      Caption         =   "&Aceptar Detalle"
      Height          =   360
      Left            =   10200
      TabIndex        =   23
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4392
      TabIndex        =   5
      Top             =   7560
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
      TabIndex        =   4
      Top             =   7560
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
      Width           =   14100
      Begin VB.CheckBox chkFiltroColeccion 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Coleccion"
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
         Left            =   8520
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroGrupo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Grupos"
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
         Left            =   8520
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroMarca 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Marca"
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
         Left            =   5760
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroLinea 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Línea"
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
         Left            =   5760
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Código"
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
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroNombre 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Nombre"
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
         Left            =   3000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcmbLinea 
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbMarca 
         Height          =   315
         Left            =   5760
         TabIndex        =   15
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbGrupo 
         Height          =   315
         Left            =   8520
         TabIndex        =   18
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbColeccion 
         Height          =   315
         Left            =   8520
         TabIndex        =   27
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupos"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8520
         TabIndex        =   28
         Top             =   1335
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   495
         Width           =   2655
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   495
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupos"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8520
         TabIndex        =   17
         Top             =   495
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   1335
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Línea"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   495
         Width           =   2655
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4560
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   9540
      _cx             =   16828
      _cy             =   8043
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPreProducto.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGDet 
      Height          =   4080
      Left            =   9720
      TabIndex        =   21
      Top             =   2880
      Width           =   4500
      _cx             =   7937
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPreProducto.frx":0475
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
      FrozenRows      =   0
      FrozenCols      =   4
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin NEED2.uctrVSFG ucrtVSFGDet 
      Height          =   375
      Left            =   9720
      TabIndex        =   22
      Top             =   2400
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   5760
      Picture         =   "frmPreProducto.frx":0599
      ToolTipText     =   "Elimina una Fila"
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPreProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Private Valor As String

Private Sub IniDato()
    Tipo = "PreProducto"
    Tipo2 = "l PreProducto"
    Me.Caption = Tipo
End Sub


Private Sub cmdAceptarDetalle_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFGDet.Select 1, VSFGDet.Cols - 1
    VSFGDet.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFGDet.Rows - 1
        'update
        If FormatoD0(VSFGDet.TextMatrix(i, VSFGDet.Cols - 1)) = 3 Then
        'prd_codigo,prd_nombre,mar_codigo,lin_codigo,gru_codigo,uni_codigo,prd_baja,'' as precio,prd_costo,prd_fechamod,prd_usumod, '0' as modi
            strSql = " UPDATE preproducto_producto " & _
                 " SET tal_codigo='" & VSFGDet.TextMatrix(i, 2) & "'," & _
                 " col_codigo='" & VSFGDet.TextMatrix(i, 4) & "'," & _
                 " prd_codigo='" & UCase(VSFGDet.TextMatrix(i, 5)) & "'," & _
                 " pre_pro_fechamod=CURRENT_TIMESTAMP," & _
                 " pre_pro_usumod='" & strUsuario & "' " & _
                 " WHERE pre_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
                 " AND emp_codigo='" & strEmpresa & "' " & _
                 " AND tal_codigo='" & VSFGDet.TextMatrix(i, 1) & "'" & _
                 " AND col_codigo='" & VSFGDet.TextMatrix(i, 3) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        
        
        ElseIf FormatoD0(VSFGDet.TextMatrix(i, VSFGDet.Cols - 1)) = 2 Then
            'controla que este lleno los datos
            If VSFGDet.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar falta Talla", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFGDet.TextMatrix(i, 4) = "" Then
                MsgBox "No puede ingresar falta el Color", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT COALESCE(COUNT(*),0) as cod " & _
                         " FROM preproducto_producto " & _
                         " WHERE pre_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
                         " AND emp_codigo='" & strEmpresa & "' " & _
                         " AND tal_codigo='" & VSFGDet.TextMatrix(i, 2) & "'" & _
                         " AND col_codigo='" & VSFGDet.TextMatrix(i, 4) & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def("cod") = 0 Then
                    strSql = " INSERT INTO preproducto_producto(emp_codigo,pre_codigo,tal_codigo,col_codigo,prd_codigo," & _
                             " pre_pro_fechamod,pre_pro_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(VSFG.Row, 1)) & "','" & VSFGDet.TextMatrix(i, 2) & "','" & VSFGDet.TextMatrix(i, 4) & "','" & UCase(VSFGDet.TextMatrix(i, 5)) & "', " & _
                             " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "La conbinacion talla color ya existe", vbInformation, "Ingreso"
                End If
             End If
        
        
        ElseIf FormatoD0(VSFGDet.TextMatrix(i, VSFGDet.Cols - 1)) = 1 Then
            
            strSql = " DELETE FROM preproducto_producto " & _
                     " WHERE pre_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
                     " AND emp_codigo='" & strEmpresa & "' " & _
                     " AND tal_codigo='" & VSFGDet.TextMatrix(i, 1) & "'" & _
                     " AND col_codigo='" & VSFGDet.TextMatrix(i, 3) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        ElseIf FormatoD0(VSFGDet.TextMatrix(i, VSFGDet.Cols - 1)) <= 0 Then
            Exit For
        End If
        
    Next i
    If control = 0 Then
        cmdConsultarDetalle_Click
    End If
End Sub

Private Sub cmdConsultarDetalle_Click()
    strSql = " SELECT tal_codigo as ta,tal_codigo,col_codigo as ca,col_codigo,prd_codigo," & _
             " pre_pro_fechamod,pre_pro_usumod, '0' as modi " & _
             " FROM preproducto_producto" & _
             " WHERE emp_codigo ='" & strEmpresa & "'" & _
             " AND pre_codigo ='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
             " ORDER BY col_codigo,tal_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFGDet.DataSource = clsCon_Def.adorec_Def.DataSource
    
    'crea combo de talla
    strSql = " SELECT tal_codigo, tal_nombre" & _
                 " FROM talla " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tal_nombre"
     clsCon_Def.Ejecutar strSql
    VSFGDet.ColComboList(2) = VSFGDet.BuildComboList(clsCon_Def.adorec_Def, "tal_codigo,*tal_nombre", "tal_codigo")
    'crea combo de color
    strSql = " SELECT col_codigo, col_nombre" & _
                 " FROM color " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY col_nombre"
     clsCon_Def.Ejecutar strSql
    VSFGDet.ColComboList(4) = VSFGDet.BuildComboList(clsCon_Def.adorec_Def, "col_codigo,*col_nombre", "col_codigo")
End Sub

Private Sub cmdFichaTecnica_Click()
    frmVerFichaTecnica.txtReferencia = VSFG.TextMatrix(VSFG.Row, 1)
    frmVerFichaTecnica.txtNombre = VSFG.TextMatrix(VSFG.Row, 2)
   'frmVerFichaTecnica.txtCostoServicio =
   'frmVerFichaTecnica.txtObservacion   =
    
    frmVerFichaTecnica.Show
End Sub

Private Sub cmdGenerarEAN_Click()
    Dim iSum As Integer
    Dim iDigit As Integer
    Dim EAN As String
    Dim Talla As String
    
    Dim iCheckSum As Integer
    
    Dim i As Long
    
    Dim Fila As Long
    strSql = " SELECT tal_codigo as ta,tal_codigo,col_codigo as ca,col_codigo,prd_codigo," & _
             " pre_pro_fechamod,pre_pro_usumod, '0' as modi " & _
             " FROM preproducto_producto" & _
             " WHERE emp_codigo ='" & strEmpresa & "'" & _
             " AND pre_codigo ='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
             " ORDER BY col_codigo,tal_codigo "
    clsCon_Def.Ejecutar strSql
    
    For i = 1 To VSFGDet.Rows - 1
        If VSFGDet.TextMatrix(i, 5) = "" Then
            VSFGDet.Select i, 5
            ucrtVSFGDet.Modificar
            iSum = 0
            iDigit = 0
            Talla = Format(i, "00")
            EAN = "500" & Format(VSFG.TextMatrix(VSFG.Row, 1), "0000000") & Format(Talla, "00")
            For ii = 1 To 12
                iDigit = Mid(EAN, ii, 1)
                If ii Mod 2 = 0 Then
                    iSum = iSum + iDigit * 3
                Else
                    iSum = iSum + iDigit
                End If
            Next
            iCheckSum = (10 - (iSum Mod 10)) Mod 10
            If ExisteEAN(EAN & iCheckSum) = True Then
                MsgBox "No puede generar el EAN para la referencia " & VSFG.TextMatrix(VSFG.Row, 1) & " " & VSFGDet.Cell(flexcpTextDisplay, i, 4) & " " & VSFGDet.Cell(flexcpTextDisplay, i, 2) & vbNewLine & _
                       "Revisar porque ya esta generado", vbCritical, "EAN"
            Else
                VSFGDet.TextMatrix(i, 5) = EAN & iCheckSum
            End If
        End If
    Next i

End Sub

Private Function ExisteEAN(EAN As String) As Boolean
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT prd_codigo FROM producto " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND prd_codigo='" & EAN & "'"
    clsAux.Ejecutar strSql, "L"
    If clsAux.adorec_Def.RecordCount > 0 Then
        ExisteEAN = True
    Else
        ExisteEAN = False
    End If
End Function

Private Sub cmdMostrar_Click()
    Carga
End Sub

Private Sub CargaCombos()
     'crea combo de marca
    strSql = " SELECT mar_codigo, mar_nombre" & _
                 " FROM marca " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY mar_nombre"
     clsCon_Def.Ejecutar strSql
    Set dcmbMarca.RowSource = clsCon_Def.adorec_Def
    dcmbMarca.BoundColumn = "mar_codigo"
    dcmbMarca.ListField = "mar_nombre"
    'crea combo de linea
    strSql = " SELECT lin_codigo, lin_nombre" & _
                 " FROM linea " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lin_nombre"
     clsCon_Def.Ejecutar strSql
    Set dcmbLinea.RowSource = clsCon_Def.adorec_Def
    dcmbLinea.BoundColumn = "lin_codigo"
    dcmbLinea.ListField = "lin_nombre"
    
    'crea combo de grupo
    strSql = " SELECT gru_codigo, CONCAT(((gru_nivel)*2),gru_nombre) as gru_nombre" & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY gru_codigo"
     clsCon_Def.Ejecutar strSql
    Set dcmbGrupo.RowSource = clsCon_Def.adorec_Def
    dcmbGrupo.BoundColumn = "gru_codigo"
    dcmbGrupo.ListField = "gru_nombre"
    
    'crea combo de coleccion
    strSql = " SELECT clc_codigo, clc_nombre" & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_nombre"
     clsCon_Def.Ejecutar strSql
    Set dcmbColeccion.RowSource = clsCon_Def.adorec_Def
    dcmbColeccion.BoundColumn = "clc_codigo"
    dcmbColeccion.ListField = "clc_nombre"
End Sub

Private Sub Carga()
    strSql = " SELECT pre_codigo,pre_nombre,pre_descripcion,mar_codigo,lin_codigo,gru_codigo,uni_codigo,clc_codigo," & _
             " pre_fechamod,pre_usumod, '0' as modi " & _
             " FROM preproducto" & _
             " WHERE emp_codigo ='" & strEmpresa & "'"
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  pre_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSql = strSql & " AND pre_nombre LIKE '%" & txtNombre.Text & "%' "
    End If
    If chkFiltroLinea.Value = 1 Then
        strSql = strSql & " AND lin_codigo = '" & dcmbLinea.BoundText & "' "
    End If
    If chkFiltroMarca.Value = 1 Then
        strSql = strSql & " AND mar_codigo = '" & dcmbMarca.BoundText & "' "
    End If
    If chkFiltroGrupo.Value = 1 Then
        strSql = strSql & " AND gru_codigo LIKE '" & dcmbGrupo.BoundText & "%' "
    End If
    If chkFiltroColeccion.Value = 1 Then
        strSql = strSql & " AND clc_codigo = '" & dcmbcolecion.BoundText & "%' "
    End If
    strSql = strSql & " ORDER BY pre_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    'crea combo de marca
    strSql = " SELECT mar_codigo, mar_nombre" & _
                 " FROM marca " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY mar_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "mar_codigo,*mar_nombre", "mar_codigo")
    'crea combo de linea
    strSql = " SELECT lin_codigo, lin_nombre" & _
                 " FROM linea " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lin_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "lin_codigo,*lin_nombre", "lin_codigo")
    'crea combo de grupo
    strSql = " SELECT gru_codigo, CONCAT(REPLICATE(' ',(gru_nivel)*2),gru_nombre) as gru_nombre" & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY gru_codigo"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "gru_codigo,*gru_nombre", "gru_codigo")
    'crea combo de unidad de medida
    strSql = " SELECT uni_codigo, uni_nombre" & _
                 " FROM unidad " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY uni_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "uni_codigo,*uni_nombre", "uni_codigo")
    'crea combo de coleccion
    strSql = " SELECT clc_codigo, clc_nombre" & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "clc_codigo,*clc_nombre", "clc_codigo")
    ucrtVSFG.PonerNum
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
        'prd_codigo,prd_nombre,mar_codigo,lin_codigo,gru_codigo,uni_codigo,prd_baja,'' as precio,prd_costo,prd_fechamod,prd_usumod, '0' as modi
            strSql = " UPDATE preproducto " & _
                 " SET pre_nombre='" & VSFG.TextMatrix(i, 2) & "'," & _
                 " pre_descripcion='" & VSFG.TextMatrix(i, 3) & "'," & _
                 " mar_codigo='" & UCase(VSFG.TextMatrix(i, 4)) & "'," & _
                 " lin_codigo='" & UCase(VSFG.TextMatrix(i, 5)) & "'," & _
                 " gru_codigo='" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                 " uni_codigo='" & UCase(VSFG.TextMatrix(i, 7)) & "'," & _
                 " clc_codigo='" & UCase(VSFG.TextMatrix(i, 8)) & "'," & _
                 " pre_fechamod=CURRENT_TIMESTAMP," & _
                 " pre_usumod='" & strUsuario & "' " & _
                 " WHERE pre_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        
        
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta referencia", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Descripcion", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 4) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Marca", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Linea", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 6) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Grupo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 7) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Unidad de Medida", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 8) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Coleccion", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT COALESCE(COUNT(*),0) as cod " & _
                         " FROM preproducto " & _
                         " WHERE pre_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def("cod") = 0 Then
                    strSql = " INSERT INTO preproducto(emp_codigo,pre_codigo,pre_nombre,pre_descripcion,mar_codigo,lin_codigo,gru_codigo,uni_codigo,clc_codigo," & _
                             " pre_fechamod,pre_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(i, 1)) & "','" & VSFG.TextMatrix(i, 2) & "','" & VSFG.TextMatrix(i, 3) & "','" & UCase(VSFG.TextMatrix(i, 4)) & "', " & _
                             " '" & VSFG.TextMatrix(i, 5) & "','" & VSFG.TextMatrix(i, 6) & "','" & VSFG.TextMatrix(i, 7) & "','" & VSFG.TextMatrix(i, 8) & "'," & _
                             " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "El código de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        
        
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
            If MsgBox("Esta seguro de Eliminar el preproducto " & VSFG.TextMatrix(i, 1), vbYesNo + vbQuestion, "Eliminacion") = vbYes Then
                strSql = " SELECT COALESCE(count(*),0) as n FROM det_backorder " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                clsCon_Def.Ejecutar strSql, "M"
                If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                    MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                Else
                    strSql = " SELECT COALESCE(count(*),0) as n FROM det_cotizacion " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                    clsCon_Def.Ejecutar strSql, "M"
                    If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                        MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                    Else
                        strSql = " SELECT COALESCE(count(*),0) as n FROM det_egreso " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                        clsCon_Def.Ejecutar strSql, "M"
                        If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                            MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                        Else
                            strSql = " SELECT COALESCE(count(*),0) as n FROM det_ingreso " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                            clsCon_Def.Ejecutar strSql, "M"
                            If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                            Else
                                strSql = " SELECT COALESCE(count(*),0) as n FROM det_ingreso_imp " & _
                                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                                         " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                clsCon_Def.Ejecutar strSql, "M"
                                If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                    MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                                Else
                                    strSql = " SELECT COALESCE(count(*),0) as n FROM det_pedido " & _
                                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                    clsCon_Def.Ejecutar strSql, "M"
                                    If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                        MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                                    Else
                                        strSql = " SELECT COALESCE(count(*),0) as n FROM det_pedido_imp " & _
                                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                 " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                        clsCon_Def.Ejecutar strSql, "M"
                                        If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                            MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                                        Else
                                            strSql = " SELECT COALESCE(count(*),0) as n FROM det_prd_com " & _
                                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                            clsCon_Def.Ejecutar strSql, "M"
                                            If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                                            Else
                                                strSql = " SELECT COALESCE(count(*),0) as n FROM existencia " & _
                                                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                         " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' AND exi_cantidad!=0 "
                                                clsCon_Def.Ejecutar strSql, "M"
                                                If FormatoD0(clsCon_Def.adorec_Def("n")) <> 0 Then
                                                    MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                                                Else
                                                    strSql = " DELETE FROM existencia " & _
                                                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                                    clsCon_Def.Ejecutar strSql, "M"
                                                    strSql = " DELETE FROM lista_precio_p " & _
                                                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                                    clsCon_Def.Ejecutar strSql, "M"
                                                    strSql = " DELETE FROM producto " & _
                                                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                                                    clsCon_Def.Ejecutar strSql, "M"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
        
    Next i
    If control = 0 Then
        Carga
    End If
    
End Sub
'
'Private Sub Form_Resize()
'    VSFG.Height = Me.Height - 4000
'    VSFG.Width = Me.Width - 500
'    cmbAceptar.Top = Me.Height - 1000
'    cmdCerrar.Top = Me.Height - 1000
'End Sub



Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        cmdConsultarDetalle_Click
    End If
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim codigo As String
    
    If VSFG.Rows - 1 < OldRow Or OldRow = 0 Or VSFG.Rows - 1 <= 1 Then Exit Sub
    
    If (Abs(VSFG.TextMatrix(OldRow, VSFG.Cols - 1)) = 2) And _
    OldCol = 1 Then
    
        strSql = " SELECT COALESCE(pre_nombre,'') as cod " & _
                 " FROM preproducto " & _
                 " WHERE pre_codigo='" & VSFG.TextMatrix(OldRow, 1) & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.BOF = True Then
            Exit Sub
        End If
        If Len(clsCon_Def.adorec_Def("cod")) > 0 Then
         Cancel = True
         'VSFG.Col = OldCol
         'VSFG.Row = OldRow
         'VSFG.TextMatrix(OldRow, OldCol) = VSFG.TextMatrix(OldRow, OldCol)
         MsgBox ("El codigo ingresado ya existe, su nombre es : " & clsCon_Def.adorec_Def("cod"))
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFGDet_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -2 Then
        VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1) = 2
    ElseIf Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -3 Then
        VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1) = 3
    End If
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

Private Sub VSFGDet_DblClick()
    Dim i As Long
    Set DAT = New frmDatos
    If VSFGDet.Row >= 1 Then
        DAT.Show
        DAT.VSFG.Rows = VSFGDet.Cols
        For i = 1 To VSFGDet.Cols - 1
            DAT.VSFG.TextMatrix(i, 0) = VSFGDet.TextMatrix(0, i)
            DAT.VSFG.Cell(flexcpText, i, 1) = VSFGDet.Cell(flexcpTextDisplay, VSFGDet.Row, i)
            If VSFGDet.ColComboList(i) <> "" Then
                DAT.VSFG.TextMatrix(i, 2) = VSFGDet.ColComboList(i)
                DAT.VSFG.Cell(flexcpText, i, 3) = VSFGDet.Cell(flexcpText, VSFGDet.Row, i)
            End If
        Next i
        DAT.VSFG.Cell(flexcpBackColor, 1, 1, DAT.VSFG.Rows - 1, 1) = VSFGDet.Cell(flexcpBackColor, VSFGDet.Row, VSFGDet.Col)
        DAT.VSFG.RowHidden(DAT.VSFG.Rows - 1) = True
        Set DAT.VSFGOrigen = VSFGDet
        DAT.VSFGOrigen.Tag = VSFGDet.Row
        DAT.Caption = Tipo
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFGDet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = 0 Or Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = 2 Or Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -2 Then
        If Col >= VSFGDet.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = 3 Or Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFGDet.Cols - 3 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.Value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
    End If
End Sub

Private Sub chkFiltroLinea_Click()
    If chkFiltroLinea.Value = 1 Then
        dcmbLinea.Enabled = True
    Else
        dcmbLinea.Enabled = False
    End If
End Sub

Private Sub chkFiltroMarca_Click()
    If chkFiltroMarca.Value = 1 Then
        dcmbMarca.Enabled = True
    Else
        dcmbMarca.Enabled = False
    End If
End Sub

Private Sub chkFiltroGrupo_Click()
    If chkFiltroGrupo.Value = 1 Then
        dcmbGrupo.Enabled = True
    Else
        dcmbGrupo.Enabled = False
    End If
End Sub

Private Sub chkFiltroColeccion_Click()
    If chkFiltroColeccion.Value = 1 Then
        dcmbColeccion.Enabled = True
    Else
        dcmbColeccion.Enabled = False
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
    Set ucrtVSFG.VSFGControl = VSFG
    Set ucrtVSFGDet.VSFGControl = VSFGDet
    ucrtVSFG.Inicializar
    ucrtVSFGDet.Inicializar , , , False, False
    IniDato
    CargaCombos
    Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFGDet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -2 Then
        VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1) = 2
    ElseIf Val(VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1)) = -3 Then
        VSFGDet.TextMatrix(Row, VSFGDet.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub

Private Sub VSFGDet_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFGDet.MouseRow > 0 Then
        ucrtVSFGDet.VerMenu
    End If
End Sub

