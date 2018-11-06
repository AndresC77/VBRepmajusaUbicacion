VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProducto 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   Icon            =   "frmProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   12345
   Begin VB.CommandButton CmdImpExcel 
      Caption         =   "Importar Excel"
      Height          =   375
      Left            =   8520
      TabIndex        =   28
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdPromociones 
      Caption         =   "Promociones"
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4392
      TabIndex        =   5
      Top             =   6960
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
      TabIndex        =   4
      Top             =   6960
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
      Width           =   12060
      Begin VB.CheckBox chkFiltroColor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Color"
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkFiltroColeccion 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Colección"
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
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
         Left            =   5760
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
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
         Left            =   3000
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
         Left            =   3000
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
         Left            =   5760
         TabIndex        =   18
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
      Begin MSDataListLib.DataCombo dcmbColeccion 
         Height          =   315
         Left            =   8520
         TabIndex        =   22
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
      Begin MSDataListLib.DataCombo dcmbColor 
         Height          =   315
         Left            =   8520
         TabIndex        =   25
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8520
         TabIndex        =   26
         Top             =   1335
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Colección"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8520
         TabIndex        =   23
         Top             =   495
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
         Left            =   5760
         TabIndex        =   17
         Top             =   1335
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
         Left            =   3000
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
      Height          =   4080
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   12060
      _cx             =   54285080
      _cy             =   54271005
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
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProducto.frx":030A
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
      Top             =   2280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   5760
      Picture         =   "frmProducto.frx":0578
      ToolTipText     =   "Elimina una Fila"
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmProducto"
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
    Tipo = "Producto"
    Tipo2 = "l Producto"
    Me.Caption = Tipo
End Sub

Private Sub chkFiltroColeccion_Click()
    If chkFiltroColeccion.value = 1 Then
        dcmbColeccion.Enabled = True
    Else
        dcmbColeccion.Enabled = False
    End If
End Sub

Private Sub chkFiltroColor_Click()
    If chkFiltroColor.value = 1 Then
        dcmbColor.Enabled = True
    Else
        dcmbColor.Enabled = False
    End If
End Sub


Private Sub CmdImpExcel_Click()
    frmMigraProductos.Show
End Sub

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
    
    'crea combo de color
    strSql = " SELECT col_codigo, col_nombre" & _
                 " FROM color " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY col_nombre"
     clsCon_Def.Ejecutar strSql
    Set dcmbColor.RowSource = clsCon_Def.adorec_Def
    dcmbColor.BoundColumn = "col_codigo"
    dcmbColor.ListField = "col_nombre"
    
End Sub


Private Sub Carga()
    strSql = " SELECT prd_codigo,prd_nombre,mar_codigo,lin_codigo,gru_codigo,uni_codigo,tal_codigo,col_codigo," & _
             " clc_codigo,prd_baja,'' as precio,'' as promo,prd_costo,prd_cambia_precio,prd_iva," & _
             " prd_sku,prd_no_comision," & _
             " prd_fechamod,prd_usumod, '0' as modi " & _
             " FROM producto" & _
             " WHERE emp_codigo ='" & strEmpresa & "'"
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & "AND  prd_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.value = 1 Then
        strSql = strSql & " AND prd_nombre LIKE '%" & txtNombre.Text & "%' "
    End If
    If chkFiltroLinea.value = 1 Then
        strSql = strSql & " AND lin_codigo = '" & dcmbLinea.BoundText & "' "
    End If
    If chkFiltroMarca.value = 1 Then
        strSql = strSql & " AND mar_codigo = '" & dcmbMarca.BoundText & "' "
    End If
    If chkFiltroGrupo.value = 1 Then
        strSql = strSql & " AND gru_codigo LIKE '" & dcmbGrupo.BoundText & "%' "
    End If
    If chkFiltroColeccion.value = 1 Then
        strSql = strSql & " AND clc_codigo LIKE '" & dcmbColeccion.BoundText & "%' "
    End If
    If chkFiltroColor.value = 1 Then
        strSql = strSql & " AND col_codigo LIKE '" & dcmbColor.BoundText & "%' "
    End If
    strSql = strSql & " ORDER BY prd_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    'crea combo de marca
    strSql = " SELECT mar_codigo, mar_nombre" & _
                 " FROM marca " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY mar_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "mar_codigo,*mar_nombre", "mar_codigo")
    'crea combo de linea
    strSql = " SELECT lin_codigo, lin_nombre" & _
                 " FROM linea " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lin_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "lin_codigo,*lin_nombre", "lin_codigo")
    'crea combo de grupo
    strSql = " SELECT gru_codigo, CONCAT(REPLICATE(' ',(gru_nivel)*2),gru_nombre) as gru_nombre" & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY gru_codigo"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "gru_codigo,*gru_nombre", "gru_codigo")
    'crea combo de unidad de medida
    strSql = " SELECT uni_codigo, uni_nombre" & _
                 " FROM unidad " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY uni_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "uni_codigo,*uni_nombre", "uni_codigo")
    'crea combo de talla
    strSql = " SELECT tal_codigo, tal_nombre" & _
                 " FROM talla " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tal_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*tal_nombre", "tal_codigo")
    'crea combo de color
    strSql = " SELECT col_codigo, col_nombre" & _
                 " FROM color " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY col_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "cal_codigo,*col_nombre", "col_codigo")
    'crea combo de coleccion
    strSql = " SELECT clc_codigo, clc_nombre" & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(9) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "clc_codigo,*clc_nombre", "clc_codigo")
    
    Set VSFG.CellButtonPicture = imgBtnUp
    VSFG.ColComboList(11) = "..."
    VSFG.ColComboList(12) = "..."
    If VSFG.Rows > 1 Then
        VSFG.Cell(flexcpPicture, 1, 11, VSFG.Rows - 1, 11) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, 1, 11, VSFG.Rows - 1, 11) = flexPicAlignRightCenter
        VSFG.Cell(flexcpPicture, 1, 12, VSFG.Rows - 1, 12) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, 1, 12, VSFG.Rows - 1, 12) = flexPicAlignRightCenter
    End If
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
            strSql = " UPDATE producto " & _
                 " SET prd_nombre='" & VSFG.TextMatrix(i, 2) & "'," & _
                 " mar_codigo='" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                 " lin_codigo='" & UCase(VSFG.TextMatrix(i, 4)) & "'," & _
                 " gru_codigo='" & UCase(VSFG.TextMatrix(i, 5)) & "'," & _
                 " uni_codigo='" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                 " tal_codigo='" & VSFG.TextMatrix(i, 7) & "'," & _
                 " col_codigo='" & VSFG.TextMatrix(i, 8) & "'," & _
                 " clc_codigo='" & VSFG.TextMatrix(i, 9) & "'," & _
                 " prd_baja='" & Abs(FormatoD0(VSFG.TextMatrix(i, 10))) & "'," & _
                 " prd_costo='" & FormatoD4(VSFG.TextMatrix(i, 13)) & "'," & _
                 " prd_cambia_precio='" & Abs(FormatoD0(VSFG.TextMatrix(i, 14))) & "'," & _
                 " prd_iva='" & Abs(FormatoD0(VSFG.TextMatrix(i, 15))) & "'," & _
                 " prd_sku='" & UCase(VSFG.TextMatrix(i, 16)) & "'," & _
                 " prd_fechamod=CURRENT_TIMESTAMP," & _
                 " prd_usumod='" & strUsuario & "', " & _
                 " PRD_NO_COMISION='" & Abs(FormatoD0(VSFG.TextMatrix(i, 17))) & "'" & _
                 " WHERE prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        
        
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                'MsgBox "No puede ingresar " & Tipo2 & " falta el Codigo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Marca", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 4) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Linea", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Grupo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 6) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Unidad de Medida", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 16) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el SKU", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT COALESCE(COUNT(*),0) as cod " & _
                         " FROM producto " & _
                         " WHERE prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def("cod") = 0 Then
                    strSql = " SELECT COALESCE(COUNT(*),0) as cod " & _
                             " FROM producto " & _
                             " WHERE prd_sku='" & UCase(VSFG.TextMatrix(i, 16)) & "'" & _
                             " AND emp_codigo='" & strEmpresa & "'"
                    clsCon_Def.Ejecutar strSql
                    If clsCon_Def.adorec_Def("cod") = 0 Then
                    'controla que no se repita el código
                        strSql = " INSERT INTO producto(emp_codigo,prd_codigo,prd_nombre,mar_codigo,lin_codigo,gru_codigo,uni_codigo,tal_codigo,col_codigo,clc_codigo,prd_baja," & _
                                 " prd_costo,prd_cambia_precio,prd_iva,prd_sku,prd_fechamod,prd_usumod,PRD_NO_COMISION) " & _
                                 " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(i, 1)) & "','" & VSFG.TextMatrix(i, 2) & "','" & UCase(VSFG.TextMatrix(i, 3)) & "','" & UCase(VSFG.TextMatrix(i, 4)) & "', " & _
                                 " '" & VSFG.TextMatrix(i, 5) & "','" & VSFG.TextMatrix(i, 6) & "','" & VSFG.TextMatrix(i, 7) & "','" & VSFG.TextMatrix(i, 8) & "','" & VSFG.TextMatrix(i, 9) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 10))) & "','" & FormatoD4(VSFG.TextMatrix(i, 13)) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 14))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 15))) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "'," & _
                                 " CURRENT_TIMESTAMP, '" & strUsuario & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 17))) & "')"
                        clsCon_Def.Ejecutar strSql, "M"
                    ' Almacenamiento de los datos del nuevo producto en las listas de precios
                        strSql = " INSERT INTO lista_precio_p " & _
                                 " SELECT lis_pre_codigo, '" & UCase(VSFG.TextMatrix(i, 1)) & "', emp_codigo, 0," & _
                                 " lis_pre_politica,0,0,CURRENT_TIMESTAMP, '" & strUsuario & "' " & _
                                 " FROM lista_precio WHERE emp_codigo='" & strEmpresa & "' "
                        clsCon_Def.Ejecutar strSql, "M"
                    ' Almacenamiento de los datos del nuevo producto en los depositos
                        strSql = " INSERT INTO existencia " & _
                                 " SELECT '" & UCase(VSFG.TextMatrix(i, 1)) & "',dep_codigo, emp_codigo, " & _
                                 " 0, CURRENT_TIMESTAMP, '" & strUsuario & "' " & _
                                 " FROM deposito WHERE emp_codigo='" & strEmpresa & "' "
                        clsCon_Def.Ejecutar strSql, "M"
                    Else
                        MsgBox "El código SKU de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                    End If
                Else
                    MsgBox "El código de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        
        
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
            If MsgBox("Esta seguro de Eliminar el producto " & VSFG.TextMatrix(i, 1), vbYesNo + vbQuestion, "Eliminacion") = vbYes Then
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


Private Sub cmdPromociones_Click()
    frmCargaPromociones.Show
End Sub

Private Sub Form_Resize()
VSFG.Height = Me.Height - 4000
VSFG.Width = Me.Width - 500
cmbAceptar.Top = Me.Height - 1000
cmdCerrar.Top = Me.Height - 1000
End Sub


Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim codigo As String
    
    If VSFG.Rows < OldRow Then Exit Sub
    If (Abs(VSFG.TextMatrix(OldRow, VSFG.Cols - 1)) = 2) And _
        OldCol = 1 Then
        
      strSql = " SELECT COALESCE(prd_nombre,'') as cod " & _
                         " FROM producto " & _
                         " WHERE prd_codigo='" & VSFG.TextMatrix(OldRow, 1) & "'" & _
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
    If (Abs(VSFG.TextMatrix(OldRow, VSFG.Cols - 1)) = 2 Or _
        Abs(VSFG.TextMatrix(OldRow, VSFG.Cols - 1)) = 3) And _
       (OldCol = 16 And Not Len(UCase(VSFG.TextMatrix(OldRow, 16))) = 0) Then
                    strSql = " SELECT COALESCE(prd_nombre,'') as nom, COALESCE(prd_CODIGO,'') as cod " & _
                             " FROM producto " & _
                             " WHERE prd_sku='" & UCase(VSFG.TextMatrix(OldRow, 16)) & "'" & _
                             " AND emp_codigo='" & strEmpresa & "'"
                    clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.BOF = True Then
                    Exit Sub
                End If
                If Len(clsCon_Def.adorec_Def("nom")) > 0 Then
                  codigo = clsCon_Def.adorec_Def("cod")
                  If codigo <> VSFG.TextMatrix(OldRow, 1) Then
                    Cancel = True
                    MsgBox ("El SKU ingresado ya existe, su nombre es : " & clsCon_Def.adorec_Def("nom"))
                  End If
                End If
    End If
    
End Sub

Private Sub VSFG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 11 Then
        frmProductoPrecio.CodPrd = VSFG.TextMatrix(Row, 1)
        frmProductoPrecio.Show
    ElseIf Col = 12 Then
        frmProductoPromo.CodPrd = VSFG.TextMatrix(Row, 1)
        frmProductoPromo.Show
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
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

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
    End If
End Sub

Private Sub chkFiltroLinea_Click()
    If chkFiltroLinea.value = 1 Then
        dcmbLinea.Enabled = True
    Else
        dcmbLinea.Enabled = False
    End If
End Sub

Private Sub chkFiltroMarca_Click()
    If chkFiltroMarca.value = 1 Then
        dcmbMarca.Enabled = True
    Else
        dcmbMarca.Enabled = False
    End If
End Sub

Private Sub chkFiltroGrupo_Click()
    If chkFiltroGrupo.value = 1 Then
        dcmbGrupo.Enabled = True
    Else
        dcmbGrupo.Enabled = False
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.value = 1 Then
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
    ucrtVSFG.Inicializar
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

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Pad As String
    Dim i As Long
    If VSFG.Col = 16 And VSFG.Row > 0 And Abs(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = 2 And KeyCode = vbKeyF4 Then
        If InStr(1, VSFG.TextMatrix(VSFG.Row, 2), " ") > 6 Then
            Pad = ""
        Else
            For i = 0 To 6 - InStr(1, VSFG.TextMatrix(VSFG.Row, 2), " ")
                Pad = Pad & "."
            Next i
        End If
        VSFG.TextMatrix(VSFG.Row, 16) = Left(VSFG.TextMatrix(VSFG.Row, 2), InStr(1, VSFG.TextMatrix(VSFG.Row, 2), " ") - 1) & Pad & VSFG.TextMatrix(VSFG.Row, 3)
        Pad = ""
        If Len(Trim(VSFG.TextMatrix(VSFG.Row, 8))) >= 3 Then
            Pad = Trim(VSFG.TextMatrix(VSFG.Row, 8))
        Else
            For i = 0 To 2 - Len(Trim(VSFG.TextMatrix(VSFG.Row, 8)))
                Pad = Pad & "0"
                
            Next i
        End If
        VSFG.TextMatrix(VSFG.Row, 16) = Trim(VSFG.TextMatrix(VSFG.Row, 16)) & Pad & Trim(VSFG.TextMatrix(VSFG.Row, 8))
        Pad = ""
        If Len(Trim(VSFG.TextMatrix(VSFG.Row, 7))) >= 3 Then
            Pad = Trim(VSFG.TextMatrix(VSFG.Row, 7))
        Else
            For i = 0 To 2 - Len(Trim(VSFG.TextMatrix(VSFG.Row, 7)))
                Pad = Pad & "0"
            Next i
        End If
        VSFG.TextMatrix(VSFG.Row, 16) = Trim(VSFG.TextMatrix(VSFG.Row, 16)) & Pad & Trim(VSFG.TextMatrix(VSFG.Row, 7))
        
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub

Private Sub LeerExcel()

'dimensiones
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long, i As Long

'abrir programa Excel
Set xlApp = New Excel.Application
xl.Visible = False

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(App.Path & "\prueba.xls", True, True, , "")
Set xlHoja = xlApp.Worksheets("Hoja1")

'2. Si no conoces el rango
'lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

For i = 1 To lngultimafial
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))
VSFG.TextMatrix(i, 1) = xlHoja.Range(xlHoja.Cells(i, 1), xlHoja.Cells(i, 1))

Next i


'utilizamos los datos…
txtLlamadas.Text = varMatriz(10, 3)

'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit

'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

End Sub
