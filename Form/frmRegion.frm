VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRegion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Region"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13140
   Icon            =   "frmRegion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   13140
   Begin VB.CommandButton cmdAceptarVendedor 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   9840
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.CommandButton cmdAsignarVendedor 
      Caption         =   "A&signar Vendedores"
      Height          =   360
      Left            =   1920
      TabIndex        =   15
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton cmdAceptarCiudad 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   9840
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.CommandButton cmdAsignarCiudad 
      Caption         =   "A&signar Ciudades"
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2322
      TabIndex        =   7
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4122
      TabIndex        =   6
      Top             =   6480
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7905
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3255
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
         Width           =   2895
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
         Left            =   4440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         TabIndex        =   5
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   495
         Width           =   3255
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3585
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   7860
      _cx             =   13864
      _cy             =   6324
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegion.frx":030A
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
      TabIndex        =   9
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGVendedor 
      Height          =   6105
      Left            =   8160
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   4860
      _cx             =   8572
      _cy             =   10769
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
      FormatString    =   $"frmRegion.frx":03C9
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGCiudad 
      Height          =   6105
      Left            =   8160
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   4860
      _cx             =   8572
      _cy             =   10769
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
      FormatString    =   $"frmRegion.frx":044F
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
End
Attribute VB_Name = "frmRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Private Sub IniDato()
    Tipo = "Region"
    Tipo2 = "la Region"
    Me.Caption = Tipo
End Sub

Private Sub cmdAceptarCiudad_Click()
    Dim i As Long
    strSql = " DELETE FROM region_ciudad " & _
         " WHERE emp_codigo='" & strEmpresa & "'" & _
         " AND reg_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
    clsCon_Def.Ejecutar strSql, "M"

    VSFGCiudad.Select 1, VSFGCiudad.Cols - 1
    VSFGCiudad.Sort = flexSortGenericDescending
    
    For i = 1 To VSFGCiudad.Rows - 1
        'insert
        If VSFGCiudad.TextMatrix(i, VSFGCiudad.Cols - 1) = 2 Then
            strSql = " INSERT INTO region_ciudad(emp_codigo,reg_codigo," & _
                    " ciu_codigo,reg_ciu_fechamod,reg_ciu_usumod) " & _
                    " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(VSFG.Row, 1)) & "'," & _
                    " '" & UCase(VSFGCiudad.TextMatrix(i, VSFGCiudad.Cols - 2)) & "', " & _
                    " CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsCon_Def.Ejecutar strSql, "M"
        ElseIf VSFGCiudad.TextMatrix(i, VSFGCiudad.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    CargarCiudades UCase(VSFG.TextMatrix(VSFG.Row, 1))
    
End Sub

Private Sub cmdAceptarVendedor_Click()
    Dim i As Long
    strSql = " DELETE FROM region_marca " & _
         " WHERE emp_codigo='" & strEmpresa & "'" & _
         " AND reg_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
    clsCon_Def.Ejecutar strSql, "M"

    VSFGVendedor.Select 1, VSFGVendedor.Cols - 1
    VSFGVendedor.Sort = flexSortGenericDescending
    
    For i = 1 To VSFGVendedor.Rows - 1
        'insert
        If VSFGVendedor.TextMatrix(i, VSFGVendedor.Cols - 1) = 2 Then
            If Trim(VSFGVendedor.TextMatrix(i, 1)) <> "-" Then
            strSql = " INSERT INTO region_marca(emp_codigo, reg_codigo, " & _
                     " mar_codigo," & _
                     " ven_codigo, reg_mar_fechamod, reg_mar_usumod) " & _
                    " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(VSFG.Row, 1)) & "'," & _
                    " '" & UCase(VSFGVendedor.TextMatrix(i, VSFGVendedor.Cols - 2)) & "', " & _
                    " '" & UCase(VSFGVendedor.TextMatrix(i, 1)) & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsCon_Def.Ejecutar strSql, "M"
            End If
        ElseIf VSFGVendedor.TextMatrix(i, VSFGVendedor.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    CargarMarcas UCase(VSFG.TextMatrix(VSFG.Row, 1))

End Sub

Private Sub cmdAsignarCiudad_Click()
    If cmdAsignarCiudad.Caption = "A&signar Ciudades" Then
        cmdAsignarCiudad.Caption = "C&rear Regiones"
        Me.Width = Me.Width + 5000
        ucrtVSFG.Visible = False
        cmbAceptar.Visible = False
        cmdAsignarVendedor.Visible = False
        cmdAceptarCiudad.Visible = True
        VSFGCiudad.Visible = True
        VSFG.SelectionMode = flexSelectionByRow
        CargarCiudades VSFG.TextMatrix(VSFG.Row, 1)
    Else
        cmdAsignarCiudad.Caption = "A&signar Ciudades"
        Me.Width = Me.Width - 5000
        ucrtVSFG.Visible = True
        cmbAceptar.Visible = True
        cmdAsignarVendedor.Visible = True
        cmdAceptarCiudad.Visible = False
        VSFGCiudad.Visible = False
        VSFG.SelectionMode = flexSelectionFree
    End If
End Sub

Private Sub CargarCiudades(strRegion As String)
    VSFGCiudad.Clear 1
    VSFGCiudad.Rows = 1
    strSql = " SELECT if(region_ciudad.reg_codigo='" & strRegion & "','1','0') as sel, " & _
             " pai_nombre,ciu_nombre,reg_nombre,ciudad.ciu_codigo,'0' as modi " & _
             " FROM ciudad INNER JOIN pais " & _
             " ON ciudad.pai_codigo=pais.pai_codigo " & _
             " LEFT JOIN region_ciudad " & _
             " ON ciudad.ciu_codigo=region_ciudad.ciu_codigo " & _
             " LEFT JOIN region " & _
             " ON region_ciudad.reg_codigo=region.reg_codigo " & _
             " AND region_ciudad.emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY pai_nombre,ciu_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFGCiudad.DataSource = clsCon_Def.adorec_Def.DataSource
End Sub

Private Sub CargarMarcas(strRegion As String)
    VSFGVendedor.Clear 1
    VSFGVendedor.Rows = 1
    strSql = " SELECT mar_nombre, " & _
             " COALESCE(region_marca.ven_codigo,'') as ven,marca.mar_codigo,'0' as modi " & _
             " FROM marca LEFT JOIN region_marca " & _
             " ON marca.emp_codigo=region_marca.emp_codigo " & _
             " AND marca.mar_codigo=region_marca.mar_codigo AND region_marca.reg_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "' " & _
             " ORDER BY mar_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFGVendedor.DataSource = clsCon_Def.adorec_Def.DataSource
    strSql = " SELECT '  ' as ven_codigo,' -' as ven UNION SELECT ven_codigo,CONCAT(ven_apellido,' ',ven_nombre) as ven " & _
             " FROM vendedor " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY ven "
    clsCon_Def.Ejecutar strSql
    VSFGVendedor.ColComboList(1) = VSFGVendedor.BuildComboList(clsCon_Def.adorec_Def, "ven", "ven_codigo")
End Sub

Private Sub cmdAsignarVendedor_Click()
    If cmdAsignarVendedor.Caption = "A&signar Vendedores" Then
        cmdAsignarVendedor.Caption = "C&rear Regiones"
        Me.Width = Me.Width + 5000
        ucrtVSFG.Visible = False
        cmbAceptar.Visible = False
        cmdAsignarCiudad.Visible = False
        cmdAceptarVendedor.Visible = True
        VSFGVendedor.Visible = True
        VSFG.SelectionMode = flexSelectionByRow
        CargarMarcas VSFG.TextMatrix(VSFG.Row, 1)
    Else
        cmdAsignarVendedor.Caption = "A&signar Vendedores"
        Me.Width = Me.Width - 5000
        ucrtVSFG.Visible = True
        cmbAceptar.Visible = True
        cmdAsignarCiudad.Visible = True
        cmdAceptarVendedor.Visible = False
        VSFGVendedor.Visible = False
        VSFG.SelectionMode = flexSelectionFree
    End If
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSql = " SELECT reg_codigo,reg_nombre,reg_fechamod, reg_usumod, '0' as modi " & _
             " FROM region " & _
             " WHERE reg_codigo LIKE '%'"
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  reg_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSql = strSql & " AND  reg_nombre LIKE '%" & txtNombre.Text & "%' "
    End If
    strSql = strSql & " ORDER BY reg_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
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
            strSql = " UPDATE region " & _
                 " SET reg_nombre='" & UCase(VSFG.TextMatrix(i, 2)) & "'," & _
                 " reg_fechamod=CURRENT_TIMESTAMP," & _
                 " reg_usumod='" & strUsuario & "' " & _
                 " WHERE reg_codigo='" & VSFG.TextMatrix(i, 1) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el código", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT reg_codigo" & _
                    " FROM region " & _
                    " WHERE reg_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    strSql = " INSERT INTO region(reg_codigo,reg_nombre,reg_fechamod,reg_usumod) " & _
                            " VALUES ('" & UCase(VSFG.TextMatrix(i, 1)) & "','" & UCase(VSFG.TextMatrix(i, 2)) & "', " & _
                            " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                 
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "El código d" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
        
            strSql = " SELECT count(ciu_codigo) as existe " & _
                    " FROM region_ciudad " & _
                    " WHERE reg_codigo = '" & VSFG.TextMatrix(i, 1) & "'"
            clsCon_Def.Ejecutar (strSql)
            ' Si existe comedor no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSql = " SELECT count(reg_codigo) as existe " & _
                        " FROM region_marca " & _
                        " WHERE reg_codigo = '" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar (strSql)
                ' Si existe comedor no puedo eliminar
                If clsCon_Def.adorec_Def("existe") > 0 Then
                    MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                Else
                    strSql = " DELETE " & _
                            " FROM region " & _
                            " WHERE reg_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                    clsCon_Def.Ejecutar (strSql), "M"
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
    If chkFiltroNombre.Value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
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
    Me.Width = Me.Width - 5000
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    IniDato
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

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        If ucrtVSFG.Visible = True Then
            ucrtVSFG.VerMenu
        End If
    End If
End Sub

Private Sub VSFG_SelChange()
    If cmdAsignarCiudad.Caption = "C&rear Regiones" Then
        CargarCiudades VSFG.TextMatrix(VSFG.Row, 1)
    ElseIf cmdAsignarVendedor.Caption = "C&rear Regiones" Then
        CargarMarcas VSFG.TextMatrix(VSFG.Row, 1)
    End If
End Sub

Private Sub VSFGCiudad_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    Else
        If Not (VSFGCiudad.TextMatrix(Row, 3) = "" Or VSFGCiudad.TextMatrix(Row, 3) = VSFG.TextMatrix(VSFG.Row, 2)) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub VSFGCiudad_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        If Abs(FormatoD0(VSFGCiudad.TextMatrix(Row, 0))) = 1 Then
            VSFGCiudad.Cell(flexcpBackColor, Row, 0, Row, VSFGCiudad.Cols - 1) = vbYellow
            VSFGCiudad.Cell(flexcpForeColor, Row, 0, Row, VSFGCiudad.Cols - 1) = &H0
            VSFGCiudad.TextMatrix(Row, VSFGCiudad.Cols - 1) = 2
        Else
            VSFGCiudad.Cell(flexcpBackColor, Row, 0, Row, VSFGCiudad.Cols - 1) = vbWhite
            If Not (VSFGCiudad.TextMatrix(Row, 3) = "" Or VSFGCiudad.TextMatrix(Row, 3) = VSFG.TextMatrix(VSFG.Row, 2)) Then
                VSFGCiudad.Cell(flexcpForeColor, Row, 0, Row, VSFGCiudad.Cols - 1) = &HCCCCCC
            Else
                VSFGCiudad.Cell(flexcpForeColor, Row, 0, Row, VSFGCiudad.Cols - 1) = &H0
            End If
            VSFGCiudad.TextMatrix(Row, VSFGCiudad.Cols - 1) = 0
        End If
    End If
End Sub

Private Sub VSFGVendedor_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGVendedor_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        If Trim(VSFGVendedor.TextMatrix(Row, 1)) <> "" Then
            VSFGVendedor.Cell(flexcpBackColor, Row, 0, Row, VSFGVendedor.Cols - 1) = vbYellow
            VSFGVendedor.TextMatrix(Row, VSFGVendedor.Cols - 1) = 2
        Else
            VSFGVendedor.Cell(flexcpBackColor, Row, 0, Row, VSFGVendedor.Cols - 1) = vbWhite
            VSFGVendedor.TextMatrix(Row, VSFGVendedor.Cols - 1) = 0
        End If
    End If
End Sub
