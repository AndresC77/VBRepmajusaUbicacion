VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLinMarConta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definiciónes de Cuentas de Inventario"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmLinMarConta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11850
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3327
      TabIndex        =   1
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5127
      TabIndex        =   0
      Top             =   6480
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   5760
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11580
      _cx             =   20426
      _cy             =   10160
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLinMarConta.frx":030A
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
      FrozenCols      =   3
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
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmLinMarConta"
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
    Tipo = " Cuenta "
    Tipo2 = " la Cuenta "
    Me.Caption = Tipo
End Sub

Private Sub Carga()
    strSql = " SELECT sucursal.suc_codigo,linea.lin_codigo,marca.mar_codigo,cen_cos_codigo," & _
             " cen_cos_lin_mar_cta_fechamod,cen_cos_lin_mar_cta_usumod, '0' as modi" & _
             " FROM sucursal INNER JOIN linea ON sucursal.emp_codigo=linea.emp_codigo " & _
             " INNER JOIN marca ON sucursal.emp_codigo=marca.emp_codigo AND linea.emp_codigo=marca.emp_codigo " & _
             " LEFT JOIN centro_costo_linea_marca ON sucursal.emp_codigo=centro_costo_linea_marca.emp_codigo AND sucursal.suc_codigo=centro_costo_linea_marca.suc_codigo " & _
             " AND linea.emp_codigo=centro_costo_linea_marca.emp_codigo AND linea.lin_codigo=centro_costo_linea_marca.lin_codigo " & _
             " AND marca.emp_codigo=centro_costo_linea_marca.emp_codigo AND marca.mar_codigo=centro_costo_linea_marca.mar_codigo " & _
             " WHERE sucursal.emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY suc_codigo,lin_codigo,mar_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    strSql = " SELECT suc_codigo,suc_nombre" & _
             " FROM sucursal " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY suc_codigo "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "suc_codigo,*suc_nombre", "suc_codigo")
    strSql = " SELECT lin_codigo,lin_nombre" & _
             " FROM linea " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY lin_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "lin_codigo,*lin_nombre", "lin_codigo")
    strSql = " SELECT mar_codigo,mar_nombre" & _
             " FROM marca " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY mar_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "mar_codigo,*mar_nombre", "mar_codigo")
    strSql = " SELECT cen_cos_codigo, CONCAT(cen_cos_codigo,' - ',cen_cos_nombre) as cen_cos_nombre" & _
             " FROM centro_costo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY cen_cos_codigo"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cen_cos_nombre", "cen_cos_codigo")
    ucrtVSFG.PonerNum
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    strSql = " DELETE FROM centro_costo_linea_marca " & _
             " WHERE emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar strSql, "M"
    For i = 1 To VSFG.Rows - 1
        strSql = " INSERT INTO centro_costo_linea_marca (emp_codigo,suc_codigo,lin_codigo,mar_codigo," & _
                 " cen_cos_codigo," & _
                 " cen_cos_lin_mar_cta_fechamod,cen_cos_lin_mar_cta_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(i, 1)) & "','" & UCase(VSFG.TextMatrix(i, 2)) & "','" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                 " '" & UCase(VSFG.TextMatrix(i, 4)) & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
    Next i
    Carga
    
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col <= 3 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
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

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
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
    ucrtVSFG.Inicializar False
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
        ucrtVSFG.VerMenu
    End If
End Sub
