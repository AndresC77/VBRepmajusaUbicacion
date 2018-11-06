VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSucursal 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sucursales"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmSucursal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10815
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3657
      TabIndex        =   1
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5457
      TabIndex        =   0
      Top             =   4800
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4080
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10620
      _cx             =   18732
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSucursal.frx":030A
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
End
Attribute VB_Name = "frmSucursal"
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
    Tipo = "Sucursales"
    Tipo2 = "la Sucursal"
    Me.Caption = Tipo
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSql = " SELECT suc_codigo,suc_nombre,dep_codigo,suc_ctaconta_ventas,suc_ctaconta_ventas_sp," & _
             " suc_ctaconta_servicios,suc_ctaconta_servicios_sp,suc_ctaconta_costoventa, " & _
             " suc_direccion,suc_telefono,suc_ciudad,suc_fechamod,suc_usumod, '0' as modi " & _
             " FROM sucursal " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY suc_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    'crea combo
    strSql = " SELECT cta_codigo, CONCAT(cta_codigo,' - ',cta_nombre) as cta_nombre" & _
            " FROM ctaconta " & _
            " WHERE emp_codigo = '" & strEmpresa & "'" & _
            " AND cta_subcta = 0 " & _
            " ORDER BY cta_codigo"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    'crea combo
    strSql = " SELECT dep_codigo, dep_nombre " & _
            " FROM deposito " & _
            " WHERE emp_codigo = '" & strEmpresa & "'" & _
            " ORDER BY dep_codigo"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " dep_codigo,*dep_nombre", "dep_codigo")
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
            strSql = " UPDATE sucursal " & _
                     " SET suc_nombre='" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "'," & _
                     " dep_codigo='" & UCase(Trim(VSFG.TextMatrix(i, 3))) & "'," & _
                     " suc_ctaconta_ventas='" & UCase(Trim(VSFG.TextMatrix(i, 4))) & "'," & _
                     " suc_ctaconta_ventas_sp='" & UCase(Trim(VSFG.TextMatrix(i, 5))) & "'," & _
                     " suc_ctaconta_servicios='" & UCase(Trim(VSFG.TextMatrix(i, 6))) & "'," & _
                     " suc_ctaconta_servicios_sp='" & UCase(Trim(VSFG.TextMatrix(i, 7))) & "'," & _
                     " suc_ctaconta_costoventa='" & UCase(Trim(VSFG.TextMatrix(i, 8))) & "'," & _
                     " suc_direccion='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "'," & _
                     " suc_telefono='" & UCase(Trim(VSFG.TextMatrix(i, 10))) & "'," & _
                     " suc_ciudad='" & UCase(Trim(VSFG.TextMatrix(i, 11))) & "'," & _
                     " suc_fechamod=CURRENT_TIMESTAMP," & _
                     " suc_usumod='" & strUsuario & "' " & _
                     " WHERE suc_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el código", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el nombre ", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la bodega", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 4) = "" Or VSFG.TextMatrix(i, 5) = "" Or VSFG.TextMatrix(i, 6) = "" Or VSFG.TextMatrix(i, 7) = "" Or VSFG.TextMatrix(i, 8) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la cuenta contable", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT suc_codigo" & _
                    " FROM sucursal " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND suc_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    strSql = " INSERT INTO sucursal(emp_codigo,suc_codigo,suc_nombre,dep_codigo,suc_ctaconta_ventas," & _
                             " suc_ctaconta_ventas_sp,suc_ctaconta_servicios,suc_ctaconta_servicios_sp,suc_ctaconta_costoventa," & _
                             " suc_direccion,suc_telefono,suc_ciudad,suc_fechamod,suc_usumod) " & _
                        " VALUES ('" & strEmpresa & "','" & UCase(Trim(VSFG.TextMatrix(i, 1))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 3))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 4))) & "', " & _
                        " '" & UCase(Trim(VSFG.TextMatrix(i, 5))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 6))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 7))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 8))) & "'," & _
                        " '" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 10))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 11))) & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "El código d" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
        
            strSql = " SELECT count(egr_codigo) as existe " & _
                    " FROM egreso " & _
                    " WHERE emp_codigo='" & strEmpresa & "' AND tip_egr_codigo IN ('FAC') " & _
                    " AND egr_codigo LIKE '" & FormatoD0(VSFG.TextMatrix(i, 1)) & "%'"
            clsCon_Def.Ejecutar (strSql)
            ' Si existe egreso
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
            
                strSql = " SELECT count(ing_codigo) as existe " & _
                        " FROM ingreso " & _
                        " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo IN ('DCL') " & _
                        " AND ing_codigo LIKE '" & FormatoD0(VSFG.TextMatrix(i, 1)) & "%'"
                clsCon_Def.Ejecutar (strSql)
                ' Si existe ingresos
                If clsCon_Def.adorec_Def("existe") > 0 Then
                    MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
                Else
                    strSql = " DELETE " & _
                      " FROM sucursal " & _
                      " WHERE emp_codigo='" & strEmpresa & "'" & _
                      " AND suc_codigo='" & VSFG.TextMatrix(i, 1) & "'"
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
