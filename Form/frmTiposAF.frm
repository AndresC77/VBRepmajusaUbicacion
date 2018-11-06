VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTiposAF 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Activos Fijos"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmTiposAF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8265
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3240
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   7980
      _cx             =   14076
      _cy             =   5715
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTiposAF.frx":030A
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
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
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
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2390
      TabIndex        =   7
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4190
      TabIndex        =   6
      Top             =   7200
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
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
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
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG1 
      Height          =   1320
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   5940
      _cx             =   10477
      _cy             =   2328
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTiposAF.frx":0498
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
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
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
Attribute VB_Name = "frmTiposAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Dim newr As Long, oldr As Long

Private Sub IniDato()
    Tipo = "Tipo Activo Fijo"
    Tipo2 = "Tipo Activo Fijo"
    Me.Caption = Tipo
End Sub

Private Sub cmdMostrar_Click()
    Carga

End Sub
Private Sub Carga()
    strSql = "Select count(*) from area where emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def(0) = 0 Then
        MsgBox "Necesita ingresar áreas primero", vbInformation, "Tipo Activo Fijo"
        Exit Sub
        Unload Me
    End If
    strSql = " SELECT tip_act_codigo,tip_act_nombre,tip_act_ctaconta,tip_act_ctaconta2," & _
             " tip_act_ctaconta3,tip_act_ctaconta4,tip_act_fechamod,tip_act_usumod, '0' as modi " & _
             " FROM tipo_activo " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & "AND  tip_act_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.value = 1 Then
        strSql = strSql & " AND  tip_act_nombre LIKE '%" & txtNombre.Text & "%' "
    End If
    strSql = strSql & " ORDER BY 1,2"
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum

    CargarCtaGasto
    'Combos
     strSql = " Select cta_codigo,cta_nombre,concat(cta_codigo,' - ',cta_nombre) as nombre From ctaconta " & _
             " Where emp_codigo = '" & strEmpresa & "' And cta_subcta = 0 " & _
             " Order By cta_codigo "
    'Ejecuta la consulta anterior
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*nombre", "cta_codigo")
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*nombre", "cta_codigo")
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*nombre", "cta_codigo")
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*nombre", "cta_codigo")
  
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long, x As Long
    Dim control As Long 'control de que esten llenos los datos
    If VSFG.Rows > 1 Then
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            If Trim(VSFG.TextMatrix(i, 2)) = "" Then
                MsgBox "No puede modificar " & Tipo2 & " falta nombre", vbInformation, "Modificación"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 3)) = "" Then
                MsgBox "No puede modificar " & Tipo2 & " falta la cuenta contable de activo", vbInformation, "Modificación"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 4)) = "" Then
                MsgBox "No puede modificar " & Tipo2 & " falta la cuenta contable de depreciación", vbInformation, "Modificación"
                control = 1
            Else
           strSql = " UPDATE tipo_activo " & _
                 " SET tip_act_nombre='" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "'," & _
                 " tip_act_ctaconta='" & VSFG.TextMatrix(i, 3) & "', " & _
                 " tip_act_ctaconta2='" & VSFG.TextMatrix(i, 4) & "', " & _
                 " tip_act_ctaconta3='" & VSFG.TextMatrix(i, 5) & "', " & _
                 " tip_act_ctaconta4='" & VSFG.TextMatrix(i, 6) & "', " & _
                 " tip_act_fechamod=CURRENT_TIMESTAMP," & _
                 " tip_act_usumod='" & strUsuario & "' " & _
                 " WHERE tip_act_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                 clsCon_Def.Ejecutar strSql, "M"
                 
                 GuardarDetGasto VSFG.TextMatrix(i, 1)
            End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If Trim(VSFG.TextMatrix(i, 1)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el código", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 2)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 3)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la cuenta contable de activo", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 4)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la cuenta contable de depreciación", vbInformation, "Ingreso"
                control = 1
            Else
                Dim conta As Long
                For x = 1 To VSFG1.Rows - 1
                    If VSFG1.TextMatrix(x, 1) = "" Then conta = conta + 1
                Next x
                
                If conta <> 0 Then
                    MsgBox "No puede ingresar " & Tipo2 & " existe datos de áreas vacíos", vbInformation, "Ingreso"
                    control = 1
                Else
                strSql = " SELECT tip_act_codigo" & _
                    " FROM tipo_activo " & _
                    " WHERE tip_act_codigo='" & Trim(VSFG.TextMatrix(i, 1)) & "' AND emp_codigo='" & strEmpresa & "'"
                    
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                
                    strSql = " INSERT INTO tipo_activo(emp_codigo,tip_act_codigo,tip_act_nombre,tip_act_ctaconta," & _
                            " tip_act_ctaconta2,tip_act_ctaconta3,tip_act_ctaconta4,tip_act_fechamod,tip_act_usumod) " & _
                            " VALUES ('" & strEmpresa & "','" & UCase(Trim(VSFG.TextMatrix(i, 1))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "'," & _
                            "'" & VSFG.TextMatrix(i, 3) & "','" & VSFG.TextMatrix(i, 4) & "','" & VSFG.TextMatrix(i, 5) & "','" & VSFG.TextMatrix(i, 6) & "', " & _
                            " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                    If VSFG1.TextMatrix(1, 0) = "" Then
                        MsgBox "No puede ingresar " & Tipo2 & " falta especificar el área", vbInformation, "Ingreso"
                        control = 1
                    Else
                    GuardarDetGasto VSFG.TextMatrix(i, 1)
                    End If
                Else
                    MsgBox "No puede ingresar " & Tipo2 & " ya existe el código", vbInformation, "Ingreso"
                    control = 1
                End If
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
            strSql = " SELECT count(*) As existe " & _
                " FROM activo_fijo " & _
                " WHERE tip_act_codigo='" & Trim(VSFG.TextMatrix(i, 1)) & "'" & _
                " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql)
        
            ' Si existe no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSql = " DELETE " & _
                    " FROM tipo_activo " & _
                    " WHERE tip_act_codigo='" & Trim(VSFG.TextMatrix(i, 1)) & "'"
                clsCon_Def.Ejecutar strSql, "M"
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    End If
    If control = 0 Then
        Carga
    End If
    
End Sub



Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    If VSFG.Rows > 1 Then
'        If NewRow <> OldRow Then
'            CargarCtaGasto
'        End If
'    End If
    newr = NewRow
    oldr = OldRow
End Sub

Private Sub VSFG_Click()
'    If VSFG.Rows > 1 Then
'            If VSFG.TextMatrix(VSFG.Row, 1) <> "" Then
''                strSql = " SELECT are_codigo,det_gas_act_are_ctaconta" & _
''                     " FROM det_gasto_act_are " & _
''                     " WHERE emp_codigo='" & strEmpresa & "' " & _
''                     " AND tip_act_codigo='" & VSFG.TextMatrix(NewRow, 1) & "' "
''                clsCon_Def.Ejecutar strSql
''                Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
''                'VSFG.TextMatrix(1, 0) = clsCon_Def.adorec_Def(0)
''                'VSFG.TextMatrix(1, 1) = clsCon_Def.adorec_Def(1)
'                CargarCtaGasto
'            End If
'    End If

     If VSFG.Rows > 1 Then
        If newr <> oldr Then
            CargarCtaGasto
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

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim x As Long
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
        VSFG1.Editable = 0
        If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
            VSFG1.Cell(flexcpBackColor, 1, 1, 1, VSFG1.Cols - 1) = &HC0C0FF
        Else
            VSFG1.Cell(flexcpBackColor, 1, 1, 1, VSFG1.Cols - 1) = vbDefault
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
'        VSFG1.TextMatrix(1, 0) = ""
'        VSFG1.TextMatrix(1, 1) = ""
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
        For x = 1 To VSFG1.Rows - 1
            VSFG1.TextMatrix(x, 1) = ""
         Next x
        VSFG1.Editable = 2
        VSFG1.Cell(flexcpBackColor, 1, 1, 1, VSFG1.Cols - 1) = &H80FFFF
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
         
        VSFG1.Editable = 2
        VSFG1.Cell(flexcpBackColor, 1, 1, 1, VSFG1.Cols - 1) = &HC0FFC0
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

Private Sub CargarCtaGasto(Optional CodSum As String)
    Dim where As String
    
    strSql = " SELECT area.are_codigo,COALESCE(det_gas_act_are_ctaconta,'') as det_gas_act_are_ctaconta" & _
         " FROM area LEFT JOIN det_gasto_act_are ON area.emp_codigo=det_gasto_act_are.emp_codigo " & _
         " AND area.are_codigo=det_gasto_act_are.are_codigo" & _
         " AND tip_act_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "' " & _
         " WHERE area.emp_codigo='" & strEmpresa & "' "

        clsCon_Def.Ejecutar strSql
        Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource

    
    strSql = " SELECT are_codigo,are_nombre " & _
             " FROM area " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql

    VSFG1.ColComboList(0) = VSFG1.BuildComboList(clsCon_Def.adorec_Def, "are_codigo, *are_nombre", "are_codigo")
    strSql = " Select cta_codigo,cta_nombre,concat(cta_codigo,' - ',cta_nombre) as nombre From ctaconta " & _
             " Where emp_codigo = '" & strEmpresa & "' And cta_subcta = 0 " & _
             " Order By cta_codigo "
    clsCon_Def.Ejecutar strSql
    VSFG1.ColComboList(1) = VSFG1.BuildComboList(clsCon_Def.adorec_Def, "*nombre", "cta_codigo")

End Sub

Private Sub GuardarDetGasto(Codigo As String)
Dim i As Long
    
    strSql = " DELETE FROM det_gasto_act_are WHERE emp_codigo='" & strEmpresa & "' AND tip_act_codigo='" & Codigo & "' "
    clsCon_Def.Ejecutar strSql, "M"
    
    For i = 1 To VSFG1.Rows - 1
    strSql = " INSERT INTO det_gasto_act_are(emp_codigo,tip_act_codigo,are_codigo,det_gas_act_are_ctaconta,det_gas_act_are_fechamod,det_gas_act_are_usumod) " & _
             " VALUES('" & strEmpresa & "','" & Codigo & "','" & VSFG1.TextMatrix(i, 0) & "','" & VSFG1.TextMatrix(i, 1) & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
    clsCon_Def.Ejecutar strSql, "M"
    Next i
End Sub


Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim valor As String
    valor = Abs(Val(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)))
    VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = valor
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub
