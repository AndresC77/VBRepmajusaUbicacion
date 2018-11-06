VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRetencion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmRetencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10155
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3327
      TabIndex        =   7
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5127
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
      Left            =   1125
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
         Caption         =   "Filtrar Nombre de la Retención"
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
      Height          =   4080
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   9900
      _cx             =   2000045110
      _cy             =   2000034845
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRetencion.frx":030A
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
Attribute VB_Name = "frmRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSQL As String
Private Tipo As String
Private Tipo2 As String
Private AuxIni As String
Private AuxFin As String
Private Sub IniDato()
    Tipo = " Retención "
    Tipo2 = " la Retención "
    Me.Caption = Tipo
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
  
    strSQL = " SELECT ret_codigo,ret_nombre,ret_descripcion,ret_ctaconta,ret_ctacontacli,ret_porcentaje,ret_gravara," & _
             " ret_fechaini, COALESCE(ret_fechafin,''), ret_activo, " & _
             " ret_fechaini, COALESCE(ret_fechafin,''),ret_fechamod, ret_usumod, '0' as modi" & _
             " FROM retencion " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             "  "
  
    If chkFiltroNombre.Value = 1 Then
        strSQL = strSQL & " AND  ret_nombre LIKE '%" & txtNombre.Text & "%' "
    End If
    strSQL = strSQL & " ORDER BY ret_codigo,ret_activo  "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    VSFG.ColComboList(7) = "IVA|IVAPRODUCTOS|IVASERVICIOS|SUBTOTAL|SUBTOTALPRODUCTOS|SUBTOTALSERVICIOS|IVA0%|TOTAL"
    'crea combo de categoria
    strSQL = " SELECT cta_codigo, CONCAT(cta_codigo,' - ',cta_nombre) as cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cta_subcta=0 " & _
                 " ORDER BY cta_codigo"
     clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_nombre", "cta_codigo")
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
            strSQL = " UPDATE retencion " & _
                     " SET ret_nombre='" & UCase(VSFG.TextMatrix(i, 2)) & "'," & _
                     " ret_descripcion='" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                     " ret_ctaconta='" & UCase(VSFG.TextMatrix(i, 4)) & "'," & _
                     " ret_ctacontacli='" & UCase(VSFG.TextMatrix(i, 5)) & "'," & _
                     " ret_porcentaje='" & VSFG.TextMatrix(i, 6) & "'," & _
                     " ret_gravara='" & VSFG.TextMatrix(i, 7) & "'," & _
                     " ret_fechaini='" & VSFG.TextMatrix(i, 8) & "'," & _
                     " ret_fechafin='" & VSFG.TextMatrix(i, 9) & "'," & _
                     " ret_activo='" & Abs(FormatoD0(VSFG.TextMatrix(i, 10))) & "'," & _
                     " ret_fechamod=CURRENT_TIMESTAMP," & _
                     " ret_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND ret_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                     " AND ret_fechaini='" & VSFG.TextMatrix(i, 11) & "'" & _
                     " AND ret_fechafin='" & VSFG.TextMatrix(i, 12) & "'"
             clsCon_Def.Ejecutar strSQL, "M"
             If Abs(FormatoD0(VSFG.TextMatrix(i, 10))) = 1 Then
                strSQL = " UPDATE retencion " & _
                         " SET ret_activo='0'," & _
                         " ret_fechamod=CURRENT_TIMESTAMP," & _
                         " ret_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND ret_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND ret_fechaini!='" & VSFG.TextMatrix(i, 11) & "'" & _
                         " AND ret_fechafin!='" & VSFG.TextMatrix(i, 12) & "'"
                 clsCon_Def.Ejecutar strSQL, "M"
             End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el codigo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 4) = "" And VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta cuenta contable", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 6) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta porcentaje", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 7) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta base imponible", vbInformation, "Ingreso"
                control = 1
            Else
                strSQL = " SELECT count(ret_codigo) as existe " & _
                     " FROM retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ret_codigo='" & UCase(VSFG.TextMatrix(i, 1)) & "'" & _
                     " AND ret_fechaini='" & VSFG.TextMatrix(i, 9) & "'" & _
                     " AND ret_fechafin='" & VSFG.TextMatrix(i, 10) & "'"
                clsCon_Def.Ejecutar (strSQL)
                ' Si existe  no puedo eliminar
                If clsCon_Def.adorec_Def("existe") <= 0 Then
                
                     strSQL = " INSERT INTO retencion " & _
                              " (emp_codigo,ret_codigo,ret_nombre,ret_descripcion,ret_ctaconta,ret_ctacontacli," & _
                              " ret_porcentaje,ret_gravara,ret_fechaini,ret_fechafin,ret_activo,ret_fechamod,ret_usumod) " & _
                              " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "'," & _
                              " '" & UCase(VSFG.TextMatrix(i, 2)) & "'," & _
                              " '" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                              " '" & VSFG.TextMatrix(i, 4) & "'," & _
                              " '" & VSFG.TextMatrix(i, 5) & "'," & _
                              " '" & VSFG.TextMatrix(i, 6) & "'," & _
                              " '" & VSFG.TextMatrix(i, 7) & "'," & _
                              " '" & VSFG.TextMatrix(i, 8) & "'," & _
                              " '" & VSFG.TextMatrix(i, 9) & "'," & _
                              " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 10))) & "'," & _
                              " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar (strSQL), "M"
                    If Abs(FormatoD0(VSFG.TextMatrix(i, 10))) = 1 Then
                       strSQL = " UPDATE retencion " & _
                                " SET ret_activo='0'," & _
                                " ret_fechamod=CURRENT_TIMESTAMP," & _
                                " ret_usumod='" & strUsuario & "' " & _
                                " WHERE emp_codigo='" & strEmpresa & "'" & _
                                " AND ret_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                                " AND ret_fechaini!='" & VSFG.TextMatrix(i, 8) & "'" & _
                                " AND ret_fechafin!='" & VSFG.TextMatrix(i, 9) & "'"
                        clsCon_Def.Ejecutar strSQL, "M"
                    End If
                Else
                    MsgBox "El código de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
            strSQL = " SELECT COALESCE(count(det_comp_ret.ret_codigo),0) as existe " & _
                     " FROM det_comp_ret INNER JOIN comprobante_retencion ON det_comp_ret.emp_codigo=comprobante_retencion.emp_codigo AND det_comp_ret.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND det_comp_ret.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo" & _
                     " INNER JOIN retencion ON det_comp_ret.emp_codigo=retencion.emp_codigo AND det_comp_ret.ret_codigo=retencion.ret_codigo WHERE det_comp_ret.emp_codigo='" & strEmpresa & "' " & _
                     " AND det_comp_ret.ret_codigo='" & UCase(VSFG.TextMatrix(i, 1)) & "' AND com_ret_fecha between '" & VSFG.TextMatrix(i, 8) & "' and '" & VSFG.TextMatrix(i, 9) & "'"
            clsCon_Def.Ejecutar (strSQL)
                ' Si existe  no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSQL = " DELETE " & _
                         " FROM retencion " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND ret_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar (strSQL), "M"
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga
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
    If Col = 8 Then
        AuxIni = VSFG.TextMatrix(VSFG.Row, 8)
    ElseIf Col = 9 Then
        AuxFin = VSFG.TextMatrix(VSFG.Row, 9)
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
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set Me.ucrtVSFG.VSFGControl = VSFG
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
    Dim i As Long
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
    If Col = 10 And Abs(FormatoD0(VSFG.TextMatrix(VSFG.Row, 10))) = 1 Then
        MsgBox "Si activa este codigo de retencion se desactivarán los otros codigos del mismo numero", vbInformation, "Retencion"
    End If
    If (Col = 9 Or Col = 8 Or Col = 1) And (VSFG.TextMatrix(VSFG.Row, 8) <> "" And VSFG.TextMatrix(VSFG.Row, 9) <> "") Then
        VSFG.TextMatrix(VSFG.Row, 8) = Format(VSFG.TextMatrix(VSFG.Row, 8), "yyyy-mm-dd")
        VSFG.TextMatrix(VSFG.Row, 9) = Format(VSFG.TextMatrix(VSFG.Row, 9), "yyyy-mm-dd")
        If VSFG.TextMatrix(VSFG.Row, 8) >= VSFG.TextMatrix(VSFG.Row, 9) Then
            MsgBox "La Fecha inicial debe ser menos a la final ", vbInformation, "Retenciones"
            Exit Sub
        End If
        For i = 1 To VSFG.Rows - 1
            If i <> VSFG.Row Then
                If VSFG.TextMatrix(VSFG.Row, 1) = VSFG.TextMatrix(i, 1) Then
                    If Col = 8 Then
                        If VSFG.TextMatrix(i, 8) <= VSFG.TextMatrix(VSFG.Row, 8) And VSFG.TextMatrix(VSFG.Row, 8) <= VSFG.TextMatrix(i, 9) Then
                            MsgBox "La Fecha inicial tiene una inconsistencia con la retencion de la fila " & i, vbInformation, "Retenciones"
                            VSFG.TextMatrix(VSFG.Row, 8) = AuxIni
                            Exit Sub
                        End If
                    ElseIf Col = 9 Then
                        If VSFG.TextMatrix(i, 8) <= VSFG.TextMatrix(VSFG.Row, 9) And VSFG.TextMatrix(VSFG.Row, 9) <= VSFG.TextMatrix(i, 9) Then
                            MsgBox "La Fecha final tiene una inconsistencia con la retencion de la fila " & i, vbInformation, "Retenciones"
                            VSFG.TextMatrix(VSFG.Row, 9) = AuxFin
                            Exit Sub
                        End If
                    End If
                    If Col = 8 Then
                        If VSFG.TextMatrix(VSFG.Row, 8) <= VSFG.TextMatrix(i, 8) And VSFG.TextMatrix(i, 8) <= VSFG.TextMatrix(VSFG.Row, 9) Then
                            MsgBox "La Fecha inicial tiene una inconsistencia con la retencion de la fila " & i, vbInformation, "Retenciones"
                            VSFG.TextMatrix(VSFG.Row, 8) = AuxIni
                            Exit Sub
                        End If
                    ElseIf Col = 9 Then
                        If VSFG.TextMatrix(VSFG.Row, 8) <= VSFG.TextMatrix(i, 9) And VSFG.TextMatrix(i, 9) <= VSFG.TextMatrix(VSFG.Row, 9) Then
                            MsgBox "La Fecha final tiene una inconsistencia con la retencion de la fila " & i, vbInformation, "Retenciones"
                            VSFG.TextMatrix(VSFG.Row, 9) = AuxFin
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next i
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

Private Sub VSFG_RowColChange()
    If Abs(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = 2 Then
        If VSFG.TextMatrix(VSFG.Row, 8) = "" Then
            VSFG.TextMatrix(VSFG.Row, 8) = HoyDia
        End If
        If VSFG.TextMatrix(VSFG.Row, 9) = "" Then
            VSFG.TextMatrix(VSFG.Row, 9) = DateAdd("d", 1, HoyDia)
        End If
    End If
End Sub
