VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComisiones 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Comisiones"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmComisiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   9240
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   360
      Left            =   3770
      TabIndex        =   9
      Top             =   9120
      Width           =   1700
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1850
      TabIndex        =   8
      Top             =   9120
      Width           =   1700
   End
   Begin VB.Frame fraDirector 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Directores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   9015
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8775
         _cx             =   15478
         _cy             =   5953
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmComisiones.frx":030A
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
   Begin VB.Frame fraGerente 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Gerentes de Zona"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   9015
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8775
         _cx             =   15478
         _cy             =   5953
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
         AllowUserResizing=   0
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
         FormatString    =   $"frmComisiones.frx":03FA
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo de Negocio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   375
         Width           =   4455
         _ExtentX        =   7858
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5690
      TabIndex        =   0
      Top             =   9120
      Width           =   1700
   End
End
Attribute VB_Name = "frmComisiones"
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
    Tipo = "Definición de Comisiones"
    Tipo2 = "Definición de Comisiones"
    Me.Caption = Tipo
End Sub

Private Sub CargaGerentes()
    strSql = " SELECT DISTINCT persona.per_codigo as codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as nombre,com_per_gz,com_per_dir,com_per_fechamod,com_per_usumod,'-3' as modi " & _
             " FROM persona " & _
             " INNER JOIN persona p1 " & _
             " ON persona.emp_codigo=p1.emp_codigo " & _
             " AND persona.per_codigo=p1.per_codigo_ref " & _
             " LEFT JOIN comision_persona " & _
             " ON comision_persona.emp_codigo=persona.emp_codigo " & _
             " AND comision_persona.per_codigo=persona.per_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = CStr(i)
    Next i
End Sub

Private Sub CargaDirectores()
    strSql = " SELECT DISTINCT persona.per_codigo as codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as nombre,com_per_dir,com_per_fechamod,com_per_usumod,'-3' as modi " & _
             " FROM persona " & _
             " INNER JOIN persona p1 " & _
             " ON persona.emp_codigo=p1.emp_codigo " & _
             " AND persona.per_codigo=p1.per_codigo_ref2 " & _
             " LEFT JOIN persona p2 " & _
             " ON persona.emp_codigo=p2.emp_codigo " & _
             " AND persona.per_codigo=p2.per_codigo_ref " & _
             " LEFT JOIN comision_persona " & _
             " ON comision_persona.emp_codigo=persona.emp_codigo " & _
             " AND comision_persona.per_codigo=persona.per_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND p2.per_codigo IS NULL " & _
             " AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set VSFG1.DataSource = clsCon_Def.adorec_Def.DataSource
    For i = 1 To VSFG1.Rows - 1
        VSFG1.TextMatrix(i, 0) = CStr(i)
    Next i
End Sub



Private Sub cmbNegocio_Change()
    
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            
            If strPtoFactura <> clsCon_Def.adorec_Def(0) Then
                strPtoFactura = clsCon_Def.adorec_Def(0)
                Limpiar
            Else
                strPtoFactura = clsCon_Def.adorec_Def(0)
            End If
        
            
        End If
    Else
        Exit Sub
    End If
  

End Sub

Private Sub Limpiar()
    CargaGerentes
    CargaDirectores
End Sub


Private Sub cmbNegocio_LostFocus()
    If cmbNegocio.BoundText = "" Then
        MsgBox "Primero seleccione un Tipo de Negocio", vbInformation, "Tipo de Negocio"
        cmbNegocio.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    'comision gerentes
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = "3" Then
            strSql = " SELECT count(*) " & _
                     " FROM comision_persona c " & _
                     " WHERE c.emp_codigo='" & strEmpresa & "' " & _
                     " AND c.per_codigo='" & VSFG.TextMatrix(i, 1) & "' "
            clsCon_Def.Ejecutar strSql
            If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                strSql = " UPDATE comision_persona " & _
                         " SET com_per_gz='" & FormatoD2(VSFG.TextMatrix(i, 3)) & "'," & _
                         " com_per_dir='" & FormatoD2(VSFG.TextMatrix(i, 4)) & "'," & _
                         " com_per_fechamod=CURRENT_TIMESTAMP," & _
                         " com_per_usumod='" & strUsuario & "' " & _
                         " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                         " AND emp_codigo='" & strEmpresa & "' "
                clsCon_Def.Ejecutar strSql, "M"
            Else
                strSql = " INSERT INTO comision_persona(emp_codigo,per_codigo,com_per_gz,com_per_dir,com_per_fechamod,com_per_usumod) VALUES('" & _
                         strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & FormatoD2(VSFG.TextMatrix(i, 3)) & "','" & _
                         FormatoD2(VSFG.TextMatrix(i, 4)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
                clsCon_Def.Ejecutar strSql, "M"
            End If
            
        End If
    Next i
    
    For i = 1 To VSFG1.Rows - 1
        If VSFG1.TextMatrix(i, VSFG1.Cols - 1) = "3" Then
            strSql = " SELECT count(*) " & _
                     " FROM comision_persona c " & _
                     " WHERE c.emp_codigo='" & strEmpresa & "' " & _
                     " AND c.per_codigo='" & VSFG1.TextMatrix(i, 1) & "' "
            clsCon_Def.Ejecutar strSql
            If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                strSql = " UPDATE comision_persona " & _
                         " SET com_per_dir='" & FormatoD2(VSFG1.TextMatrix(i, 3)) & "'," & _
                         " com_per_fechamod=CURRENT_TIMESTAMP," & _
                         " com_per_usumod='" & strUsuario & "' " & _
                         " WHERE per_codigo='" & VSFG1.TextMatrix(i, 1) & "' " & _
                         " AND emp_codigo='" & strEmpresa & "' "
                clsCon_Def.Ejecutar strSql, "M"
            Else
                strSql = " INSERT INTO comision_persona(emp_codigo,per_codigo,com_per_dir,com_per_fechamod,com_per_usumod) VALUES('" & _
                         strEmpresa & "','" & VSFG1.TextMatrix(i, 1) & "','" & FormatoD2(VSFG1.TextMatrix(i, 3)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
                clsCon_Def.Ejecutar strSql, "M"
            End If
            
        End If
    Next i
    MsgBox "Se guardó los cambios realizados", vbInformation, "Definición de Comisiones"
    Limpiar
End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= 2 Or Col >= VSFG.Cols - 3 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= 2 Or Col >= VSFG1.Cols - 3 Then
        Cancel = True
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
    IniDato
    
    cargarTipoPedido
    Limpiar
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
    If Col = 3 Or Col = 4 Then
        If VSFG.TextMatrix(Row, Col) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, Col)) Then
            VSFG.TextMatrix(Row, Col) = "0"
        End If
    End If
End Sub

Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG1.TextMatrix(Row, VSFG1.Cols - 1)) = -2 Then
        VSFG1.TextMatrix(Row, VSFG1.Cols - 1) = 2
    ElseIf Val(VSFG1.TextMatrix(Row, VSFG1.Cols - 1)) = -3 Then
        VSFG1.TextMatrix(Row, VSFG1.Cols - 1) = 3
    End If
    If Col = 3 Then
        If VSFG1.TextMatrix(Row, Col) <> "" And Not IsNumeric(VSFG1.TextMatrix(Row, Col)) Then
            VSFG1.TextMatrix(Row, Col) = "0"
        End If
    End If
End Sub
