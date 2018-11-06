VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCtaBanco 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta. Banco"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmCtaBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10155
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
         Caption         =   "Filtrar Código Banco"
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
         Caption         =   "FiltrarNúmero de Cuenta"
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
      Begin MSDataListLib.DataCombo dcmbTipoBanco 
         Height          =   330
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código de Banco Banco"
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
         Caption         =   "Número de Cuenta"
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
      _cx             =   17462
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCtaBanco.frx":030A
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
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmCtaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSQL As String
Private Tipo As String
Private Tipo2 As String
Private Sub IniDato()
    Tipo = " Cuentas Bancarias "
    Tipo2 = " Cuentas Bancarias "
    Me.Caption = Tipo
End Sub
Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo = 0 Then
       dcmbTipoBanco.Enabled = True
    Else
       strSQL = " SELECT ban_nombre,ban_codigo " & _
                " FROM banco " & _
                " ORDER BY ban_nombre "
       clsCon_Def.Ejecutar strSQL
       Set dcmbTipoBanco.RowSource = clsCon_Def.adorec_Def.DataSource
       dcmbTipoBanco.ListField = "ban_nombre"
       dcmbTipoBanco.BoundColumn = "ban_codigo"
    End If
    
    If chkFiltroCodigo.Value = 1 Then
       dcmbTipoBanco.Enabled = True
    Else
       dcmbTipoBanco.Enabled = False
    End If
    
End Sub
Private Sub cmdMostrar_Click()
    Carga
End Sub

Private Sub Carga()
  
     strSQL = " SELECT cta_ban_numero,ban_codigo,cta_ban_ctaconta, cta_ban_ch_ultimo," & _
              " cta_ban_saldoreal,cta_ban_saldodisponible,cta_ban_saldoprevisto," & _
              " cta_ban_observacion," & _
              " cta_ban_fechamod, cta_ban_usumod, '0' as modi" & _
              " FROM cta_banco " & _
              " WHERE emp_codigo='" & strEmpresa & "' "
    If chkFiltroCodigo.Value = 1 Then
        strSQL = strSQL & "AND  cta_banco.ban_codigo LIKE  '" & dcmbTipoBanco.BoundText & "'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSQL = strSQL & " AND  cta_ban_numero LIKE '%" & txtNombre.Text & "%' "
    End If
    strSQL = strSQL & " ORDER BY cta_banco.ban_codigo "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    'crea combo
    strSQL = " SELECT ban_codigo, ban_nombre " & _
             " FROM banco " & _
             " ORDER BY ban_nombre "
    clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *ban_nombre", "ban_codigo")
    'crea combo
    strSQL = " SELECT cta_codigo, cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cta_subcta=0 " & _
             " ORDER BY cta_codigo "
    clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cta_codigo,cta_nombre", "cta_codigo")
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
        
            strSQL = " UPDATE cta_banco " & _
                     " SET cta_ban_ctaconta='" & VSFG.TextMatrix(i, 3) & "'," & _
                     " cta_ban_ch_ultimo ='" & VSFG.TextMatrix(i, 4) & "'," & _
                     " cta_ban_saldoreal='" & Replace(VSFG.TextMatrix(i, 5), ",", ".") & "'," & _
                     " cta_ban_saldodisponible='" & Replace(VSFG.TextMatrix(i, 6), ",", ".") & "'," & _
                     " cta_ban_saldoprevisto='" & Replace(VSFG.TextMatrix(i, 7), ",", ".") & "'," & _
                     " cta_ban_observacion='" & UCase(VSFG.TextMatrix(i, 8)) & "', " & _
                     " cta_ban_fechamod = CURRENT_TIMESTAMP, " & _
                     " cta_ban_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo = '" & strEmpresa & "'" & _
                     " AND ban_codigo = '" & VSFG.TextMatrix(i, 2) & "'" & _
                     " AND cta_ban_numero = '" & VSFG.TextMatrix(i, 1) & "'"
            clsCon_Def.Ejecutar strSQL, "M"
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta banco", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta número de cuenta", vbInformation, "Ingreso"
                control = 1
            Else
                strSQL = "SELECT(cta_ban_numero)  " & _
                         "FROM cta_banco " & _
                         "WHERE emp_codigo = '" & strEmpresa & "'" & _
                         "AND ban_codigo = '" & VSFG.TextMatrix(i, 2) & "'" & _
                         "AND cta_ban_numero = '" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar strSQL
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    strSQL = " INSERT INTO cta_banco" & _
                             "(cta_ban_numero,ban_codigo,emp_codigo,cta_ban_ctaconta," & _
                             " cta_ban_ch_ultimo,cta_ban_saldoreal,cta_ban_saldodisponible,cta_ban_saldoprevisto," & _
                             " cta_ban_observacion,cta_ban_fechamod,cta_ban_usumod) " & _
                             " VALUES (UPPER('" & VSFG.TextMatrix(i, 1) & "')," & _
                             " '" & VSFG.TextMatrix(i, 2) & "','" & strEmpresa & "'," & _
                             " '" & VSFG.TextMatrix(i, 3) & "','" & VSFG.TextMatrix(i, 4) & "'," & _
                             " '" & 0 & "','" & 0 & "','" & 0 & "','" & UCase(VSFG.TextMatrix(i, 8)) & "'," & _
                             " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar (strSQL), "M"
                Else
                    MsgBox "El número de cuenta ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
            
            strSQL = " SELECT count(*) As existe " & _
                     " FROM egreso_comun " & _
                     " WHERE cta_ban_numero = '" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSQL)
       
               
            ' Si existe no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSQL = " DELETE " & _
                         " FROM cta_banco " & _
                         " WHERE cta_ban_numero='" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND ban_codigo = '" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND emp_codigo = '" & strEmpresa & "'"
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

Private Sub VSFG_DblClick()
    Dim i As Long
    If VSFG.Row >= 1 Then
        frmDatos.Show
        frmDatos.VSFG.Rows = VSFG.Cols
        For i = 1 To VSFG.Cols - 1
            frmDatos.VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(0, i)
            frmDatos.VSFG.TextMatrix(i, 1) = VSFG.Cell(flexcpTextDisplay, VSFG.Row, i)
            If VSFG.ColComboList(i) <> "" Then
                frmDatos.VSFG.TextMatrix(i, 2) = VSFG.ColComboList(i)
            End If
        Next i
        frmDatos.VSFG.Cell(flexcpBackColor, 1, 1, frmDatos.VSFG.Rows - 1, 1) = VSFG.Cell(flexcpBackColor, VSFG.Row, VSFG.Col)
        frmDatos.VSFG.RowHidden(frmDatos.VSFG.Rows - 1) = True
        Set frmDatos.VSFGOrigen = VSFG
        frmDatos.VSFGOrigen.Tag = VSFG.Row
        frmDatos.Caption = Tipo
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
        If Col <= 2 Or Col >= VSFG.Cols - 3 Then
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


