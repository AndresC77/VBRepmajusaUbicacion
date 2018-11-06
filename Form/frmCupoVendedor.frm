VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCupoVendedor 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cupo por Vendedor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmCupoVendedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6870
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6615
      _cx             =   11668
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
      FormatString    =   $"frmCupoVendedor.frx":030A
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
   Begin VB.CommandButton cmbLimpiar 
      Caption         =   "&Limpiar"
      Height          =   360
      Left            =   1685
      TabIndex        =   2
      Top             =   4680
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3485
      TabIndex        =   1
      Top             =   4680
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Vendedores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6585
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   375
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un Vendedor:"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmCupoVendedor"
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
    Tipo = "Cupo por Vendedor"
    Tipo2 = "Cupo por Vendedor"
    Me.Caption = Tipo
End Sub

Private Sub CargaVendedores()
    strSQL = " SELECT CONCAT(ven_apellido,' ',ven_nombre) as nombre,ven_codigo as codigo " & _
             " FROM vendedor " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 1 "
    clsCon_Def.Ejecutar strSQL
    Set cmbVendedor.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbVendedor.ListField = "nombre"
    cmbVendedor.BoundColumn = "codigo"
End Sub

Private Sub Carga()
    strSQL = " SELECT ven_codigo,linea.lin_codigo,lin_nombre,ISNULL(cup_ven_cupo,0),cup_ven_fechamod,cup_ven_usumod, '3' as modi " & _
             " FROM linea " & _
             " LEFT JOIN cupo_vendedor " & _
             " ON linea.emp_codigo=cupo_vendedor.emp_codigo" & _
             " AND linea.lin_codigo=cupo_vendedor.lin_codigo " & _
             " AND ven_codigo='" & cmbVendedor.BoundText & "' " & _
             " WHERE linea.emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY lin_nombre "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
    Next i
    
    If VSFG.Rows <= 2 Then
        If VSFG.TextMatrix(1, 1) = "" Then
            VSFG.Editable = flexEDKbd
        Else
            VSFG.Editable = flexEDKbdMouse
        End If
    Else
        VSFG.Editable = flexEDKbdMouse
    End If
End Sub


Private Sub cmbLimpiar_Click()
    cmbVendedor.BoundText = ""
    VSFG.Clear 1
    VSFG.Rows = 2
    If VSFG.Rows <= 2 Then
        If VSFG.TextMatrix(1, 1) = "" Then
            VSFG.Editable = flexEDKbd
        Else
            VSFG.Editable = flexEDKbdMouse
        End If
    Else
        VSFG.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub cmbVendedor_Change()
    Carga
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFG.Editable = flexEDKbdMouse Then
        If Col = 1 Or Col = 2 Or Col = 3 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    Else
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
    CargaVendedores
 
    If VSFG.Rows <= 2 Then
        If VSFG.TextMatrix(1, 1) = "" Then
            VSFG.Editable = flexEDKbd
        Else
            VSFG.Editable = flexEDKbdMouse
        End If
    Else
        VSFG.Editable = flexEDKbdMouse
    End If
 
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
    If Col = 4 Then
        If Trim(VSFG.TextMatrix(Row, 1)) <> "" Then
            strSQL = " UPDATE cupo_vendedor " & _
                     " SET cup_ven_cupo='" & FormatoD4(VSFG.TextMatrix(Row, 4)) & "', " & _
                     " cup_ven_fechamod=CURRENT_TIMESTAMP, " & _
                     " cup_ven_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ven_codigo='" & cmbVendedor.BoundText & "' " & _
                     " AND lin_codigo='" & VSFG.TextMatrix(Row, 2) & "' "
            clsCon_Def.Ejecutar strSQL
        Else
            strSQL = " INSERT INTO cupo_vendedor(emp_codigo,ven_codigo,lin_codigo,cup_ven_cupo,cup_ven_fechamod,cup_ven_usumod) VALUES('" & _
                     strEmpresa & "','" & cmbVendedor.BoundText & "','" & VSFG.TextMatrix(Row, 2) & "','" & FormatoD4(VSFG.TextMatrix(Row, 4)) & "', " & _
                     "CURRENT_TIMESTAMP,'" & strUsuario & "') "
            clsCon_Def.Ejecutar strSQL
        End If
        Carga
    End If
End Sub


