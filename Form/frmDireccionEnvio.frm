VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDireccionEnvio 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direccion de Envio"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   Icon            =   "frmDireccionEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   11550
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6225
      TabIndex        =   1
      Top             =   3000
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   8025
      TabIndex        =   0
      Top             =   3000
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2265
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11340
      _cx             =   1980255778
      _cy             =   1980239771
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDireccionEnvio.frx":030A
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
Attribute VB_Name = "frmDireccionEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Public strCliente As String
Public strPedido As String

Private Sub IniDato()
    Tipo = "Direccion"
    Tipo2 = "Direccion"
    Me.Caption = Tipo
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
     strSql = " SELECT per_dir_default,per_dir_codigo,ciu_codigo,per_dir_direccion,per_dir_fechamod,per_dir_usumod,'0'as modi " & _
              " FROM persona_direccion " & _
              " WHERE emp_codigo='" & strEmpresa & "' " & _
              " AND per_codigo='" & strCliente & "' "
    clsCon_Def.Ejecutar strSql, "M"
'    While Not clsCon_Def.adorec_Def.EOF
'        VSFG.AddItem clsCon_Def.adorec_Def("per_dir_default") & vbTab & clsCon_Def.adorec_Def("per_dir_codigo") & vbTab & clsCon_Def.adorec_Def("ciu_codigo") & vbTab & clsCon_Def.adorec_Def("per_dir_direccion") & vbTab & clsCon_Def.adorec_Def("per_dir_fechamod") & vbTab & clsCon_Def.adorec_Def("per_dir_usumod") & vbTab & "0"
'        clsCon_Def.adorec_Def.MoveNext
'    Wend
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    strSql = " SELECT ciu_codigo,CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre) as ciucanpai" & _
             " FROM ciudad INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo" & _
             " INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo" & _
             " ORDER BY CONCAT(ciu_nombre,'-',can_nombre,'-',pai_nombre)"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*ciucanpai", "ciu_codigo")
    
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim num As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            strSql = " UPDATE persona_direccion " & _
                 " SET ciu_codigo='" & UCase(VSFG.TextMatrix(i, 2)) & "'," & _
                 " per_dir_direccion='" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                 " per_dir_fechamod=CURRENT_TIMESTAMP," & _
                 " per_dir_usumod='" & strUsuario & "' " & _
                 " WHERE per_dir_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                 " AND per_codigo='" & strCliente & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la ciudad", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta Direccion", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT COALESCE(MAX(persona_direccion.per_dir_codigo),0) as n " & _
                    " FROM persona_direccion " & _
                    " WHERE cli_codigo='" & strCliente & "'" & _
                    " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    num = FormatoD0(clsCon_Def.adorec_Def("n")) + 1
                Else
                    num = 0
                End If
                strSql = " INSERT INTO persona_direccion (emp_codigo,per_codigo,per_dir_codigo,ciu_codigo, " & _
                         " per_dir_direccion,per_dir_default,per_dir_fechamod,per_dir_usumod)" & _
                         " VALUES ('" & strEmpresa & "','" & strCliente & "','" & num & "','" & VSFG.TextMatrix(i, 2) & "'," & _
                         " '" & UCase(VSFG.TextMatrix(i, 3)) & "','0',CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
        
            strSql = " DELETE " & _
                     " FROM persona_direccion " & _
                     " WHERE per_dir_codigo= '" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND per_codigo= '" & strCliente & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql), "M"
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga
    End If
    cmdSeleccionar_Click
End Sub

Private Sub cmdSeleccionar_Click()
    Dim i As Long
    If strPedido <> "" Then
        For i = 1 To VSFG.Rows - 1
            If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
                strSql = " UPDATE pedido " & _
                         " SET ped_direccion_envio='" & UCase(VSFG.Cell(flexcpTextDisplay, i, 2) & " - " & VSFG.TextMatrix(i, 3)) & "'" & _
                         " WHERE ped_codigo='" & strPedido & "' " & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql, "M"
            End If
        Next i
    End If
    Unload Me
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
    If Col > 0 Then
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
    If Col > 0 Then
        If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
            VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
        ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
            VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
        End If
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


