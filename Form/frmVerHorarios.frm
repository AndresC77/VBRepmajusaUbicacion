VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerHorarios 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmVerHorarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8460
   Begin VB.CommandButton btnEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   2420
      TabIndex        =   3
      Top             =   6600
      Width           =   1700
   End
   Begin VB.CommandButton btnModificar 
      Caption         =   "&Modificar"
      Height          =   360
      Left            =   4340
      TabIndex        =   2
      Top             =   6120
      Width           =   1700
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   2420
      TabIndex        =   1
      Top             =   6120
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4340
      TabIndex        =   4
      Top             =   6600
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3240
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8220
      _cx             =   14499
      _cy             =   5715
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerHorarios.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG1 
      Height          =   2040
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   8220
      _cx             =   14499
      _cy             =   3598
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerHorarios.frx":0402
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
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmVerHorarios"
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
    Tipo = "Horario"
    Tipo2 = "Horarios"
    Me.Caption = Tipo
End Sub

Public Sub Carga()
    strSql = " SELECT horario.hor_codigo,'' as ver,hor_descripcion,SUM(TIME_FORMAT(det_hor_salida,'%H')-TIME_FORMAT(det_hor_entrada,'%H'))+(SUM(TIME_FORMAT(det_hor_salida,'%i')-TIME_FORMAT(det_hor_entrada,'%i')))/60, " & _
             " hor_fechamod,hor_usumod, '0' as modi " & _
             " FROM horario " & _
             " INNER JOIN det_horario " & _
             " ON horario.emp_codigo=det_horario.emp_codigo " & _
             " AND horario.hor_codigo=det_horario.hor_codigo " & _
             " WHERE horario.emp_codigo ='" & strEmpresa & "' " & _
             " GROUP BY horario.hor_codigo " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum
    VSFG1.Clear 1
    VSFG1.Rows = 1
    
End Sub


Private Sub CargarEsquema()
    Dim i As Long, horas As String
    VSFG1.Clear 1
    VSFG1.Rows = 1
    i = 1
    horas = ""
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    For i = 1 To VSFG.Rows - 1
        If CBool(FormatoD0(VSFG.TextMatrix(i, 2))) = True Then
            horas = horas & VSFG.TextMatrix(i, 1) & "','"
        End If
    Next i
    If horas <> "" Then
        horas = "('" & Left(horas, Len(horas) - 2) & ")"
    Else
        Exit Sub
    End If
         
            strSql = " SELECT DISTINCT hor_codigo,RIGHT(det_hor_entrada,8) as entrada,RIGHT(det_hor_salida,8) as salida " & _
                     " FROM det_horario " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND hor_codigo IN " & horas & " " & _
                     " ORDER BY 2,3"
            clsCon_Def.Ejecutar strSql
            i = 1
            
            While Not clsCon_Def.adorec_Def.EOF
                VSFG1.AddItem ""
                VSFG1.TextMatrix(i, 0) = Left(clsCon_Def.adorec_Def("entrada"), 5) & "-" & Left(clsCon_Def.adorec_Def("salida"), 5)
                strSql = " SELECT det_hor_dia " & _
                         " FROM det_horario " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND hor_codigo ='" & clsCon_Def.adorec_Def("hor_codigo") & "' " & _
                         " AND det_hor_entrada= '" & clsCon_Def.adorec_Def("entrada") & "' " & _
                         " AND det_hor_salida= '" & clsCon_Def.adorec_Def("salida") & "' " & _
                         " ORDER BY 1 "
                clsAux.Ejecutar strSql
                
                While Not clsAux.adorec_Def.EOF
                    VSFG1.Cell(flexcpBackColor, i, FormatoD0(clsAux.adorec_Def("det_hor_dia")) + 1) = &H80FFFF
                    VSFG1.TextMatrix(i, FormatoD0(clsAux.adorec_Def("det_hor_dia")) + 1) = VSFG.TextMatrix(VSFG.FindRow(clsCon_Def.adorec_Def("hor_codigo"), , 1), 3)
                    clsAux.adorec_Def.MoveNext
                Wend
                
                i = i + 1
                clsCon_Def.adorec_Def.MoveNext
            Wend
    VSFG1.AutoSizeMode = flexAutoSizeRowHeight
    VSFG1.WordWrap = True
    VSFG1.AutoSize 0, VSFG1.Cols - 1
        
    Set clsAux = Nothing
End Sub

Private Sub btnEliminar_Click()
    If VSFG.Rows > 1 Then
        Dim horCodigo As String, mas As String
        Dim respuesta As Integer
        horCodigo = VSFG.TextMatrix(VSFG.Row, 1)
        'Controlar que no esten asigandos empleados
        strSql = " SELECT COUNT(epl_codigo) " & _
                 " FROM horario_empleado " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND hor_codigo='" & horCodigo & "' "
        clsCon_Def.Ejecutar strSql
        respuesta = vbYes
        mas = ""
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                mas = "El horario posee empleados asginados a éste." & vbCrLf
            End If
        End If
        respuesta = MsgBox(mas & "Está seguro que desea eliminar el horario: " & VSFG.TextMatrix(VSFG.Row, 3) & "?", vbQuestion + vbYesNo, "Horarios")
        If respuesta = vbYes Then
            strSql = " DELETE FROM horario " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND hor_codigo='" & horCodigo & "' "
            clsCon_Def.Ejecutar strSql, "M"
            
            strSql = " DELETE FROM det_horario " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND hor_codigo='" & horCodigo & "' "
            clsCon_Def.Ejecutar strSql, "M"
            
            strSql = " DELETE FROM horario_empleado " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND hor_codigo='" & horCodigo & "' "
            clsCon_Def.Ejecutar strSql, "M"
            Carga
        End If
    End If
End Sub

Private Sub btnModificar_Click()
    If VSFG.Rows > 1 Then
        frmHorarios.Actualizar = True
        frmHorarios.horariocodigo = VSFG.TextMatrix(VSFG.Row, 1)
        frmHorarios.Show 1
    End If
End Sub

Private Sub btnNuevo_Click()
    frmHorarios.Actualizar = False
    frmHorarios.horariocodigo = ""
    frmHorarios.Show 1
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
    ucrtVSFG.Inicializar False, False, False
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
    CargarEsquema
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then
        Cancel = True
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


