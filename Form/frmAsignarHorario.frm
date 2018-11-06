VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAsignarHorario 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Horarios"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmAsignarHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8175
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Información general"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin MSDataListLib.DataCombo dcmbHorarios 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   1200
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   7500
         _cx             =   13229
         _cy             =   2117
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
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
         FormatString    =   $"frmAsignarHorario.frx":030A
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
      Begin VB.Label lblFacultad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horario:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame fraProfesor 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Empleados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   7935
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4080
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   7500
         _cx             =   13229
         _cy             =   7197
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAsignarHorario.frx":03F4
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
   End
   Begin VB.Frame fraBotones 
      BackColor       =   &H00DDDDDD&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   7935
      Begin VB.CommandButton btnAgregar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   2345
         TabIndex        =   4
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   4145
         TabIndex        =   5
         Top             =   240
         Width           =   1700
      End
   End
End
Attribute VB_Name = "frmAsignarHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private clsSql As New clsConsulta
Private strSql As String
Private i As Long

Private Sub btnAgregar_Click()
    If ComprobarDatos = True Then
        If AgregarDatos = True Then
            Dim Contar As Long: Contar = 0
            For i = 1 To VSFG.Rows - 1
                If CBool(VSFG.TextMatrix(i, 1)) = True Then Contar = Contar + 1
            Next i
            If Contar <= 1 Then
                MsgBox "Se ha asignado " & CStr(Contar) & " empleado al horario " & dcmbHorarios.Text, vbInformation, "Asignación de Horarios"
            Else
                MsgBox "Se han asignado " & CStr(Contar) & " empleados al horario " & dcmbHorarios.Text, vbInformation, "Asignación de Horarios"
            End If
            Limpiar
            dcmbHorarios.SetFocus
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub dcmbHorarios_Change()
    CargarEsquema
    CargarHorEmp
End Sub


Private Sub CargarEsquema()
    VSFG1.Clear 1
    VSFG1.Rows = 1
    i = 1
    If dcmbHorarios.Text <> "" Then
        Dim clsAux As New clsConsulta
        clsAux.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT DISTINCT RIGHT(det_hor_entrada,8) as entrada,RIGHT(det_hor_salida,8) as salida " & _
                 " FROM det_horario " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND hor_codigo='" & dcmbHorarios.BoundText & "' " & _
                 " ORDER BY 1,2"
        clsSql.Ejecutar strSql
        
        While Not clsSql.adorec_Def.EOF
            VSFG1.AddItem ""
            VSFG1.TextMatrix(i, 0) = Format(clsSql.adorec_Def("entrada"), "HH:mm") & "-" & Format(clsSql.adorec_Def("salida"), "HH:mm")
            strSql = " SELECT det_hor_dia " & _
                     " FROM det_horario " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND hor_codigo='" & dcmbHorarios.BoundText & "' " & _
                     " AND det_hor_entrada= '" & Format(clsSql.adorec_Def("entrada"), "HH:MM:SS") & "' " & _
                     " AND det_hor_salida= '" & Format(clsSql.adorec_Def("salida"), "HH:MM:SS") & "' " & _
                     " ORDER BY 1 "
            clsAux.Ejecutar strSql
            
            While Not clsAux.adorec_Def.EOF
                VSFG1.Cell(flexcpBackColor, i, FormatoD0(clsAux.adorec_Def("det_hor_dia")) + 1) = &H80FFFF
                clsAux.adorec_Def.MoveNext
            Wend
            
            i = i + 1
            clsSql.adorec_Def.MoveNext
        Wend
        
        Set clsAux = Nothing
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    
        
    CargarHorarios
    
End Sub

Private Sub CargarHorarios()
    strSql = " SELECT hor_codigo as codigo,hor_descripcion as nombre " & _
             " FROM horario " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND hor_disponible='1' "
    clsSql.Ejecutar strSql
    Set dcmbHorarios.RowSource = clsSql.adorec_Def.DataSource
    dcmbHorarios.ListField = "nombre"
    dcmbHorarios.BoundColumn = "codigo"
End Sub

Private Sub Limpiar()

    dcmbHorarios.Text = ""
    'CargarHorEmp
End Sub

Private Sub CargarHorEmp()
    VSFG.Rows = 1
    If dcmbHorarios.Text <> "" Then
        strSql = " SELECT IF(horario_empleado.epl_codigo IS NULL,0,1) as asignado,empleado.epl_codigo,CONCAT(empleado.epl_apellidos,' ',empleado.epl_nombres) as nombre,horario_empleado.hor_epl_fechamod,horario_empleado.hor_epl_usumod " & _
                 " FROM empleado " & _
                 " LEFT JOIN horario_empleado " & _
                 " ON empleado.emp_codigo=horario_empleado.emp_codigo " & _
                 " AND empleado.epl_codigo=horario_empleado.epl_codigo " & _
                 " AND horario_empleado.hor_codigo='" & dcmbHorarios.BoundText & "' " & _
                 " WHERE empleado.emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY epl_apellidos,epl_nombres "
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        ucrtVSFG.PonerNum
    End If
End Sub

Private Function ComprobarDatos() As Boolean
    ComprobarDatos = True
End Function

Private Function AgregarDatos() As Boolean
    On Error GoTo noagregar

    strSql = " DELETE FROM horario_empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND hor_codigo='" & dcmbHorarios.BoundText & "' "
    clsSql.Ejecutar strSql, "M"

    For i = 1 To VSFG.Rows - 1
        If CBool(VSFG.TextMatrix(i, 1)) = True Then
        
            strSql = " INSERT INTO horario_empleado(emp_codigo,hor_codigo,epl_codigo,hor_epl_fechamod,hor_epl_usumod) VALUES('" & _
                     strEmpresa & "','" & dcmbHorarios.BoundText & "','" & VSFG.TextMatrix(i, 2) & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
            clsSql.Ejecutar strSql, "M"
        End If
    Next i
    
    AgregarDatos = True
    Exit Function
    
noagregar:
    MsgBox "Ocurrió un problema al intentar agregar, intente nuevamente", vbCritical, "Asignación de Horarios"
    AgregarDatos = False
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim clsAux As New clsConsulta
        clsAux.Inicializar AdoConn, AdoConnMaster
        
        strSql = " SELECT det_horario.hor_codigo,hor_descripcion,det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s') as entrada,TIME_FORMAT(det_hor_salida,'%H:%i:%s') as salida " & _
                 " FROM horario " & _
                 " INNER JOIN det_horario " & _
                 " ON horario.emp_codigo=det_horario.emp_codigo " & _
                 " AND horario.hor_codigo=det_horario.hor_codigo " & _
                 " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                 " AND horario.hor_codigo='" & dcmbHorarios.BoundText & "' " & _
                 " ORDER BY 1,3,4,5 "
        clsAux.Ejecutar strSql
        
        

    
        strSql = " SELECT det_horario.hor_codigo,hor_descripcion,det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s') as entrada,TIME_FORMAT(det_hor_salida,'%H:%i:%s') as salida " & _
                 " FROM horario " & _
                 " INNER JOIN det_horario " & _
                 " ON horario.emp_codigo=det_horario.emp_codigo " & _
                 " AND horario.hor_codigo=det_horario.hor_codigo " & _
                 " INNER JOIN horario_empleado " & _
                 " ON horario_empleado.emp_codigo=det_horario.emp_codigo " & _
                 " AND horario_empleado.hor_codigo=det_horario.hor_codigo " & _
                 " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                 " AND horario.hor_codigo!='" & dcmbHorarios.BoundText & "' " & _
                 " AND horario_empleado.epl_codigo='" & VSFG.TextMatrix(Row, 2) & "' " & _
                 " ORDER BY 1,3,4,5 "
        clsSql.Ejecutar strSql
                 
        If clsSql.adorec_Def.RecordCount > 0 Then
            While Not clsAux.adorec_Def.EOF
                While Not clsSql.adorec_Def.EOF
                    If clsAux.adorec_Def("det_hor_dia") = clsSql.adorec_Def("det_hor_dia") Then
                                           
                        If (clsAux.adorec_Def("entrada") < clsSql.adorec_Def("entrada") And clsAux.adorec_Def("salida") <= clsSql.adorec_Def("entrada")) _
                        Or (clsAux.adorec_Def("entrada") >= clsSql.adorec_Def("salida") And clsAux.adorec_Def("salida") > clsSql.adorec_Def("salida")) Then
                        Else
                            MsgBox "No puede asignarle este horario al empleado " & VSFG.TextMatrix(Row, 3) & "," & vbCrLf & "debido a que se cruza con el horario: """ & clsSql.adorec_Def("hor_descripcion") & """"
                            VSFG.TextMatrix(Row, 1) = 0
                            VSFG.Cell(flexcpBackColor, Row, 1, Row, VSFG.Cols - 1) = &H80000005
                            Exit Sub
                        End If
                    End If
                    clsSql.adorec_Def.MoveNext
                Wend
                clsAux.adorec_Def.MoveNext
            Wend
        
        End If
        VSFG.Cell(flexcpBackColor, Row, 1, Row, VSFG.Cols - 1) = &HC0FFC0
    
    Set clsAux = Nothing
        
        
        
    
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
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

