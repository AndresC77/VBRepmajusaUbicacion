VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerHorEmp 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Horarios por Empleado"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmVerHorEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   9735
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
      TabIndex        =   6
      Top             =   120
      Width           =   9495
      Begin VB.CheckBox chkFiltroHorario 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Horario"
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
         Left            =   6240
         TabIndex        =   14
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
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar CI/RUC"
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
         Left            =   3720
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dcmbHorarios 
         Height          =   315
         Left            =   6240
         TabIndex        =   16
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horario"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   495
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CI/RUC"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   495
         Width           =   2295
      End
   End
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Horario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   9495
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   1920
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   9060
         _cx             =   15981
         _cy             =   3387
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
         FormatString    =   $"frmVerHorEmp.frx":030A
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
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   9495
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3240
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   9060
         _cx             =   15981
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
         BackColorSel    =   -2147483638
         ForeColorSel    =   128
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerHorEmp.frx":03FB
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
         TabIndex        =   17
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
      TabIndex        =   2
      Top             =   8880
      Width           =   9495
      Begin VB.CommandButton btnCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1700
      End
   End
End
Attribute VB_Name = "frmVerHorEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private clsSql As New clsConsulta
Private strSql As String
Private i As Long

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.value = 1 Then
        txtCodigo.Enabled = True
    Else
        txtCodigo.Enabled = False
    End If
End Sub

Private Sub chkFiltroHorario_Click()
    If chkFiltroHorario.value = 1 Then
        dcmbHorarios.Enabled = True
    Else
        dcmbHorarios.Enabled = False
    End If
End Sub

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    CargarEmpleados
End Sub


Private Sub CargarEsquema(Emp As String)
    VSFG1.Clear 1
    VSFG1.Rows = 1
    i = 1
    If Emp <> "" Then
        Dim clsAux As New clsConsulta
        clsAux.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT DISTINCT TIME_FORMAT(det_hor_entrada,'%H:%i:%s') as entrada,TIME_FORMAT(det_hor_salida,'%H:%i:%s') as salida " & _
                 " FROM det_horario " & _
                 " INNER JOIN horario_empleado " & _
                 " ON horario_empleado.emp_codigo=det_horario.emp_codigo " & _
                 " AND horario_empleado.hor_codigo=det_horario.hor_codigo " & _
                 " WHERE det_horario.emp_codigo='" & strEmpresa & "' " & _
                 " AND horario_empleado.epl_codigo='" & Emp & "' " & _
                 " ORDER BY 1,2 "
        clsSql.Ejecutar strSql
        
        While Not clsSql.adorec_Def.EOF
            VSFG1.AddItem ""
            VSFG1.TextMatrix(i, 0) = Format(clsSql.adorec_Def("entrada"), "HH:mm") & "-" & Format(clsSql.adorec_Def("salida"), "HH:mm")
            strSql = " SELECT det_hor_dia,hor_descripcion " & _
                     " FROM horario " & _
                     " INNER JOIN det_horario " & _
                     " ON horario.emp_codigo=det_horario.emp_codigo " & _
                     " AND horario.hor_codigo=det_horario.hor_codigo " & _
                     " INNER JOIN horario_empleado " & _
                     " ON horario_empleado.emp_codigo=det_horario.emp_codigo " & _
                     " AND horario_empleado.hor_codigo=det_horario.hor_codigo " & _
                     " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                     " AND horario_empleado.epl_codigo='" & Emp & "' " & _
                     " AND det_hor_entrada= '" & clsSql.adorec_Def("entrada") & "' " & _
                     " AND det_hor_salida= '" & clsSql.adorec_Def("salida") & "' " & _
                     " ORDER BY 1 "
            clsAux.Ejecutar strSql
            
            While Not clsAux.adorec_Def.EOF
                VSFG1.Cell(flexcpBackColor, i, FormatoD0(clsAux.adorec_Def("det_hor_dia")) + 1) = &H80FFFF
                VSFG1.TextMatrix(i, FormatoD0(clsAux.adorec_Def("det_hor_dia")) + 1) = clsAux.adorec_Def("hor_descripcion")
                clsAux.adorec_Def.MoveNext
            Wend
            
            i = i + 1
            clsSql.adorec_Def.MoveNext
        Wend
        
        Set clsAux = Nothing
        VSFG1.AutoSizeMode = flexAutoSizeRowHeight
        VSFG1.WordWrap = True
        VSFG1.AutoSize 0, VSFG1.Cols - 1
        'VSFG1.AutoSizeMode = flexAutoSizeColWidth
        'VSFG1.AutoSize 0
   
       ' VSFG1.AutoSize VSFG1.Cols - 1
    
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    CargarHorarios
    CargarEmpleados

End Sub

Private Sub CargarHorarios()
    strSql = " SELECT hor_codigo as codigo,hor_descripcion as nombre " & _
             " FROM horario " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSql
    Set dcmbHorarios.RowSource = clsSql.adorec_Def.DataSource
    dcmbHorarios.ListField = "nombre"
    dcmbHorarios.BoundColumn = "codigo"
End Sub

Private Sub CargarEmpleados()
    VSFG.Rows = 1
    VSFG1.Clear 1
    VSFG1.Rows = 1
    strSql = " SELECT DISTINCT empleado.epl_codigo,epl_nombres,epl_apellidos,epl_cedula " & _
             " FROM empleado " & _
             " LEFT JOIN horario_empleado " & _
             " ON horario_empleado.emp_codigo=empleado.emp_codigo " & _
             " AND horario_empleado.epl_codigo=empleado.epl_codigo " & _
             " WHERE empleado.emp_codigo='" & strEmpresa & "' "
             
    If chkFiltroNombre.value = 1 Then
        strSql = strSql & " AND (epl_nombres LIKE '%" & txtNombre.Text & "%' OR epl_apellidos LIKE '%" & txtNombre.Text & "%') "
    End If
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & " AND epl_cedula LIKE '%" & txtCodigo.Text & "%' "
    End If
    If chkFiltroHorario.value = 1 Then
        strSql = strSql & " AND hor_codigo LIKE '%" & dcmbHorarios.BoundText & "%' "
    End If
    strSql = strSql & " ORDER BY 3,2 "
    clsSql.Ejecutar strSql
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    ucrtVSFG.PonerNum
    fraDatos.Caption = "Horario"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub VSFG_DblClick()
    If VSFG.Rows > 1 Then
        CargarEsquema VSFG.TextMatrix(VSFG.Row, 1)
        fraDatos.Caption = "Horario de " & VSFG.TextMatrix(VSFG.Row, 2) & " " & VSFG.TextMatrix(VSFG.Row, 3)
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

