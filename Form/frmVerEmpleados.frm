VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmVerEmpleados 
   BackColor       =   &H00DDCCBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Empleados"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   Icon            =   "frmVerEmpleados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   13110
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDCCBB&
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
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   12855
      Begin VB.CommandButton cmdNueva 
         Caption         =   "&Nuevo Empleado"
         Height          =   375
         Left            =   7680
         TabIndex        =   21
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   10800
         TabIndex        =   20
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   5640
         Width           =   6375
         Begin VB.CommandButton cmdImprimirF 
            Caption         =   "&Imprimir Formulario"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdRestaurar 
            Caption         =   "&Restaurar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdLiquidar 
            Caption         =   "&Liquidar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbTiempo 
         Height          =   315
         ItemData        =   "frmVerEmpleados.frx":030A
         Left            =   7800
         List            =   "frmVerEmpleados.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Left            =   10080
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFG 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   12615
         _cx             =   22251
         _cy             =   6165
         _ConvInfo       =   1
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerEmpleados.frx":0347
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
      End
      Begin MSDataListLib.DataCombo dcmbArea 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VSFlex7Ctl.VSFlexGrid VSFG2 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Visible         =   0   'False
         Width           =   6135
         _cx             =   10821
         _cy             =   2355
         _ConvInfo       =   1
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerEmpleados.frx":0574
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
      End
      Begin MSDataListLib.DataCombo dcmbCiudad 
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         Top             =   600
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcmbEmpleado 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleado:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   3720
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tiempo en la empresa"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7800
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   360
         Width           =   3660
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   11040
         Picture         =   "frmVerEmpleados.frx":05D5
         ToolTipText     =   "Elimina una Fila"
         Top             =   360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   11280
         Picture         =   "frmVerEmpleados.frx":070B
         Top             =   360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblLineas 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Líneas"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   5640
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Área Laboral"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3660
      End
      Begin VB.Label lblNumRows 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleados"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   12615
      End
   End
End
Attribute VB_Name = "frmVerEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsArea As New clsConsulta
Private clsLinea As New clsConsulta
Private clsCiudad As New clsConsulta
Private clsCargo As New clsConsulta
Private strSql As String
Private CodigoEmpleado As Long
Private NombreEmpleado As String
Private Hacer As Boolean
Private FilaActual As Long

Private Sub cmbTiempo_Click()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub cmdBuscar_Click()
    BuscarEmpleados
End Sub

Private Sub cmdEliminar_Click()
    Dim Mensaje As String
    
    'Para ver si se puede borrar
    strSql = " SELECT count(des_codigo) AS Num " & _
             " FROM descuento " & _
             " WHERE epl_codigo ='" & CodigoEmpleado & "' AND emp_codigo = '" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    
    If clsSql.adorec_Def("Num") > 0 Then
        If clsSql.adorec_Def(0) = 1 Then
            Mensaje = "Hay 1 registro del módulo de recursos humanos relacionado"
        Else
            Mensaje = "Hay " & clsSql.adorec_Def(0) & " registros del módulo de recursos humanos relacionados"
        End If
        MsgBox "No puede eliminar " & NombreEmpleado & " de la tabla EMPLEADO." & _
                vbNewLine & Mensaje & ".", vbCritical, "Eliminación"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de eliminar a " & NombreEmpleado & " del registro de empleados?" & vbNewLine & "Código de empleado: " & CodigoEmpleado, vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar empleado") = vbNo Then Exit Sub
'    strSql = " DELETE FROM empleado_linea WHERE" & _
'             " epl_codigo='" & CodigoEmpleado & "' AND emp_codigo='" & strEmpresa & "'"
'    clsSql.Ejecutar (strSql)
    strSql = " DELETE FROM empleado WHERE" & _
             " epl_codigo='" & CodigoEmpleado & "' AND emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    VSFG.RemoveItem (VSFG.Row)
    CambiarEtiqueta False
    Numerar
    VSFG_AfterSelChange VSFG.Row, 1, VSFG.Row, 1
End Sub

Private Sub cmdExportar_Click()
    SeleccionarFlexGrid VSFG
    CopiarFlexGrid VSFG
    If VSFG.Rows > 1 And VSFG.Col > 1 Then
        VSFG.Select 1, 1, 1, 1
    End If
    MsgBox "Se ha copiado la lista de Empleados al portapapeles.", vbInformation, "Información"
End Sub

Private Sub cmdImprimirF_Click()
    drptFormulario107.Fecha = VSFG.TextMatrix(VSFG.Row, 16)
    drptFormulario107.CodigoEmpleado = CodigoEmpleado
    drptFormulario107.Show
End Sub

Private Sub cmdLiquidar_Click()
    If MsgBox("¿Está seguro de liquidar al empleado " & NombreEmpleado & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Liquidar empleado") = vbNo Then Exit Sub
    Dim FechaSalida As String
    
    Set frmFecha.Objeto = VSFG
    frmFecha.Caption = "Fecha Salida"
    frmFecha.Fecha = Date
    frmFecha.Show vbModal
    FechaSalida = VSFG.Tag
    
'    'Llamar a pantalla de liquidaciones
'    frmLiquidacion.txtEmpleado = NombreEmpleado
'    frmLiquidacion.txtEmpleado.Tag = CodigoEmpleado
'    frmLiquidacion.Show
    
    strSql = "UPDATE empleado SET epl_fec_salida='" & FechaSalida & "' WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & CodigoEmpleado & "'"
    clsSql.Ejecutar (strSql)
    VSFG.TextMatrix(VSFG.Row, 16) = FechaSalida
    cmdRestaurar.Enabled = True
    cmdImprimirF.Enabled = True
    cmdLiquidar.Enabled = False
End Sub

Private Sub cmdNueva_Click()
    Dim Maximo As String
Repetir:
    strSql = " SELECT ISNULL(MAX(epl_codigo),0)+1 FROM empleado WHERE emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    Maximo = clsSql.adorec_Def(0)
    
    'Hacer insert
    strSql = " INSERT INTO empleado (epl_codigo, emp_codigo, epl_nombres, epl_apellidos, epl_sueldo, epl_fec_ingreso, epl_fechamod, epl_usumod) VALUES" & _
             " ('" & Maximo & "', '" & strEmpresa & "', '', '', 0,'" & Date & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    
    If clsSql.EjecutarSeguro(strSql) = False Then GoTo Repetir
    VSFG.AddItem VSFG.Rows & vbTab & Maximo
    VSFG.TextMatrix(VSFG.Rows - 1, 13) = Date
    CambiarEtiqueta True
End Sub

Private Sub cmdRestaurar_Click()
    If MsgBox("¿Está seguro de restaurar al empleado " & NombreEmpleado & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Restaurar empleado") = vbNo Then Exit Sub
    strSql = "UPDATE empleado SET epl_fec_salida=NULL WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & CodigoEmpleado & "'"
    clsSql.Ejecutar (strSql)
    VSFG.TextMatrix(VSFG.Row, 16) = ""
    cmdRestaurar.Enabled = False
    cmdImprimirF.Enabled = False
    cmdLiquidar.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub dcmbArea_Change()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub dcmbCiudad_Change()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub dcmbEmpleado_Change()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - mdiPrincipal.Height / 40
    
    'Inicializa la clase con la conexión activa a la base de datos
    clsSql.Inicializar AdoConn
    clsArea.Inicializar AdoConn
    clsLinea.Inicializar AdoConn
    clsCiudad.Inicializar AdoConn
    clsCargo.Inicializar AdoConn
    
    VSFG.FrozenCols = 3
    
    strSql = " SELECT are_lab_codigo AS codigo, are_lab_nombre AS nombre FROM area_laboral " & _
             " WHERE emp_codigo='" & strEmpresa & "' ORDER BY nombre"
    clsArea.Ejecutar (strSql)
    
    strSql = " SELECT ciu_codigo AS codigo, ciu_nombre AS nombre FROM ciudad " & _
             " ORDER BY nombre"
    clsCiudad.Ejecutar (strSql)
    
    strSql = " SELECT car_codigo AS codigo, car_nombre AS nombre FROM cargo " & _
             " ORDER BY car_nombre"
    clsCargo.Ejecutar (strSql)
    
    strSql = " SELECT '' AS codigo, ' ' AS nombre UNION" & _
             " SELECT lin_codigo AS codigo, lin_nombre AS nombre FROM linea " & _
             " WHERE emp_codigo='" & strEmpresa & "' ORDER BY nombre"
    clsLinea.Ejecutar (strSql)
    
    strSql = " SELECT '%' AS codigo, ' --Todas Las Áreas Laborales--' AS nombre UNION" & _
             " SELECT are_lab_codigo AS codigo, are_lab_nombre AS nombre FROM area_laboral " & _
             " WHERE emp_codigo='" & strEmpresa & "' ORDER BY codigo"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        Set dcmbArea.RowSource = clsSql.adorec_Def.DataSource
        dcmbArea.ListField = "nombre"
        dcmbArea.BoundColumn = "codigo"
        Hacer = False
        dcmbArea.BoundText = clsSql.adorec_Def(0)
    End If
    
    strSql = " SELECT '%' AS codigo, ' --Todas Las Ciudades--' AS nombre UNION" & _
             " SELECT ciu_codigo AS codigo, ciu_nombre AS nombre FROM ciudad " & _
             " ORDER BY nombre"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        Set dcmbCiudad.RowSource = clsSql.adorec_Def.DataSource
        dcmbCiudad.ListField = "nombre"
        dcmbCiudad.BoundColumn = "codigo"
        Hacer = True
        dcmbCiudad.BoundText = clsSql.adorec_Def(0)
    End If
    cmbTiempo.ListIndex = 0
    
    strSql = " SELECT '%' AS epl_codigo, ' --Todos Los Empleados--' AS nombre UNION" & _
             " SELECT CAST(epl_codigo AS VARCHAR) AS epl_codigo, epl_apellidos+' '+epl_nombres AS nombre " & _
             " FROM empleado WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY nombre"
    clsSql.Ejecutar strSql
    Set Me.dcmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    dcmbEmpleado.ListField = "nombre"
    dcmbEmpleado.BoundColumn = "epl_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbEmpleado.BoundText = clsSql.adorec_Def("epl_codigo")
    End If
End Sub

Private Sub BuscarEmpleados()
    If Hacer = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim strWhere As String
    Dim strWhere2 As String
    If dcmbArea.BoundText <> "%" Then
        strWhere = " AND are_lab_codigo='" & dcmbArea.BoundText & "'"
    End If
    If dcmbCiudad.BoundText <> "%" Then
        strWhere2 = " AND empleado.ciu_codigo='" & dcmbCiudad.BoundText & "'"
    End If
    Select Case Me.cmbTiempo.ListIndex
    Case 1
        strWhere2 = strWhere2 & " AND empleado.epl_fec_ingreso<='" & DateAdd("yyyy", -1, Date) & "'"
    Case 2
        strWhere2 = strWhere2 & " AND empleado.epl_fec_ingreso>'" & DateAdd("yyyy", -1, Date) & "'"
    End Select
    
    
    'LEFT JOIN ciudad ON empleado.ciu_codigo=ciudad.ciu_codigo
    
    strSql = " SELECT epl_codigo, epl_apellidos, epl_nombres, ISNULL(epl_cedula,''), ISNULL(epl_sexo,''), ISNULL(are_lab_codigo,''), car_codigo, ISNULL(epl_sueldo,0), ISNULL(epl_direccion,''), ISNULL(CAST(epl_direccion_num AS VARCHAR),''), ISNULL(epl_telefono,''), ISNULL(ciu_codigo,''), epl_fec_ingreso, ' ', epl_baja, epl_fec_salida, empleado.asi_numasiento" & _
             " FROM empleado" & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND empleado.epl_codigo LIKE '" & Me.dcmbEmpleado.BoundText & "' " & strWhere & strWhere2 & _
             " ORDER BY epl_baja, epl_apellidos, epl_nombres"
    clsSql.Ejecutar (strSql)
    NumRows = clsSql.adorec_Def.RecordCount
    If NumRows > 0 Then
        If NumRows = 1 Then
            lblNumRows.Caption = "1 empleado"
        Else
            lblNumRows.Caption = NumRows & " empleados"
        End If
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        VSFG.ColComboList(5) = "M|F"
        VSFG.ColComboList(6) = VSFG.BuildComboList(clsArea.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(7) = VSFG.BuildComboList(clsCargo.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(12) = VSFG.BuildComboList(clsCiudad.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(13) = "..."
        VSFG.Select 0, 0
        VSFG.Select 1, 1
    Else
        lblNumRows.Caption = "0 empleados"
        VSFG.Clear 1
        VSFG.Rows = 1
        VSFG2.Clear 1
        VSFG2.Rows = 1
        VSFG.ColComboList(5) = "M|F"
        VSFG.ColComboList(6) = VSFG.BuildComboList(clsArea.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(7) = VSFG.BuildComboList(clsCargo.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(12) = VSFG.BuildComboList(clsCiudad.adorec_Def, "codigo, *nombre", "codigo")
        VSFG.ColComboList(13) = "..."
        lblLineas.Caption = "Líneas"
        cmdLiquidar.Enabled = False
        cmdEliminar.Enabled = False
        cmdRestaurar.Enabled = False
        cmdImprimirF.Enabled = False
    End If
    Numerar
    Me.cmdBuscar.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub CambiarEtiqueta(SubirUno As Boolean)
    Dim Valorcito As Long
    Valorcito = FormatoInt(lblNumRows.Caption)
    If SubirUno = True Then
        Valorcito = Valorcito + 1
    Else
        Valorcito = Valorcito - 1
    End If
    If Valorcito = 1 Then
        lblNumRows.Caption = "1 empleado"
    Else
        lblNumRows.Caption = Valorcito & " empleados"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Numerar()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 6 Then
        strSql = " UPDATE empleado SET are_lab_codigo='" & VSFG.TextMatrix(Row, Col) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        clsSql.Ejecutar (strSql)
    End If
    If Col = 7 Then
        strSql = " UPDATE empleado SET car_codigo='" & VSFG.TextMatrix(Row, Col) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        clsSql.Ejecutar (strSql)
    End If
    If Col = 12 Then
        strSql = " UPDATE empleado SET ciu_codigo='" & VSFG.TextMatrix(Row, Col) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        clsSql.Ejecutar (strSql)
    End If
End Sub

Private Sub VSFG_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If NewRowSel <> 0 Then
        Screen.MousePointer = vbHourglass
        If VSFG.Rows - 1 > 0 Then
            VSFG.Cell(flexcpBackColor, 1, 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(255, 255, 255)
        End If
        VSFG.Cell(flexcpBackColor, NewRowSel, 1, NewRowSel, VSFG.Cols - 1) = &HC0FFFF
        CodigoEmpleado = VSFG.TextMatrix(NewRowSel, 1)
        NombreEmpleado = VSFG.TextMatrix(NewRowSel, 2) & " " & VSFG.TextMatrix(NewRowSel, 3)
        Frame1.Caption = StrConv(NombreEmpleado, vbProperCase)
        lblLineas.Caption = "Líneas de " & NombreEmpleado
        cmdEliminar.Enabled = True
        If Trim(VSFG.TextMatrix(NewRowSel, 16)) = "" Then
            cmdLiquidar.Enabled = True
            cmdRestaurar.Enabled = False
            cmdImprimirF.Enabled = False
        Else
            cmdLiquidar.Enabled = False
            cmdImprimirF.Enabled = True
            'Si ya está de baja ya no dejar restaurar
            If Val(VSFG.TextMatrix(NewRowSel, 15)) = 0 Then
                cmdRestaurar.Enabled = True
            Else
                cmdRestaurar.Enabled = False
                cmdEliminar.Enabled = False
            End If
        End If
'        'Buscar líneas relacionadas con el empleado
'        strSql = " SELECT lin_codigo, lin_codigo FROM empleado_linea " & _
'             " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & CodigoEmpleado & "' ORDER BY lin_codigo"
'        clsSql.Ejecutar (strSql)
'        Set VSFG2.DataSource = clsSql.adorec_Def.DataSource
'        Dim i As Long
'        For i = 1 To VSFG2.Rows - 1
'            VSFG2.TextMatrix(i, 0) = i
'        Next i
'        VSFG2.AddItem VSFG2.Rows
'        VSFG2.ColComboList(1) = VSFG2.BuildComboList(clsLinea.adorec_Def, "codigo, *nombre", "codigo")
        FilaActual = NewRowSel
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub VSFG_AfterSort(ByVal Col As Long, Order As Integer)
    Numerar
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Or Col = 14 Or Col = 15 Then
        Cancel = True
    Else
        If CInt(Val(VSFG.TextMatrix(Row, 15))) <> 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_BeforeScrollTip(ByVal Row As Long)
    VSFG.ScrollTipText = "Empleado: " & VSFG.TextMatrix(Row, 2) & " " & VSFG.TextMatrix(Row, 3)
End Sub

Private Sub VSFG_BeforeSort(ByVal Col As Long, Order As Integer)
'    If FilaActual > 0 And VSFG.Rows > 0 Then
'        VSFG.Cell(flexcpBackColor, OldRowSel, 1, OldRowSel, VSFG.Cols - 1) = RGB(255, 255, 255)
'        'Call VSFG_AfterSelChange(FilaActual, Col, 1, Col)
'    End If
End Sub

Private Sub VSFG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 13 Then
        Set frmFecha.Objeto = VSFG
        frmFecha.Fecha = VSFG.TextMatrix(Row, Col)
        frmFecha.Show vbModal
        VSFG.TextMatrix(Row, Col) = VSFG.Tag
        strSql = " UPDATE empleado SET epl_fec_ingreso='" & VSFG.TextMatrix(Row, Col) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        clsSql.Ejecutar (strSql)
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 15 And Row > 0 Then
        If CInt(VSFG.TextMatrix(Row, 15)) <> 0 Then
            'Poner en gris si está de baja
            VSFG.Cell(flexcpForeColor, Row, 0, Row, VSFG.Cols - 1) = RGB(150, 150, 150)
        Else
            VSFG.Cell(flexcpForeColor, Row, 0, Row, VSFG.Cols - 4) = RGB(0, 0, 0)
        End If
    End If
    If Col = 13 Or Col = 14 And Row > 0 Then
        If Trim(VSFG.TextMatrix(Row, 13)) <> "" And IsDate(VSFG.TextMatrix(Row, 13)) Then
            Dim Diferencia As Integer
            Diferencia = DateDiff("yyyy", VSFG.TextMatrix(Row, 13), Date)
            If Diferencia > 5 Then
                '1 día más a partir del 5 año
                Diferencia = 15 + Diferencia - 5
                If Diferencia > 30 Then Diferencia = 30
                VSFG.TextMatrix(Row, 14) = CStr(Diferencia)
                VSFG.Cell(flexcpForeColor, Row, 14) = RGB(160, 0, 0)
            Else
                '15 días de vacaciones
                VSFG.TextMatrix(Row, 14) = "15"
            End If
        End If
    End If
End Sub

Private Sub VSFG_DblClick()
    'MsgBox VSFG.TextMatrix(VSFG.Row, VSFG.Col)
End Sub

Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    strSql = " UPDATE empleado SET "
    If Col = 2 Or Col = 3 Then
        VSFG.EditText = Left(UCase(Trim(VSFG.EditText)), 40)
        'Verificar que no haya nombres y apellidos repetidos
        Dim i As Long
        Dim NombreCompleto As String
        Dim Encontro As Boolean
        If Col = 2 Then
            NombreCompleto = VSFG.EditText & " " & VSFG.TextMatrix(Row, 3)
        ElseIf Col = 3 Then
            NombreCompleto = VSFG.TextMatrix(Row, 2) & " " & VSFG.EditText
        End If
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 2) & " " & VSFG.TextMatrix(i, 3) = NombreCompleto And i <> Row Then
                Encontro = True
                Exit For
            End If
        Next i
        If Encontro = True Then
            MsgBox "El nombre " & NombreCompleto & " ya le pertenece a otro empleado.", vbInformation, "Información"
            Cancel = True
        End If
        VSFG.EditText = Left(VSFG.EditText, 40)
        If Col = 2 Then
            strSql = strSql & " epl_apellidos='" & VSFG.EditText & "'"
        ElseIf Col = 3 Then
            strSql = strSql & " epl_nombres='" & VSFG.EditText & "'"
        End If
    End If
    'Cédula
    If Col = 4 Then
        VSFG.EditText = Left(VSFG.EditText, 10)
        strSql = strSql & " epl_cedula='" & VSFG.EditText & "'"
    End If
    'Sexo
    If Col = 5 Then
        strSql = strSql & " epl_sexo='" & VSFG.EditText & "'"
    End If
    'Sueldo
    If Col = 8 Then
        VSFG.EditText = FormatoD(VSFG.EditText)
        strSql = strSql & " epl_sueldo='" & VSFG.EditText & "'"
    End If
    'Dirección
    If Col = 9 Then
        VSFG.EditText = UCase(Left(VSFG.EditText, 100))
        strSql = strSql & " epl_direccion='" & VSFG.EditText & "'"
    End If
    'Número de dirección
    If Col = 10 Then
        VSFG.EditText = FormatoInt(VSFG.EditText)
        strSql = strSql & " epl_direccion_num='" & VSFG.EditText & "'"
    End If
    'Teléfono
    If Col = 11 Then
        VSFG.EditText = Left(VSFG.EditText, 50)
        strSql = strSql & " epl_telefono='" & VSFG.EditText & "'"
    End If
    If Col = 6 Or Col = 7 Or Col = 12 Or Col = 13 Then
        'Están en el evento AfterEdit
        Exit Sub
    End If
    strSql = strSql & " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
    clsSql.Ejecutar (strSql)
End Sub

Private Sub VSFG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        If Trim(VSFG2.TextMatrix(Row, Col)) <> "" Then
            strSql = " INSERT INTO empleado_linea(epl_codigo,emp_codigo,lin_codigo) VALUES" & _
                     " ('" & CodigoEmpleado & "','" & strEmpresa & "','" & VSFG2.TextMatrix(Row, Col) & "')"
            clsSql.Ejecutar (strSql)
            If VSFG2.TextMatrix(Row, 2) = "" Then
                VSFG2.AddItem VSFG2.Rows
            End If
            VSFG2.TextMatrix(Row, 2) = VSFG2.TextMatrix(Row, 1)
        Else
            'Borrar item si es que no es nuevo
            If VSFG2.TextMatrix(Row, 2) <> "" Then
                VSFG2.RemoveItem Row
                Dim i As Integer
                For i = 1 To VSFG2.Rows - 1
                    VSFG2.TextMatrix(i, 0) = i
                Next i
            End If
        End If
    End If
End Sub

Private Sub VSFG2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Then
        'Verificar si no es línea repetida
        Dim i As Long
        For i = 1 To VSFG2.Rows - 1
            If VSFG2.EditText = VSFG2.Cell(flexcpTextDisplay, i, Col) And Row <> i And Trim(VSFG2.EditText) <> "" Then
                MsgBox "La línea " & VSFG2.EditText & " ya fue ingresada anteriormente.", vbInformation, "Información"
                Cancel = True
                Exit Sub
            End If
        Next i
        'Borrar cuando se edite
        If VSFG2.TextMatrix(Row, 2) <> "" Then
            strSql = " DELETE FROM empleado_linea WHERE" & _
                     " epl_codigo='" & CodigoEmpleado & "' AND emp_codigo='" & strEmpresa & "' AND lin_codigo='" & VSFG2.TextMatrix(Row, 2) & "'"
            clsSql.Ejecutar (strSql)
        End If
    End If
End Sub
