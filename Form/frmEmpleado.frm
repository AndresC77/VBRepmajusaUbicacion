VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEmpleado 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleado"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   Icon            =   "frmEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12840
   Begin VB.CommandButton cmdRestaurar 
      Caption         =   "&Restaurar"
      Height          =   360
      Left            =   6830
      TabIndex        =   12
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdLiquidar 
      Caption         =   "&Liquidar"
      Height          =   360
      Left            =   8750
      TabIndex        =   11
      Top             =   7080
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4560
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   12540
      _cx             =   22119
      _cy             =   8043
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
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEmpleado.frx":030A
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
      ExplorerBar     =   3
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
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2390
      TabIndex        =   7
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4310
      TabIndex        =   6
      Top             =   7080
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
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
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
      TabIndex        =   13
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmEmpleado"
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
    Tipo = "Empleados "
    Tipo2 = "Empleados"
    Me.Caption = Tipo
End Sub


Private Sub cmdLiquidar_Click()
    If MsgBox("¿Está seguro de liquidar al empleado " & VSFG.TextMatrix(VSFG.Row, 2) & " " & VSFG.TextMatrix(VSFG.Row, 3) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Liquidar empleado") = vbNo Then Exit Sub
    Dim FechaSalida As String
    
    Set frmFecha.objeto = VSFG
    frmFecha.Caption = "Fecha Salida"
    frmFecha.Fecha = Date
    frmFecha.Show vbModal
    FechaSalida = VSFG.Tag
    
'    'Llamar a pantalla de liquidaciones
'    frmLiquidacion.txtEmpleado = NombreEmpleado
'    frmLiquidacion.txtEmpleado.Tag = CodigoEmpleado
    'frmLiquidacion.Show
    
    strSql = "UPDATE empleado SET epl_fec_salida='" & FechaSalida & "' WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
    clsCon_Def.Ejecutar strSql
    VSFG.TextMatrix(VSFG.Row, 16) = FechaSalida
    Carga
End Sub

Private Sub cmdRestaurar_Click()
    If MsgBox("¿Está seguro de restaurar al empleado " & VSFG.TextMatrix(VSFG.Row, 2) & " " & VSFG.TextMatrix(VSFG.Row, 3) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Restaurar empleado") = vbNo Then Exit Sub
    strSql = "UPDATE empleado SET epl_fec_salida=NULL WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
    clsCon_Def.Ejecutar strSql
    VSFG.TextMatrix(VSFG.Row, 16) = ""
    Carga
End Sub



Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    
    strSql = " SELECT epl_codigo, epl_apellidos, epl_nombres, epl_cedula, epl_sexo, are_lab_codigo, " & _
             " car_codigo, epl_sueldo, epl_direccion, epl_direccion_num, epl_telefono, ciu_codigo," & _
             " epl_fec_ingreso, ' ' as vacacion, epl_baja, epl_fec_salida, asi_numasiento," & _
             " epl_fechamod,epl_usumod,'0' as modi" & _
             " FROM empleado" & _
             " WHERE emp_codigo = '" & strEmpresa & "' "
             
    
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & "AND  epl_codigo LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.value = 1 Then
        strSql = strSql & " AND  epl_nombres LIKE '%" & txtNombre.Text & "%' "
    End If
    strSql = strSql & " ORDER BY epl_baja, epl_apellidos, epl_nombres"
    clsCon_Def.Ejecutar strSql
    
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum

    VSFG_AfterRowColChange 1, 1, 1, 1

    'crea combo de area
    strSql = " SELECT are_lab_codigo, are_lab_nombre" & _
             " FROM area_laboral " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY are_lab_nombre"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*are_lab_nombre", "are_lab_codigo")
    
    
    'crea combo de cargo
    strSql = " SELECT car_codigo, car_nombre" & _
             " FROM cargo " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY car_nombre"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*car_nombre", "car_codigo")
    
    'crea combo de ciudad
    strSql = " SELECT ciu_codigo, ciu_nombre" & _
             " FROM ciudad " & _
             " ORDER BY ciu_nombre"
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(12) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*ciu_nombre", "ciu_codigo")
    
    
    VSFG.ColComboList(5) = "M|F"
    'VSFG.ColComboList(24) = "PRIMARIA|SECUNDARIA|SUPERIOR"
    'VSFG.ColComboList(26) = "SOLTERO|CASADO|DIVORCIADO|VIUDO|FALLECIDO"
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long, maximo As String
    Dim control As Long 'control de que esten llenos los datos
    If VSFG.Rows > 1 Then
        VSFG.Select 1, VSFG.Cols - 1
        VSFG.Sort = flexSortGenericDescending
    End If
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            If Trim(VSFG.TextMatrix(i, 1)) = "" Then
            ElseIf Not IsNumeric(Trim(VSFG.TextMatrix(i, 1))) Then
                MsgBox "No puede ingresar " & Tipo2 & " el código debe ser numérico", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 2)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 8)) <> "" And Not IsNumeric(Trim(VSFG.TextMatrix(i, 8))) Then
                MsgBox "No puede ingresar " & Tipo2 & " el sueldo es inválido", vbInformation, "Ingreso"
                control = 1
            ElseIf Not IsDate(Trim(VSFG.TextMatrix(i, 13))) Then
                MsgBox "No puede ingresar " & Tipo2 & " la fecha de ingreso es inválida", vbInformation, "Ingreso"
                control = 1
            Else
           strSql = " UPDATE empleado " & _
                 " SET epl_apellidos='" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "'," & _
                 " epl_nombres='" & UCase(Trim(VSFG.TextMatrix(i, 3))) & "'," & _
                 " epl_cedula='" & Trim(VSFG.TextMatrix(i, 4)) & "'," & _
                 " epl_sexo='" & VSFG.TextMatrix(i, 5) & "'," & _
                 " are_lab_codigo='" & VSFG.TextMatrix(i, 6) & "'," & _
                 " car_codigo='" & VSFG.TextMatrix(i, 7) & "'," & _
                 " epl_sueldo='" & Trim(VSFG.TextMatrix(i, 8)) & "'," & _
                 " epl_direccion='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "'," & _
                 " epl_direccion_num='" & Trim(VSFG.TextMatrix(i, 10)) & "'," & _
                 " epl_telefono='" & Trim(VSFG.TextMatrix(i, 11)) & "'," & _
                 " ciu_codigo='" & VSFG.TextMatrix(i, 12) & "'," & _
                 " epl_fec_ingreso='" & Trim(VSFG.TextMatrix(i, 13)) & "'," & _
                 " epl_baja='" & VSFG.TextMatrix(i, 15) & "'," & _
                 " epl_fec_salida='" & Trim(VSFG.TextMatrix(i, 16)) & "'," & _
                 " asi_numasiento='" & Trim(VSFG.TextMatrix(i, 17)) & "'," & _
                 " epl_fechamod=CURRENT_TIMESTAMP," & _
                 " epl_usumod='" & strUsuario & "' " & _
                 " WHERE epl_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
           clsCon_Def.Ejecutar strSql
           End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If Trim(VSFG.TextMatrix(i, 2)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 8)) <> "" And Not IsNumeric(Trim(VSFG.TextMatrix(i, 8))) Then
                MsgBox "No puede ingresar " & Tipo2 & " el sueldo es inválido", vbInformation, "Ingreso"
                control = 1
            ElseIf Not IsDate(Trim(VSFG.TextMatrix(i, 13))) Then
                MsgBox "No puede ingresar " & Tipo2 & " la fecha de ingreso es inválida", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT COALESCE(MAX(epl_codigo),0)+1 FROM empleado WHERE emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    maximo = FormatoD0(clsCon_Def.adorec_Def(0))
                Else
                    maximo = 0
                End If
                    
                strSql = " SELECT epl_codigo" & _
                    " FROM empleado " & _
                    " WHERE epl_codigo='" & maximo & "' " & _
                    " AND emp_codigo='" & strEmpresa & "' "
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                   'Hacer insert
                    
    
                    strSql = " INSERT INTO empleado (emp_codigo,epl_codigo,epl_apellidos,epl_nombres,epl_cedula,epl_sexo,are_lab_codigo," & _
                    "car_codigo,epl_sueldo,epl_direccion,epl_direccion_num,epl_telefono,ciu_codigo,epl_fec_ingreso,epl_baja,epl_fec_salida,asi_numasiento,epl_fechamod,epl_usumod) VALUES('" & _
                    strEmpresa & "','" & maximo & "','" & UCase(Trim(VSFG.TextMatrix(i, 2))) & "','" & UCase(Trim(VSFG.TextMatrix(i, 3))) & "','" & Trim(VSFG.TextMatrix(i, 4)) & "','" & VSFG.TextMatrix(i, 5) & "','" & VSFG.TextMatrix(i, 6) & "','" & _
                    VSFG.TextMatrix(i, 7) & "','" & Trim(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & Trim(VSFG.TextMatrix(i, 10)) & "','" & Trim(VSFG.TextMatrix(i, 11)) & "','" & VSFG.TextMatrix(i, 12) & "','" & Trim(VSFG.TextMatrix(i, 13)) & "','" & _
                    VSFG.TextMatrix(i, 15) & "','" & Trim(VSFG.TextMatrix(i, 16)) & "','" & Trim(VSFG.TextMatrix(i, 17)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql
                Else
                    MsgBox "No puede ingresar " & Tipo2 & " el código ya existe", vbInformation, "Ingreso"
                    control = 1
                End If
            End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then

            strSql = " SELECT count(des_codigo) as existe " & _
                    " FROM descuento " & _
                    " WHERE epl_codigo ='" & VSFG.TextMatrix(i, 1) & "' AND emp_codigo = '" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql)
    

            ' Si existe no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                If clsCon_Def.adorec_Def(0) = 1 Then
                    Mensaje = "Hay 1 registro del módulo de recursos humanos relacionado"
                Else
                    Mensaje = "Hay " & clsCon_Def.adorec_Def(0) & " registros del módulo de recursos humanos relacionados"
                End If
                MsgBox "No puede eliminar " & VSFG.TextMatrix(i, 2) & " " & VSFG.TextMatrix(i, 3) & _
                        vbNewLine & Mensaje, vbInformation, "Eliminación"
                control = 1
            Else
                strSql = " DELETE " & _
                    " FROM empleado " & _
                    " WHERE epl_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                    " AND emp_codigo = '" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga
    End If
     
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If VSFG.Rows > 1 Then
        If VSFG.TextMatrix(NewRow, 16) <> "" Then
            cmdLiquidar.Enabled = False
            cmdRestaurar.Enabled = True
        Else
            cmdLiquidar.Enabled = True
            cmdRestaurar.Enabled = False
        End If
    Else
        cmdLiquidar.Enabled = False
        cmdRestaurar.Enabled = False
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
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col = 1 Or Col = 14 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col = 14 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    End If
    If VSFG.TextMatrix(Row, 13) = "" Then VSFG.TextMatrix(Row, 13) = Format(Date, "yyyy-MM-dd")
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


Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub


