VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProveedor 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12285
   Begin VB.CommandButton cmdCrearUsuarioWeb 
      Caption         =   "Crear/Modif. Usuario Web"
      Height          =   360
      Left            =   9720
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4392
      TabIndex        =   7
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
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
         TabIndex        =   11
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
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
         Caption         =   "CI/RUC"
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
      Width           =   12060
      _cx             =   95114008
      _cy             =   95099933
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
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProveedor.frx":030A
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
      WordWrap        =   -1  'True
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
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   6720
      Picture         =   "frmProveedor.frx":05D4
      ToolTipText     =   "Elimina una Fila"
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmProveedor"
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
    Tipo = "Proveedor"
    Tipo2 = "el Proveedor"
    Me.Caption = Tipo
End Sub

Private Sub cmdCrearUsuarioWeb_Click()
    Dim strUsuWeb As String
    Dim strPasWeb As String
    Dim strEmaWeb As String
    Dim clsAux As clsConsulta
    Set clsAux = New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If Trim(VSFG.TextMatrix(VSFG.Row, 18)) = "" Then
        strUsuWeb = Trim(InputBox("Ingrese el usuario para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Crear usuario"))
        strEmaWeb = Trim(InputBox("Ingrese el Email para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Crear usuario"))
        strPasWeb = Trim(InputBox("Ingrese la Clave para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Crear usuario"))
        strSQL = " INSERT INTO usuario_web_proveedor (emp_codigo, per_codigo, usu_web_pro_usuario, " & _
                 " usu_web_pro_clave, usu_web_pro_email," & _
                 " usu_web_pro_fechamod, usu_web_pro_usuwebmod)" & _
                 " VALUES('" & strEmpresa & "','" & VSFG.TextMatrix(VSFG.Row, 1) & "','" & strUsuWeb & "'," & _
                 " SUBSTRING(sys.fn_sqlvarbasetostr(HASHBYTES('MD5','" & strPasWeb & "')),3,32),'" & strEmaWeb & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
    Else
        strUsuWeb = Trim(InputBox("Ingrese el usuario para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Modificar usuario", VSFG.TextMatrix(VSFG.Row, 18)))
        strEmaWeb = Trim(InputBox("Ingrese el Email para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Modificar usuario", VSFG.TextMatrix(VSFG.Row, 19)))
        strPasWeb = Trim(InputBox("Ingrese la Clave para le proveedor" & vbNewLine & _
                             VSFG.TextMatrix(VSFG.Row, 3) & " " & VSFG.TextMatrix(VSFG.Row, 4), "Modificar usuario"))
        strSQL = " UPDATE usuario_web_proveedor " & _
                 " SET usu_web_pro_usuario='" & strUsuWeb & "', " & _
                 " usu_web_pro_clave=SUBSTRING(sys.fn_sqlvarbasetostr(HASHBYTES('MD5','" & strPasWeb & "')),3,32)," & _
                 " usu_web_pro_email='" & strEmaWeb & "'," & _
                 " usu_web_pro_fechamod=CURRENT_TIMESTAMP," & _
                 " usu_web_pro_usuwebmod='" & strUsuario & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
    End If
    If strUsuWeb <> "" And strEmaWeb <> "" And strPasWeb <> "" Then
        clsAux.Ejecutar strSQL, "M"
    Else
        MsgBox "No se procesó el usuario pues no puede tener datos vacios", vbInformation, "Usuario Web"
    End If
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSQL = " SELECT persona.per_codigo,per_tipo,per_apellido,per_nombre,cat_p_codigo,per_ruc,ciu_codigo,zon_codigo,per_direccion,per_telf,per_fax,per_email,per_credito,per_observacion,'' as con,'' as ret,'' as cta,COALESCE(usu_web_pro_usuario,'') as usuario,COALESCE(usu_web_pro_email,'') as emailusuario,per_fechamod,per_usumod, '0' as modi " & _
             " FROM persona LEFT JOIN usuario_web_proveedor " & _
             " ON persona.emp_codigo=usuario_web_proveedor.emp_codigo " & _
             " AND persona.per_codigo=usuario_web_proveedor.per_codigo " & _
             " WHERE persona.emp_codigo ='" & strEmpresa & "'" & _
             " AND persona.cat_p_tipo='P'"
    If chkFiltroCodigo.Value = 1 Then
        strSQL = strSQL & "AND  per_ruc LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSQL = strSQL & " AND CONCAT(per_apellido,' ',per_nombre) LIKE '%" & txtNombre.Text & "%' "
    End If
    strSQL = strSQL & " ORDER BY CONCAT(per_apellido,' ',per_nombre) "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    'crea combo de categoria
    strSQL = " SELECT cat_p_codigo, cat_p_nombre" & _
                 " FROM categoria_p " & _
                 " WHERE cat_p_tipo='P' " & _
                 " AND emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY cat_p_nombre"
     clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cat_p_nombre", "cat_p_codigo")
    'crea combo de ciudad
    strSQL = " SELECT ciu_codigo,pai_nombre,ciu_nombre" & _
                 " FROM ciudad INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo" & _
                 " ORDER BY pai_nombre,ciu_nombre"
     clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " pai_nombre, *ciu_nombre", "ciu_codigo")
    'crea combo de zona
    strSQL = " SELECT zon_codigo, zon_nombre" & _
                 " FROM zona " & _
                 " ORDER BY zon_nombre"
     clsCon_Def.Ejecutar strSQL
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *zon_nombre", "zon_codigo")
    Set VSFG.CellButtonPicture = imgBtnUp
    VSFG.ColComboList(15) = "..."
    VSFG.ColComboList(16) = "..."
    VSFG.ColComboList(17) = "..."
    If VSFG.Rows > 1 Then
        VSFG.Cell(flexcpPicture, 1, 15, VSFG.Rows - 1, 17) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, 1, 15, VSFG.Rows - 1, 17) = flexPicAlignRightCenter
    End If
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
            strSQL = " SELECT count(per_ruc) " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='P' " & _
                         " AND per_ruc='" & UCase(VSFG.TextMatrix(i, 6)) & "' " & _
                         " AND per_codigo!='" & VSFG.TextMatrix(i, 1) & "' "
                clsCon_Def.Ejecutar strSQL
                
                If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                    MsgBox "No puede modificar " & Tipo2 & " el CI/RUC ya existe", vbInformation, "Modificar"
                    control = 1
                Else
                    strSQL = " UPDATE persona " & _
                         " SET per_tipo='" & VSFG.TextMatrix(i, 2) & "'," & _
                         " per_apellido='" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                         " per_nombre='" & UCase(VSFG.TextMatrix(i, 4)) & "'," & _
                         " cat_p_codigo='" & UCase(VSFG.TextMatrix(i, 5)) & "'," & _
                         " per_ruc='" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                         " ciu_codigo='" & UCase(VSFG.TextMatrix(i, 7)) & "'," & _
                         " zon_codigo='" & UCase(VSFG.TextMatrix(i, 8)) & "'," & _
                         " per_direccion='" & UCase(VSFG.TextMatrix(i, 9)) & "'," & _
                         " per_telf='" & UCase(VSFG.TextMatrix(i, 10)) & "'," & _
                         " per_fax='" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
                         " per_email='" & VSFG.TextMatrix(i, 12) & "'," & _
                         " per_credito='" & FormatoD2(VSFG.TextMatrix(i, 13)) & "'," & _
                         " per_observacion='" & UCase(VSFG.TextMatrix(i, 14)) & "'," & _
                         " per_fechamod=CURRENT_TIMESTAMP," & _
                         " per_usumod='" & strUsuario & "' " & _
                         " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                         " AND emp_codigo='" & strEmpresa & "'" & _
                         " AND cat_p_tipo='P'"
                    clsCon_Def.Ejecutar strSQL, "M"
                End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el tipo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre o Apellido", vbInformation, "Ingreso"
                control = 1
'            ElseIf VSFG.TextMatrix(i, 4) = "" Then
'                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre o Apellido", vbInformation, "Ingreso"
'                control = 1
            ElseIf VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Categoria", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 6) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el CI/RUC", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 7) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Ciudad", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 8) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Zona", vbInformation, "Ingreso"
                control = 1
            Else
                strSQL = " SELECT count(per_ruc) " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='P' " & _
                         " AND per_ruc='" & UCase(VSFG.TextMatrix(i, 6)) & "' "
                clsCon_Def.Ejecutar strSQL
                
                If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                    MsgBox "No puede ingresar " & Tipo2 & " el CI/RUC ya existe", vbInformation, "Ingreso"
                    control = 1
                Else
                    strSQL = " SELECT CONCAT('P',FORMAT(ROUND(COALESCE(MAX(RIGHT(per_codigo,LEN(per_codigo)-1)),0)+1,0),'00000')) as cod " & _
                             " FROM persona " & _
                             " WHERE cat_p_tipo='P'" & _
                             " AND emp_codigo='" & strEmpresa & "'" & _
                             " GROUP BY emp_codigo"
                    clsCon_Def.Ejecutar strSQL
                    If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("cod")
                    Else
                    VSFG.TextMatrix(i, 1) = "P00000"
                    End If
                    'controla que no se repita el código
                        strSQL = " INSERT INTO persona(emp_codigo,per_codigo,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
                                 " cat_p_codigo,per_ruc,ciu_codigo,zon_codigo," & _
                                 " per_direccion,per_telf,per_fax,per_email," & _
                                 " per_credito,per_observacion," & _
                                 " per_fechamod,per_usumod) " & _
                                 " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','P','" & VSFG.TextMatrix(i, 2) & "','" & UCase(VSFG.TextMatrix(i, 3)) & "','" & UCase(VSFG.TextMatrix(i, 4)) & "', " & _
                                 " '" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "','" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 9)) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "','" & VSFG.TextMatrix(i, 12) & "'," & _
                                 " '" & FormatoD2(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "'," & _
                                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                        clsCon_Def.Ejecutar strSQL, "M"
                        If MsgBox("Desea registrar contactos al proveedor" & vbNewLine & UCase(VSFG.TextMatrix(i, 3)) & " " & UCase(VSFG.TextMatrix(i, 4)) & "?", vbQuestion + vbYesNo, "Proveedores") = vbYes Then
                            frmContacto.CodPer = VSFG.TextMatrix(i, 1)
                            frmContacto.Show vbModal
                        End If
                End If
             End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga
    End If
    
End Sub

Private Sub VSFG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 15 Then
        frmContacto.CodPer = VSFG.TextMatrix(Row, 1)
        frmContacto.Show
    ElseIf Col = 16 Then
        frmPerRetencion.CodPer = VSFG.TextMatrix(Row, 1)
        frmPerRetencion.Show
    ElseIf Col = 17 Then
        frmPerCtaBancaria.CodPer = VSFG.TextMatrix(Row, 1)
        frmPerCtaBancaria.Show
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
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 3 Then
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
    ucrtVSFG.Inicializar , False
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
