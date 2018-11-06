VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClienteModSACJefe 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Clientes (Jefe SAC)"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmClienteModSACJefe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12285
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4392
      TabIndex        =   7
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11985
      Begin VB.CheckBox chkEmprendedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Emprendedor"
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkDirector 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Director"
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
         Left            =   8640
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkGerente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Gerente de Zona"
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
         Left            =   8640
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
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
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
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
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   315
         Left            =   8640
         TabIndex        =   14
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   8640
         TabIndex        =   16
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   315
         Left            =   4440
         TabIndex        =   19
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emprendedor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   1455
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Director"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8640
         TabIndex        =   17
         Top             =   1455
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gerente de Zona"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8640
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
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4080
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   12060
      _cx             =   69554968
      _cy             =   69540893
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
      Cols            =   49
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClienteModSACJefe.frx":030A
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
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   5760
      Picture         =   "frmClienteModSACJefe.frx":0905
      ToolTipText     =   "Elimina una Fila"
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmClienteModSACJefe"
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
    Tipo = "Cliente"
    Tipo2 = "el Cliente"
    Me.Caption = Tipo
End Sub

Private Sub chkGerente_Click()
    If chkGerente.value = 1 Then
        cmbGerente.Enabled = True
    Else
        cmbGerente.Enabled = False
    End If
End Sub

Private Sub chkDirector_Click()
    If chkDirector.value = 1 Then
        cmbDirector.Enabled = True
    Else
        cmbDirector.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSql = " SELECT per_codigo,per_cm,per_rcm,per_tipo,per_apellido,per_nombre,cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
             " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
             " per_credito,per_dcto,for_pag_codigo,CONCAT(ven_codigo,'') as vend,tip_ped_codigo,per_codigo_ref," & _
             " per_codigo_ref2,per_codigo_ref3,per_observacion,'' as datos,per_fac_flete,per_especial,per_bloqueado," & _
             " per_sec_publico,per_siniva,per_inactivo,per_perdesde,per_es_gz,per_es_di,per_es_em,sac_codigo,cob_codigo,per_aplica_nc,per_fechamod,per_usumod,per_fechaing,per_usuing, '0' as modi " & _
             " FROM persona" & _
             " WHERE persona.emp_codigo ='" & strEmpresa & "'" & _
             " AND persona.cat_p_tipo='C'"
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & "AND  per_ruc LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.value = 1 Then
        strSql = strSql & " AND CONCAT(per_apellido,' ',per_nombre) LIKE '%" & txtNombre.Text & "%' "
    End If
    If chkGerente.value = 1 Then
        strSql = strSql & " AND per_codigo_ref LIKE '" & cmbGerente.BoundText & "' "
    End If
    If chkDirector.value = 1 Then
        strSql = strSql & " AND per_codigo_ref2 LIKE '" & cmbDirector.BoundText & "' "
    End If
    If chkEmprendedor.value = 1 Then
        strSql = strSql & " AND per_codigo_ref3 LIKE '" & cmbEmprendedor.BoundText & "' "
    End If
    strSql = strSql & " ORDER BY CONCAT(per_apellido,' ',per_nombre) "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    'crea combo de categoria
    strSql = " SELECT cat_p_codigo, cat_p_nombre" & _
                 " FROM categoria_p " & _
                 " WHERE cat_p_tipo='C' " & _
                 " AND emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY cat_p_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cat_p_nombre", "cat_p_codigo")
    'crea combo de canal
    strSql = " SELECT can_codigo, can_nombre" & _
                 " FROM canal " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY can_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *can_nombre", "can_codigo")
    'crea combo de ciudad
    strSql = " SELECT ciu_codigo,pai_nombre,ciu_nombre" & _
                 " FROM ciudad INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo" & _
                 " ORDER BY pai_nombre,ciu_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(10) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " pai_nombre, *ciu_nombre", "ciu_codigo")
    'crea combo de zona
    strSql = " SELECT zon_codigo, zon_nombre" & _
                 " FROM zona " & _
                 " ORDER BY zon_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(11) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *zon_nombre", "zon_codigo")
    'crea combo de forma de entrega
    strSql = " SELECT for_ent_codigo, for_ent_nombre " & _
                 " FROM forma_entrega " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY for_ent_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(20) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *for_ent_nombre", "for_ent_codigo")
    'crea combo de forma de pago
    strSql = " SELECT for_pag_codigo, for_pag_nombre " & _
                 " FROM forma_pago " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY for_pag_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(23) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *for_pag_nombre", "for_pag_codigo")
    'crea combo de vendedor
    strSql = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as ven " & _
                 " FROM vendedor " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY CONCAT(ven_apellido,' ',ven_nombre) "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(24) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *ven", "ven_codigo")
    'crea combo de tipo negocio
    strSql = " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " ORDER BY tip_ped_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(25) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *tip_ped_nombre", "tip_ped_codigo")
    'crea combo de gerente de zona
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
            " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
            " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
            " AND cat_p_tipo='C' AND per_es_gz=1 " & _
            " ORDER BY per_apellido,per_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(26) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *nombre", "per_codigo")
    'crea combo de director
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
            " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
            " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
            " AND cat_p_tipo='C' AND per_es_di=1 " & _
            " ORDER BY per_apellido,per_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(27) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *nombre", "per_codigo")
    'crea combo de emprendedor
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
            " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
            " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
            " AND cat_p_tipo='C' AND per_es_em=1 " & _
            " ORDER BY per_apellido,per_nombre "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(28) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *nombre", "per_codigo")
    Set VSFG.CellButtonPicture = imgBtnUp
    VSFG.ColComboList(30) = "..."
    If VSFG.Rows > 1 Then
        VSFG.Cell(flexcpPicture, 1, 30, VSFG.Rows - 1, 30) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, 1, 30, VSFG.Rows - 1, 30) = flexPicAlignRightCenter
    End If
    
    'crea combo de sac
    strSql = " SELECT sac_codigo, CONCAT(sac_apellido,' ',sac_nombre) as sacn " & _
                 " FROM sac " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY CONCAT(sac_apellido,' ',sac_nombre) "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(41) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *sacn", "sac_codigo")
    
    'crea combo de cobrador
    strSql = " SELECT cob_codigo, CONCAT(cob_apellido,' ',cob_nombre) as cobn " & _
                 " FROM cobrador " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY CONCAT(cob_apellido,' ',cob_nombre) "
    clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(42) = VSFG.BuildComboList(clsCon_Def.adorec_Def, " *cobn", "cob_codigo")
    ucrtVSFG.PonerNum
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    Dim entra As Integer
    entra = 0
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            strSql = " SELECT count(per_ruc) " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "' " & _
                     " AND tip_ped_codigo='" & VSFG.TextMatrix(i, 25) & "' " & _
                     " AND per_codigo!='" & VSFG.TextMatrix(i, 1) & "' "
            clsCon_Def.Ejecutar strSql
            
            If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                If MsgBox("El CI/RUC " & UCase(Trim(VSFG.TextMatrix(i, 9))) & " ya existe, desea continuar?", vbQuestion + vbYesNo, "Modificar") = vbNo Then
                    entra = 1
                    control = 1
                Else
                    entra = 0
                End If
            End If
            
            If entra = 0 Then
            
            
                strSql = " UPDATE persona " & _
                     " SET per_cm='" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "'," & _
                     " per_rcm='" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "'," & _
                     " per_tipo='" & VSFG.TextMatrix(i, 4) & "'," & _
                     " per_apellido='" & UCase(VSFG.TextMatrix(i, 5)) & "'," & _
                     " per_nombre='" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                     " cat_p_codigo='" & UCase(VSFG.TextMatrix(i, 7)) & "'," & _
                     " can_codigo='" & UCase(VSFG.TextMatrix(i, 8)) & "'," & _
                     " per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "'," & _
                     " ciu_codigo='" & UCase(VSFG.TextMatrix(i, 10)) & "'," & _
                     " zon_codigo='" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
                     " per_direccion='" & UCase(VSFG.TextMatrix(i, 12)) & "'," & _
                     " per_ubicacion='" & UCase(VSFG.TextMatrix(i, 13)) & "'," & _
                     " per_telf='" & UCase(VSFG.TextMatrix(i, 14)) & "'," & _
                     " per_fax='" & UCase(VSFG.TextMatrix(i, 15)) & "'," & _
                     " per_celular='" & UCase(VSFG.TextMatrix(i, 16)) & "'," & _
                     " per_email='" & VSFG.TextMatrix(i, 17) & "'," & _
                     " per_fechacumplea='" & VSFG.TextMatrix(i, 18) & "'," & _
                     " per_direccion2='" & UCase(VSFG.TextMatrix(i, 19)) & "'," & _
                     " for_ent_codigo='" & VSFG.TextMatrix(i, 20) & "'," & _
                     " per_credito='" & FormatoD2(VSFG.TextMatrix(i, 21)) & "'," & _
                     " per_dcto='" & FormatoD4(VSFG.TextMatrix(i, 22)) & "'," & _
                     " for_pag_codigo='" & VSFG.TextMatrix(i, 23) & "',"

                strSql = strSql & " ven_codigo='" & VSFG.TextMatrix(i, 24) & "'," & _
                     " tip_ped_codigo='" & VSFG.TextMatrix(i, 25) & "'," & _
                     " per_codigo_ref='" & VSFG.TextMatrix(i, 26) & "'," & _
                     " per_codigo_ref2='" & VSFG.TextMatrix(i, 27) & "'," & _
                     " per_codigo_ref3='" & VSFG.TextMatrix(i, 28) & "'," & _
                     " per_observacion='" & UCase(VSFG.TextMatrix(i, 29)) & "'," & _
                     " per_fac_flete='" & Abs(FormatoD0(VSFG.TextMatrix(i, 31))) & "'," & _
                     " per_especial='" & Abs(FormatoD0(VSFG.TextMatrix(i, 32))) & "'," & _
                     " per_bloqueado='" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "'," & _
                     " per_sec_publico='" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "'," & _
                     " per_siniva='" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "'," & _
                     " per_inactivo='" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "'," & _
                     " per_perdesde='" & VSFG.TextMatrix(i, 37) & "'," & _
                     " per_es_gz='" & Abs(FormatoD0(VSFG.TextMatrix(i, 38))) & "'," & _
                     " per_es_di='" & Abs(FormatoD0(VSFG.TextMatrix(i, 39))) & "'," & _
                     " per_es_em='" & Abs(FormatoD0(VSFG.TextMatrix(i, 40))) & "'," & _
                     " sac_codigo='" & VSFG.TextMatrix(i, 41) & "'," & _
                     " cob_codigo='" & VSFG.TextMatrix(i, 42) & "'," & _
                     " per_aplica_nc='" & Abs(FormatoD0(VSFG.TextMatrix(i, 43))) & "'," & _
                     " per_fechamod=CURRENT_TIMESTAMP," & _
                     " per_usumod='" & strUsuario & "' " & _
                     " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C'"
                clsCon_Def.Ejecutar strSql, "M"
            End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 4) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el tipo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre o Apellido", vbInformation, "Ingreso"
                control = 1
'            ElseIf VSFG.TextMatrix(i, 6) = "" Then
'                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre o Apellido", vbInformation, "Ingreso"
'                control = 1
            ElseIf VSFG.TextMatrix(i, 7) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Categoria", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 8) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Canal", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 9)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el CI/RUC", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 10) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Ciudad", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 11) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Zona", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 23) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Forma de Pago", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 24) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Vendedor", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 25) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Tipo de Negocio", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT count(per_ruc) " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND tip_ped_codigo='" & VSFG.TextMatrix(i, 25) & "' " & _
                         " AND per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "' "
                clsCon_Def.Ejecutar strSql
                
                If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                    MsgBox "No puede ingresar " & Tipo2 & " el CI/RUC ya existe", vbInformation, "Ingreso"
                    control = 1
                Else
                    strSql = " SELECT CONCAT('C',LPAD(ROUND(COALESCE(MAX(REPLACE(per_codigo,'C','0')+0),0)+1,0),6,'0')) as cod " & _
                             " FROM persona " & _
                             " WHERE cat_p_tipo='C'" & _
                             " AND emp_codigo='" & strEmpresa & "'" & _
                             " GROUP BY emp_codigo"
                    clsCon_Def.Ejecutar strSql
                    VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("cod")
                    'controla que no se repita el código
                        strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
                                 " cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
                                 " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
                                 " per_credito,per_dcto,for_pag_codigo,ven_codigo,tip_ped_codigo,per_codigo_ref,per_codigo_ref2,per_codigo_ref3," & _
                                 " per_observacion,per_fac_flete,per_especial,per_bloqueado,per_sec_publico,per_siniva, " & _
                                 " per_fechamod,per_usumod,per_fechaing,per_usuing,per_perdesde,per_es_gz,per_es_di,per_es_em,sac_codigo,cob_codigo,per_inactivo,per_aplica_nc) " & _
                                 " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 4) & "','" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "', " & _
                                 " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & VSFG.TextMatrix(i, 17) & "','" & VSFG.TextMatrix(i, 18) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & VSFG.TextMatrix(i, 20) & "','" & FormatoD2(VSFG.TextMatrix(i, 21)) & "','" & FormatoD4(VSFG.TextMatrix(i, 22)) & "','" & VSFG.TextMatrix(i, 23) & "','" & _
                                 VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "','" & VSFG.TextMatrix(i, 26) & "','" & VSFG.TextMatrix(i, 27) & "','" & VSFG.TextMatrix(i, 28) & "','" & UCase(VSFG.TextMatrix(i, 29)) & "'," & _
                                 " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 31))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 32))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "'," & _
                                 " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "','" & VSFG.TextMatrix(i, 37) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 38))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 39))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 40))) & "', " & _
                                 " '" & VSFG.TextMatrix(i, 41) & "','" & VSFG.TextMatrix(i, 42) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 43))) & "')"
                        clsCon_Def.Ejecutar strSql, "M"
'                        If MsgBox("Desea registrar contactos al cliente" & vbNewLine & UCase(VSFG.TextMatrix(i, 3)) & " " & UCase(VSFG.TextMatrix(i, 4)) & "?", vbQuestion + vbYesNo, "Clientes") = vbYes Then
'                            frmContacto.CodPer = VSFG.TextMatrix(i, 1)
'                            frmContacto.Top = frmClienteMod.Top
'                            frmContacto.Show 1
'                        End If
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
    frmContacto.CodPer = VSFG.TextMatrix(Row, 1)
    frmContacto.Top = frmClienteMod.Top
    frmContacto.Show 1
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
        If Col <> 30 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 5 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col = 3 Or Col = 7 Or Col = 8 Or (Col >= 21 And Col <= 23) Or (Col >= 32 And Col <= 43) Or Col >= VSFG.Cols - 5 Then
            Cancel = True
        End If
    End If
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
    Unload frmContacto
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
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_gz=1" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbGerente.RowSource = clsCon_Def.adorec_Def
    cmbGerente.BoundColumn = "codigo"
    cmbGerente.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_di=1" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbDirector.RowSource = clsCon_Def.adorec_Def
    cmbDirector.BoundColumn = "codigo"
    cmbDirector.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_em=1" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbEmprendedor.RowSource = clsCon_Def.adorec_Def
    cmbEmprendedor.BoundColumn = "codigo"
    cmbEmprendedor.ListField = "nombre"
  
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
        If VSFG.TextMatrix(Row, 18) = "" Then
            VSFG.TextMatrix(Row, 18) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 37) = "" Then
            VSFG.TextMatrix(Row, 37) = HoyDia
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
        If VSFG.TextMatrix(Row, 37) = "" Then
            VSFG.TextMatrix(Row, 37) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 18) = "" Then
            VSFG.TextMatrix(Row, 18) = HoyDia
        End If
    End If
    If Col = 18 Then
        If Not IsDate(VSFG.TextMatrix(Row, 18)) Then
            VSFG.TextMatrix(Row, 18) = HoyDia
        End If
    End If
    If Col = 37 Then
        If Not IsDate(VSFG.TextMatrix(Row, 37)) Then
            VSFG.TextMatrix(Row, 37) = HoyDia
        End If
    End If
    If Col = 9 Then
        If VSFG.TextMatrix(Row, 9) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, 9)) Then
            If MsgBox("Está ingresando en el campo CI/RUC valores no numéricos, desea continuar?", vbQuestion + vbYesNo, "CI/RUC") = vbNo Then
                VSFG.TextMatrix(Row, 9) = ""
            End If
        End If
    End If
End Sub

Private Sub VSFG_EnterCell()
    If VSFG.Col = 36 And VSFG.TextMatrix(VSFG.Row, VSFG.Col) = "" And (Val(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = -2) Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Col) = HoyDia
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
