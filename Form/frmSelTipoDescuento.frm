VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelTipoDescuento 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Descuento"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "frmSelTipoDescuento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8670
   Begin VB.OptionButton Option1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingresos Rol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2228
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Egresos Rol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4628
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1508
      TabIndex        =   6
      Tag             =   "3"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2948
      TabIndex        =   7
      Tag             =   "4"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4388
      TabIndex        =   8
      Tag             =   "5"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5828
      TabIndex        =   9
      Tag             =   "6"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Catálogos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   8415
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   7935
         _cx             =   13996
         _cy             =   2778
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
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelTipoDescuento.frx":030A
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
         Editable        =   0
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Define variable Sueldo IESS"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   18
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtProvision 
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Provisión"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Define variable Sueldo Mes"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Define variable Impuesto Renta"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtOrden 
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox txtCtaConta 
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "1"
         Top             =   2160
         Width           =   3840
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Préstamo o anticipo"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1305
         Width           =   1815
      End
      Begin VB.TextBox txtFactor 
         Height          =   315
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "1"
         Top             =   2160
         Width           =   3960
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Sólo para grupos"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1140
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Tag             =   "1"
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
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
      Begin VB.Label lblCtaConta 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Contable"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor de Cálculo"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label lblOrden 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6960
         TabIndex        =   22
         Top             =   360
         Width           =   1150
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   360
         Width           =   1150
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   5415
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   8040
         Picture         =   "frmSelTipoDescuento.frx":03BA
         ToolTipText     =   "Elimina una Fila"
         Top             =   120
         Visible         =   0   'False
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmSelTipoDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TablaBDD As String
Public Tabla As String
Public TablaHija As String
Public Tablas As String
Public strTabla As String
Public intLargoCodigo As Integer
Public booCodigoUsuario As Boolean
Private clsSql As New clsConsulta
Private strSql As String
Private Hacer As Boolean
Public Ingreso As Boolean
Private HacerActivate As Boolean
Private CuentaNomina As String
Private NumRegistros As String
Private RegistroVisible As Integer

Private Sub Check1_Click(Index As Integer)
    If Hacer = True Then Exit Sub
    If Check1(Index).value = 1 Then
        Hacer = True
        Check1(Index).value = 0
        Hacer = False
    ElseIf Check1(Index).value = 0 Then
        Hacer = True
        Check1(Index).value = 1
        Hacer = False
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim Mensaje As String
    
    'Para ver si se puede borrar
    strSql = " SELECT count(*) as Num " & _
             " FROM " & TablaHija & _
             " WHERE " & strTabla & "_codigo ='" & txtCodigo.Text & "' AND emp_codigo = '" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    
    If clsSql.adorec_Def("Num") > 0 Then
        If clsSql.adorec_Def(0) = 1 Then
            Mensaje = "Hay 1 registro relacionado"
        Else
            Mensaje = "Hay " & clsSql.adorec_Def(0) & " registros relacionados"
        End If
        MsgBox "No puede eliminar este registro de la tabla " & UCase(Tabla) & "." & _
                vbNewLine & Mensaje & " en la tabla " & UCase(TablaHija) & ".", vbInformation, "Eliminación"
        Exit Sub
    Else
        If MsgBox("¿Está seguro de eliminar este registro de la tabla " & UCase(Tabla) & "?", vbQuestion + vbYesNo, Capitulo & "Eliminar") = vbNo Then Exit Sub
        strSql = " DELETE FROM det_tip_descuento " & _
                 " WHERE det_tip_descuento.emp_codigo='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbNombre.BoundText & "'"
        clsSql.Ejecutar strSql, "M"
        strSql = " DELETE " & _
                 " FROM " & TablaBDD & _
                 " WHERE " & strTabla & "_codigo='" & txtCodigo.Text & "' AND emp_codigo = '" & strEmpresa & "'"
        clsSql.Ejecutar strSql, "M"
    End If
    'MsgBox "Registro " & dcmbNombre & " eliminado.", vbInformation, "Información"
    BuscarCatalogo
End Sub

Private Sub cmdModificar_Click()
    frmTipoDescuento.Tag = "M"
    Set frmTipoDescuento.Objeto = Me
    
    frmTipoDescuento.txtCodigo.Text = Me.txtCodigo
    frmTipoDescuento.txtNombre.Text = Me.dcmbNombre.Text
    
    frmTipoDescuento.txtFactor = Me.txtFactor(0)
    frmTipoDescuento.txtCtaConta = Me.txtCtaConta.Tag
    frmTipoDescuento.Check1(0).value = Me.Check1(0).value
    frmTipoDescuento.Check1(1).value = Me.Check1(1).value
    frmTipoDescuento.Check1(2).value = Me.Check1(2).value
    frmTipoDescuento.Check1(3).value = Me.Check1(3).value
    frmTipoDescuento.Check1(4).value = Me.Check1(4).value
    frmTipoDescuento.Check1(5).value = Me.Check1(5).value
    frmTipoDescuento.Provision = CInt(Me.txtProvision.Tag)
    frmTipoDescuento.Ingreso = Me.Ingreso
    frmTipoDescuento.txtOrden = Me.txtOrden
    frmTipoDescuento.Show 'vbModal
    HacerActivate = True
    'BuscarCatalogo
End Sub

Private Sub cmdNuevo_Click()
    RegistroVisible = 0
    Me.Hide
    frmTipoDescuento.Tag = "N"
    Set frmTipoDescuento.Objeto = Me
    frmTipoDescuento.Ingreso = Me.Ingreso
    frmTipoDescuento.Show 'vbModal
    HacerActivate = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbNombre_Change()
    If dcmbNombre.BoundText = "" Then
        limpiar
        Exit Sub
    End If
    txtCodigo = dcmbNombre.BoundText
    RegistroVisible = txtCodigo
    
     strSql = " SELECT COALESCE((concat(cta1.cta_codigo,' - ',cta1.cta_nombre)),'') AS cta_nombre, COALESCE(cta1.cta_codigo,'') AS cta_codigo," & _
             " COALESCE(tipo_descuento.tip_des_factor,'') AS tip_des_factor, tipo_descuento.tip_des_solo_grupos, tipo_descuento.tip_des_prestamo, tipo_descuento.tip_des_sueldo_mes, tipo_descuento.tip_des_iess, tipo_descuento.tip_des_impuesto_renta, tipo_descuento.tip_des_orden, COALESCE(B.tip_des_codigo,0) AS cod_provision, COALESCE(B.tip_des_nombre,'') AS provision, tipo_descuento.tip_des_fechamod, tipo_descuento.tip_des_usumod " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_provision=B.tip_des_codigo AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & txtCodigo & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)

    Me.txtCtaConta = clsSql.adorec_Def("cta_nombre")
    Me.txtCtaConta.Tag = clsSql.adorec_Def("cta_codigo")
    Me.txtFactor(0) = clsSql.adorec_Def("tip_des_factor")
    Me.txtOrden = clsSql.adorec_Def("tip_des_orden")
    Hacer = True
    If Abs(FormatoD0(clsSql.adorec_Def("tip_des_solo_grupos"))) = 1 Then
        Check1(0).value = 1
    Else
        Check1(0).value = 0
    End If
    If Abs(FormatoD0(clsSql.adorec_Def("tip_des_prestamo"))) = 1 Then
        Check1(1).value = 1
    Else
        Check1(1).value = 0
    End If
    If Abs(FormatoD0(clsSql.adorec_Def("tip_des_sueldo_mes"))) = 1 Then
        Check1(2).value = 1
    Else
        Check1(2).value = 0
    End If
    If Abs(FormatoD0(clsSql.adorec_Def("tip_des_impuesto_renta")) = 1) Then
        Check1(3).value = 1
    Else
        Check1(3).value = 0
    End If
    
    If Trim(clsSql.adorec_Def("provision")) = "" Then
        Check1(4).value = 0
    Else
        Check1(4).value = 1
    End If
    If Abs(FormatoD0(clsSql.adorec_Def("tip_des_iess"))) = 1 Then
        Check1(5).value = 1
    Else
        Check1(5).value = 0
    End If
    txtProvision = clsSql.adorec_Def("provision")
    txtProvision.Tag = clsSql.adorec_Def("cod_provision")
    Hacer = False
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 2) = ""
        VSFG.TextMatrix(i, 3) = ""
    Next i
    'Buscar configuraciones de cuentas contables
    strSql = " SELECT cta_nombre, det_tip_descuento.cta_codigo, are_lab_codigo " & _
             " FROM det_tip_descuento INNER JOIN ctaconta ON det_tip_descuento.cta_codigo=ctaconta.cta_codigo AND det_tip_descuento.emp_codigo=ctaconta.emp_codigo" & _
             " WHERE det_tip_descuento.emp_codigo='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbNombre.BoundText & "'"
    clsSql.Ejecutar (strSql)
    While clsSql.adorec_Def.EOF = False
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 1) = clsSql.adorec_Def("are_lab_codigo") Then
                VSFG.TextMatrix(i, 2) = clsSql.adorec_Def("cta_codigo") & " - " & clsSql.adorec_Def("cta_nombre")
                Exit For
            End If
        Next i
        clsSql.adorec_Def.MoveNext
    Wend
    
    VSFG.Editable = flexEDKbdMouse
    'lblRealizado.Caption = "Realizado por: " & clsSql.adorec_Def(strTabla & "_usumod") & " el " & Left(clsSql.adorec_Def(strTabla & "_fechamod"), 10) & " a las " & Mid(clsSql.adorec_Def(strTabla & "_fechamod"), 12, 8)
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
End Sub

Private Sub limpiar()
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 2) = ""
        VSFG.TextMatrix(i, 3) = ""
    Next i
    VSFG.Editable = flexEDNone
    txtCodigo = ""
    Hacer = True
    Check1(0).value = 0
    Check1(1).value = 0
    Hacer = False
    Me.txtFactor(0) = ""
    Me.txtCtaConta = ""
    Me.txtCtaConta.Tag = ""
    Me.txtOrden = ""
    'lblRealizado.Caption = "Realizado por:"
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
End Sub

Private Sub Form_Activate()
    If HacerActivate = True Then
        Me.Caption = Tablas
        Frame1.Caption = Tablas
        BuscarCatalogo
        If Ingreso = True Then
            Me.Caption = "Tipos de Ingresos Rol"
            Me.Frame1 = "Tipos de Ingresos Rol - " & NumRegistros
            Me.lblCtaConta.Caption = "Cuenta Contable HABER:"
        Else
            Me.Caption = "Tipos de Egresos Rol"
            Me.Frame1 = "Tipos de Egresos Rol - " & NumRegistros
            Me.lblCtaConta.Caption = "Cuenta Contable HABER:"
        End If
        HacerActivate = False
    End If
End Sub

Private Sub Form_Load()
    HacerActivate = True
    RegistroVisible = 0
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
   
   
   
    strSql = "SELECT are_lab_nombre, are_lab_codigo,'','','' FROM area_laboral WHERE emp_codigo='" & strEmpresa & "' ORDER BY are_lab_codigo"
    clsSql.Ejecutar (strSql)
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    VSFG.ColComboList(2) = "..."
    
    
    VSFG.ColComboList(4) = "..."
    For i = 1 To VSFG.Rows - 1
        VSFG.Cell(flexcpPictureAlignment, i, 4) = flexPicAlignRightCenter
        VSFG.Cell(flexcpPicture, i, 4) = Me.imgBtnUp
    Next i
    
    strSql = " SELECT par_con_cta_codigo FROM parametro_contable" & _
            " WHERE emp_codigo='" & strEmpresa & "' AND par_con_tipo='RRHH' AND par_con_codigo='1' "
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        CuentaNomina = clsSql.adorec_Def(0)
    Else
        CuentaNomina = ""
    End If
    
    Hacer = False
End Sub

Private Sub BuscarCatalogo()
    Screen.MousePointer = vbHourglass
    strSql = " SELECT " & strTabla & "_codigo, " & strTabla & "_nombre " & _
             " FROM " & TablaBDD & " " & _
             " WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_ingreso=" & Abs(CInt(Ingreso)) & "" & _
             " ORDER BY " & strTabla & "_orden "
    clsSql.Ejecutar (strSql)
    dcmbNombre = ""
    dcmbNombre.BoundText = ""
    Set dcmbNombre.RowSource = clsSql.adorec_Def.DataSource
    dcmbNombre.ListField = strTabla & "_nombre"
    dcmbNombre.BoundColumn = strTabla & "_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        If clsSql.adorec_Def.RecordCount = 1 Then
            NumRegistros = "1 registro"
        Else
            NumRegistros = clsSql.adorec_Def.RecordCount & " registros"
        End If
        If RegistroVisible = 0 Then
            dcmbNombre.BoundText = clsSql.adorec_Def(0)
        Else
            dcmbNombre.BoundText = RegistroVisible
        End If
    Else
        NumRegistros = "Ningún registro"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Option1_Click()
    RegistroVisible = 0
    Ingreso = True
    HacerActivate = True
    Call Form_Activate
End Sub

Private Sub Option2_Click()
    RegistroVisible = 0
    Ingreso = False
    HacerActivate = True
    Call Form_Activate
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 And Col <> 4 Then Cancel = True
    If Col = 4 Then
        VSFG.CellButtonPicture = Me.imgBtnUp
    End If
    If Col = 2 Then
        VSFG.CellButtonPicture = Nothing
    End If
End Sub

Private Sub VSFG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        Screen.MousePointer = vbHourglass
        frmSelecCtaConta.Tag = "UN"
        Screen.MousePointer = vbDefault
        Set frmSelecCtaConta.objEscribir = Me.VSFG
        frmSelecCtaConta.UserCol = 3
        frmSelecCtaConta.UserRow = Row
        frmSelecCtaConta.Normal = False
        'frmSelecCtaConta.Normal1 = False
        frmSelecCtaConta.Show
    End If
    If Col = 4 Then
        If MsgBox("¿Está seguro de quitar " & VSFG.TextMatrix(Row, 2) & " como la cuenta contable del área " & VSFG.TextMatrix(Row, 0) & "?", vbQuestion + vbYesNo, "Eliminar") = vbNo Then Exit Sub
        VSFG.TextMatrix(Row, 2) = ""
        VSFG.TextMatrix(Row, 3) = ""
        strSql = "DELETE FROM det_tip_descuento WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbNombre.BoundText & "' AND are_lab_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        clsSql.Ejecutar strSql, "M"
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 Then
        If VSFG.TextMatrix(Row, Col) <> "" Then
            strSql = "SELECT cta_nombre FROM ctaconta WHERE emp_codigo='" & strEmpresa & "' AND cta_codigo='" & VSFG.TextMatrix(Row, Col) & "'"
            clsSql.Ejecutar (strSql)
            If clsSql.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(Row, 3) & " - " & clsSql.adorec_Def(0)
            End If
            'Si es egreso, restringir a la misma cuenta si la cuenta de nómina
            If Ingreso = False And VSFG.TextMatrix(Row, Col) = CuentaNomina Then
                strSql = "DELETE FROM det_tip_descuento WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbNombre.BoundText & "'"
                clsSql.Ejecutar strSql, "M"
                For i = 1 To VSFG.Rows - 1
                    strSql = " INSERT INTO det_tip_descuento (tip_des_codigo, are_lab_codigo, emp_codigo, cta_codigo) VALUES ('" & Me.dcmbNombre.BoundText & "','" & VSFG.TextMatrix(i, 1) & "','" & strEmpresa & "', '" & CuentaNomina & "')"
                    clsSql.Ejecutar strSql, "M"
                    VSFG.TextMatrix(i, 2) = VSFG.TextMatrix(Row, 2)
                Next i
                Exit Sub
            End If
            
            strSql = "DELETE FROM det_tip_descuento WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbNombre.BoundText & "' AND are_lab_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
            clsSql.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_tip_descuento (tip_des_codigo, are_lab_codigo, emp_codigo, cta_codigo) VALUES ('" & Me.dcmbNombre.BoundText & "','" & VSFG.TextMatrix(Row, 1) & "','" & strEmpresa & "', '" & VSFG.TextMatrix(Row, 3) & "')"
            clsSql.Ejecutar strSql, "M"
        End If
    End If
End Sub
