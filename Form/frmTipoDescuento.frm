VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTipoDescuento 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Descuento"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmTipoDescuento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Catálogo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8535
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Define variable Sueldo IESS"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcmbProvision 
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Sólo para grupos"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Préstamo o anticipo"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1305
         Width           =   1815
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Define variable Sueldo Mes"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Provisión"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtOrden 
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1140
      End
      Begin VB.CommandButton cmdCtaConta 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   5
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtCtaConta 
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "1"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CommandButton cmdFormula 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtFactor 
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "1"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "1"
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   240
         MaxLength       =   4
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor de Cálculo"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   3615
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
         TabIndex        =   20
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   360
         Width           =   1150
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   745
      End
      Begin VB.Label LblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4448
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3128
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmTipoDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsSql1 As New clsConsulta
Private strSql As String
Public Objeto As Object
Public Ingreso As Boolean
Public Provision As Integer

Private Sub Check1_Click(Index As Integer)
    If Index = 4 Then
        If Check1(Index).value = 1 Then
            Me.dcmbProvision.Enabled = True
        Else
            Me.dcmbProvision.Enabled = False
        End If
    End If
End Sub

Private Sub cmbAceptar_Click()
    Dim CuentaContable As String
    Dim LaProvision As String
    If Trim(Me.txtCtaConta) = "" Then
        CuentaContable = "NULL"
    Else
        CuentaContable = "'" & Trim(Me.txtCtaConta) & "'"
    End If
    If Check1(4).value = 1 Then
        LaProvision = "'" & Me.dcmbProvision.BoundText & "'"
    Else
        LaProvision = "NULL"
    End If
    If Me.Tag = "N" Then
        strSql = " INSERT INTO tipo_descuento(tip_des_codigo, tip_des_nombre, emp_codigo, tip_des_factor, tip_des_solo_grupos, tip_des_prestamo, tip_des_sueldo_mes, tip_des_impuesto_renta, tip_des_iess, tip_des_provision, tip_des_ingreso, cta_codigo, tip_des_orden, tip_des_fechamod, tip_des_usumod) " & _
                 " VALUES ('" & txtCodigo.Text & "', '" & txtNombre.Text & "', '" & strEmpresa & "', '" & Me.txtFactor & "', '" & Abs(Me.Check1(0).value) & "', '" & Abs(Me.Check1(1).value) & "', '" & Abs(Me.Check1(2).value) & "', '" & Abs(Me.Check1(3).value) & "','" & Abs(Me.Check1(5).value) & "', " & LaProvision & " ,'" & Abs(CInt(Ingreso)) & "'," & CuentaContable & ",'" & txtOrden & "'," & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
        'MsgBox "Registro de la tabla " & UCase(Objeto.tabla) & " ingresado.", vbInformation, "Ingreso"
    ElseIf Me.Tag = "M" Then
        strSql = " UPDATE tipo_descuento " & _
                 " SET tip_des_nombre='" & txtNombre.Text & "'," & _
                 " tip_des_factor='" & Me.txtFactor & "', tip_des_solo_grupos='" & Abs(Me.Check1(0).value) & "', tip_des_prestamo='" & Abs(Me.Check1(1).value) & "',tip_des_iess='" & Abs(Me.Check1(5).value) & "', tip_des_sueldo_mes='" & Abs(Me.Check1(2).value) & "', tip_des_impuesto_renta='" & Abs(Me.Check1(3).value) & "', tip_des_provision=" & LaProvision & ",cta_codigo=" & CuentaContable & ", tip_des_orden='" & txtOrden & "'," & _
                 " tip_des_fechamod=CURRENT_TIMESTAMP, tip_des_usumod='" & strUsuario & "' " & _
                 " WHERE tip_des_codigo='" & txtCodigo.Text & "' AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSql, "M"
        Dim CodigoOrden As Long
        CodigoOrden = txtOrden
        'Actualizar los siguientes registros con los nuevos valores de orden
        strSql = " SELECT tip_des_codigo FROM tipo_descuento WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_ingreso=" & Abs(CInt(Ingreso)) & _
                 " AND tip_des_orden>='" & txtOrden & "' AND tip_des_codigo<>'" & txtCodigo.Text & "' ORDER BY tip_des_orden"
        clsSql.Ejecutar (strSql)
        While clsSql.adorec_Def.EOF = False
            CodigoOrden = CodigoOrden + 1
            strSql = " UPDATE tipo_descuento SET tip_des_orden='" & CodigoOrden & "' WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_ingreso=" & Abs(CInt(Ingreso)) & _
                 " AND tip_des_codigo='" & clsSql.adorec_Def(0) & "'"
            clsSql1.Ejecutar strSql, "M"
            'Hacer el update de todos estos
            clsSql.adorec_Def.MoveNext
        Wend
        'MsgBox "Registro de la tabla " & UCase(Objeto.tabla) & " modificado.", vbInformation, "Modificación"
    End If
    Unload Me
    frmSelTipoDescuento.Show
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
    frmSelTipoDescuento.Show
End Sub



Private Sub cmdCtaConta_Click()
    Screen.MousePointer = vbHourglass
    'frmSelecCtaConta.booBoundText = False
    
    frmSelecCtaConta.Tag = "UN"
    frmSelecCtaConta.Normal1 = False
    frmSelecCtaConta.Normal = True
    Screen.MousePointer = vbDefault
    Set frmSelecCtaConta.objEscribir = txtCtaConta
    frmSelecCtaConta.Show
End Sub

Private Sub cmdFormula_Click(Index As Integer)
    Set frmFormula.Objeto = Me.txtFactor
    frmFormula.txtFormula = Me.txtFactor
    If Trim(frmFormula.txtFormula) <> "" Then
        frmFormula.Alternar (False)
    Else
        frmFormula.Alternar (True)
    End If
    frmFormula.Show vbModal
    'Me.Show
End Sub

Private Sub Form_Activate()
    If Me.Tag = "M" Then
        Me.txtOrden.Locked = False
        txtCodigo.Enabled = False
        If Ingreso = True Then
            Me.Caption = "Modificar Tipos de Ingresos Rol"
        Else
            Me.Caption = "Modificar Tipos de Egresos Rol"
        End If
    ElseIf Me.Tag = "N" Then
        Me.txtOrden.Enabled = False
        If Ingreso = True Then
            strSql = " SELECT COALESCE(max(tip_des_orden),0)+1 as num " & _
                     " FROM tipo_descuento WHERE emp_codigo='" & strEmpresa & "' AND tip_des_ingreso=1" & _
                     " GROUP BY emp_codigo"
            clsSql.Ejecutar (strSql)
            txtOrden = clsSql.adorec_Def(0)
            Me.Caption = "Ingresar Tipos de Ingresos Rol"
        Else
            strSql = " SELECT  COALESCE(max(tip_des_orden),0)+1 as num " & _
                     " FROM tipo_descuento WHERE emp_codigo='" & strEmpresa & "' AND tip_des_ingreso=0" & _
                     " GROUP BY emp_codigo"
            clsSql.Ejecutar (strSql)
            txtOrden = clsSql.adorec_Def(0)
            Me.Caption = "Ingresar Tipos de Egresos Rol"
        End If
        If Objeto.booCodigoUsuario = True Then
            txtCodigo.Enabled = True
            txtCodigo.MaxLength = Objeto.intLargoCodigo
            txtCodigo.TabIndex = 0
        Else
            Dim Codigo As String
            txtCodigo.Enabled = False
            If Ingreso = True Then
                strSql = " SELECT  COALESCE(max(" & Objeto.strTabla & "_codigo),0)+1 as num " & _
                     " FROM " & Objeto.TablaBDD & _
                     " WHERE " & Objeto.strTabla & "_codigo<1000 AND emp_codigo='" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
            Else
                strSql = " SELECT  COALESCE(max(" & Objeto.strTabla & "_codigo),1000)+1 as num " & _
                     " FROM " & Objeto.TablaBDD & " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
                     
            End If
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                Codigo = clsSql.adorec_Def("num")
                If Ingreso = False Then
                    If Val(Codigo) < 1000 Then Codigo = 1001
                End If
                While Len(Codigo) < Objeto.intLargoCodigo
                    Codigo = "0" & Codigo
                Wend
                txtCodigo = Codigo
            End If
        End If
    End If
    If Ingreso = True Then
        Me.Frame1.Caption = "Tipo de Ingreso Rol"
        Me.lblCtaConta.Caption = "Cuenta Contable HABER:"
    Else
        Me.Frame1.Caption = "Tipo de Egreso Rol"
        Me.lblCtaConta.Caption = "Cuenta Contable HABER:"
    End If
    'Poner provisiones
    strSql = " SELECT tip_des_codigo, tip_des_nombre " & _
             " FROM tipo_descuento " & _
             " WHERE emp_codigo ='" & strEmpresa & "' AND tip_des_ingreso=" & Abs(CInt(Ingreso)) & "" & _
             " ORDER BY tip_des_orden "
    clsSql.Ejecutar (strSql)
    Set dcmbProvision.RowSource = clsSql.adorec_Def.DataSource
    dcmbProvision.ListField = "tip_des_nombre"
    dcmbProvision.BoundColumn = "tip_des_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        If Me.Tag = "M" And Provision <> 0 Then
            dcmbProvision.BoundText = Provision
        Else
            dcmbProvision.BoundText = clsSql.adorec_Def(0)
        End If
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Label6_Click()

End Sub



Private Sub txtCodigo_Validate(Cancel As Boolean)
    txtCodigo = Trim(UCase(txtCodigo))
    
    strSql = " SELECT " & Objeto.strTabla & "_codigo FROM " & Objeto.TablaBDD & _
             " WHERE " & Objeto.strTabla & "_codigo='" & txtCodigo & "' AND emp_codigo='" & strEmpresa & "'"
'    If Me.Tag = "M" Then
'        strSql = strSql & " AND " & Objeto.strTabla & "_codigo <> '" & txtCodigo & "'"
'    End If
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        MsgBox "El código " & txtCodigo & " ya le pertenece a otro registro de la tabla " & UCase(Objeto.Tabla) & ".", vbInformation, "Información"
        Cancel = True
    End If
End Sub

Private Sub txtCtaConta_Change()
    Verificar
End Sub

Private Sub txtFactor_Validate(Cancel As Boolean)
    txtFactor = FormatoD0(txtFactor)
End Sub

Private Sub txtNombre_Change()
    Verificar
End Sub

Private Sub Verificar()
    If Trim(txtCodigo.Text) = "" Or Trim(txtNombre.Text) = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtNombre_Validate(Cancel As Boolean)
    txtNombre = Trim(txtNombre)
    If Objeto.Tabla <> "Unidad" Then
        txtNombre = UCase(txtNombre)
    End If
    strSql = " SELECT " & Objeto.strTabla & "_codigo FROM " & Objeto.TablaBDD & _
             " WHERE " & Objeto.strTabla & "_nombre='" & txtNombre & "' AND emp_codigo='" & strEmpresa & "'"
    If Me.Tag = "M" Then
        strSql = strSql & " AND " & Objeto.strTabla & "_codigo <> '" & txtCodigo & "'"
    End If
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        MsgBox "El nombre " & txtNombre & " ya le pertenece a otro registro de la tabla " & UCase(Objeto.Tabla) & ".", vbInformation, "Información"
        Cancel = True
    End If
End Sub

Private Sub txtOrden_Validate(Cancel As Boolean)
    txtOrden = CInt(txtOrden)
End Sub
