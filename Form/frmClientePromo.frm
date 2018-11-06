VERSION 5.00
Begin VB.Form frmClientePromo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente Promociones"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmClientePromo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4253
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   6495
   End
   Begin VB.TextBox txtCelular 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   6495
   End
   Begin VB.TextBox txtTelefono 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   6495
   End
   Begin VB.TextBox txtDireccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   6495
   End
   Begin VB.TextBox txtCIRUC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2453
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   2565
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Telf:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2085
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1605
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   668
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1148
      Width           =   1005
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C3DBD1&
      BackStyle       =   0  'Transparent
      Caption         =   "CI/RUC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   165
      Width           =   930
   End
End
Attribute VB_Name = "frmClientePromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Private strSql As String
Public strTipoPed As String

Private Sub cmdAceptar_Click()
    Dim CodPersona As String
    strSql = " BEGIN TRAN "
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT CONCAT('C',FORMAT(ROUND(COALESCE(MAX(REPLACE(RIGHT(per_codigo,LEN(per_codigo)-1),'C','')+0),0)+1,0),'000000'),'C') as cod " & _
             " FROM persona WITH (TABLOCKX) " & _
             " WHERE cat_p_tipo='C'" & _
             " AND per_codigo like 'C%C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    
    CodPersona = clsCon_Def.adorec_Def("cod")
    strSql = " INSERT INTO persona " & _
             " SELECT TOP(1) '" & CodPersona & "' , emp_codigo, zon_codigo, cat_p_codigo, cat_p_tipo, can_codigo, ciu_codigo, " & _
             " ven_codigo , sac_codigo, cob_codigo, for_pag_codigo, for_pag_codigo_imp, for_ent_codigo, tip_gar_codigo," & _
             " gar_aut_codigo , fid_codigo, per_tipo, '" & UCase(txtCIRUC.Text) & "', '" & UCase(txtNombre.Text) & "', '" & UCase(txtApellido.Text) & "', '" & UCase(txtDireccion.Text) & "', '" & UCase(txtDireccion.Text) & "'," & _
             " per_ubicacion , '" & UCase(txtTelefono.Text) & "', '', '" & UCase(txtCelular.Text) & "', '" & UCase(txtEmail.Text) & "', per_credito, per_dcto, per_pagare," & _
             " per_garantiasolidariareal , per_nombregarante, per_codigo_ref, per_codigo_ref2, per_codigo_ref3," & _
             " per_codigo_ref4 , per_codigo_ref5, per_codigo_ref6, per_codigo_ref7, per_codigo_ref8, per_codigo_ref9," & _
             " per_codigo_ref10 , per_codigo_resp, per_fac_flete, '" & strTipoPed & "', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP," & _
             " per_observacion , per_sec_publico, per_siniva, per_especial, per_bloqueado, per_bloqueado_g, CURRENT_TIMESTAMP," & _
             " '" & strUsuario & "' , per_inactivo, per_aplica_nc, per_es_em, per_es_di, per_es_gz, per_es_ee, per_es_n5, per_es_n6," & _
             " per_es_n7 , per_es_n8, per_es_n9, per_es_n10, per_padre, dis_pol_codigo, per_sexo, est_civ_codigo," & _
             " ori_ing_codigo , per_codigo_postal , per_aplica_ret, '', '" & strUsuario & "', CURRENT_TIMESTAMP," & _
             " per_cm , per_rcm, for_pag_codigo_aux, per_modificado_fp, per_direccion_act, per_ubicacion_act," & _
             " per_telf_act , per_fax_act, per_celular_act, per_email_act, per_red_contado" & _
             " FROM persona WHERE per_apellido like 'CONSUMI%' and tip_ped_codigo='" & strTipoPed & "'" & _
             " ORDER BY per_codigo DESC "
    clsCon_Def.Ejecutar strSql, "M"
    
    strSql = " COMMIT TRAN "
    clsCon_Def.Ejecutar strSql, "M"
    frmV_PedBod.txtRuc.Text = UCase(txtCIRUC.Text)
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    If strPtoFactura = "" Then
        Cancel = vbCancel
        Exit Sub
    End If
    
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub Form_Load()
 Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = (mdiPrincipal.Height - Me.Height) / 2
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta los documentos que estan disponibles
'        strSql = " SELECT tip_ped_codigo,tip_ped_nombre " & _
'                 " FROM tipo_pedido " & _
'                 " ORDER BY tip_ped_nombre"
'
'        clsCon_Def.Ejecutar (strSql)
'
'        'Muestra los datos de cada agente en los combobox
'
'        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
'        dcmbNombre.ListField = "tip_ped_nombre"
'        dcmbNombre.BoundColumn = "tip_ped_codigo"
        Exit Sub

errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtCIRUC_Validate(Cancel As Boolean)
    Dim strCed As String
    strCed = UCase(Trim(txtCIRUC.Text))
    If Left(strCed, 1) <> "P" Then
        If VerificaCedula(strCed) = False Then
            MsgBox "La CI/RUC no es valido", vbInformation, "CI/RUC"
            Cancel = True
        End If
    End If
    txtCIRUC.Text = strCed
    strSql = " SELECT count(per_ruc) " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " AND tip_ped_codigo='" & strTipoPed & "' " & _
             " AND per_ruc='" & UCase(Trim(txtCIRUC.Text)) & "' "
    clsCon_Def.Ejecutar strSql
    If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
        MsgBox "Este CI/RUC " & txtCIRUC.Text & " ya existe", vbInformation, "Ingreso"
        txtCIRUC.Text = ""
        Cancel = True
    Else
        strSql = " SELECT per_apellido,per_nombre,per_direccion,per_telf,per_celular,per_email " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND tip_ped_codigo != '" & strTipoPed & "' " & _
                 " AND per_ruc='" & UCase(Trim(txtCIRUC.Text)) & "' " & _
                 " ORDER BY per_fechamod DESC"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtApellido.Text = clsCon_Def.adorec_Def("per_apellido")
            txtNombre.Text = clsCon_Def.adorec_Def("per_nombre")
            txtDireccion.Text = clsCon_Def.adorec_Def("per_direccion")
            txtTelefono.Text = clsCon_Def.adorec_Def("per_telf")
            txtCelular.Text = clsCon_Def.adorec_Def("per_celular")
            txtEmail.Text = clsCon_Def.adorec_Def("per_email")
        End If
    End If
    
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
    If revisarEmail(txtEmail.Text) = False Then
        MsgBox "El email no tiene un formato valido", vbInformation, "Email"
        Cancel = True
    End If
End Sub
