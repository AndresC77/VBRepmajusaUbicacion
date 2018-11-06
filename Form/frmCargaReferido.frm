VERSION 5.00
Begin VB.Form frmCargaReferido 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Referido"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaReferido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtReferente 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox txtCI 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtCliente 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4013
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2573
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Referente:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   855
      Width           =   1365
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CI/RUC Referente:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   300
      TabIndex        =   6
      Top             =   525
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   1080
      TabIndex        =   4
      Top             =   135
      Width           =   525
   End
End
Attribute VB_Name = "frmCargaReferido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta

Private Sub cmbAceptar_Click()
    Dim strSql As String
    
    If txtReferente.Tag <> "" Then
        strSql = " INSERT INTO persona_referida (emp_codigo,per_codigo,per_referido," & _
                 " per_ref_fechamod,per_ref_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & txtReferente.Tag & "','" & Me.txtCliente.Tag & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                 
        clsCon_Def.Ejecutar strSql, "M"
        
        MsgBox "Referido asignado", vbInformation, "Referido"
        Unload Me
    Else
        MsgBox "No tiene los campos llenos", vbCritical, "Referido"
    End If
    
End Sub

Private Sub cmdBuscar_Click()
    BuscarReferente UCase(txtCI.Text), txtCliente.Tag
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    INICIO = False
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
       
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtCI" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub txtCI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtCI.Text <> "" Then
            BuscarReferente UCase(txtCI.Text), txtCliente.Tag
        End If
    End If
End Sub
Private Sub BuscarReferente(strCIReferente As String, strCodigoCliente As String)
    Dim strSql As String
    Dim strTipoPedido As String
    If strCIReferente <> "" Then
        strSql = " SELECT tip_ped_codigo " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & strCodigoCliente & "'" & _
                 " AND cat_p_tipo='C'"
        clsCon_Def.Ejecutar strSql
        strTipoPedido = clsCon_Def.adorec_Def("tip_ped_codigo")
        
        strSql = " SELECT per_codigo,concat(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as cli " & _
                 " FROM persona INNER JOIN tipo_pedido " & _
                 " ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                 " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                 " AND persona.tip_ped_codigo='" & strTipoPedido & "'" & _
                 " AND persona.per_ruc='" & strCIReferente & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtReferente.Text = clsCon_Def.adorec_Def("cli")
            txtReferente.Tag = clsCon_Def.adorec_Def("per_codigo")
            cmbAceptar.Enabled = True
        Else
            MsgBox "No encuenta el Referente", vbInformation, "Referente"
            txtReferente.Text = ""
            txtReferente.Tag = ""
            cmbAceptar.Enabled = False
        End If
    End If
End Sub
