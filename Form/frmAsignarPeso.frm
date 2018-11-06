VERSION 5.00
Begin VB.Form frmAsignarPeso 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Peso en Lista de Embarque"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsignarPeso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3030
   Begin VB.TextBox txtPeso 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtContenedor 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtGuia 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   855
      Width           =   405
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenedor:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "No.Guia:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "frmAsignarPeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Public INICIO As Boolean


Private Sub cmbAceptar_Click()
    Dim strSql As String
    
    If txtContenedor.Text <> "" And txtPeso.Text <> "" Then
        strSql = " UPDATE contenedor SET " & _
                 " con_peso='" & txtPeso.Text & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND con_codigo='" & txtContenedor.Text & "'"
                 
        clsCon_Def.Ejecutar strSql, "M"
        
        MsgBox "Contenedor Actualizado", vbInformation, "Contenedor"
        txtContenedor.Text = ""
        txtPeso.Text = ""
        txtGuia.Text = ""
        txtGuia.SetFocus
    Else
        MsgBox "no tiene los campos llenos", vbCritical, "Contenedor"
    End If
    
    
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
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtGuia" And Screen.ActiveControl.Name <> "txtContenedor" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtContenedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarContenedor "", txtContenedor
    End If
End Sub

Private Sub txtGuia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtGuia.Text <> "" Then
            BuscarContenedor UCase(txtGuia.Text), ""
        Else
            txtContenedor.SetFocus
        End If
    End If
End Sub
Private Sub BuscarContenedor(strGuia As String, strContenedor As String)
    Dim strSql As String
    If strGuia <> "" Then
        strSql = " SELECT con_codigo " & _
                 " FROM contenedor " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND con_guia='" & strGuia & "'"
    Else
        strSql = " SELECT con_guia " & _
                 " FROM contenedor " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND con_codigo='" & strContenedor & "'"
    
    End If
    clsCon_Def.Ejecutar strSql
    
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        If strGuia <> "" Then
            txtContenedor.Text = clsCon_Def.adorec_Def("con_codigo")
        Else
            txtGuia.Text = clsCon_Def.adorec_Def("con_guia")
        End If
        txtPeso.Text = ""
        txtPeso.SetFocus
    Else
        MsgBox "No encuentra esa guia", vbInformation, "Despacho"
        txtContenedor.Text = ""
        txtPeso.Text = ""
        txtGuia.Text = ""
        txtGuia.SetFocus
    End If
End Sub
