VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmModListaEmbarque 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Lista de Embarque"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModListaEmbarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContenedor 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtGuia 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2032
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   487
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo cmbCourier 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenedor:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   165
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
      TabIndex        =   5
      Top             =   855
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   525
      Width           =   735
   End
End
Attribute VB_Name = "frmModListaEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Public INICIO As Boolean


Private Sub cmbAceptar_Click()
    Dim strSQL As String
    Dim Sec As Long
    strSQL = " BEGIN TRAN " '" LOCK TABLES contenedor WRITE,courier WRITE"
    clsCon_Def.Ejecutar strSQL, "M"

    If txtGuia.Locked = True Then
        strSQL = " SELECT cou_prefijo_secuencial,cou_secuencial_actual,cou_secuencial_mascara" & _
                 " FROM courier WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & cmbCourier.BoundText & "'"
        clsCon_Def.Ejecutar strSQL, "M"
        If clsCon_Def.adorec_Def("cou_secuencial_mascara") <> "" Then
            txtGuia.Text = clsCon_Def.adorec_Def("cou_prefijo_secuencial") & Format(clsCon_Def.adorec_Def("cou_secuencial_actual"), clsCon_Def.adorec_Def("cou_secuencial_mascara"))
            txtGuia.Tag = "T"
            Sec = clsCon_Def.adorec_Def("cou_secuencial_actual") + 1
        End If
    End If
    
    strSQL = " UPDATE contenedor SET " & _
             " cou_codigo='" & cmbCourier.BoundText & "', " & _
             " con_guia='" & UCase(txtGuia.Text) & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_codigo='" & txtContenedor.Text & "'"
             
    clsCon_Def.Ejecutar strSQL, "M"
    
    
    If txtGuia.Locked = True Then
        strSQL = " UPDATE courier " & _
                 " SET cou_secuencial_actual='" & FormatoD0(Sec) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & cmbCourier.BoundText & "'"
        clsCon_Def.Ejecutar strSQL, "M"
    End If
    
    strSQL = " COMMIT TRAN "
    clsCon_Def.Ejecutar strSQL, "M"
    
    MsgBox "Contenedor Actualizado", vbInformation, "Contenedor"
    Unload Me
    
End Sub

Private Sub cmbCourier_Validate(Cancel As Boolean)
    txtGuia.Locked = False
    txtGuia.Text = ""
    txtGuia.Tag = ""
    strSQL = " SELECT cou_prefijo_secuencial,cou_secuencial_actual,cou_secuencial_mascara" & _
             " FROM courier " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND cou_codigo='" & cmbCourier.BoundText & "'"
    clsCon_Def.Ejecutar strSQL, "M"
    If clsCon_Def.adorec_Def("cou_secuencial_mascara") <> "" Then
        txtGuia.Locked = True
        txtGuia.Text = clsCon_Def.adorec_Def("cou_prefijo_secuencial") & Format(clsCon_Def.adorec_Def("cou_secuencial_actual"), clsCon_Def.adorec_Def("cou_secuencial_mascara"))
        txtGuia.Tag = "T"
    End If

End Sub

Private Sub cmdcancelar_Click()
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
    
    
    
    strSQL = " SELECT cou_codigo, cou_nombre " & _
             " FROM courier " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSQL
    Set cmbCourier.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbCourier.ListField = "cou_nombre"
    cmbCourier.BoundColumn = "cou_codigo"
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
