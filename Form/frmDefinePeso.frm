VERSION 5.00
Begin VB.Form frmDefinePeso 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peso de producto"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ControlBox      =   0   'False
   Icon            =   "frmDefinePeso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5385
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   855
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   495
      Width           =   735
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmDefinePeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Private strSql As String

Private Sub cmdAceptar_Click()
    
    If FormatoD4(txtPeso.Text) > 0 Then
        strSql = " UPDATE producto " & _
                 " SET prd_peso='" & FormatoD4(txtPeso.Text) & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & txtCodigo.Text & "' "
        clsCon_Def.Ejecutar strSql, "M"
        
        Unload Me
    Else
        MsgBox "Ingrese un peso MAYOR a cero", vbInformation, "Producto"
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

Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = (mdiPrincipal.Height - Me.Height) / 2
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
