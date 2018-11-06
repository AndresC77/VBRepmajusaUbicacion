VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelNegocio 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Negocios"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ControlBox      =   0   'False
   Icon            =   "frmSelNegocio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5385
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1965
      TabIndex        =   1
      Top             =   550
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dcmbNombre 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Negocio:"
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
Attribute VB_Name = "frmSelNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Private strSql As String

Private Sub cmdAceptar_Click()
    
    If dcmbNombre.MatchedWithList = True Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & dcmbNombre.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        
        strPtoFactura = clsCon_Def.adorec_Def("tip_ped_ptofac")
        
        Unload Me
    Else
        MsgBox "Seleccione un Negocio", vbInformation, "NEGOCIO"
    End If
    
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
        
      
        'Muestra los datos de cada agente en los combobox
        
        Set dcmbNombre.RowSource = ComboNegocioDataSource.DataSource
        dcmbNombre.ListField = "tip_ped_nombre"
        dcmbNombre.BoundColumn = "tip_ped_codigo"
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
