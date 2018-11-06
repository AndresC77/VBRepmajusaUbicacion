VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCambiarColeccion 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Colección"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambiarColeccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3720
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin MSDataListLib.DataCombo dcmbMarca 
         Height          =   330
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Colecciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
      Begin MSDataListLib.DataCombo dcmbOrigen 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDestino 
         Height          =   330
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   585
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1913
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   353
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmCambiarColeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Dim strSQL As String

Private Sub cmbAceptar_Click()
    If MsgBox("Esta seguro de cambiar la colección?", vbYesNo + vbQuestion, "Inventario") = vbYes Then
        strSQL = " UPDATE producto " & _
                 " SET clc_codigo='" & dcmbDestino.BoundText & "' " & _
                 " WHERE clc_codigo='" & dcmbOrigen.BoundText & "' " & _
                 " AND mar_codigo like '" & Trim(dcmbMarca.BoundText) & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        MsgBox "El cambio fue realizado con exito", vbInformation, "Inventario"
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT clc_codigo,clc_nombre " & _
             " FROM coleccion " & _
             " ORDER BY clc_nombre "
    clsCon_Def.Ejecutar strSQL
    Set dcmbOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbOrigen.ListField = "clc_nombre"
    dcmbOrigen.BoundColumn = "clc_codigo"
    Set dcmbDestino.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbDestino.ListField = "clc_nombre"
    dcmbDestino.BoundColumn = "clc_codigo"
    strSQL = " CREATE TABLE #marT (mar_codigo char(3),mar_nombre varchar(50))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO #marT VALUES('%',' ---Todas las Marcas--- ')"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO #marT" & _
             " SELECT mar_codigo, mar_nombre " & _
             " FROM marca WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY mar_nombre"
    clsCon_Def.Ejecutar strSQL
    strSQL = " SELECT mar_codigo, mar_nombre " & _
             " FROM #marT " & _
             " ORDER BY mar_nombre"
    clsCon_Def.Ejecutar strSQL
    Set Me.dcmbMarca.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbMarca.ListField = "mar_nombre"
    dcmbMarca.BoundColumn = "mar_codigo"
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        dcmbMarca.BoundText = clsCon_Def.adorec_Def("mar_codigo")
    End If
    strSQL = " DROP TABLE #marT "
    clsCon_Def.Ejecutar strSQL
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

