VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImpCompCont 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Contabilidad"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpCompCont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   4545
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtFecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   600
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1710
         TabIndex        =   0
         Top             =   240
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Comprobante"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2325
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   765
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmImpCompCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As clsConsulta

Private Sub dcmbCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 118 Then
        frmRecosteo.Show
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
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    frmReporte.strAsiento = dcmbCodigo.BoundText
    frmReporte.strReporte = "rptAsiento"
    frmReporte.Show
End Sub

Private Sub dcmbCodigo_Change()
' Chequea la zona seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            clsCon_Def.adorec_Def.MoveFirst
        End If
        strComparar = "asi_numasiento = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            txtDescripcion.Text = clsCon_Def.adorec_Def("asi_descripcion")
            txtFecha.Text = clsCon_Def.adorec_Def("asi_fecha")
            cmdVistaPrevia.Enabled = True
        Else
            txtDescripcion.Text = ""
            txtFecha.Text = ""
            cmdVistaPrevia.Enabled = False
        End If
        dcmbCodigo.Tag = ""
        Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub Form_Activate()
' Actualiza la lista de zonas al volver al formulario
    Dim strCod As String
    strCod = dcmbCodigo.BoundText
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "asi_numasiento"
    dcmbCodigo.BoundColumn = "asi_numasiento"
    dcmbCodigo.BoundText = strCod
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las zonas que estan disponibles
        strSql = " SELECT asi_numasiento,asi_fecha,asi_descripcion " & _
                 " FROM asiento " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY asiento.asi_numasiento "
        clsCon_Def.Ejecutar (strSql)
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "asi_numasiento"
        dcmbCodigo.BoundColumn = "asi_numasiento"
    Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
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
