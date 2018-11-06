VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDefCXP 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Cuentas por Pagar"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefCXP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3600
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cuentas por Pagar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   113
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   570
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCtaConta 
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parametro:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.Contable:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1853
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   293
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmDefCXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la definición de IVA, cta. contable y porcentaje para los        #
'#  cobros.                                                                     #
'#  frmDefIVA V1.0                                                              #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la definición del porcentaje y del la cuenta contable          #
'#  a la que pertenece el IVA.                                                  #
'#  Esto se almacenará en la tabla PARAMETRO con el codigo IVA                  #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    parametro: En esta tabla se almacenan los datos del IVA y otros parametros#
'#               pero para lo que nos interesa se maneja el codio IVA           #
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
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmbAceptar_Click()
    Dim strSql As String
    'Almacenamiento de los cambios realizados a CXP
    strSql = " UPDATE parametro " & _
             " SET par_texto='" & dcmbCtaConta.Text & _
             "',par_fechamod=CURRENT_TIMESTAMP,par_usumod='" & strUsuario & "' " & _
             " WHERE par_codigo='CXP' AND emp_codigo='" & strEmpresa & "'"
    On Error GoTo errhandler
        clsCon_Def.Ejecutar (strSql), "M"
        Unload Me
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

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbCtaConta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "UN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCtaConta
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta los datos del IVA de la tabla
        strSql = " SELECT par_nombre,par_descripcion,par_numero,par_texto " & _
                 " FROM parametro " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND par_codigo='CXP' "
        clsCon_Def.Ejecutar (strSql)
        txtNombre.Text = clsCon_Def.adorec_Def("par_nombre")
        txtDescripcion.Text = clsCon_Def.adorec_Def("par_descripcion")
        dcmbCtaConta.Text = clsCon_Def.adorec_Def("par_texto")
        ' Extrae todas las cuentas de último nivel de una empresa
        strSql = " SELECT cta_codigo FROM ctaconta " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' AND cta_subcta = 0 " & _
                 " ORDER BY cta_codigo "
        'Ejecuta la consulta anterior
        clsCon_Def.Ejecutar (strSql)
        'Muestra los datos de los códigos de las cuentas en un datacombo
        Set dcmbCtaConta.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCtaConta.ListField = "cta_codigo"
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

