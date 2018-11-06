VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMovIngreso 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Cta. Contable para Movimientos Ingreso"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovIngreso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3480
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cta. Contable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   113
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1920
      End
      Begin VB.CheckBox chkIVA 
         BackColor       =   &H00DDDDDD&
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   195
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   690
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCtaConta 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
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
      Begin MSDataListLib.DataCombo dcmbCtaConta2 
         Height          =   330
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servicios:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Productos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimiento:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.Contable"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   2670
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1808
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   218
      TabIndex        =   5
      Top             =   3225
      Width           =   1455
   End
End
Attribute VB_Name = "frmMovIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de modificación de la Cta.Contable de los movimientos de ingreso de   #
'#  mercaderia y del IVA.                                                       #
'#  frmMovIngreso V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la modificación de los datos de los movimientos de ingreso     #
'#  de mercadería.                                                              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  tipo_ingreso:En esta tabla se almacenan los tipos de ingresos y se          #
'#               modifican los datos de estos.                                  #
'#      ctaconta:En esta tabla se encuentran las cuentas que se pueden asignar  #
'#               a los movimientos.                                             #
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
    'EL INGRESO NO ESTA IMPLEMENTADO. NO SE NECESITA, PERO SI LA MODIFICACION
    ' Si se esta ingresando un nuevo movimiento
    If Me.Tag = "N" Then
    ' Almacenamiento de los datos de la ueva lista
        strSql = " INSERT INTO lista_precio(lis_pre_codigo,emp_codigo,lis_pre_descripcion,lis_pre_politica,lis_pre_fijada,lis_pre_fechamod,lis_pre_usumod) " & _
                 " VALUES ('" & txtCodigo.Text & "','" & strEmpresa & "','" & txtDescripcion.Text & "','" & txtPolitica.Text & "',0, " & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
    ' Si se esta modificando la Lista
    ElseIf Me.Tag = "M" Then
    'Almacenamiento de los cambios realizados al movimiento
        strSql = " UPDATE tipo_ingreso " & _
                 " SET tip_ing_ctaconta='" & dcmbCtaConta.Text & "',tip_ing_impuesto=" & chkIVA.value & _
                 ", tip_ing_ctaconta2='" & dcmbCtaConta2.Text & "' " & _
                 ",tip_ing_fechamod=CURRENT_TIMESTAMP,tip_ing_usumod='" & strUsuario & "' " & _
                 " WHERE tip_ing_codigo='" & txtCodigo.Text & "' AND emp_codigo='" & strEmpresa & "'"
    End If
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
Private Sub dcmbCtaConta2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "UN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCtaConta2
    End If
End Sub
Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar datos de la Cta. Contable de Movimiento"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nueva Lista"
        txtCodigo.Enabled = False
        'Consulta el codigo de la lista que tocaría ingresar (autonumerico)
        strSql = " SELECT COALESCE(max(lis_pre_codigo),0) as num " & _
                 " FROM lista_precio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " GROUP BY emp_codigo"
        clsCon_Def.Ejecutar (strSql)
        txtCodigo.Text = clsCon_Def.adorec_Def("num") + 1
    End If
    ' Extrae todas las cuentas de último nivel de una empresa
    strSql = " SELECT cta_codigo FROM ctaconta " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND cta_subcta = 0 " & _
             " ORDER BY cta_codigo "
    'Ejecuta la consulta anterior
    clsCon_Def.Ejecutar (strSql)
    'Muestra los datos de los códigos de las cuentas en un datacombo
    Set dcmbCtaConta.RowSource = clsCon_Def.adorec_Def.DataSource
    Set dcmbCtaConta2.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCtaConta.ListField = "cta_codigo"
    dcmbCtaConta2.ListField = "cta_codigo"
    
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

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

