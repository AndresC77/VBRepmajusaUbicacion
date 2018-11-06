VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSuministro 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Suministros"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuministro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   8880
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   4498
      TabIndex        =   8
      Top             =   2040
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2683
      TabIndex        =   7
      Top             =   2040
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Suministro"
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
      TabIndex        =   9
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2265
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   3105
      End
      Begin VB.TextBox txtdescripcion 
         Height          =   795
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtExistencia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtUltimo_precio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtPrecio_prom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   330
         Left            =   5400
         TabIndex        =   4
         Top             =   600
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Último Precio:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   638
         Width           =   375
      End
      Begin VB.Label Label15 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   165
         TabIndex        =   13
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label9 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia Suministro:"
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   120
         TabIndex        =   12
         Top             =   1372
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Promedio:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   990
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSuministro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Suministros con los que trabajará la     #
'#  empresa.                                                                    #
'#  frmSuministro V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de los suministros.                 #
'#  Permitirá almacenar en la base de datos nuevos suministros y modificar      #
'#  sus datos, esto dependiendo de la propiedad Tag, la cual se cambiará en la  #
'#  ventana frmSelsuministro y desde esta se llamará a esta ventana.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     Tipo_Suministro: En esta tabla se almacenan los tipos de Suministros     #
'#               con sus codigos parar cada suministro.                         #
'#  Procedimientos INTERNOS:                                                    #
'#     LlenarListaGrupo(strCod As String, intNiv As Integer)                    #
'#               Proceso para llenar la lista el grupo y sub grupos a los       #
'#               que pertenece el Suministro.                                   #
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
Private clsCon_nivel As clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_nivel = Nothing
End Sub

Private Sub cmbAceptar_Click()
    Dim strSql As String
'    ' Si todos los campos estan llenos
    If txtCodigo.Text <> "" _
        And txtNombre.Text <> "" _
        And txtDescripcion.Text <> "" _
        And txtExistencia.Text <> "" _
        And txtPrecio_prom.Text <> "" _
        And dcmbTipo.Text <> "" _
        And txtUltimo_precio.Text <> "" Then
       
    ' Si se esta ingresando un nuevo Suministro
        If Me.Tag = "N" Then
         On Error GoTo errhandler
        'verifico que no se repita el codigo del Suministro
            strSql = " select sum_codigo from suministro where sum_codigo='" & txtCodigo.Text & "'  "
            clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount <= 0 Then
        
        ' Almacenamiento de los datos del nuevo Suministro
                     
            strSql = " INSERT INTO suministro(sum_codigo,emp_codigo,tip_sum_codigo, sum_descripcion, " & _
                     " sum_nombre,sum_existencia,sum_precio_prom, sum_ultimo_precio, sum_fechamod,sum_usumod) " & _
                     " VALUES ('" & UCase(txtCodigo.Text) & "','" & strEmpresa & "', " & _
                     " '" & dcmbTipo.BoundText & "','" & UCase(txtDescripcion.Text) & "', '" & UCase(txtNombre.Text) & "', " & _
                     " " & txtExistencia.Text & ", " & txtPrecio_prom.Text & ", " & txtUltimo_precio.Text & "," & _
                     " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsCon_Def.Ejecutar (strSql), "M"
        
          Else
            MsgBox "El Suministro que ingresó, ya existe." & vbCrLf & "Por favor cambie el código", vbExclamation, "Error Suministro"
            txtCodigo.SetFocus
            txtCodigo.SelStart = 0
            txtCodigo.SelLength = Len(txtCodigo)
            Exit Sub
        End If
        
        ' Si se esta modificando al Suministro
        ElseIf Me.Tag = "M" Then
        'Almacenamiento de los cambios realizados al Suminstro
        strSql = " UPDATE suministro " & _
                     " SET sum_nombre='" & UCase(txtNombre.Text) & "', " & _
                     " sum_descripcion='" & UCase(txtDescripcion.Text) & "', " & _
                     " tip_sum_codigo='" & dcmbTipo.BoundText & "', " & _
                     " sum_existencia= " & txtExistencia.Text & ", " & _
                     " sum_precio_prom= " & txtPrecio_prom.Text & ", " & _
                     " sum_ultimo_precio= " & txtUltimo_precio.Text & ", " & _
                     " sum_fechamod=CURRENT_TIMESTAMP,sum_usumod='" & strUsuario & "' " & _
                     " WHERE sum_codigo='" & txtCodigo.Text & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        
        End If
        On Error GoTo errhandler
            Unload Me
    Else ' Si no estan llenos todos los campos
        MsgBox "Alguno de los campos esta vacío", vbExclamation, "ERROR"
    End If
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

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar los datos del Suministro"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nuevo Suministro"
    End If
    
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
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT tip_sum_codigo FROM tipo_suministro " & _
             " ORDER BY tip_sum_codigo "
    'Ejecuta la consulta anterior
    clsCon_Def.Ejecutar (strSql)
   ' Extrae los tipos de suministros
    Set dcmbTipo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbTipo.ListField = "tip_sum_codigo"
        
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

Private Sub txtExistencia_GotFocus()
   Seleccionar_Contenido
End Sub
Private Sub txtExistencia_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
                KeyAscii = 0
    End If
End Sub

Private Sub txtPrecio_prom_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtPrecio_prom_Validate(Cancel As Boolean)
  ' Pone los decimales en el txt de precio promedio
  If (txtPrecio_prom.Text <> "") Then
    txtPrecio_prom.Text = Format(CDbl(Val(txtPrecio_prom.Text)), "###0.00")
    End If
End Sub

Private Sub txtPrecio_prom_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> Asc(".")) Then
                KeyAscii = 0
    End If
End Sub

Private Sub txtUltimo_precio_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtUltimo_precio_Validate(Cancel As Boolean)
  ' Pone los decimales en el txt del ultimo precio
  If (txtUltimo_precio.Text <> "") Then
    txtUltimo_precio.Text = Format(CDbl(txtUltimo_precio.Text), "###0.00")
    End If
End Sub



Private Sub txtUltimo_precio_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> Asc(".")) Then
                KeyAscii = 0
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtCodigo_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtNombre_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub


