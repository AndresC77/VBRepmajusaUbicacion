VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelTipoSum 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Suministros"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelTipoSum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   3975
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo Suministro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtCuenta 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.Contable:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label LblCodigo 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   405
         Width           =   615
      End
      Begin VB.Label LblNombre 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Tipo:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   765
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSelTipoSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de los tipos de suministros, y poder modificar o    #
'#  crear o eliminar los tipos                                                  #
'#  frmSelTipoSum V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los tipos de suministro que al momento estan         #
'#  ingresadas en el sistema. Desde esta ventana se puede crear un nuevo        #
'#  tipo de suministro o modificar o eliminar las unidades ya creadas.          #
'#  Desde esta ventana se llama a la ventana frmTipoAF en la que se crea        #
'#  y modifica los tipos                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#       tipo_suministro: En esta tabla se almacenan los nuevos tipos , se      #
'#               modifican los datos de los tipos y se eliminan              #
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

Private Sub cmdEliminar_Click()
  ' Consulta para conocer si existe un producto de la marca a eliminar
    strSql = " SELECT count(tip_sum_codigo) As suministro " & _
             " FROM tipo_suministro " & _
             " WHERE tip_sum_codigo='" & dcmbCodigo.Text & "'" & _
             " AND emp_codigo='" & strEmpresa & "'"
             
    clsCon_Def.Ejecutar (strSql)
    ' Si existen produtos del este tipo a eliminar
    If clsCon_Def.adorec_Def("suministro") > 0 Then
        MsgBox "No Puede eliminar esta tipo", vbInformation, "Eliminación"
    Else ' Si no existen productos de ese tipo se elimina
        strSql = " DELETE " & _
                 " FROM tipo_suministro " & _
                 " WHERE tip_sum_codigo='" & dcmbCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"

        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Tipo eliminado", vbInformation, "Eliminación"
    End If
    ' Consulta para actualizar los combos
    strSql = " SELECT tip_sum_codigo,tip_sum_nombre,tip_sum_ctaconta " & _
             " FROM tipo_suministro " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY tip_sum_nombre "
    clsCon_Def.Ejecutar (strSql)
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "tip_sum_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_sum_nombre"
    dcmbNombre.BoundColumn = "tip_sum_codigo"
    dcmbCodigo.Text = ""
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de una marca, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código de la marca que se modificará
    frmTipoSum.Show
    frmTipoSum.txtCodigo.Text = Me.dcmbCodigo.Text
    frmTipoSum.txtNombre.Text = Me.dcmbNombre.Text
    frmTipoSum.dcmbCuentas.Text = Me.txtCuenta.Text
    frmTipoSum.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea un nueva marca se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nueva marca
    frmTipoSum.Show
    frmTipoSum.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea la marca seleccionada y escribe su nombre en el combo
    Dim strComparar As String
    'On Error GoTo errhandler
        If dcmbCodigo.Text = "" Then
            borrar_datos
            dcmbNombre.Text = ""
            Exit Sub
        End If
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            borrar_datos
            dcmbNombre.Text = ""
            Exit Sub
        End If
        
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "tip_sum_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
    If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("tip_sum_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            txtCuenta.Text = clsCon_Def.adorec_Def("tip_sum_ctaconta")
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            borrar_datos
            dcmbNombre = ""
            dcmbNombre.BoundText = ""
            txtCuenta.Text = ""
    
        End If
        dcmbCodigo.Tag = ""
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

Private Sub dcmbNombre_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbNombre.Text = "" Then
        borrar_datos
        dcmbCodigo.Text = ""
        Exit Sub
    'Else
    '    dcmbNombre.Text = UCase(dcmbNombre.Text)
    End If
    
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
        dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
    
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub Form_Activate()
' Actualiza la lista de marcas al volver al formulario
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "tip_sum_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_sum_nombre"
    dcmbNombre.BoundColumn = "tip_sum_codigo"
    If clsCon_Def.adorec_Def.RecordCount < 0 Then
      MsgBox "No tiene Tipos ingresados", vbInformation, "Tipos de Suministro"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
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
    'Consulta las marcas que estan disponibles
        strSql = " SELECT tip_sum_codigo,tip_sum_nombre,tip_sum_ctaconta " & _
                 " FROM tipo_suministro " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY tip_sum_nombre "
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "tip_sum_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "tip_sum_nombre"
        dcmbNombre.BoundColumn = "tip_sum_codigo"
        Else
            MsgBox "No tiene Tipos ingresados", vbInformation, "Seleccionar Tipos"
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
Public Sub borrar_datos()
     cmdModificar.Enabled = False
     cmdEliminar.Enabled = False
End Sub
