VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelAgenteAfianzado 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agentes Afianzados"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmSelAgenteAfianzado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   4065
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Agentes Afianzados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtFax 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   1920
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2040
         Width           =   1920
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   1920
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   960
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbldireccion 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
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
         TabIndex        =   16
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label lblTelefono 
         BackStyle       =   0  'Transparent
         Caption         =   "Número Telefónico:"
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
         TabIndex        =   15
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
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
         TabIndex        =   14
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label lblFax 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
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
         TabIndex        =   13
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNombre 
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
         TabIndex        =   11
         Top             =   630
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   525
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   2085
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   525
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2085
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelAgenteAfianzado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de agentes afianzados, y poder modificar, ingresar o eliminar   #
'#  agentes afianzados                                                                      #
'#  frmSelAgenteAfianzado V1.0                                                              #
'#  Copyright (C) 2002                                                                      #
'#                                                                                          #
'#  Ventana para consultar los ag. afianzados que hasta el momento estan ingresados en      #
'#  en el sistema. Desde esta ventana se puede añadir un nuevo agente, modificar            #
'#  o eliminar los ag. afianzados ya ingresados.                                            #
'#  Esta ventana se llama a la ventana frmAgenteAfianzado en la que se añade y modifica     #
'#  los agentes                                                                             #
'#                                                                                          #
'#  Tablas que se maneja:                                                                   #
'#    agente_afianzado: En esta tabla se almacenan los nuevos agentes, se modifican los     #
'#            datos y se eliminan los ya ingresados.                                        #
'#                                                                                          #
'#  Procedimientos INTERNOS:                                                                #
'#  Procedimientos EXTERNOS:                                                                #
'#                                                                                          #
'#  Objetos de la forma:                                                                    #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos                      #
'#                                                                                          #
'#                                                                                          #
'############################################################################################
'/*****************************************************************************************/'

Private clsCon_Def As clsConsulta
Private strSql As String
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
'Elimina los agentes existentes
  Dim strSql As String
  If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un agente", vbInformation, "Agente Afianzado"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
   Else
    ' Consulta para conocer si existen pedidos asignadas a dicho agente
    strSql = " SELECT count(age_afi_codigo) as Num " & _
             " FROM pedido_importacion" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND age_afi_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen pedidos con  este agente no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este Agente Afianzado", vbInformation, "Eliminación"
    Else ' Si no existen pedidos con ese agente, se procede a eliminar
        strSql = " DELETE " & _
                 " FROM agente_afianzado " & _
                 " WHERE age_afi_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Agente Afianzado eliminado", vbInformation, "Eliminación"
    End If
    
    ' Consulta para actualizar los combobox, luego de eliminar
 
    strSql = "SELECT age_afi_codigo,age_afi_nombre,age_afi_direccion," & _
             "age_afi_telf,age_afi_fax,age_afi_email" & _
                 " FROM agente_afianzado " & _
                 " ORDER BY age_afi_codigo"
        
        clsCon_Def.Ejecutar (strSql)
        
        'Muestra los datos de los códigos del depósito
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "age_afi_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "age_afi_nombre"
        dcmbNombre.BoundColumn = "age_afi_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("age_afi_codigo")
        End If
    End If
'    strSql = " SELECT cat_p_codigo,cat_p_nombre " & _
'             " FROM categoria_p " & _
'             " WHERE cat_p_tipo='P' " & _
'             " AND emp_codigo='" & strEmpresa & "'"
'    clsCon_Def.Ejecutar (strSql)
'    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
'    dcmbCodigo.ListField = "cat_p_codigo"
'    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
'    dcmbNombre.ListField = "cat_p_nombre"
'    dcmbNombre.BoundColumn = "cat_p_codigo"
'    dcmbCodigo.Text = ""
End Sub

Private Sub cmdModificar_Click()

' Modifica los datos del 'agente' seleccionado, se manda a la variable Tag del formulario una bandera para
' que indique que se va a modificar del 'agente', además se envia como datos a la forma fmrbanco el código y el nombre
    Dim intPos As Integer
    'Verifica si se ha seleccionado un agente para ser modificado
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un agente afianzado", vbInformation, "Agente Afianzado"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmAgenteAfianzado.Tag = "M"
    frmAgenteAfianzado.txtCodigo.Text = Me.dcmbCodigo.Text
    frmAgenteAfianzado.txtNombre.Text = Me.dcmbNombre.Text
    frmAgenteAfianzado.txtDireccion.Text = Me.txtDireccion
    frmAgenteAfianzado.txtEmail.Text = Me.txtEmail
    frmAgenteAfianzado.txtTelefono.Text = Me.txtTelefono
    frmAgenteAfianzado.txtFax.Text = Me.txtFax
    frmAgenteAfianzado.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa un nuevo agente, se manda a la variable Tag del formulario una bandera para
' que indique que se va a ingresar un nuevo agente
    frmAgenteAfianzado.Tag = "N"
    frmAgenteAfianzado.Show
End Sub
Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
'Muestra el nombre relacionado con el código del agente en el momento de seleccionar uno del combobox
    clsCon_Def.adorec_Def.MoveFirst
    clsCon_Def.adorec_Def.Find "age_afi_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsCon_Def.adorec_Def.EOF = True Then
        dcmbNombre = ""
        dcmbNombre.BoundText = ""
        txtDireccion = ""
        txtTelefono.Text = ""
        txtFax = ""
        txtEmail = ""
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        dcmbNombre = clsCon_Def.adorec_Def("age_afi_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        txtDireccion = clsCon_Def.adorec_Def("age_afi_direccion")
        txtTelefono.Text = clsCon_Def.adorec_Def("age_afi_telf")
        txtFax = clsCon_Def.adorec_Def("age_afi_fax")
        txtEmail = clsCon_Def.adorec_Def("age_afi_email")
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbNombre_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub


Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub Form_Activate()
 
    'Muestra la lista de datos actualizada
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "age_afi_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "age_afi_nombre"
    dcmbNombre.BoundColumn = "age_afi_codigo"
    If Me.Tag <> "" Then
        dcmbCodigo = ""
        dcmbCodigo = Me.Tag
    ElseIf Not clsCon_Def.adorec_Def.EOF Then
        dcmbCodigo_Change
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
    'Consulta los documentos que estan disponibles
        strSql = " SELECT age_afi_codigo,age_afi_nombre," & _
                 " age_afi_direccion,age_afi_telf,age_afi_fax,age_afi_email" & _
                 " FROM agente_afianzado " & _
                 " ORDER BY age_afi_codigo"
        
        clsCon_Def.Ejecutar (strSql)
      
        'Muestra los datos de cada agente en los combobox
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "age_afi_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "age_afi_nombre"
        dcmbNombre.BoundColumn = "age_afi_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("age_afi_codigo")
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

