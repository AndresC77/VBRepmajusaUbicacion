VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelEmbarcador 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embarcadores"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "frmSelEmbarcador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   3510
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Embarcadores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   128
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtFax 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   1920
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   1920
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   1920
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   720
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
         Left            =   240
         TabIndex        =   16
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblTelefono 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1470
         Width           =   855
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
         Left            =   240
         TabIndex        =   14
         Top             =   2190
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
         Left            =   240
         TabIndex        =   13
         Top             =   1830
         Width           =   495
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
         Left            =   240
         TabIndex        =   12
         Top             =   405
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
         Left            =   240
         TabIndex        =   11
         Top             =   750
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   248
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1808
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   248
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1808
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelEmbarcador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Embarcadores, y poder modificar, ingresar o eliminar     #
'#  embarcadores                                                                        #
'#  frmSelEmbarcador V1.0                                                               #
'#  Copyright (C) 2002                                                                  #
'#                                                                                      #
'#  Ventana para consultar los embarcadores que hasta el momento estan ingresados en    #
'#  en el sistema. Desde esta ventana se puede añadir un nuevo embarcador, modificar    #
'#  o eliminar los embarcadores ya ingresados.                                          #
'#  Esta ventana se llama a la ventana frmEmbarcador en la que se añade y modifica      #
'#  los embarcadores                                                                    #
'#                                                                                      #
'#  Tablas que se maneja:                                                               #
'#    Embarcador: En esta tabla se almacenan los nuevos embarcadores, se modifican los  #
'#            datos y se eliminan los ya ingresados.                                    #
'#                                                                                      #
'#  Procedimientos INTERNOS:                                                            #
'#  Procedimientos EXTERNOS:                                                            #
'#                                                                                      #
'#  Objetos de la forma:                                                                #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos                  #
'#                                                                                      #
'#                                                                                      #
'########################################################################################
'/*************************************************************************************/'

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
'Elimina los embarcadores existentes
  Dim strSql As String
  If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un embarcador", vbInformation, "Embarcador"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
   Else
    ' Consulta para conocer si existen pedidos asignadas a dicho embarcador
    strSql = " SELECT count(emb_codigo) as Num " & _
             " FROM pedido_importacion" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND emb_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen pedidos con este embarcador, no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este embarcador", vbInformation, "Eliminación"
    Else ' Si no existen pedidos con ese embarcador, se procede a eliminar
        strSql = " DELETE " & _
                 " FROM embarcador " & _
                 " WHERE emb_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Embarcador eliminado", vbInformation, "Eliminación"
    End If
    
    ' Consulta para actualizar los combos
 
    strSql = " SELECT emb_codigo,emb_nombre,emb_direccion,emb_telf,emb_fax,emb_email" & _
                 " FROM embarcador " & _
                 " ORDER BY emb_codigo"
        
        clsCon_Def.Ejecutar (strSql)
        
        'Muestra los datos de los códigos del depósito
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "emb_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "emb_nombre"
        dcmbNombre.BoundColumn = "emb_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("emb_codigo")
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

' Modifica los datos del embarcador seleccionado, se manda a la variable Tag del formulario una bandera para
' que indique que se va a modificar el embarcador, además se envia como datos a la forma frmEmbarcador el código y el nombre
    Dim intPos As Integer
    'Verifica si se ha seleccionado un embarcador para ser modificado
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un embarcador", vbInformation, "Embarcador"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmEmbarcador.Tag = "M"
    frmEmbarcador.txtCodigo.Text = Me.dcmbCodigo.Text
    frmEmbarcador.txtNombre.Text = Me.dcmbNombre.Text
    frmEmbarcador.txtDireccion.Text = Me.txtDireccion
    frmEmbarcador.txtEmail.Text = Me.txtEmail
    frmEmbarcador.txtTelefono.Text = Me.txtTelefono
    frmEmbarcador.txtFax.Text = Me.txtFax
    frmEmbarcador.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa un nuevo embarcador, se manda a la variable Tag del formulario una bandera para
' que indique se se va a ingresar un nuevo embarcador
    frmEmbarcador.Tag = "N"
    frmEmbarcador.Show
End Sub
Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
'Muestra el nombre relacionado con el código del embarcador en el momento de seleccionar uno del combobox
    clsCon_Def.adorec_Def.MoveFirst
    clsCon_Def.adorec_Def.Find "emb_codigo = '" & dcmbCodigo & "'", , adSearchForward
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
        dcmbNombre = clsCon_Def.adorec_Def("emb_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        txtDireccion = clsCon_Def.adorec_Def("emb_direccion")
        txtTelefono.Text = clsCon_Def.adorec_Def("emb_telf")
        txtFax = clsCon_Def.adorec_Def("emb_fax")
        txtEmail = clsCon_Def.adorec_Def("emb_email")
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
    dcmbCodigo.ListField = "emb_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "emb_nombre"
    dcmbNombre.BoundColumn = "emb_codigo"
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
        strSql = " SELECT emb_codigo,emb_nombre,emb_direccion,emb_telf,emb_fax,emb_email" & _
                 " FROM embarcador " & _
                 " ORDER BY emb_codigo"
        
        clsCon_Def.Ejecutar (strSql)
      
        'Muestra los datos de cada embarcador en los combos
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "emb_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "emb_nombre"
        dcmbNombre.BoundColumn = "emb_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("emb_codigo")
        End If
        Exit Sub
'        ,ban_nombre,ban_direccion,ban_telefono,ban_email,ban_url
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

