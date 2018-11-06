VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelVerificadora 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificadoras"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmSelVerificadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   3405
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Verificadoras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   135
      TabIndex        =   10
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtFax 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   1920
         Width           =   1920
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   2280
         Width           =   1920
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   1560
         Width           =   1920
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   840
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
         Top             =   1230
         Width           =   1695
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
         Left            =   120
         TabIndex        =   15
         Top             =   1590
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
         Left            =   120
         TabIndex        =   14
         Top             =   2310
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
         Top             =   1950
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
         Top             =   525
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
         Top             =   870
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   195
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1755
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   195
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1755
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelVerificadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################'
'#  Forma para la selección de Verificadoras, y poder modificar, ingresar o eliminar    #
'#  verificadoras                                                                       #
'#  frmSelVerificadora V1.0                                                             #
'#  Copyright (C) 2002                                                                  #
'#                                                                                      #
'#  Ventana para consultar las verificadoras que hasta el momento estan ingresados en   #
'#  en el sistema. Desde esta ventana se puede añadir una nueva verificadora, modificar #
'#  o eliminar las verificadoras ya ingresados.                                         #
'#  Esta ventana se llama a la ventana frmVerificadora en la que se añade y modifica    #
'#  las verificadoras                                                                   #
'#                                                                                      #
'#  Tablas que se maneja:                                                               #
'#    verificadora: En esta tabla se almacenan las nuevas verificadoras, se modifican   #
'#                  los datos y se eliminan los ya ingresados.                          #
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
'Elimina las verificadoras existentes
  Dim strSql As String
  If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione una verificadora", vbInformation, "Verificadora"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
   Else
    ' Consulta para conocer si existen pedidos asignadas a dicha verificadora
    strSql = " SELECT count(ver_codigo) as Num " & _
             " FROM pedido_importacion" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND ver_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen pedidos con esta verificadora, no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar esta verificadora", vbInformation, "Eliminación"
    Else ' Si no existen pedidos con esa verificadora, se procede a eliminar
        strSql = " DELETE " & _
                 " FROM verificadora " & _
                 " WHERE ver_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Verificadora eliminada", vbInformation, "Eliminación"
    End If

    
    ' Consulta para actualizar los combos
     strSql = " SELECT ver_codigo,ver_nombre,ver_direccion,ver_telf,ver_fax,ver_email" & _
                 " FROM verificadora" & _
                 " ORDER BY ver_codigo"
        
        clsCon_Def.Ejecutar (strSql)
        
        'Muestra los datos de los códigos de la verificadora
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "ver_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "ver_nombre"
        dcmbNombre.BoundColumn = "ver_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("ver_codigo")
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

' Modifica los datos de la verificadora seleccionada, se manda a la variable Tag del formulario una bandera para
' que indique qué se va a modificar de la verificadora, además se envia como datos a la forma frmVerificadora el código y el nombre
    Dim intPos As Integer
    'Verifica si se ha seleccionado una verificadora para ser modificada
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione una verificadora", vbInformation, "Verificadora"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmVerificadora.Tag = "M"
    frmVerificadora.txtCodigo.Text = Me.dcmbCodigo.Text
    frmVerificadora.txtNombre.Text = Me.dcmbNombre.Text
    frmVerificadora.txtDireccion.Text = Me.txtDireccion
    frmVerificadora.txtEmail.Text = Me.txtEmail
    frmVerificadora.txtTelefono.Text = Me.txtTelefono
    frmVerificadora.txtFax.Text = Me.txtFax
    frmVerificadora.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa una nueva verificadora, se manda a la variable Tag del formulario una bandera para
' que indique que se va a ingresar una nueva verificadora
    frmVerificadora.Tag = "N"
    frmVerificadora.Show
End Sub
Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
'Muestra el nombre relacionado con el código de la verificadora en el momento de seleccionar una del combobox
    clsCon_Def.adorec_Def.MoveFirst
    clsCon_Def.adorec_Def.Find "ver_codigo = '" & dcmbCodigo & "'", , adSearchForward
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
        dcmbNombre = clsCon_Def.adorec_Def("ver_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        txtDireccion = clsCon_Def.adorec_Def("ver_direccion")
        txtTelefono.Text = clsCon_Def.adorec_Def("ver_telf")
        txtFax = clsCon_Def.adorec_Def("ver_fax")
        txtEmail = clsCon_Def.adorec_Def("ver_email")
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
    dcmbCodigo.ListField = "ver_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "ver_nombre"
    dcmbNombre.BoundColumn = "ver_codigo"
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
        strSql = " SELECT ver_codigo,ver_nombre,ver_direccion,ver_telf," & _
                 " ver_fax,ver_email" & _
                 " FROM verificadora" & _
                 " ORDER BY ver_codigo"
        
        clsCon_Def.Ejecutar (strSql)
      
        'Muestra los datos de cada verificadora en los combobox
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "ver_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "ver_nombre"
        dcmbNombre.BoundColumn = "ver_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("ver_codigo")
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

