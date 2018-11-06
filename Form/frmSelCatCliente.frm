VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmSelCatCliente 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorias de Clientes"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelCatCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1193
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2753
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1193
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2753
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Categorías Clientes"
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
      Left            =   113
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkDcto 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Permite Dcto."
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Lista de Precios de la Categoría:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   4815
         Begin VB.TextBox txtDescripcion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   720
            Width           =   2730
         End
         Begin VB.TextBox txtListaPrecio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label lblListaPrecio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   11
            Top             =   412
            Width           =   540
         End
         Begin VB.Label Label1 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   750
            Width           =   1095
         End
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   600
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
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   660
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSelCatCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Categoria de Clientes, y poder modificar o    #
'#  crear o eliminar categorias                                                 #
'#  frmSelCatCliente V1.0                                                       #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las categorias de clientes que al momento estan      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  categoria o modificar o eliminar las categorias ya creadas.                 #
'#  Desde esta ventana se llama a la ventana frmCatCliente en la que se crea    #
'#  y modifica las cateorias                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    categoria_p: En esta tabla se almacenan las nuevas cateorias, se          #
'#                 modifican los datos de las categorias y se eliminan.         #
'#                                                                              #
'#    lista_precio: De esta tabla se obtiene las listas de precios que estan    #
'#                  disponibles para asignar a cada categoria de clientes.      #
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
    Dim strSql As String
    ' Consulta para conocer si existen personas con la categoría a eliminar
    strSql = " SELECT count(per_codigo) as Num " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C' " & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen personas de esta categoria no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar esta categoría", vbInformation, "Eliminación"
    Else ' Si no existen personas de esta categoria se elimina
        strSql = " DELETE " & _
                 " FROM categoria_p " & _
                 " WHERE cat_p_tipo='C' " & _
                 " AND emp_codigo='" & strEmpresa & "'" & _
                 " AND cat_p_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql)
        MsgBox "Categoría Eliminada", vbInformation, "Eliminación"
    End If
    ' Consulta para actualizar los combos
    strSql = " SELECT cat_p_codigo,cat_p_nombre,categoria_p.lis_pre_codigo,lista_precio.lis_pre_descripcion " & _
             " FROM categoria_p LEFT JOIN lista_precio " & _
             " ON categoria_p.lis_pre_codigo=lista_precio.lis_pre_codigo " & _
             " AND categoria_p.emp_codigo=lista_precio.emp_codigo " & _
             " WHERE cat_p_tipo='C' " & _
             " AND categoria_p.emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "cat_p_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "cat_p_nombre"
    dcmbNombre.BoundColumn = "cat_p_codigo"
    dcmbCodigo.Text = ""
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de una categoria, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código de la categoria que se modificará
    frmCatCliente.Show
    frmCatCliente.txtCodigo.Text = Me.dcmbCodigo.Text
    frmCatCliente.txtNombre.Text = Me.dcmbNombre.Text
    frmCatCliente.chkDcto.Value = Me.chkDcto.Value
    frmCatCliente.dcmbListaPrecio.Text = Me.txtListaPrecio.Text
    frmCatCliente.dcmbDescripcion.Text = Me.txtDescripcion.Text
    frmCatCliente.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea una nueva categoria, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará una nuevo categoria
    frmCatCliente.Show
    frmCatCliente.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea la categoria seleccionada y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "cat_p_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("cat_p_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            chkDcto.Value = clsCon_Def.adorec_Def("cat_p_dcto")
            If IsNull(clsCon_Def.adorec_Def("lis_pre_codigo")) Then
                txtListaPrecio.Text = ""
            Else
                txtListaPrecio.Text = clsCon_Def.adorec_Def("lis_pre_codigo")
            End If
            If IsNull(clsCon_Def.adorec_Def("lis_pre_descripcion")) Then
                txtDescripcion.Text = ""
            Else
                txtDescripcion.Text = clsCon_Def.adorec_Def("lis_pre_descripcion")
            End If
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            dcmbNombre.Text = ""
            dcmbNombre.BoundText = ""
            chkDcto.Value = 0
            txtListaPrecio.Text = ""
            txtDescripcion.Text = ""
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
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
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
        dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
' Actualiza la lista de categorias al volver al formulario
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "cat_p_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "cat_p_nombre"
    dcmbNombre.BoundColumn = "cat_p_codigo"

End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn
    ' Consulta para actualizar los combos
        strSql = " SELECT cat_p_codigo,cat_p_nombre,categoria_p.lis_pre_codigo,lista_precio.lis_pre_descripcion,cat_p_dcto " & _
                 " FROM categoria_p LEFT JOIN lista_precio " & _
                 " ON categoria_p.lis_pre_codigo=lista_precio.lis_pre_codigo " & _
                 " AND categoria_p.emp_codigo=lista_precio.emp_codigo " & _
                 " WHERE cat_p_tipo='C' " & _
                 " AND categoria_p.emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql)
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "cat_p_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "cat_p_nombre"
        dcmbNombre.BoundColumn = "cat_p_codigo"
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
        SendKeys "{TAB}"
    End If
End Sub


