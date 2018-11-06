VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelListaPrecio 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precio"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelListaPrecio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   3630
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listas de Precio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   128
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtPolitica 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1080
         Width           =   1920
      End
      Begin VB.TextBox txtDesPolitica 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDescripcion 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   720
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
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Política:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Política Fijada:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1485
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   278
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   278
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1898
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelListaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
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

Private Sub cmdCargar_Click()
    frmCargaListaPrecio.Show
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

Private Sub cmdEliminar_Click()
 Dim strSQL As String
    ' Consulta para conocer si existe un producto asignado a la lista de precios  a eliminar
    strSQL = " SELECT count(*) As Prod " & _
             " FROM categoria_p " & _
             " WHERE lis_pre_codigo='" & dcmbCodigo.Text & "'" & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSQL)
    ' Si existen precios asignados a un producto no se elimina
    If clsCon_Def.adorec_Def("Prod") > 0 Then
        MsgBox "No Puede eliminar esta lista de precio", vbInformation, "Eliminación"
    Else ' Si no existen precios asignados a un producto se elimina
        strSQL = " DELETE " & _
                 " FROM lista_precio_p " & _
                 " WHERE lis_pre_codigo='" & dcmbCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSQL), "M"
        strSQL = " DELETE " & _
                 " FROM lista_precio " & _
                 " WHERE lis_pre_codigo='" & dcmbCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSQL), "M"
        MsgBox "Lista de Precio eliminada", vbInformation, "Eliminación"
    End If
    ' Consulta para actualizar los combos
    strSQL = " SELECT lis_pre_codigo,lis_pre_descripcion,lis_pre_politica,lis_pre_fijada " & _
             " FROM lista_precio " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY lis_pre_codigo "
    clsCon_Def.Ejecutar (strSQL)
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "lis_pre_codigo"
    Set dcmbDescripcion.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbDescripcion.ListField = "lis_pre_descripcion"
    dcmbDescripcion.BoundColumn = "lis_pre_codigo"
    dcmbCodigo.Text = ""
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de una lista, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código de la lista que se modificará
    frmListaPrecio.Show
    frmListaPrecio.txtCodigo.Text = Me.dcmbCodigo.Text
    frmListaPrecio.txtDescripcion.Text = Me.dcmbDescripcion.Text
    frmListaPrecio.txtPolitica.Text = Me.txtPolitica.Text
    frmListaPrecio.txtDesPolitica.Text = Me.txtDesPolitica.Text
    frmListaPrecio.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea una nueva lista se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará una nueva lista
    frmListaPrecio.Show
    frmListaPrecio.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea la lista seleccionada y escribe su descripcion en el combo
    Dim strComparar As String
    Dim intEstadoL As Integer
    On Error GoTo errhandler
        clsCon_Def.adorec_Def.MoveFirst
        If dcmbCodigo.Text <> "" Then
            strComparar = "lis_pre_codigo = " & dcmbCodigo.Text
        Else
            strComparar = "lis_pre_codigo = 0"
        End If
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbDescripcion.Text = clsCon_Def.adorec_Def("lis_pre_descripcion")
            dcmbDescripcion.BoundText = dcmbCodigo.Text
            txtPolitica.Text = clsCon_Def.adorec_Def("lis_pre_politica")
            intEstadoL = clsCon_Def.adorec_Def("lis_pre_fijada")
            If intEstadoL = 0 Then
                txtDesPolitica.Text = "NADA"
            ElseIf intEstadoL = 1 Then
                txtDesPolitica.Text = "TOTAL"
            ElseIf intEstadoL = 2 Then
                txtDesPolitica.Text = "DEFINIDO"
            Else
                txtDesPolitica.Text = "PARCIAL"
            End If
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            dcmbDescripcion.Text = ""
            dcmbDescripcion.BoundText = ""
            txtPolitica.Text = ""
            txtDesPolitica.Text = ""
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

Private Sub dcmbDescripcion_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbCodigo.Tag <> "A" Then
        If dcmbDescripcion.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbDescripcion.BoundText
        End If
    End If
End Sub

Private Sub dcmbdescripcion_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbDescripcion.BoundText
End Sub

Private Sub dcmbDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbDescripcion.BoundText
    End If
End Sub
Private Sub Form_Activate()
' Actualiza la lista de listas al volver al formulario
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "lis_pre_codigo"
    Set dcmbDescripcion.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbDescripcion.ListField = "lis_pre_descripcion"
    dcmbDescripcion.BoundColumn = "lis_pre_codigo"
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        strSQL = " SELECT (lis_pre_codigo) AS lis_pre_codigo,lis_pre_descripcion,lis_pre_politica,lis_pre_fijada " & _
                 " FROM lista_precio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lis_pre_codigo "
        clsCon_Def.Ejecutar (strSQL)
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "lis_pre_codigo"
        Set dcmbDescripcion.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbDescripcion.ListField = "lis_pre_descripcion"
        dcmbDescripcion.BoundColumn = "lis_pre_codigo"
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

Private Sub txtPolitica_Change()
    ' Pone los decimales en el txt de la política
    txtPolitica.Text = FormatoD2(txtPolitica.Text)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

