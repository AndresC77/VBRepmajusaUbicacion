VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelSuministro 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suministro"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelSuministro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   6335
      TabIndex        =   10
      Top             =   2150
      Width           =   1700
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   450
      Left            =   4505
      TabIndex        =   9
      Top             =   2150
      Width           =   1700
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   450
      Left            =   2690
      TabIndex        =   8
      Top             =   2150
      Width           =   1700
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   905
      TabIndex        =   7
      Top             =   2150
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Suministro"
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
      Left            =   143
      TabIndex        =   11
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtExistencia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   675
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox txtUltimo_precio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtTipo 
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox TxtPrecio_prom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label8 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Último precio:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Promedio:"
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   4200
         TabIndex        =   18
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia  Suministro:"
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   240
         TabIndex        =   17
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label4 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   13
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label3 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   405
         Width           =   615
      End
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C3DBD1&
      Caption         =   "Descripción:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmSelSuministro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion del Suministro y poder modificarlo o                #
'#  crear o eliminar                                                            #
'#  frmSelSuministro V1.0                                                       #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los Suministros que al momento estan                 #
'#  ingresadas en el sistema. Desde esta ventana se puede crear un nuevo        #
'#  Suministro o modificar o eliminar los Suministros ya creados.               #
'#  Desde esta ventana se llama a la ventana frmSuministro en la que se crea    #
'#  y modifica los Suministros                                                  #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     Tipo_Suministro: En esta tabla se almacenan los tipos de Suministros     #
'#               con sus codigos parar cada suministro.                         #
'#  Procedimientos INTERNOS:                                                    #
'#     LlenarListaGrupo(strCod As String, intNiv As Integer)                    #
'#               Proceso para llenar la lista el grupo y sub grupos a los       #
'#               que pertenece el Suministro.                                   #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#    strGrupo String: Variable que tiene el codio del grupo del producto       #
'#    intNivel Integer: Variable que tiene el nivel máximo de los grupos        #
'#                                                                              #
'#                                                                              #
'################################################################################
'/******************************************************************************/'

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
    ' Consulta para conocer si hay existencia del Suministro
    strSql = " SELECT count(*) As Num " & _
             " FROM det_gasto_suministro " & _
             " WHERE sum_codigo = '" & dcmbCodigo.Text & "' " & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen Sumnistros no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este Suministro", vbInformation, "Eliminación"
    ' Si no existen Suministros se elimina
    Else
        ' Consulta para conocer si hay Suministros en el detalle de adquisicion de sumnistros
        strSql = " SELECT count(*) As Ing " & _
                 " FROM det_adquisicion_su " & _
                 " WHERE sum_codigo = '" & dcmbCodigo.Text & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql)
        ' Si existen Suministros no se elimina
        If clsCon_Def.adorec_Def("Ing") > 0 Then
            MsgBox "No Puede eliminar este Suministro", vbInformation, "Eliminación"
        ' Si no existen Suministros se elimina
                     
            Else
                'Se elimina de Suministros
                        strSql = " DELETE " & _
                                 " FROM suministro " & _
                                 " WHERE sum_codigo = '" & dcmbCodigo.Text & "' " & _
                                 " AND emp_codigo='" & strEmpresa & "'"
                        clsCon_Def.Ejecutar (strSql), "M"
                        
                        MsgBox "Suministro Eliminado", vbInformation, "Eliminación"
            
        End If
    End If
      
      'Consulta los Suministros que estan disponibles
                       strSql = " SELECT sum_codigo,tip_sum_codigo, sum_nombre,  " & _
                                " sum_descripcion, sum_existencia, " & _
                                " sum_precio_prom, sum_ultimo_precio " & _
                                " FROM suministro " & _
                                " WHERE  emp_codigo='" & strEmpresa & "' "
                        clsCon_Def.Ejecutar (strSql)
                        
                        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
                        dcmbCodigo.ListField = "sum_codigo"
                        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
                        dcmbNombre.ListField = "sum_nombre"
                        dcmbNombre.BoundColumn = "sum_codigo"
                        dcmbCodigo.Text = ""
    
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de un Suministro, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del Suministro que se modificará
    Dim i As Integer
    Dim intPos As Integer
    Dim strCodAux As String
    frmSuministro.Show
    frmSuministro.txtCodigo.Text = Me.dcmbCodigo.Text
    frmSuministro.txtNombre.Text = Me.dcmbNombre.Text
    frmSuministro.dcmbTipo.Text = Me.txtTipo.Text
    frmSuministro.txtDescripcion.Text = Me.txtDescripcion.Text
    frmSuministro.txtExistencia.Text = Me.txtExistencia.Text
    frmSuministro.txtPrecio_prom.Text = Me.txtPrecio_prom.Text
    frmSuministro.txtUltimo_precio.Text = Me.txtUltimo_precio.Text
            
    frmSuministro.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea un nuevo Suministro, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nuevo producto
    frmSuministro.Show
    frmSuministro.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea el Suministro seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
         If dcmbCodigo.Text = "" Then
            borrar_datos
            dcmbNombre.Text = ""
            Exit Sub
        End If
        
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "sum_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("sum_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            txtDescripcion.Text = clsCon_Def.adorec_Def("sum_descripcion")
            txtExistencia.Text = clsCon_Def.adorec_Def("sum_existencia")
            txtTipo.Text = clsCon_Def.adorec_Def("tip_sum_codigo")
            txtPrecio_prom.Text = clsCon_Def.adorec_Def("sum_precio_prom")
            txtUltimo_precio.Text = clsCon_Def.adorec_Def("sum_ultimo_precio")
            
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            borrar_datos
            dcmbNombre.Text = ""
            dcmbNombre.BoundText = ""
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
' Actualiza la lista de Suministros al volver al formulario
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "sum_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "sum_nombre"
    dcmbNombre.BoundColumn = "sum_codigo"
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        Set clsCon_nivel = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
            
            strSql = " SELECT sum_codigo, sum_nombre, tip_sum_codigo, " & _
                     " sum_descripcion, sum_existencia , sum_precio_prom, sum_ultimo_precio " & _
                     " FROM suministro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' "
            clsCon_Def.Ejecutar (strSql)
                Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbCodigo.ListField = "sum_codigo"
                Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbNombre.ListField = "sum_nombre"
                dcmbNombre.BoundColumn = "sum_codigo"
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

Private Sub txtPrecio_prom_change()
    ' Pone los decimales en el txt del precio promedio
    txtPrecio_prom.Text = FormatoD2(txtPrecio_prom.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Public Sub borrar_datos()
        txtDescripcion.Text = ""
        txtPrecio_prom.Text = ""
        txtUltimo_precio.Text = ""
        txtExistencia.Text = ""
        txtTipo.Text = ""
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
End Sub

Private Sub txtUltimo_precio_change()
 ' Pone los decimales en el ultimo precio
txtUltimo_precio.Text = FormatoD2(txtUltimo_precio.Text)
End Sub

