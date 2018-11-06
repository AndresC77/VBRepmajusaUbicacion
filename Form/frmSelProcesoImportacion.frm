VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelProcesoImportacion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos de Importación"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmSelProcesoImportacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   3615
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Procesos de Importación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtOrden 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   795
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   1200
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
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblOrden 
         BackStyle       =   0  'Transparent
         Caption         =   "Orden:"
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
         Top             =   1950
         Width           =   1935
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   750
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   300
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1860
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   300
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1860
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelProcesoImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################'
'#  Forma para la seleccion de Procesos de Importación, y poder modificar, ingresar     #
'#  o eliminar dichos procesos                                                          #
'#  frmSelProcesoImportacion V1.0                                                       #
'#  Copyright (C) 2002                                                                  #
'#                                                                                      #
'#  Ventana para consultar los procesos que hasta el momento estan ingresados en        #
'#  en el sistema. Desde esta ventana se puede añadir un nuevo proceso, modificar       #
'#  o eliminar los procesos ya ingresados.                                              #
'#  Esta ventana se llama a la ventana frmProcesoImportacion en la que se añade y       #
'#  modifica los procesos                                                               #
'#                                                                                      #
'#  Tablas que se maneja:                                                               #
'#    proceso_importación: En esta tabla se almacenan los nuevos procesos, se           #
'#                         modifican los datos y se eliminan los ya ingresados.         #
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
'Elimina procesos de importacion existentes
  Dim strSql As String
  If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un proceso", vbInformation, "Proceso de Importación"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
   Else
    ' Consulta para conocer si existen historiales de importaciones con
    ' dicho proceso de importación
    strSql = " SELECT count(pro_imp_codigo) as Num " & _
             " FROM historial_importacion" & _
             " WHERE pro_imp_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen historiales con este proceso de importacion, no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este Proceso de importación", vbInformation, "Eliminación"
    Else ' Si no existen historiales con ese proceso, se procede a eliminar
        strSql = " DELETE " & _
                 " FROM proceso_importacion " & _
                 " WHERE pro_imp_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSql)
        MsgBox "Proceso de Importación Eliminado", vbInformation, "Eliminación"
    End If
    
    ' Consulta para actualizar los combos
 
    strSql = " SELECT pro_imp_codigo,pro_imp_nombre,pro_imp_descripcion,pro_imp_orden" & _
                 " FROM  proceso_importacion" & _
                 " ORDER BY pro_imp_codigo"
        
        clsCon_Def.Ejecutar (strSql)
        
        'Muestra los datos correspondiente al código
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "pro_imp_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "pro_imp_nombre"
        dcmbNombre.BoundColumn = "pro_imp_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("pro_imp_codigo")
        End If
    End If


End Sub

Private Sub cmdModificar_Click()

' Modifica los datos del Proceso de importación seleccionado,
' se manda a la variable Tag del formulario una bandera para
' que indique que se va a modificar el proceso, además se
' envia como datos a la forma frmProcesoImportación el código y el nombre
    Dim intPos As Integer
    'Verifica si se ha seleccionado un proceso para ser modificado
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un proceso", vbInformation, "Proceso de Importación"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmProcesoImportacion.Tag = "M"
    frmProcesoImportacion.txtCodigo.Text = Me.dcmbCodigo.Text
    frmProcesoImportacion.txtNombre.Text = Me.dcmbNombre.Text
    frmProcesoImportacion.TxtDescripcion.Text = Me.TxtDescripcion
    frmProcesoImportacion.txtOrden.Text = Me.txtOrden
    frmProcesoImportacion.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa un nuevo proceso, se manda a la variable Tag del formulario una bandera para
' que indique se se va a ingresar un nuevo proceso de importación
    frmProcesoImportacion.Tag = "N"
    frmProcesoImportacion.Show
End Sub
Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Muestra el nombre relacionado con el código del proceso de imp. en el momento
' de seleccionar uno del combobox
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        clsCon_Def.adorec_Def.MoveFirst
    End If
    clsCon_Def.adorec_Def.Find "pro_imp_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsCon_Def.adorec_Def.EOF = True Then
        dcmbNombre = ""
        dcmbNombre.BoundText = ""
        TxtDescripcion = ""
        txtOrden.Text = ""
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        dcmbNombre = clsCon_Def.adorec_Def("pro_imp_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        TxtDescripcion = clsCon_Def.adorec_Def("pro_imp_descripcion")
        txtOrden.Text = clsCon_Def.adorec_Def("pro_imp_orden")
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

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub Form_Activate()
 
    'Muestra la lista de datos actualizada
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "pro_imp_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "pro_imp_nombre"
    dcmbNombre.BoundColumn = "pro_imp_codigo"
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
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn
    'Consulta los documentos que estan disponibles
        strSql = " SELECT pro_imp_codigo,pro_imp_nombre,pro_imp_descripcion," & _
                 " pro_imp_orden" & _
                 " FROM proceso_importacion " & _
                 " ORDER BY pro_imp_orden,pro_imp_nombre"
        
        clsCon_Def.Ejecutar (strSql)
      
        'Muestra los datos de cada proceso de imp. en los combobox
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "pro_imp_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "pro_imp_nombre"
        dcmbNombre.BoundColumn = "pro_imp_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("pro_imp_codigo")
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
        SendKeys "{TAB}"
    End If
End Sub

