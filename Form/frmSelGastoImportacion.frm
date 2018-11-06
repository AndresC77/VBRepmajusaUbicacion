VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelGastoImportacion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos de Importación"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelGastoImportacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   3600
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Gastos de Importación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   113
      TabIndex        =   12
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtProrra 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   2520
         Width           =   1920
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   2880
         Width           =   720
      End
      Begin VB.TextBox txtTipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1800
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin VB.TextBox txtCalcula 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   2160
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
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
         TabIndex        =   19
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         TabIndex        =   18
         Top             =   1845
         Width           =   345
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         TabIndex        =   17
         Top             =   2925
         Width           =   435
      End
      Begin VB.Label lblProrra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prorratea a:"
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
         TabIndex        =   16
         Top             =   2565
         Width           =   855
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
         TabIndex        =   15
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   765
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calcula según:"
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
         Left            =   135
         TabIndex        =   13
         Top             =   2205
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   285
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   285
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblPorcentaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      Caption         =   "( En porcentaje )"
      ForeColor       =   &H00644017&
      Height          =   210
      Left            =   6360
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmSelGastoImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################'
'#  Forma para la seleccion de Gastos de Importación, y poder modificar, ingresar       #
'#  o eliminar gastos de importación.                                                   #
'#  frmSelGastoImportacion V1.0                                                         #
'#  Copyright (C) 2002                                                                  #
'#                                                                                      #
'#  Ventana para consultar los gastos que hasta el momento estan ingresados en          #
'#  en el sistema. Desde esta ventana se puede añadir un nuevo gasto, modificar         #
'#  o eliminar los gastos de importación ya ingresados.                                 #
'#  Esta ventana se llama a la ventana frmGastoImportacion en la que se añade y         #
'#  modifica los gastos                                                                 #
'#                                                                                      #
'#  Tablas que se maneja:                                                               #
'#    gasto_importacion: En esta tabla se almacenan los nuevos gastos, se modifican los #
'#                       datos y se eliminan los ya ingresados.                         #
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
Private strSQL As String
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
'Elimina un gasto de importación existente escogido por el usuario
  Dim strSQL As String
  If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un gasto de importación", vbInformation, "Gasto de importación"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
   Else
    ' Consulta para conocer si existen detalles de gastos de importación
    ' con el gasto de importación escogido para borrar
    strSQL = " SELECT count(gas_imp_codigo) as Num " & _
             " FROM det_gasto_imp " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND gas_imp_codigo='" & dcmbCodigo.Text & "'"
    clsCon_Def.Ejecutar (strSQL)
    ' Si existen detalles con este gasto, no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No se puede eliminar este gasto de importación", vbInformation, "Eliminación"
    Else
    ' Si no existen detalles con ese gasto, se procede a eliminar
        strSQL = " DELETE " & _
                 " FROM gasto_importacion " & _
                 " WHERE gas_imp_codigo='" & dcmbCodigo.Text & "'"
        clsCon_Def.Ejecutar (strSQL), "M"
        MsgBox "Gasto de importación eliminado", vbInformation, "Eliminación"
    End If
    
    ' Consulta para actualizar los combobox
 
    strSQL = " SELECT gas_imp_codigo,gas_imp_nombre,gas_imp_descripcion,gas_imp_tipo," & _
             "gas_imp_valor,gas_imp_porcentaje,gas_imp_prorra_a" & _
             " FROM gasto_importacion " & _
             " ORDER BY gas_imp_codigo"
        
        clsCon_Def.Ejecutar (strSQL)
        
        'Muestra los datos en los combobox
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "gas_imp_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "gas_imp_nombre"
        dcmbNombre.BoundColumn = "gas_imp_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("gas_imp_codigo")
        End If
    End If

End Sub

Private Sub cmdModificar_Click()

' Modifica los datos del gasto seleccionado, se manda a la variable Tag del formulario una bandera para
' indicar que se va a modificar el gasto, además se envia como datos a la forma
' frmGastoImportacion el código y el nombre
    Dim intPos As Integer
    'Verifica si se ha seleccionado un gasto de importacion para ser modificado
    If dcmbCodigo = "" And dcmbNombre = "" Then
        MsgBox "Seleccione un gasto de importación", vbInformation, "Gasto de Importación"
        dcmbCodigo.SetFocus
        'cmdModificar.Enabled = False
    Exit Sub
    End If
    frmGastoImportacion.Tag = "M"
    frmGastoImportacion.txtCodigo.Text = Me.dcmbCodigo.Text
    frmGastoImportacion.txtNombre.Text = Me.dcmbNombre.Text
    frmGastoImportacion.txtDescripcion.Text = Me.txtDescripcion
    frmGastoImportacion.cmbCalcula.Text = Me.txtCalcula.Text
    frmGastoImportacion.cmbTipo.Text = Me.txtTipo
    frmGastoImportacion.cmbProrra.Text = Me.txtProrra
    frmGastoImportacion.txtValor.Text = Me.txtValor
    ' Para llenar o vaciar la casila del checkbox.
    If Me.lblValor.Caption = "Porcentaje:" Then
        frmGastoImportacion.chkPorcentaje.Value = Checked
    Else
        frmGastoImportacion.chkPorcentaje.Value = Unchecked
    End If
      
    frmGastoImportacion.Show
End Sub

Private Sub cmdNuevo_Click()
' Ingresa un nuevo gasto de importacion, se manda a la variable Tag
' del formulario una bandera para
' indicar que se va a ingresar un nuevo gasto
    frmGastoImportacion.Tag = "N"
    frmGastoImportacion.Show
End Sub

Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Muestra el nombre relacionado con el código del gasto de importacion
' en el momento de seleccionar uno del combobox
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        clsCon_Def.adorec_Def.MoveFirst
    End If
    clsCon_Def.adorec_Def.Find "gas_imp_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsCon_Def.adorec_Def.EOF = True Then
        dcmbNombre = ""
        dcmbNombre.BoundText = ""
        txtDescripcion.Text = ""
        txtTipo.Text = ""
        txtValor.Text = ""
        txtProrra.Text = ""
        txtCalcula.Text = ""
        lblPorcentaje.Caption = ""
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        dcmbNombre = clsCon_Def.adorec_Def("gas_imp_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        txtDescripcion = clsCon_Def.adorec_Def("gas_imp_descripcion")
        txtCalcula.Text = clsCon_Def.adorec_Def("calcula")
        ' escribe la palabra correspondiente a la variable char de gas_imp_tipo
            If clsCon_Def.adorec_Def("gas_imp_tipo") = "S" Then
            txtTipo.Text = "SEGURO"
            ElseIf clsCon_Def.adorec_Def("gas_imp_tipo") = "F" Then
            txtTipo.Text = "FLETE"
            ElseIf clsCon_Def.adorec_Def("gas_imp_tipo") = "O" Then
            txtTipo.Text = "OTRO"
            Else
            txtTipo.Text = clsCon_Def.adorec_Def("gas_imp_tipo")
            End If
        ' escribe la palabra correspondiente a la variable char de gas_imp_prorra_a
            If clsCon_Def.adorec_Def("gas_imp_prorra_a") = "C" Then
            txtProrra.Text = "FOB"
            ElseIf clsCon_Def.adorec_Def("gas_imp_prorra_a") = "P" Then
            txtProrra.Text = "PESO"
            Else
            txtProrra.Text = clsCon_Def.adorec_Def("gas_imp_prorra_a")
            End If
        
        txtValor.Text = clsCon_Def.adorec_Def("gas_imp_valor")
        If clsCon_Def.adorec_Def("gas_imp_porcentaje") = 1 Then
        lblValor.Caption = "Porcentaje:"
        Else
        lblValor.Caption = "Valor:"
        End If
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

Private Sub txtValor_Change()
    ' Pone los decimales en el txt de Valor
    txtValor.Text = FormatoD2(txtValor.Text)
End Sub

Private Sub Form_Activate()
 
    'Muestra la lista de datos actualizada
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "gas_imp_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "gas_imp_nombre"
    dcmbNombre.BoundColumn = "gas_imp_codigo"
    If Me.Tag <> "" Then
        dcmbCodigo = ""
        dcmbCodigo = Me.Tag
    ElseIf Not clsCon_Def.adorec_Def.EOF Then
        dcmbCodigo_Change
    End If
End Sub

Private Sub Form_Load()
 Dim strSQL As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta los documentos que estan disponibles
        strSQL = " SELECT gas_imp_codigo,gas_imp_nombre,gas_imp_descripcion,iif(gas_imp_calcula_a='C','CIF','FOB') as calcula," & _
                 " gas_imp_tipo,gas_imp_prorra_a,gas_imp_valor,gas_imp_porcentaje" & _
                 " FROM gasto_importacion " & _
                 " ORDER BY gas_imp_codigo"
        
        clsCon_Def.Ejecutar (strSQL)
      
        'Muestra los datos de cada gasto de importacion en los combobox
        
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "gas_imp_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "gas_imp_nombre"
        dcmbNombre.BoundColumn = "gas_imp_codigo"
        If Not clsCon_Def.adorec_Def.EOF Then
            dcmbCodigo = clsCon_Def.adorec_Def("gas_imp_codigo")
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

