VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerRetencion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Retenciones"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerRetencion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7485
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Retenciones"
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
      Left            =   135
      TabIndex        =   9
      Top             =   120
      Width           =   7215
      Begin VB.TextBox TxtCtaConta 
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox TxtGravarA 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   5895
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   4800
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   540
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   15
         Top             =   420
         Width           =   600
      End
      Begin VB.Label LblPorcentaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   765
         Width           =   810
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1485
         Width           =   900
      End
      Begin VB.Label LblCtaContable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Contable:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   12
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label LblGravarA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravar a:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3000
         TabIndex        =   10
         Top             =   765
         Width           =   150
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4372
      TabIndex        =   8
      Top             =   2160
      Width           =   1080
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nueva"
      Height          =   405
      Left            =   2032
      TabIndex        =   6
      Top             =   2160
      Width           =   1020
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3112
      TabIndex        =   7
      Top             =   2160
      Width           =   1080
   End
End
Attribute VB_Name = "frmVerRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para ver las retenciones que puede manejar una empresa                #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Esta ventana muestra todos los tipos de retenciones que tiene una           #
'#  determinada empresa.                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  retencion:                                                                  #
'#      * Tabla que contiene toda la información necesaria de todas las         #
'#        posibles retenciones que maneja una determinada empresa.              #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#      * Se puede llamar al formulario que nos permite modificar o ingresar    #
'#        los datos para una retención.                                         #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#  clsConsu  - Objeto para consultar a la base de datos todas las posiles      #
'#              retenciones que maneja una empresa y desplegarlas en un         #
'#              combobox tanto su código como su nombre repectivamente.         #
'#                                                                              #
'#                                                                              #
'################################################################################
'/*****************************************************************************/'

Private clsConsu As New clsConsulta
Private strSql As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
End Sub

Private Sub cmdModificar_Click()
    'Verifica que se haya seleccionado un depósito para poder modificarla
    If dcmbCodigo = "" Then
        MsgBox "Seleccione un depósito primero.", vbInformation, "Depósito"
        'dcmbCodigo.SetFocus
        Exit Sub
    End If
    'Llama a la forma depósito con la opción de modificar depósito
    frmRetencion.Tag = "Mod"
    frmRetencion.txtCodigo = Me.dcmbCodigo
    frmRetencion.txtNombre = Me.dcmbNombre
    frmRetencion.dcmbCtaConta = Me.txtCtaConta
    frmRetencion.CmbGravarA = Me.TxtGravarA
    frmRetencion.txtDescripcion = Me.txtDescripcion
    frmRetencion.TxtPorcentaje = Me.TxtPorcentaje
    frmRetencion.Show
End Sub

Private Sub cmdNuevo_Click()
    'Llama a la forma depósito con la opción de nuevo depósito
    frmRetencion.Tag = "New"
    frmRetencion.Show
End Sub

Private Sub CmdSalir_Click()
    'Cierra el formulario actual
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
    'Muestra el nombre relacionado con el código del depósito en el momento de seleccionar uno del combobox
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "ret_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsConsu.adorec_Def.EOF = True Then
        dcmbNombre = ""
        dcmbNombre.BoundText = ""
    Else
        dcmbNombre = clsConsu.adorec_Def("ret_nombre")
        dcmbNombre.BoundText = dcmbCodigo.Text
        txtDescripcion = clsConsu.adorec_Def("ret_descripcion")
        TxtGravarA = clsConsu.adorec_Def("ret_gravara")
        TxtPorcentaje = clsConsu.adorec_Def("ret_porcentaje")
        txtCtaConta = clsConsu.adorec_Def("ret_ctaconta")
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
    Dim aux As String
    'Muestra la lista de depósitos actualizada
    clsConsu.Actualizar
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    Set dcmbNombre.RowSource = clsConsu.adorec_Def.DataSource
    If Me.Tag <> "" Then
        dcmbCodigo = ""
        dcmbCodigo = Me.Tag
    ElseIf Not clsConsu.adorec_Def.EOF Then
        dcmbCodigo_Change
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn
    'Ejecuta un SQL contra la base de datos
    strSql = " SELECT ret_codigo,ret_nombre,ret_porcentaje,ret_ctaconta,ret_descripcion,ret_gravara " & _
             " From retencion " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY ret_codigo"
    clsConsu.Ejecutar strSql
    'Muestra los datos de los códigos del depósito en un datacombo
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "ret_codigo"
    Set dcmbNombre.RowSource = clsConsu.adorec_Def.DataSource
    dcmbNombre.ListField = "ret_nombre"
    dcmbNombre.BoundColumn = "ret_codigo"
    'Muestra el primer registro de la consulta
    If Not clsConsu.adorec_Def.EOF Then
        dcmbCodigo = clsConsu.adorec_Def(0)
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
