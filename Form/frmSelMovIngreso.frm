VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelMovIngreso 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta. Contable para Movimientos Ingreso"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelMovIngreso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cta. Contable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   113
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtCtaContable2 
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
         TabIndex        =   15
         Top             =   2280
         Width           =   1920
      End
      Begin VB.TextBox txtCtaContable 
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1920
      End
      Begin VB.CheckBox chkIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00BAA892&
         Enabled         =   0   'False
         Height          =   210
         Left            =   1320
         TabIndex        =   4
         Top             =   2640
         Width           =   195
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   690
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
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
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.Contable:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servicios:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Productos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1965
         Width           =   780
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimiento:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1913
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   353
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1913
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   353
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelMovIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de movimientos de ingreso de mecaderia  #
'#  y poder modificar la cuenta contable de estos, asi como si tienen iva       #
'#  frmSelMovIngreso V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los movimientos de ingreso de mercaderia del sistema #
'#  y conocer la cuenta contable de estos y si generan IVA.                     #
'#  Se podrá modificar lo que tiene que ver a la cuenta contable y al IVA       #
'#  Desde esta ventana se llama a la ventana frmMovIngreso en la que se         #
'#  modifica estos datos.                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  tipo_ingreso:En esta tabla se almacenan los tipos de ingresos que el sistema#
'#               puede emitir, asi como sus datos.                              #
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
' Consulta para conocer si existe el tipo de ingreso asignado a ingresos
    strSql = " SELECT count(ing_codigo) As Ing " & _
             " FROM ingreso " & _
             " WHERE tip_ing_codigo='" & dcmbCodigo.Text & "'" & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen tipos de ingresos asignados a ingresos no se elimina
    If clsCon_Def.adorec_Def("Ing") > 0 Then
        MsgBox "No Puede eliminar este tipo de ingreso", vbInformation, "Eliminación"
    Else ' Si no existen tipos de ingresos asignados a ingresos se elimina
        strSql = " DELETE " & _
                 " FROM tipo_ingreso " & _
                 " WHERE tip_ing_codigo='" & dcmbCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Tipo de Ingreso eliminado", vbInformation, "Eliminación"
    End If
    ' Consulta para actualizar los combos
    strSql = " SELECT tip_ing_codigo,tip_ing_nombre,tip_ing_descripcion,tip_ing_ctaconta,tip_ing_impuesto " & _
             " FROM tipo_ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY tip_ing_nombre "
    clsCon_Def.Ejecutar (strSql)
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "tip_ing_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_ing_nombre"
    dcmbNombre.BoundColumn = "tip_ing_codigo"
    dcmbCodigo.Text = ""
    
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos del ingreso, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del movimiento que se modificará
    frmMovIngreso.Show
    frmMovIngreso.txtCodigo.Text = Me.dcmbCodigo.Text
    frmMovIngreso.txtNombre.Text = Me.dcmbNombre.Text
    frmMovIngreso.dcmbCtaConta.Text = Me.txtCtaContable.Text
    frmMovIngreso.dcmbCtaConta2.Text = Me.txtCtaContable2.Text
    frmMovIngreso.txtDescripcion.Text = Me.txtDescripcion.Text
    frmMovIngreso.chkIVA.value = Me.chkIVA.value
    frmMovIngreso.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
'NO VISIBLE....NO ES NECESARIO ESTO
' Crea un nuevo movimiento se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará una nueva lista
    frmMovIngreso.Show
    frmMovIngreso.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea el movimiento seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    Dim strAux As String
    On Error GoTo errhandler
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "tip_ing_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("tip_ing_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            strAux = clsCon_Def.adorec_Def("tip_ing_ctaconta")
            If strAux = "0" Then
                txtCtaContable.Text = ""
            Else
                txtCtaContable.Text = strAux
            End If
            strAux = clsCon_Def.adorec_Def("tip_ing_ctaconta2")
            If strAux = "0" Then
                txtCtaContable2.Text = ""
            Else
                txtCtaContable2.Text = strAux
            End If
            txtDescripcion.Text = clsCon_Def.adorec_Def("tip_ing_descripcion")
            chkIVA.value = clsCon_Def.adorec_Def("tip_ing_impuesto")
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            dcmbNombre.Text = ""
            dcmbNombre.BoundText = ""
            txtCtaContable.Text = ""
            txtCtaContable2.Text = ""
            txtDescripcion.Text = ""
            chkIVA.value = 0
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
' Actualiza la lista de movimientos al volver al formulario
    clsCon_Def.Actualizar
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "tip_ing_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_ing_nombre"
    dcmbNombre.BoundColumn = "tip_ing_codigo"
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta los movimientos que estan disponibles
        strSql = " SELECT tip_ing_codigo,tip_ing_nombre,tip_ing_descripcion,tip_ing_ctaconta,COALESCE(tip_ing_ctaconta2,'-') as tip_ing_ctaconta2,tip_ing_impuesto " & _
                 " FROM tipo_ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tip_ing_nombre "
        clsCon_Def.Ejecutar (strSql)
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "tip_ing_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "tip_ing_nombre"
        dcmbNombre.BoundColumn = "tip_ing_codigo"
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
