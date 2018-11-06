VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelMovEgreso 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta. Contable para Movimientos Egreso"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelMovEgreso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   3750
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
      Left            =   128
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
         Left            =   1440
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1920
         Width           =   1920
      End
      Begin VB.CheckBox chkIVA 
         BackColor       =   &H00DDDDDD&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   2640
         Width           =   315
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   690
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.Contable"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servicios:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
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
         Left            =   240
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
         Left            =   240
         TabIndex        =   13
         Top             =   2670
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Productos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
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
         Left            =   240
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
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   308
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1988
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   308
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1988
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelMovEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de movimientos de egreso de mecaderia   #
'#  y poder modificar la cuenta contable de estos, asi como si tienen iva       #
'#  frmSelMovEgreso V1.0                                                        #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los movimientos de egreso de mercaderia del sistema  #
'#  y conocer la cuenta contable de estos y si generan IVA.                     #
'#  Se podrá modificar lo que tiene que ver a la cuenta contable y al IVA       #
'#  Desde esta ventana se llama a la ventana frmMovEgreso en la que se          #
'#  modifica estos datos.                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  tipo_egreso :En esta tabla se almacenan los tipos de egresos que el sistema #
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
 ' Consulta para conocer si existe el tipo de egreso asignado a un egresos
    strSql = " SELECT count(egr_codigo) As Egr " & _
             " FROM egreso " & _
             " WHERE tip_egr_codigo='" & dcmbCodigo.Text & "'" & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen tipos de egresos asignados a egresos no se elimina
    If clsCon_Def.adorec_Def("Egr") > 0 Then
        MsgBox "No Puede eliminar este tipo de egreso", vbInformation, "Eliminación"
    Else ' Si no existen tipos de egresos asignados a egresos se elimina
        strSql = " DELETE " & _
                 " FROM tipo_egreso " & _
                 " WHERE tip_egr_codigo='" & dcmbCodigo.Text & "'" & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Tipo de Egreso eliminado", vbInformation, "Eliminación"

    End If
    ' Consulta para actualizar los combos
    strSql = " SELECT tip_egr_codigo,tip_egr_nombre,tip_egr_descripcion,tip_egr_ctaconta,tip_egr_impuesto " & _
             " FROM tipo_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY tip_egr_nombre "
    clsCon_Def.Ejecutar (strSql)
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.ListField = "tip_egr_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_egr_nombre"
    dcmbNombre.BoundColumn = "tip_egr_codigo"
    dcmbCodigo.Text = ""
    
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos del egreso, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del movimiento que se modificará
    frmMovEgreso.Show
    frmMovEgreso.txtCodigo.Text = Me.dcmbCodigo.Text
    frmMovEgreso.txtNombre.Text = Me.dcmbNombre.Text
    frmMovEgreso.dcmbCtaConta.Text = Me.txtCtaContable.Text
    frmMovEgreso.dcmbCtaConta2.Text = Me.txtCtaContable2.Text
    frmMovEgreso.txtDescripcion.Text = Me.txtDescripcion.Text
    frmMovEgreso.chkIVA.value = Me.chkIVA.value
    frmMovEgreso.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
'NO VISIBLE....NO ES NECESARIO ESTO
' Crea un nuevo movimiento se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará una nueva lista
    frmMovEgreso.Show
    frmMovEgreso.Tag = "N"
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
        strComparar = "tip_egr_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("tip_egr_nombre")
            strAux = clsCon_Def.adorec_Def("tip_egr_ctaconta")
            dcmbNombre.BoundText = dcmbCodigo.Text
            If strAux = "0" Then
                txtCtaContable.Text = ""
            Else
                txtCtaContable.Text = strAux
            End If
            strAux = clsCon_Def.adorec_Def("tip_egr_ctaconta2")
            dcmbNombre.BoundText = dcmbCodigo.Text
            If strAux = "0" Then
                txtCtaContable2.Text = ""
            Else
                txtCtaContable2.Text = strAux
            End If
            txtDescripcion.Text = clsCon_Def.adorec_Def("tip_egr_descripcion")
            chkIVA.value = clsCon_Def.adorec_Def("tip_egr_impuesto")
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
    dcmbCodigo.ListField = "tip_egr_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "tip_egr_nombre"
    dcmbNombre.BoundColumn = "tip_egr_codigo"
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
        strSql = " SELECT tip_egr_codigo,tip_egr_nombre,tip_egr_descripcion,tip_egr_ctaconta,COALESCE(tip_egr_ctaconta2,'-') as tip_egr_ctaconta2,tip_egr_impuesto " & _
                 " FROM tipo_egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tip_egr_nombre "
        clsCon_Def.Ejecutar (strSql)
        Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodigo.ListField = "tip_egr_codigo"
        Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbNombre.ListField = "tip_egr_nombre"
        dcmbNombre.BoundColumn = "tip_egr_codigo"
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
