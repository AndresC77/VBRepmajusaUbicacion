VERSION 5.00
Begin VB.Form frmGastoImportacion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de datos de Gasto de Importación"
   ClientHeight    =   4245
   ClientLeft      =   5970
   ClientTop       =   4650
   ClientWidth     =   3825
   Icon            =   "frmGastoImportacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   3825
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Gasto de Importación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1305
         TabIndex        =   1
         Top             =   720
         Width           =   2040
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   2040
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   795
         Left            =   1305
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2040
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmGastoImportacion.frx":030A
         Left            =   1290
         List            =   "frmGastoImportacion.frx":0317
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cmbProrra 
         Height          =   315
         ItemData        =   "frmGastoImportacion.frx":0330
         Left            =   1290
         List            =   "frmGastoImportacion.frx":033A
         TabIndex        =   5
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   3000
         Width           =   870
      End
      Begin VB.CheckBox chkPorcentaje 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Porcentaje"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   3030
         Width           =   1095
      End
      Begin VB.ComboBox cmbCalcula 
         Height          =   315
         ItemData        =   "frmGastoImportacion.frx":0349
         Left            =   1290
         List            =   "frmGastoImportacion.frx":0353
         TabIndex        =   4
         Top             =   2280
         Width           =   2055
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
         Left            =   135
         TabIndex        =   17
         Top             =   765
         Width           =   600
      End
      Begin VB.Label lblCodio 
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
         Left            =   135
         TabIndex        =   16
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   1080
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
         Left            =   135
         TabIndex        =   14
         Top             =   1965
         Width           =   345
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
         Left            =   135
         TabIndex        =   13
         Top             =   2685
         Width           =   855
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
         Left            =   135
         TabIndex        =   12
         Top             =   3045
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calcula Según:"
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
         TabIndex        =   11
         Top             =   2325
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1965
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   405
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmGastoImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################################'
'#  Forma para el ingreso y modificación de Gastos de Importación                    #
'#  frmGastoImportación V1.0                                                         #
'#  Copyright (C) 2002                                                               #
'#                                                                                   #
'#  Ventana para el ingreso y modificación de Gastos de Importación.                 #
'#  Permitirá almacenar en la base de datos nuevos gastos y modificar sus            #
'#  nombres, dependiendo de la propiedad Tag, la cual se cambiará en la              #
'#  ventana frmSelGastoImportacion y desde esta se llamará a esta ventana.           #
'#                                                                                   #
'#  Tablas que se maneja:                                                            #
'#    gasto_importcion: En esta tabla se almacenan los nuevos gastos y se modifican  #
'#                      los datos de estos.                                          #
'#                                                                                   #
'#  Procedimientos INTERNOS:                                                         #
'#  Procedimientos EXTERNOS:                                                         #
'#                                                                                   #
'#  Objetos de la forma:                                                             #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos               #
'#                                                                                   #
'#                                                                                   #
'#####################################################################################
'/**********************************************************************************/'

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

Private Sub cmbProrra_Validate(Cancel As Boolean)
' verifica si ingresa valores erroneos en el comobobox
If cmbProrra.Text = "PESO" Or cmbProrra.Text = "FOB" Then
 Cancel = False
Else
  MsgBox "Opcion no permitida", vbOKOnly, "Error"
  Cancel = True
End If
End Sub

Private Sub cmbTipo_Validate(Cancel As Boolean)

' verifica si ingresa valores erroneos en el comobobox
If cmbTipo.Text = "SEGURO" Or cmbTipo.Text = "FLETE" Or cmbTipo.Text = "OTRO" Then
 Cancel = False
Else
  MsgBox "Opcion no permitida", vbOKOnly, "Error"
  Cancel = True
End If

End Sub

Private Sub cmdAceptar_Click()
 Dim strSql As String
 Dim t As String * 1 ' para insertar en la BDD el char correspondiente a gas_imp_tipo
 Dim p As String * 1 ' para insertar en la BDD el char correspondiente a gas_imp_prorro_a
 Dim strCalc As String * 1
  
    If cmbTipo.Text = "SEGURO" Then
    t = "S"
    ElseIf cmbTipo.Text = "FLETE" Then
    t = "F"
    Else
    t = "-"
    End If
  
    If cmbProrra.Text = "FOB" Then
    p = "C"
    Else
    p = "P"
    End If
    If cmbCalcula.Text = "CIF" Then
        strCalc = "C"
    Else
        strCalc = "F"
    End If
             
    ' Si la variable Tag es N ingresa un nuevo Gasto de Importacion
    If Me.Tag = "N" Then
    
    'Consulta para ingresar los datos a la base de datos
   
    strSql = " INSERT INTO gasto_importacion(gas_imp_codigo,gas_imp_nombre," & _
             " gas_imp_descripcion,gas_imp_calcula_a,gas_imp_tipo,gas_imp_prorra_a," & _
             " gas_imp_valor,gas_imp_porcentaje,gas_imp_fechamod,gas_imp_usumod) " & _
             " VALUES ('" & UCase(txtCodigo.Text) & "','" & UCase(txtNombre.Text) & _
             "','" & UCase(txtDescripcion.Text) & "','" & strCalc & "','" & t & "','" & _
             p & "','" & txtValor.Text & "'," & chkPorcentaje.value & "," & _
             " CURRENT_TIMESTAMP, '" & strUsuario & "')"
    
    ' Si la variable Tag es M se modifican los datos del gasto de importacion
    ElseIf Me.Tag = "M" Then
    
    'Consulta para modificar los datos del gasto de imp. seleccionado
  
      strSql = " UPDATE gasto_importacion " & _
               " SET gas_imp_nombre='" & UCase(txtNombre.Text) & "',gas_imp_descripcion='" & _
               UCase(txtDescripcion.Text) & "',gas_imp_tipo='" & t & _
               "',gas_imp_calcula_a='" & strCalc & _
               "',gas_imp_prorra_a='" & p & "',gas_imp_valor='" & txtValor.Text & _
               "',gas_imp_porcentaje=" & chkPorcentaje.value & ",gas_imp_fechamod=CURRENT_TIMESTAMP,gas_imp_usumod='" & strUsuario & "' " & _
               " WHERE gas_imp_codigo='" & txtCodigo.Text & "'"
    End If
    
    
    On Error GoTo errhandler
        clsCon_Def.Ejecutar (strSql), "M"
        Unload Me
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

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar Datos de Gasto de Importación"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nuevo Gasto de Importación"
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

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub txtCodigo_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtNombre_Change()
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
' Verifica si el dato ingresado es numérico
    If IsNumeric(txtValor.Text) = False Then
        MsgBox "Solo se permiten valores numéricos", vbOKOnly + vbInformation, "ERROR"
        Cancel = True
    Else
        ' Pone dos decimales al valor
        txtValor.Text = FormatoD2(txtValor.Text)
        Cancel = False
    End If
End Sub
