VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmModCuenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Cuenta"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4230
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   128
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.TextBox TxtCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox ChkPyG 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pérdidas Y Ganancias"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   24
         PromptChar      =   "_"
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Cuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Cuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   510
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmModCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para modificar la descripción de una cuenta
'#  Copyright (C) 2002
'#
'#  En esta ventana podemos modificar la descripción realacionada con una
'#  cuenta específica.
'#
'#  Tablas que se maneja:
'#  - ctaconta:
'#      Se registran los cambios realizados en la descripción de la cuenta
'#      así como también el nombre del usuario que hizo la modificación y que
'#      día y a que hora lo hizo.
'#
'#  Objetos de la forma:
'#  - clsModCuenta:
'#      Objeto que ejecuta el SQL contra la tabla de cuentas para actualizar
'#      los datos mensionados anteriormente.
'#
'#  Consideraciones:
'#  - Esta ventana solo verifica que no se ingrese una descripción en blanco
'#    para la cuenta o subcuenta.
'#
'################################################################################
'/****************************************************************************/'

Private clsModCuenta As New clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsModCuenta = Nothing
End Sub

Private Sub cmdAceptar_Click()
    Dim SqlStr As String
    If txtCuenta = "" Then
        MsgBox "Ingrese la descripción de la cuenta.", vbInformation, "Descripción"
        txtCuenta.SetFocus
        Exit Sub
    End If
    'SQl para modificar la descripción de una cuenta.
    SqlStr = "UPDATE ctaconta SET cta_nombre= '" & UCase(txtCuenta) & "'," & _
            "cta_fechamod = CURRENT_TIMESTAMP," & _
            "cta_usumod = '" & strUsuario & "' " & _
            "WHERE cta_codigo='" & MskCodigo & "' and emp_codigo = '" & strEmpresa & "';"
    clsModCuenta.Inicializar AdoConn, AdoConnMaster
    clsModCuenta.Ejecutar SqlStr, "M"
    'Verifica si es una cuenta de nivel 1 para actualizar el reporte a que pertenecen sus subcuentas
    If intNivCta = 1 Then
        If ChkPyG.value = 1 Then
            strEstadoPYG = "PYG"
        Else
            strEstadoPYG = "G"
        End If
        SqlStr = "UPDATE ctaconta SET cta_interviene= '" & strEstadoPYG & "'," & _
        "cta_fechamod = CURRENT_TIMESTAMP," & _
        "cta_usumod = '" & strUsuario & "' " & _
        "WHERE cta_codigo like '" & MskCodigo & "%' and emp_codigo = '" & strEmpresa & "';"
        clsModCuenta.Inicializar AdoConn, AdoConnMaster
        clsModCuenta.Ejecutar SqlStr, "M"
    End If
    MsgBox "La cuenta fue modificada con éxito.", vbInformation, "Cuenta"
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If MskCodigo.Text = "" Or txtCuenta.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
    'Inicializa la clase con la conexión activa a la base de datos
    MskCodigo = strCodCuenta
    MskCodigo.Enabled = False
    txtCuenta = strDescCuenta
    'Muestra a que reporte pertenece la cuenta a modificar
    If strEstadoPYG = "PYG" Then
        ChkPyG.value = 1
    End If
    If intNivCta > 1 Then
        ChkPyG.Enabled = False
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

Private Sub MskCodigo_Change()
    If MskCodigo.Text = "" Or txtCuenta.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub MskCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub TxtCuenta_Change()
    If MskCodigo.Text = "" Or txtCuenta.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub TxtCuenta_GotFocus()
    Seleccionar_Contenido
End Sub
