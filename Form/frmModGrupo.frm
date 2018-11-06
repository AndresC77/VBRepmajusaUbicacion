VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmModGrupo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Grupo"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   5040
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   113
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox TxtGrupo 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   4575
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   315
         Left            =   1440
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
         Caption         =   "Código Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2573
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1013
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmModGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para modificar la descripción y nombre de un grupo (frmModGrupo.frm)  #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  En esta ventana podemos modificar la descripción y el nombre realacionado   #
'#  con un grupo específico.                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  grupo:                                                                      #
'#      Aquí se registran los cambios realizados en la descripción o el nombre  #
'#      del grupo, así como también el nombre del usuario que hizo la modifi_   #
'#      cación y que día y a que hora lo hizo.                                  #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#  clsModCuenta:                                                               #
'#      Objeto que ejecuta el SQL contra la tabla de grupos para actualizar     #
'#      los datos mensionados anteriormente.                                    #
'#                                                                              #
'#  Consideraciones:                                                            #
'#      Esta ventana solo verifica que no se ingrese un nombre en blanco para   #
'#      el grupo o subgrupo.                                                    #
'#                                                                              #
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
    If TxtGrupo = "" Then
        MsgBox "Ingrese el nombre del Grupo.", vbInformation, "Nombre"
        TxtGrupo.SetFocus
        Exit Sub
    End If
    'SQl para modificar la descripción de una cuenta.
    SqlStr = " UPDATE grupo SET gru_nombre= '" & UCase(TxtGrupo) & "', " & _
             " gru_descripcion='" & UCase(txtDescripcion) & "', " & _
             " gru_fechamod = CURRENT_TIMESTAMP, " & _
             " gru_usumod = '" & strUsuario & "' " & _
             " WHERE gru_codigo='" & MskCodigo & "' and emp_codigo = '" & strEmpresa & "'"
    clsModCuenta.Inicializar AdoConn, AdoConnMaster
    clsModCuenta.Ejecutar SqlStr, "M"
    MsgBox "El grupo fue modificado con éxito.", vbInformation, "Grupo"
    frmVerGrupos.Tag = MskCodigo
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
    If MskCodigo.Text = "" Or TxtGrupo.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
    'Inicializa la clase con la conexión activa a la base de datos
    MskCodigo.Enabled = False
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
    If MskCodigo.Text = "" Or TxtGrupo.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub MskCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub TxtGrupo_Change()
    If MskCodigo.Text = "" Or TxtGrupo.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub TxtGrupo_GotFocus()
    Seleccionar_Contenido
End Sub
