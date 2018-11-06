VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNuevaCuenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Cuenta"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNuevaCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   4470
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle Cuenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   128
      TabIndex        =   12
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox TxtCuenta 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox ChkPyG 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pérdidas Y Ganancias"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código SubCuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   532
         Width           =   1380
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc SubCuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   892
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   368
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton OptCuenta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cuenta"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptSubCuenta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "SubCuenta"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
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
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2288
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   728
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmNuevaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para agragar una nueva cuenta.                                        #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Esta ventana nos permite agregar una cuenta o subcuenta con su respectiva   #
'#  descripción, a una empresa determinada con anterioridad.                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  - ctaconta:  Utilizada para extraer los códigos de las cuentas ya ingresdas #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#  - clsConsu:                                                                 #
'#        Objeto de consulta a la tabla apara extraer los códigos máximos       #
'#        de ciertas cuentas seleccionadas por el usuario.                      #
'#  - clsInsCuenta:                                                             #
'#        Objeto que ejecuta SQl de inserción y actualización de registros en   #
'#        la tabla de cuentas.                                                  #
'#                                                                              #
'#  Consideracionesdes:                                                         #
'#  - Esta ventana está en la capacidad de sugerir un código para una nueva     #
'#    cuenta o subcuenta, para lo cual se realizan consultas sobre la tabla     #
'#    ctaconta con el fin de extraer el código máximo relacionado con una       #
'#    cuenta específica.                                                        #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsInsCuenta As New clsConsulta
Private intCtaNivel As Integer
Private InterV As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsInsCuenta = Nothing
End Sub

Private Sub ChkPyG_Click()
    If ChkPyG.value = 1 Then
        InterV = "PYG"
    Else
        InterV = "G"
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim SqlStr As String, SqlActCta As String
    'Verifica que la descripción de la cuenta no esté en blanco
    If txtCuenta = "" Then
        MsgBox "Ingrese la descripción de la cuenta.", vbInformation, "Descripción"
        Exit Sub
    End If
    'Verifica que tipo de cuenta se quiere crear
    'Para nueva cuenta
    If OptCuenta = True Then
        'Verifica que se ingrese solo un número entero para la cuenta
        If (Not IsNumeric(MskCodigo)) Or (InStr(MskCodigo, ".") <> 0) Then
            MsgBox "Código de cuenta no válido", vbInformation, "Código"
            Exit Sub
        End If
        'SQl para insertar una nueva cuenta en la base de datos
        SqlStr = "Insert Into ctaconta " & _
                "(cta_codigo, emp_codigo, cta_nombre, cta_nivel, cta_interviene, cta_debe, cta_haber, cta_fechamod, cta_usumod) " & _
                "values ('" & UCase(MskCodigo) & "','" & strEmpresa & "','" & UCase(txtCuenta) & "',1,'" & InterV & "',0,0,CURRENT_TIMESTAMP,'" & strUsuario & "')"
        SqlActCta = "Update ctaconta Set cta_subcta = 0 " & _
                    "Where cta_codigo='" & UCase(MskCodigo.FormattedText) & "' AND emp_codigo='" & strEmpresa & "'"
    End If
    'Para nueva subcuenta
    If OptSubCuenta = True Then
        If Not (Trim(MskCodigo.FormattedText) Like (dcmbCodigo & ".#*")) Then
            MsgBox "Código de Subcuenta no válido.", vbInformation, "Subcuenta"
            Exit Sub
        End If
        'SQl para insertar una nueva subcuenta en la base de datos
        SqlStr = "Insert Into ctaconta " & _
                "(cta_codigo, emp_codigo, cta_nombre, cta_nivel, cta_interviene, cta_debe, cta_haber, cta_fechamod, cta_usumod) " & _
                "values ('" & Trim(UCase(MskCodigo.FormattedText)) & "','" & strEmpresa & "','" & UCase(txtCuenta) & "'," & intCtaNivel + 1 & ",'" & InterV & "',0,0,CURRENT_TIMESTAMP,'" & strUsuario & "')"
        SqlActCta = "Update ctaconta Set cta_subcta = 1 " & _
                    "Where cta_codigo='" & UCase(dcmbCodigo) & "' AND emp_codigo='" & strEmpresa & "'"
    End If
    clsInsCuenta.Inicializar AdoConn, AdoConnMaster
    'Verifica que la cuenta que se quiere insertar no exista previamente
    clsInsCuenta.Ejecutar "Select * from ctaconta where cta_codigo = '" & Trim(UCase(MskCodigo.FormattedText)) & "' and emp_codigo = '" & strEmpresa & "';"
    If Not clsInsCuenta.adorec_Def.EOF Then
        MsgBox "La cuenta ya existe en esta empresa!!, ingrese otro Código.", vbInformation, "Error Cuenta"
        Exit Sub
    End If
    'Ejecuta la inserción de la cuenta
    clsInsCuenta.Ejecutar SqlStr, "M"
    clsInsCuenta.Ejecutar SqlActCta, "M"
    MsgBox "La cuenta fue ingresada con éxito.", vbInformation, "Nueva Cuenta."
    Unload Me
'    clsConsu.Actualizar
'    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
'    Set dcmbCuenta.RowSource = clsConsu.adorec_Def.DataSource
'    OptCuenta = False
'    OptSubCuenta = False
'    dcmbCodigo.Enabled = True
'    dcmbCuenta.Enabled = True
'    dcmbCodigo = ""
'    dcmbCuenta = ""
'    MskCodigo.Mask = ""
'    MskCodigo = ""
'    TxtCuenta = ""
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
    MskCodigo.Mask = ""
    MskCodigo = ""
    txtCuenta = ""
    OptCuenta = False
    OptSubCuenta = False
    'Muestra el nombre relacionado con el código de la cuenta en el momento de seleccionar otra cuenta del combobox
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "cta_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsConsu.adorec_Def.EOF = True Then
        dcmbCuenta = ""
        dcmbCuenta.BoundText = ""
        intCtaNivel = 0
        InterV = ""
    Else
        dcmbCuenta = clsConsu.adorec_Def("cta_nombre")
        dcmbCuenta.BoundText = dcmbCodigo.Text
        intCtaNivel = clsConsu.adorec_Def("cta_nivel")
        InterV = clsConsu.adorec_Def("cta_interviene")
        OptSubCuenta = True
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "CN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCodigo
    End If
End Sub

Private Sub dcmbCuenta_Change()
    'Muestra el código relacionado con el nombre de la cuenta en el momento de seleccionar otra cuenta del combobox
    If dcmbCodigo.Tag <> "A" Then
        If dcmbCuenta.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbCuenta.BoundText
        End If
    End If
End Sub
Private Sub dcmbCuenta_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbCuenta.BoundText
End Sub

Private Sub dcmbCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbCuenta.BoundText
    End If
End Sub

Private Sub Form_Activate()
    'MskCodigo.SetFocus
    MskCodigo.SelStart = 0
    MskCodigo.SelLength = Len(MskCodigo)
    ChkPyG.value = 0
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
    clsConsu.Inicializar AdoConn, AdoConnMaster
    'Ejecuta un SQL contra la base de datos
    clsConsu.Ejecutar ("select cta_codigo,cta_nombre,cta_nivel,cta_interviene from ctaconta where emp_codigo = '" & strEmpresa & "' order by cta_codigo")
    'Muestra los datos de una columna del resultado del SQL en un data combo
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "cta_codigo"
    Set dcmbCuenta.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCuenta.ListField = "cta_nombre"
    dcmbCuenta.BoundColumn = "cta_codigo"
    OptCuenta = True
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

Private Sub OptCuenta_Click()
    dcmbCodigo.Enabled = False
    dcmbCuenta.Enabled = False
    ChkPyG.Enabled = True
    'Ejecuta un SQL para obtener el mayor código de la tabla de cuentas
    Dim clsMaxCta As New clsConsulta
    clsMaxCta.Inicializar AdoConn, AdoConnMaster
    clsMaxCta.Ejecutar ("Select max(ROUND(cta_codigo,0,1)) from ctaconta where emp_codigo = '" & strEmpresa & "' and cta_nivel=1 GROUP BY emp_codigo")
    txtCuenta = ""
    MskCodigo.Mask = ""
    If IsNull(clsMaxCta.adorec_Def(0)) Then
        MskCodigo = 1
    Else
        MskCodigo = clsMaxCta.adorec_Def(0) + 1
    End If
End Sub

Private Sub OptSubCuenta_Click()
    dcmbCodigo.Enabled = True
    dcmbCuenta.Enabled = True
    'Verifica que se haya seleccionado una cuenta
    If dcmbCodigo = "" Or dcmbCuenta = "" Then
        'MsgBox "Seleccione alguna cuenta.", vbInformation, "Cuenta"
        OptSubCuenta = False
        OptCuenta = False
        Exit Sub
    End If
    
    Dim clsMaxCta As New clsConsulta
    Dim strCodMax As String
    Dim intPosPto As Integer, i As Integer
    Dim strNumSug As String
    Dim strMascara As String
    'Selecciona el código máximo relacionado con una cuenta específica
    clsMaxCta.Inicializar AdoConn, AdoConnMaster
    clsMaxCta.Ejecutar ("select max(FLOOR(RIGHT(cta_codigo,LEN(cta_codigo)-" & 1 + Len(dcmbCodigo) & "))) as n, max(LEN(cta_codigo)-" & 1 + Len(dcmbCodigo) & ") as c from ctaconta where cta_codigo like '" & dcmbCodigo & ".%' and emp_codigo = '" & strEmpresa & "' and cta_nivel=" & intCtaNivel + 1 & " GROUP BY emp_codigo")
    i = 1
    If (Not clsMaxCta.adorec_Def.EOF) Then
        If clsMaxCta.adorec_Def("n") <> "Nulo" Then
            strCodMax = clsMaxCta.adorec_Def("n")
            i = clsMaxCta.adorec_Def("c")
        Else
            strCodMax = dcmbCodigo & ".0"
        End If
    Else
        strCodMax = dcmbCodigo & ".0"
    End If
    'Verifica si la subcuenta inteviene en el reporte de pérdidas y ganancias
    ChkPyG.Enabled = False
    If InterV = "PYG" Then
        ChkPyG.value = 1
    Else
        ChkPyG.value = 0
    End If
    'Se suguiere y muestra un código para la nueva cuenta
    strMascara = Replace(dcmbCodigo, "9", "\9")
    MskCodigo.Mask = strMascara & ".9999"
    intPosPto = InStrRev(strCodMax, ".")
    strNumSug = Format(Mid(strCodMax, intPosPto + 1, Len(strCodMax)) + 1, String(i, "0"))
    MskCodigo.Text = dcmbCodigo.Text & "." & strNumSug
    MskCodigo.SelStart = intPosPto + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
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
