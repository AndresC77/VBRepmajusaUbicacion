VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNuevoGrupo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Grupo"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNuevoGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle de Grupo:"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6015
      Begin VB.TextBox TxtGrupo 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1320
         Width           =   5775
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   360
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
         Caption         =   "Código Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   405
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
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Grupo:"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton OptGrupo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Grupo:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptSubGrupo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "SubGrupo:"
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
      Begin MSDataListLib.DataCombo dcmbGrupo 
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
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1140
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3180
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1620
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "frmNuevoGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para agregar una nuevo grupo de productos (frmNuevoGrupo.frm)         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Esta ventana nos permite agregar un grupo o subgrupo de productos con su    #
'#  respectivo nombre y descripción, a una empresa determinada con anterioridad.#
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  grupo:                                                                      #
'#      Utilizada para extraer los códigos de los grupos ya ingresdos en una    #
'#      empresa.                                                                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  1.  Esta ventana está en la capacidad de sugerir un código para un grupo o  #
'#      subgrupo, para lo cual se realizan consultas sobre la tabla ctaconta    #
'#      con el fin de extraer el código máximo relacionado con un grupo         #
'#      específico.                                                             #
'#  2.  El usuario no puede cambiar el inicio de un código sugerido por el      #
'#      sistema, solamente su último valor.                                     #
'#  3.  En el momento de guardar el nuevo grupo se verifica que el código       #
'#      seleccionado por el usuario ya no exista en la base de datos para       #
'#      permitir su inserción.                                                  #
'#  4.  Una vez que se inserta un nuevo subgrupo se actualiza el campo subgrupo #
'#      del grupo padre al valor de 1.                                          #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#  clsConsu:                                                                   #
'#        Objeto de consulta a la tabla para extraer los códigos máximos de     #
'#        ciertos grupos seleccionados por el usuario.                          #
'#  clsGrupo:                                                                   #
'#        Objeto que ejecuta SQl de inserción y actualización de registros en   #
'#        la tabla de grupos.                                                   #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsGrupo As New clsConsulta
Private intCtaNivel As Integer 'Variable global que contiene el nivel del grupo padre
Private strSql As String

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsGrupo = Nothing
End Sub

Private Sub cmdAceptar_Click()
    Dim SqlStr As String, SqlActGrupo As String
    'Verifica que la descripción de la cuenta no esté en blanco
    If TxtGrupo = "" Then
        MsgBox "Ingrese la descripción del grupo.", vbInformation, "Descripción"
        Exit Sub
    End If
    'Verifica que tipo de cuenta se quiere crear
    'Para nueva cuenta
    If OptGrupo = True Then
        'Verifica que se ingrese solo un número entero para el código de grupo
        If (Not IsNumeric(MskCodigo)) Or (InStr(MskCodigo, ".") <> 0) Then
            MsgBox "Código de cuenta no válido", vbInformation, "Código"
            Exit Sub
        End If
        'SQl para insertar un nuevo grupo en la base de datos
        SqlStr = " INSERT INTO grupo " & _
                 " (gru_codigo, emp_codigo, gru_nombre, gru_nivel, gru_subgrupo, gru_descripcion, gru_fechamod, gru_usumod) " & _
                 " Values ('" & UCase(MskCodigo) & "','" & strEmpresa & "','" & UCase(TxtGrupo) & "',1,0,'" & UCase(txtDescripcion) & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
        'Sql que actualiza el campo gru_subgrupo para indicar que no se tiene subgrupos relacionados
        SqlActGrupo = " Update grupo Set gru_subgrupo = 0 " & _
                      " Where gru_codigo='" & UCase(MskCodigo.FormattedText) & "' AND emp_codigo='" & strEmpresa & "'"
    End If
    'Para nueva subcuenta
    If OptSubGrupo = True Then
        If Not (Trim(MskCodigo.FormattedText) Like (dcmbCodigo & ".#*")) Then
            MsgBox "Código de Subgrupo no válido.", vbInformation, "Subgrupo"
            Exit Sub
        End If
        'SQl para insertar una nueva subcuenta en la base de datos
        SqlStr = " Insert Into grupo " & _
                 " (gru_codigo, emp_codigo, gru_nombre, gru_nivel, gru_subgrupo, gru_descripcion, gru_fechamod, gru_usumod) " & _
                 " Values ('" & Trim(UCase(MskCodigo.FormattedText)) & "','" & strEmpresa & "','" & UCase(TxtGrupo) & "'," & intCtaNivel + 1 & ",0,'" & UCase(txtDescripcion) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        SqlActGrupo = " Update grupo Set gru_subgrupo = 1 " & _
                      " Where gru_codigo='" & UCase(dcmbCodigo) & "' AND emp_codigo='" & strEmpresa & "'"
    End If
    clsGrupo.Inicializar AdoConn, AdoConnMaster
    'Verifica que la cuenta que se quiere insertar no exista previamente
    clsGrupo.Ejecutar "Select * From grupo where gru_codigo = '" & Trim(MskCodigo.FormattedText) & "' and emp_codigo = '" & strEmpresa & "';"
    If Not clsGrupo.adorec_Def.EOF Then
        MsgBox "El Grupo ya existe en esta empresa!!, ingrese otro Código.", vbInformation, "Error Grupo"
        Exit Sub
    End If
    'Ejecuta la inserción de la cuenta
    clsGrupo.Ejecutar SqlStr, "M"
    'Ejecuta la actualización de nivel superior
    clsGrupo.Ejecutar SqlActGrupo, "M"
    MsgBox "El grupo fue ingresado con éxito.", vbInformation, "Nuevo Grupo."
    'Manda el código del grupo actualmente ingresado a la forma ver grupos
    frmVerGrupos.Tag = MskCodigo
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
    MskCodigo.Mask = ""
    MskCodigo = ""
    TxtGrupo = ""
    OptGrupo = False
    OptSubGrupo = False
    'Muestra el nombre relacionado con el código de la cuenta en el momento de seleccionar otra cuenta del combobox
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "gru_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsConsu.adorec_Def.EOF = True Then
        dcmbGrupo = ""
        dcmbGrupo.BoundText = ""
        intCtaNivel = 0
        InterV = ""
    Else
        dcmbGrupo = clsConsu.adorec_Def("gru_nombre")
        dcmbGrupo.BoundText = dcmbCodigo.Text
        intCtaNivel = clsConsu.adorec_Def("gru_nivel")
        OptSubGrupo = True
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbGrupo_Change()
    'Muestra el código relacionado con el nombre de la cuenta en el momento de seleccionar otra cuenta del combobox
     If dcmbCodigo.Tag <> "A" Then
        If dcmbGrupo.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbGrupo.BoundText
        End If
    End If
    
End Sub
Private Sub dcmbGrupo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbGrupo.BoundText
End Sub

Private Sub dcmbGrupo_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbGrupo.BoundText
    End If
End Sub

Private Sub Form_Activate()
    MskCodigo.SelStart = 0
    MskCodigo.SelLength = Len(MskCodigo)
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
    clsConsu.Inicializar AdoConn, AdoConnMaster
    'Ejecuta un SQL contra la base de datos
    strSql = " Select gru_codigo,gru_nombre,gru_nivel " & _
             " From grupo " & _
             " Where emp_codigo = '" & strEmpresa & "' " & _
             " Order by gru_codigo "
    clsConsu.Ejecutar strSql
    'Muestra los datos de una columna del resultado del SQL en un data combo
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "gru_codigo"
    Set dcmbGrupo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbGrupo.ListField = "gru_nombre"
    dcmbGrupo.BoundColumn = "gru_codigo"
    OptGrupo = True
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
    If MskCodigo.Text = "" Or TxtGrupo.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub MskCodigo_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub OptGrupo_Click()
    dcmbCodigo.Enabled = False
    dcmbGrupo.Enabled = False
    'Ejecuta un SQL para obtener el mayor código de la tabla de cuentas
    Dim clsMaxCta As New clsConsulta
    clsMaxCta.Inicializar AdoConn, AdoConnMaster
    strSql = " Select Max(ROUND(gru_codigo,0,1)) " & _
             " From grupo " & _
             " Where emp_codigo = '" & strEmpresa & "' and gru_nivel=1" & _
             " GROUP BY emp_codigo"
    clsMaxCta.Ejecutar strSql
    TxtGrupo = ""
    MskCodigo.Mask = ""
    If IsNull(clsMaxCta.adorec_Def(0)) Then
        MskCodigo = 1
    Else
        MskCodigo = clsMaxCta.adorec_Def(0) + 1
    End If
End Sub

Private Sub OptSubGrupo_Click()
    dcmbCodigo.Enabled = True
    dcmbGrupo.Enabled = True
    'Verifica que se haya seleccionado una cuenta
    If dcmbCodigo = "" Or dcmbGrupo = "" Then
        MsgBox "Seleccione un grupo.", vbInformation, "Grupo"
        OptSubGrupo = False
        OptGrupo = False
        Exit Sub
    End If
    
    Dim clsMaxCta As New clsConsulta
    Dim strCodMax As String
    Dim intPosPto As Integer, i As Integer
    Dim strNumSug As String
    Dim strMascara As String
    'Selecciona el código máximo relacionado con una cuenta específica
    clsMaxCta.Inicializar AdoConn, AdoConnMaster
    strSql = " Select COALESCE(Max(gru_codigo),'" & dcmbCodigo & ".0') " & _
             " From grupo " & _
             " Where gru_codigo like '" & dcmbCodigo & ".%' " & _
             " AND emp_codigo = '" & strEmpresa & "' AND " & _
             " gru_nivel=" & intCtaNivel + 1 & _
             " GROUP BY emp_codigo"
    clsMaxCta.Ejecutar strSql
    If Not clsMaxCta.adorec_Def.EOF Then
        strCodMax = clsMaxCta.adorec_Def(0)
    Else
        strCodMax = dcmbCodigo & ".0"
    End If
        'Se suguiere y muestra un código para la nueva cuenta
    strMascara = Replace(dcmbCodigo, "9", "\9")
    MskCodigo.Mask = strMascara & ".9999"
    intPosPto = InStrRev(strCodMax, ".")
    strNumSug = Mid(strCodMax, intPosPto + 1, Len(strCodMax)) + 1
    MskCodigo.Text = dcmbCodigo.Text & "." & strNumSug
    MskCodigo.SelStart = intPosPto + 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
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
