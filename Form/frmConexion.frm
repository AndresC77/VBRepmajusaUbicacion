VERSION 5.00
Begin VB.Form frmConexion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NEED - Login"
   ClientHeight    =   2640
   ClientLeft      =   11325
   ClientTop       =   4665
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConexion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtClave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   945
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1815
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   945
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   450
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdConfigurar 
      Caption         =   "C&onfigurar Conexión"
      Height          =   375
      Left            =   142
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1627
      TabIndex        =   3
      Top             =   1470
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1470
      Width           =   1455
   End
End
Attribute VB_Name = "frmConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de Conexion a la base de datos                                        #
'#  frmConexion V1.0                                                            #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana inicial, en la que se ingresará el usuario y la clave de acceso     #
'#  a la base de datos, con la que se habre la conexión a esta a traves del     #
'#  objeto adoconn,adoconnmaster que es un objeto ADODB.Connection el cual es   #
'#  PUBLICO y se encuentra en el modulo modconexion.                            #
'#  Esta ventana al conseguir la conexion con la base de datos abrirá la        #
'#  la ventana mdiPrincipal que es la ventana principal del sistema.            #
'#  Adicionalmente abre la ventana frmEmpresa, ventana para la selección de la  #
'#  empresa en la que se trabajara.                                             #
'#                                                                              #
'#                                                                              #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'


Dim clsRSComp As New clsRegSet

Public Sub cmdAceptar_Click()
    
    On Error GoTo errhandler
        Set AdoConn = New ADODB.Connection
        Set AdoConnMaster = New ADODB.Connection
        txtUsuario.Text = UCase(txtUsuario.Text)
        strClave = txtClave.Text
        strUsuario = txtUsuario.Text
        'Set adoC_Def = New Adodc
        ' Cadena de conexión a la base de datos, esta esta para el uso de MyODBC
'        AdoConn.ConnectionString = "driver={SQLOLEDB.1};" _
'                                 & "server=" & strServidorBDDLocal & ";" _
'                                 & "uid=" & txtUsuario.Text & ";" _
'                                 & "pwd=" & txtClave.Text & ";" _
'                                 & "database=" & strBDD
        Conectar
        
        If Trim(s_instanciaSQL) <> "" Then
        
            cadena_conexion = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & LCase(s_userSql) & ";" _
                           & "Password=" & LCase(s_passwordSql) & ";" _
                           & "Initial Catalog=" & s_catalogoSQL & ";" _
                           & "Data Source=" & s_instanciaSQL
        End If
        ' Comprueba si se logro la conexion a la base de datos
        'MsgBox "Estado de acoConn: " & GetState(adoconn.State)
        ' Almacena el usuario y la clave para enviar estos datos al browser
        ' Muestra la ventana principal y la ventana para seleccionar la empresa
        mdiPrincipal.Show
        frmSelEmpresa.Show
        Unload Me
        Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub cmdcancelar_Click()
    Dim clsRSComp As New clsRegSet
    clsRSComp.SetRegionalSetting strForNumDec, strForNumMil, strForMonDec, strForMonMil, strForFecha
    Unload Me
End Sub

Public Function GetState(intState As Integer) As String
    Select Case intState
        Case adStateClosed
            GetState = "adStateClosed"
        Case adStateOpen
            GetState = "adStateOpen"
    End Select
End Function

Private Sub cmdConfigurar_Click()
    Unload Me
    frmConfig.Show
End Sub

Private Sub Form_Load()
    Dim strLinea As String
    Dim strVar As String
    Dim strVal As String
    Dim strPath As String
    Dim i As Integer
    
    LeerImpresoras
    LeerBalanza
    
    strServidorBDD = "prog2"
    strBDD = "sisadmin"
    strServidorWeb = "prog2"
    strPath = Trim(App.Path)
    If Right(strpad, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    If ImpresoraEtiqueta <> "" And ImpresoraTicket <> "" And ImpresoraPorDefecto <> "" Then
        If Dir(strPath & "config") <> "" Then
            Open strPath & "config" For Input As #1
                For i = 0 To 12
                    Input #1, strLinea
                    strVar = LCase(Trim(Left(strLinea, InStr(strLinea, "=") - 1)))
                    strVal = Trim(Right(strLinea, Len(strLinea) - InStr(strLinea, "=")))
                    If Len(strVal) <> 0 Then
                        Select Case strVar
                            Case "servidorbddmaster"
                                strServidorBDDMaster = strVal
                            Case "puertomaster"
                                strPuertoMaster = strVal
                            Case "servidorbddlocal"
                                strServidorBDDLocal = strVal
                            Case "puertolocal"
                                strPuertoLocal = strVal
                            Case "bdd"
                                strBDD = strVal
                            Case "servidorweb"
                                strServidorWeb = strVal
                            Case "ptofactura"
                                'strPtoFactura = strVal
                                strPtoFacturaOriginal = strVal
                            Case "autorfactura"
                                strAutorFactura = strVal
                            Case "caducafactura"
                                strCaducaFactura = strVal
                            Case "mssql"
                                s_instanciaSQL = strVal
                            Case "mssqlcatalogo"
                                s_catalogoSQL = strVal
                            Case "mssqlusuario"
                                s_userSql = strVal
                            Case "mssqlclave"
                                s_passwordSql = strVal
        
                        End Select
                    End If
                    If EOF(1) Then
                        Exit For
                    End If
                Next
            Close #1
        Else
            MsgBox "Archivo de configuracion"
            frmConfig.Show
            Unload Me
        End If
    Else
        MsgBox "Impresoras no configuradas"
        frmConfig.Show
        Unload Me
    End If
    
    strForNumDec = clsRSComp.DecimalSymbol
    strForMonDec = clsRSComp.MonDecimalSymbol
    strForNumMil = clsRSComp.ThousandSeparator
    strForMonMil = clsRSComp.MonThousandSeparator
    strForFecha = clsRSComp.ShortDate
    If clsRSComp.DecimalSymbol <> "." Or clsRSComp.MonDecimalSymbol <> "." Or clsRSComp.ThousandSeparator <> "," Or clsRSComp.MonThousandSeparator <> "," Or clsRSComp.ShortDate <> "yyyy-MM-dd" Then
        If MsgBox("La configuración regional esta mal seleccionada" & vbNewLine & _
               "Símbolo decimal en Número: ' " & clsRSComp.DecimalSymbol & " '" & vbNewLine & _
               "Símbolo de separación de miles en Número: ' " & clsRSComp.ThousandSeparator & " '" & vbNewLine & _
               "Símbolo decimal en Moneda: ' " & clsRSComp.MonDecimalSymbol & " '" & vbNewLine & _
               "Símbolo de separación de miles en Moneda: ' " & clsRSComp.MonThousandSeparator & " '" & vbNewLine & _
               "Formato fecha corta: ' " & clsRSComp.ShortDate & " '" & vbNewLine & vbNewLine & _
               "Para que el sistema funcione adecuadamente hay que cambiar a:" & vbNewLine & _
               "Símbolo decimal en Número: ' . '" & vbNewLine & _
               "Símbolo de separación de miles en Número: ' , '" & vbNewLine & _
               "Símbolo decimal en Moneda: ' . '" & vbNewLine & _
               "Símbolo de separación de miles en Moneda: ' , '" & vbNewLine & _
               "Formato fecha corta: ' yyyy-MM-dd '" & vbNewLine & vbNewLine & _
               "¿Desea Cambiarlo ahora?", vbCritical + vbDefaultButton1 + vbYesNo, "Configuración Regional") = vbYes Then
            clsRSComp.SetRegionalSetting ".", ",", ".", ",", "yyyy-MM-dd"
        Else
            MsgBox "El sistema se cerrará, hasta que la fonfiguración regional sea cambiada"
            Unload Me
        End If
    End If
    Path = App.Path
    PathCpp = Path & "\Programa\HuellaDigital.exe"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtClave_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtUsuario_GotFocus()
    Seleccionar_Contenido
End Sub


