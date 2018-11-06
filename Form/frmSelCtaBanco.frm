VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmSelCtaBanco 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelCtaBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   8175
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cuentas Bancarias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   1410
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   2880
      End
      Begin VB.TextBox txtReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin VB.TextBox txtDisponible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1920
      End
      Begin VB.TextBox txtCtaConta 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox txtChUltimo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   4920
         TabIndex        =   6
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox txtPrevisto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   1920
      End
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   330
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbBanco 
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
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Width           =   510
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   19
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblReal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Real:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Disponible:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   1500
         Width           =   1230
      End
      Begin VB.Label lblobservaciones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   16
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Cheque:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   14
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label lblprevisto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Previsto:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5700
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4140
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2580
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1020
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelCtaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion la Cuenta Bancaria para poder eliminar, modificar   #
'#  o crear nuevas cuentas                                                      #
'#  frmSelMarca V1.0                                                            #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las cuentas bancarias que tienen los bancos          #
'#  ingresadas en el sistema. Desde esta ventana se puede crear, modificar      #
'#  o eliminar las cuentas bancarias creadas.                                   #
'#  Desde esta ventana se llama a la ventana frmCtaBanco en la que se crea      #
'#  y modifica las cuentas bancarias                                            #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  cta_banco: En esta tabla se almacenan las nuevas cuentas, se                #
'#               modifican y eliminan los datos de las cuentas bancarias        #
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

Private clsCon_Def As New clsConsulta
Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsSql As New clsConsulta
Dim strSql As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsSql = Nothing
End Sub

Private Sub limpiar()
    'limpia los datos de la forma
    txtChUltimo.Tag = ""
    txtDescripcion.Text = ""
    txtCtaConta = ""
    txtChUltimo = ""
    txtReal = 0
    txtDisponible = 0
    txtPrevisto = 0
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    dcmbCuenta = ""
End Sub
Private Sub saldodisponible()
    
    'Calcula el saldo disponible de la cuenta bancaria
    
    strSql = " SELECT  sum(com_egr_ch_valor) as valor, com_egr_ch_fecha" & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND com_egr_ch_estado = 'GIRADO' AND cta_ban_numero = '" & dcmbCuenta.Text & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'  AND com_egr_ch_fecha <= CURRENT_TIMESTAMP " & _
             " GROUP BY cta_ban_numero "
    clsSql.Ejecutar strSql
    
    If Not IsNull(clsSql.adorec_Def("valor")) And clsSql.adorec_Def.EOF = False Then
        Valor = clsSql.adorec_Def("valor")
        disponible = Val(txtReal) - Val(Valor)
        txtDisponible = disponible
    Else
        Valor = 0
        disponible = Val(txtReal) - Val(Valor)
        txtDisponible = disponible
    End If
End Sub

Private Sub cmdEliminar_Click()

    If MsgBox("Desea eliminar la cuenta bancaria?", vbYesNo, "Eliminación") = vbYes Then
          
    ' Consultar las cuentas a eliminar
       strSql = " SELECT count(*) As Egr " & _
                " FROM egreso_comun " & _
                " WHERE cta_ban_numero = '" & dcmbCuenta.Text & "' " & _
                " AND emp_codigo='" & strEmpresa & "'"
       clsCon_Def.Ejecutar (strSql)
          
       ' Si existen existen tablas relacionadas con cuentas bancarias no se elimina
       If clsCon_Def.adorec_Def("egr") > 0 Then
           MsgBox "No Puede eliminar esta Cuenta Bancaria", vbInformation, "Eliminación"
       
       Else ' Si no existen existen cuentas bancarias se elimina
           strSql = " DELETE " & _
                    " FROM cta_banco " & _
                    " WHERE cta_ban_numero='" & dcmbCuenta.Text & "' AND ban_codigo = '" & txtChUltimo.Tag & "'" & _
                    " AND emp_codigo = '" & strEmpresa & "'"
           clsCon_Def.Ejecutar (strSql)
           
           
          MsgBox "Cuenta Bancaria eliminada", vbInformation, "Eliminación"
           
       End If
    End If
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de una cuenta bancaria, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia los datos de la cuenta que se va a modificar
    frmCtaBanco.Show
    frmCtaBanco.txtNumero.Text = dcmbCuenta.Text
    frmCtaBanco.dcmbCodigo.Text = txtChUltimo.Tag
    frmCtaBanco.txtBanco.Text = dcmbBanco.Text
    frmCtaBanco.txtObservaciones.Text = txtDescripcion.Text
    frmCtaBanco.dcmbCuentas.Text = txtCtaConta.Text
    frmCtaBanco.txtUltimo.Text = txtChUltimo.Text
    frmCtaBanco.txtReal.Text = txtReal
    frmCtaBanco.txtDisponible.Text = txtDisponible.Text
    frmCtaBanco.txtPrevisto.Text = txtPrevisto.Text
    frmCtaBanco.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea un nueva cuenta bancaria, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nueva cuenta bancaria
    frmCtaBanco.Show
    frmCtaBanco.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbBanco_Change()
    dcmbCuenta.Text = ""
    'Consulta las cuentas bancarias que estan disponibles
    strSql = " SELECT cta_ban_numero, ban_codigo " & _
             " FROM cta_banco " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'" & _
             " ORDER BY cta_ban_numero "
    clsCta.Ejecutar strSql
    
    If clsCta.adorec_Def.EOF = False Then
        Set dcmbCuenta.RowSource = clsCta.adorec_Def.DataSource
        dcmbCuenta.ListField = "cta_ban_numero"
        dcmbCuenta.BoundColumn = "ban_codigo"
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
    Else
        Set dcmbCuenta.RowSource = Nothing
        limpiar
    End If
End Sub


Private Sub dcmbCuenta_Change()
    
    'Consulta los datos de la cuenta seleccionada
     strSql = " SELECT ban_codigo,cta_ban_ctaconta, cta_ban_ch_ultimo, cta_ban_observacion, cta_ban_saldoreal, cta_ban_saldoprevisto " & _
              " FROM cta_banco " & _
              " WHERE emp_codigo='" & strEmpresa & "' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND cta_ban_numero = '" & dcmbCuenta.Text & "' " & _
              " ORDER BY cta_ban_numero "
     clsCta.Ejecutar (strSql)

    If clsCta.adorec_Def.EOF = False Then
        txtChUltimo.Tag = clsCta.adorec_Def("ban_codigo")
        txtDescripcion.Text = clsCta.adorec_Def("cta_ban_observacion")
        txtCtaConta = clsCta.adorec_Def("cta_ban_ctaconta")
        txtChUltimo = clsCta.adorec_Def("cta_ban_ch_ultimo")
        txtReal = clsCta.adorec_Def("cta_ban_saldoreal")
        txtPrevisto = clsCta.adorec_Def("cta_ban_saldoprevisto")
        saldodisponible
    Else
        Set dcmbCuenta.RowSource = Nothing
        limpiar
    End If
End Sub


Private Sub Form_Activate()
' Actualiza la lista de cuentas bancarias al volver al formulario
    dcmbBanco.Text = ""
    Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
    dcmbBanco.ListField = "ban_nombre"
    dcmbBanco.BoundColumn = "ban_codigo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    'On Error GoTo errhandler
    
        clsCon_Def.Inicializar AdoConn
        clsBan.Inicializar AdoConn
        clsCta.Inicializar AdoConn
        clsSql.Inicializar AdoConn
        
    'Consulta los bancos existentes en el sistema
        strSql = " select ban_codigo, ban_nombre " & _
                 " from banco "
        clsBan.Ejecutar strSql
        
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
End Sub

Private Sub txtDisponible_Change()
    txtDisponible = FormatoD2(txtDisponible)
End Sub

Private Sub txtPrevisto_Change()
    txtPrevisto = FormatoD2(txtPrevisto)
End Sub

Private Sub txtReal_Change()
    txtReal = FormatoD2(txtReal)
End Sub


