VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmV_ImpCotizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Cotizaciones"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_ImpCotizacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdFormato 
      Caption         =   "&Formato"
      Height          =   375
      Left            =   2153
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impresión de Cotizaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtAtencion 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCotizacion 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cotización"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atención:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1125
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3833
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   473
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmV_ImpCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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
Private strSql As String
Private clsSql As New clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub cmbCliente_Change()
    strSql = " SELECT cotizacion.cot_codigo " & _
             " FROM proyecto_venta INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo) " & _
             " WHERE cotizacion.cot_estado BETWEEN 0 AND 2 AND proyecto_venta.emp_codigo='" & strEmpresa & "' AND proyecto_venta.per_codigo='" & cmbCliente.BoundText & "'" & _
             " ORDER BY cotizacion.cot_codigo "
    clsSql.Ejecutar (strSql)
    cmbCotizacion = ""
    Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
    cmbCotizacion.ListField = "cot_codigo"
End Sub

Private Sub cmdFormato_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
'        drptCotizacionFormato.Tag = cmbCotizacion.BoundText
'        drptCotizacionFormato.Atencion = txtAtencion
'        drptCotizacionFormato.Show
        Dim frmCotF As New frmReporte
        frmCotF.Atencion = txtAtencion
        frmCotF.strNumero = cmbCotizacion.BoundText
        frmCotF.strReporte = "rptCotizacionF"
        frmCotF.Show
    Else
        MsgBox "No ha seleccionado una cotización", vbInformation, "Cotización"
    End If

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
'        drptCotizacion.Tag = cmbCotizacion.BoundText
'        drptCotizacion.Atencion = txtAtencion
'        drptCotizacion.Show
        Dim frmCot As New frmReporte
        frmCot.Atencion = txtAtencion
        frmCot.strNumero = cmbCotizacion.BoundText
        frmCot.strReporte = "rptCotizacion"
        frmCot.Show
    Else
        MsgBox "No ha seleccionado una cotización", vbInformation, "Cotización"
    End If
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
'    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre) as nombC, cotizacion.cot_codigo " & _
'             " FROM (persona INNER JOIN proyecto_venta ON (persona.per_codigo = proyecto_venta.per_codigo) AND (persona.emp_codigo = proyecto_venta.emp_codigo)) " & _
'             " INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo) " & _
'             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
'             " ORDER BY nombC, cotizacion.cot_codigo "
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
End Sub
