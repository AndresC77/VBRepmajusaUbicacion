VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmV_ReImpCobro 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re Impresión de Cobros"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_ReImpCobro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   6390
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impresíon de Cobros"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCotizacion 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   4665
         _ExtentX        =   8229
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
         TabIndex        =   6
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Documento:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3308
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   1628
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmV_ReImpCobro"
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
Private strSQL As String
Private clsSql As New clsConsulta

Private Sub cmbCliente_Validate(Cancel As Boolean)
    strSQL = " SELECT doc_pag_codigo,CONCAT(doc_pag_codigo,' - Recepción: ',LEFT(doc_pag_fecha_recepcion,10),' - Tipo: '," & _
             " IIF(doc_pago.tip_doc_pag_codigo IS NULL or doc_pago.tip_doc_pag_codigo='','EFECTIVO',tip_doc_pag_nombre)) as descr " & _
             " FROM doc_pago INNER JOIN persona ON (doc_pago.emp_codigo = persona.emp_codigo) AND (doc_pago.per_codigo = persona.per_codigo) " & _
             " LEFT JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo " & _
             " WHERE doc_pago.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' " & _
             " ORDER BY doc_pag_fecha_recepcion DESC,doc_pag_codigo "
    clsSql.Ejecutar (strSQL)
    cmbCotizacion = ""
    Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
    cmbCotizacion.ListField = "descr"
    cmbCotizacion.BoundColumn = "doc_pag_codigo"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        Dim RepCobro As New frmReporte
        RepCobro.strNumero = cmbCotizacion.BoundText
        RepCobro.strReporte = "rptReciboCaja"
        RepCobro.Show
    Else
        MsgBox "No ha seleccionado un documento de cobro", vbInformation, "Factura"
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
    strSQL = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombC " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSQL)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
End Sub
