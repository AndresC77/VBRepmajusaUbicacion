VERSION 5.00
Begin VB.Form frmV_ReImpNV 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re Impresión de Facturas"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_ReImpNV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   5655
   Begin VB.TextBox txtNumero 
      Height          =   315
      Left            =   1080
      MaxLength       =   20
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   705
      Width           =   3255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   540
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000050&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota de Venta"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmV_ReImpNV"
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

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.MatchedWithList = True Then
        strSql = " SELECT egr_codigo " & _
                 " FROM egreso INNER JOIN persona ON (egreso.emp_codigo = persona.emp_codigo) AND (egreso.per_codigo = persona.per_codigo) " & _
                 " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' AND cat_p_tipo='C' " & _
                 " ORDER BY egreso.egr_codigo "
        clsSql.Ejecutar (strSql)
        cmbCotizacion = ""
        Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
        cmbCotizacion.ListField = "egr_codigo"
    End If
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        LimpiarTodo
'''        strSql = " SELECT tip_ped_ptofac " & _
'''                 " FROM tipo_pedido " & _
'''                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
'''        clsSql.Ejecutar strSql
'''        If clsSql.adorec_Def.RecordCount > 0 Then
'''
'''        End If
    Else
        Exit Sub
    End If
    
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
    cmbCotizacion.BoundText = ""
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

Private Sub cmdImpGuia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        frmReporte.strNumero = cmbCotizacion.BoundText
        frmReporte.strTipo = "FAC"
        frmReporte.strReporte = "rptGuiaRemision"
        frmReporte.Show
    Else
        MsgBox "No ha seleccionado una factura", vbInformation, "Factura"
    End If

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If txtNumero.Text <> "" Then

            Dim RepNotaVenta As New frmReporte
            RepNotaVenta.strNumero = txtNumero.Text
            RepNotaVenta.strReporte = "rptNotaVenta"
            RepNotaVenta.Show
    Else
        MsgBox "No ha seleccionado una Nota de Venta", vbInformation, "Nota de Venta"
    End If
End Sub
Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub cargarTipoPedido()
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    cargarTipoPedido
    
End Sub
