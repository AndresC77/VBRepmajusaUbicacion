VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAnularCompra 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Compras Locales"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnularCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   6990
   Begin VB.TextBox txtNo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1480
      Width           =   1935
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   1199
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   2759
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4337
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox TxtSubTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TxtDesc 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox TxtIva 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtRecargo 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtro de Compras Locales"
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
      Left            =   788
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCotizacion 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Compra:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   885
      End
   End
   Begin MSDataListLib.DataCombo CmbFpago 
      Height          =   330
      Left            =   1410
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   360
      TabIndex        =   24
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO ANULADO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   3435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   21
      Top             =   2550
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   750
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recargos:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   17
      Top             =   2190
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   16
      Top             =   1710
      Width           =   825
   End
   Begin VB.Label LblIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA X%"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   15
      Top             =   1950
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   14
      Top             =   1470
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1845
      Width           =   1125
   End
End
Attribute VB_Name = "frmAnularCompra"
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
Private clsFPago As New clsConsulta

Private Sub cmbCotizacion_Change()
    If cmbCotizacion.MatchedWithList = True Then
        strSQL = " SELECT ing_anulado,CONCAT(per_apellido,' ',per_nombre) as nombC,IIF(LTRIM(RTRIM(ing_factura))='',CONCAT(ing_serie,'-',FORMAT(ing_numero,'0000000')),ing_factura) as numero,ing_fecha, for_pag_nombre,ing_subtotal,ing_dcto,ing_impuesto,ing_subtotal_o,ing_total,COALESCE(ing_numasiento,'') as num " & _
                 " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND  ingreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON ingreso.emp_codigo=forma_pago.emp_codigo AND  ingreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND ingreso.ing_codigo='" & cmbCotizacion.Text & "' " & _
                 " AND ingreso.tip_ing_codigo='COM' "
        clsSql.Ejecutar strSQL
        txtNo.Text = clsSql.adorec_Def("numero")
        CmbFpago.Text = clsSql.adorec_Def("for_pag_nombre")
        txtFecha.Text = clsSql.adorec_Def("ing_fecha")
        TxtSubTotal.Text = Format(clsSql.adorec_Def("ing_subtotal"), "###0.00")
        TxtDesc.Text = Format(clsSql.adorec_Def("ing_dcto"), "###0.00")
        TxtIva.Text = Format(clsSql.adorec_Def("ing_impuesto"), "###0.00")
        TxtRecargo.Text = Format(clsSql.adorec_Def("ing_subtotal_o"), "###0.00")
        TxtTotal.Text = Format(clsSql.adorec_Def("ing_total"), "###0.00")
        If FormatoD0(clsSql.adorec_Def("ing_anulado")) = 1 Then
            lblEstado.Caption = "ANULADO"
            cmdCambiar.Enabled = False
        Else
            lblEstado.Caption = ""
            cmdCambiar.Enabled = True
            If MesCerrado(txtFecha.Text) = True Then
                cmdCambiar.Enabled = False
            End If
        End If
        cmbCotizacion.Tag = clsSql.adorec_Def("num")
    End If
End Sub

Private Sub cmdCambiar_Click()
    Dim Motivo As String
    Dim anula As Boolean
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    Dim Puede As Boolean: Puede = False
    If Left(txtFecha.Text, 7) = Left(HoyDia, 7) Then
        Puede = True
    ElseIf Left(txtFecha.Text, 7) = Left(DateAdd("m", -1, HoyDia), 7) And Right(Left(HoyDia, 10), 2) + 0 <= 5 Then
        Puede = True
    End If
    If Puede = False Then
        frmClave.strClaveMAESTRA = strClaveMAESTRA
        frmClave.dblPrecio = "Anulacion"
        frmClave.Show vbModal
        If frmClave.Ret = False Then
            Puede = False
        Else
            Puede = True
        End If
    End If
    
    If Puede = True Then
        Motivo = ""
        While Motivo = ""
            Motivo = InputBox("Motivo de Anulacion", "Contabilidad")
            Motivo = Motivo & vbNewLine & strUsuario & vbNewLine & HoyDia & " " & Format(Ahora, "HH:MM:SS")
            If Motivo = "" Then
                If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el Asiento?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                    anula = False
                    Motivo = "NO ANULAR"
                End If
            Else
                anula = True
            End If
        Wend
        If anula = True Then
            clsAsiento.NumAsiento = ""
            If cmbCotizacion.Tag = "" Then
                clsAsiento.AnularCX "P", Motivo, txtNo.Text
                Dim clsIngreso As New clsInventario
                clsIngreso.Inicializar AdoConn, AdoConnMaster
                clsIngreso.AnularIng cmbCotizacion.BoundText, "COM", clsAsiento.NumAsiento, Motivo
                Set clsIngreso = Nothing
            Else
                strSQL = " SELECT COALESCE(asi_descripcion,'') as descripcion FROM asiento WHERE emp_codigo='" & strEmpresa & "' AND asi_numasiento='" & cmbCotizacion.Tag & "' "
                clsSql.Ejecutar strSQL
                clsAsiento.NumAsiento = Right(cmbCotizacion.Tag, 14)
                clsAsiento.AnularAsientoYOtros UCase(Motivo), clsSql.adorec_Def("descripcion")
            End If
            MsgBox "Compra # " & cmbCotizacion.BoundText & " anulada", vbInformation, "Anular"
            cmbCotizacion_Change
        End If
    End If
    Set clsAsiento = Nothing
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

Private Sub cmbCliente_Change()
    If cmbCliente.MatchedWithList = True Then
    strSQL = " SELECT ing_codigo " & _
             " FROM ingreso INNER JOIN persona ON (ingreso.emp_codigo = persona.emp_codigo) AND (ingreso.per_codigo = persona.per_codigo) " & _
             " WHERE tip_ing_codigo='COM' AND ingreso.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' AND cat_p_tipo='P' " & _
             " ORDER BY ingreso.ing_codigo "
    clsSql.Ejecutar strSQL
    cmbCotizacion = ""
    Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
    cmbCotizacion.ListField = "ing_codigo"
    cmbCotizacion.Tag = ""
    lblEstado.Caption = ""
    End If
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        frmReporte.strNumero = cmbCotizacion.BoundText
        frmReporte.strTipo = "COM"
        frmReporte.strReporte = "rptIngresoMercaderia"
        frmReporte.Show
    Else
        MsgBox "No ha seleccionado un No. Compra", vbInformation, "Factura"
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    cmbCotizacion.Tag = ""
    lblEstado.Caption = ""
    cmdCambiar.Enabled = False
    
    strSQL = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSQL)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")
    
    'Coloca los datos de los vendedores en un listado
    strSQL = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='P' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar strSQL
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSQL = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar strSQL
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
End Sub


