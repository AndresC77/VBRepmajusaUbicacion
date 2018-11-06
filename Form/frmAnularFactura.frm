VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAnularFactura 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Facturas"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnularFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdCambiarNumero 
      Caption         =   "&Cambiar Numero"
      Height          =   375
      Left            =   3548
      TabIndex        =   27
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1988
      TabIndex        =   26
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   428
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   1988
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5108
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImpGuia 
      Caption         =   "Imprimir Guía"
      Height          =   375
      Left            =   3548
      TabIndex        =   6
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   21
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   16
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   14
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
         Size            =   8.25
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
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtro de Facturas"
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
      TabIndex        =   8
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
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   840
      End
   End
   Begin MSDataListLib.DataCombo CmbFpago 
      Height          =   330
      Left            =   1410
      TabIndex        =   3
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
   Begin MSDataListLib.DataCombo cmbVendedor 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   25
      Top             =   2520
      Width           =   3435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   24
      Top             =   2550
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recargos:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   1470
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1485
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   1845
      Width           =   1125
   End
End
Attribute VB_Name = "frmAnularFactura"
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
        strSQL = " SELECT egr_anulado,CONCAT(per_apellido,' ',per_nombre) as nombC,COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV,egr_fecha, for_pag_nombre,egr_subtotal,egr_dcto,egr_impuesto,egr_subtotal_o,egr_total,egr_numasiento as num " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND  egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON egreso.emp_codigo=forma_pago.emp_codigo AND  egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " LEFT JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND  egreso.ven_codigo=vendedor.ven_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.egr_codigo='" & cmbCotizacion.Text & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' "
        clsSql.Ejecutar strSQL
        cmbVendedor.Text = clsSql.adorec_Def("nombV")
        CmbFpago.Text = clsSql.adorec_Def("for_pag_nombre")
        txtFecha.Text = clsSql.adorec_Def("egr_fecha")
        TxtSubTotal.Text = Format(clsSql.adorec_Def("egr_subtotal"), "###0.00")
        TxtDesc.Text = Format(clsSql.adorec_Def("egr_dcto"), "###0.00")
        TxtIva.Text = Format(clsSql.adorec_Def("egr_impuesto"), "###0.00")
        TxtRecargo.Text = Format(clsSql.adorec_Def("egr_subtotal_o"), "###0.00")
        TxtTotal.Text = Format(clsSql.adorec_Def("egr_total"), "###0.00")
        If FormatoD0(clsSql.adorec_Def("egr_anulado")) = 1 Then
            lblEstado.Caption = "ANULADO"
            cmdCambiar.Enabled = False
            cmdEliminar.Enabled = True
        Else
            lblEstado.Caption = ""
            cmdCambiar.Enabled = True
            cmdEliminar.Enabled = False
            If MesCerrado(txtFecha.Text) = True Then
                cmdCambiar.Enabled = False
            End If
        End If
        cmbCotizacion.Tag = clsSql.adorec_Def("num")
    End If
End Sub

Private Function RevisionCobro(strFac As String) As Boolean
    Dim clsConsultaPago As New clsConsulta
    Dim strSqlPago As String
    Dim strMensajeAnul As String
    clsConsultaPago.Inicializar AdoConn, AdoConnMaster
    strSqlPago = " SELECT count(*) as n " & _
                 " FROM cuenta_p_c INNER JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                 " AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
                 " AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_egr_codigo='" & strFac & "' AND cuenta_p_c.cue_p_c_tipo='C'" & _
                 " AND pago.pag_monto!=0 "
    clsConsultaPago.Ejecutar strSqlPago
    RevisionCobro = True
    strMensajeAnul = ""
    If FormatoD0(clsConsultaPago.adorec_Def("n")) <> 0 Then
        RevisionCobro = False
        strMensajeAnul = "La Factura tiene registrado COBROS."
    End If
    strSqlPago = " SELECT count(*) as n " & _
                 " FROM cuenta_p_c INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                 " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_egr_codigo='" & strFac & "' AND cuenta_p_c.cue_p_c_tipo='C' " & _
                 " AND comprobante_retencion.com_ret_total!=0 "
    clsConsultaPago.Ejecutar strSqlPago
    If FormatoD0(clsConsultaPago.adorec_Def("n")) <> 0 Then
        RevisionCobro = False
        strMensajeAnul = strMensajeAnul & vbNewLine & "La Factura tiene registrado RETENCION."
    End If
    If Trim(strMensajeAnul) <> "" Then
        MsgBox "La Factura no puede ser ANULADA" & vbNewLine & vbNewLine & strMensajeAnul, vbInformation + vbOKOnly, "Anulación Factura"
    End If
End Function

Private Sub cmdCambiar_Click()
    Dim Motivo As String
    Dim Anula As Boolean
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    Dim Puede As Boolean: Puede = False
    If RevisionCobro(cmbCotizacion.BoundText) = False Then Exit Sub
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
                    Anula = False
                    Motivo = "NO ANULAR"
                End If
            Else
                Anula = True
            End If
        Wend
        If Anula = True Then
            clsAsiento.NumAsiento = ""
            If cmbCotizacion.Tag = "" Then
                clsAsiento.AnularCX "C", Motivo, cmbCotizacion.BoundText
                Dim clsEgreso As New clsInventario
                clsEgreso.Inicializar AdoConn, AdoConnMaster
                clsEgreso.AnularEgr cmbCotizacion.BoundText, "FAC", clsAsiento.NumAsiento, Motivo
                Set clsEgreso = Nothing
            Else
                strSQL = " SELECT COALESCE(asi_descripcion,'') as descripcion FROM asiento WHERE emp_codigo='" & strEmpresa & "' AND asi_numasiento='" & cmbCotizacion.Tag & "' "
                clsSql.Ejecutar strSQL
                clsAsiento.NumAsiento = Right(cmbCotizacion.Tag, 14)
                clsAsiento.AnularAsientoYOtros UCase(Motivo), clsSql.adorec_Def("descripcion")
            End If
            MsgBox "Factura # " & cmbCotizacion.BoundText & " anulada", vbInformation, "Anular"
            cmbCotizacion_Change
        End If
    End If
    Set clsAsiento = Nothing
End Sub

Private Sub cmdCambiarNumero_Click()
    Dim NuevoNumero As Double
    NuevoNumero = Val(InputBox("FACTURA A GENERAR (00100X000000X)", "Numeración de Documento", NuevoNuemro))
    If NuevoNumero = 0 Then Exit Sub
    strSQL = " SELECT egreso.per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli " & _
             " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
             " AND egreso.per_codigo=persona.per_codigo " & _
             " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & NuevoNumero & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        MsgBox "La factura esta utilizada por el cliente: " & clsSql.adorec_Def("cli") & vbNewLine & "No puede usar esta factura", vbInformation + vbOKOnly, "FACTURA"
        Exit Sub
    Else
        strSQL = " UPDATE egreso " & _
                 " SET egr_codigo='" & NuevoNumero & "',egr_numero='" & Right(NuevoNumero, 7) & "' " & _
                 " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE det_egreso " & _
                 " SET egr_codigo='" & NuevoNumero & "' " & _
                 " WHERE tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE det_egreso_c " & _
                 " SET egr_codigo='" & NuevoNumero & "' " & _
                 " WHERE tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE cuenta_p_c " & _
                 " SET cue_p_c_descripcion = CONCAT('FACTURA # " & NuevoNumero & vbNewLine & strUsuario & vbNewLine & "ANTERIOR ',cue_p_c_descripcion)," & _
                 " cue_p_c_egr_codigo='" & NuevoNumero & "',cue_p_c_numero='" & Right(NuevoNumero, 7) & "', " & _
                 " cue_p_c_usumod='" & strUsuario & "',cue_p_c_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE cue_p_c_tipo='C' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND cue_p_c_egr_codigo='" & cmbCotizacion.Text & "' " & _
                 " AND asi_numasiento='" & cmbCotizacion.Tag & "'"
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE asiento " & _
                 " SET asi_descripcion = CONCAT('FACTURA " & NuevoNumero & vbNewLine & strUsuario & vbNewLine & "ANTERIOR ',asi_descripcion)," & _
                 " asi_usumod='" & strUsuario & "',asi_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & cmbCotizacion.Tag & "'"
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE pedido" & _
                 " set ped_egr_codigo='" & NuevoNumero & "',ped_observacion=CONCAT('FACTURA MODIFICADA " & cmbCotizacion.Text & vbNewLine & strUsuario & vbNewLine & " ',ped_observacion)," & _
                 " ped_usumod='" & strUsuario & "',ped_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE ped_tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND ped_egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE det_contenedor" & _
                 " set egr_codigo='" & NuevoNumero & "'," & _
                 " det_con_usumod='" & strUsuario & "',det_con_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE cambio" & _
                 " SET cam_factura='" & NuevoNumero & "',cam_observacion=CONCAT('FACTURA MODIFICADA " & cmbCotizacion.Text & vbNewLine & strUsuario & vbNewLine & " ',cam_observacion)," & _
                 " cam_usumod='" & strUsuario & "',cam_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cam_factura='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        MsgBox "Factura CAMBIADA.", vbInformation + vbOKOnly, "Cambio"
        cmbCliente_Change
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lblEstado.Caption = "ANULADO" Then
        strSQL = " DELETE " & _
                 " FROM egreso " & _
                 " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " DELETE " & _
                 " FROM det_egreso " & _
                 " WHERE tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " DELETE " & _
                 " FROM det_egreso_c " & _
                 " WHERE tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE cuenta_p_c " & _
                 " SET cue_p_c_descripcion = CONCAT('FACTURA ELIMINADA " & vbNewLine & strUsuario & vbNewLine & " ',cue_p_c_descripcion)," & _
                 " cue_p_c_egr_codigo='',cue_p_c_numero='0',tip_doc_cue_codigo='-10', " & _
                 " cue_p_c_usumod='" & strUsuario & "',cue_p_c_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE cue_p_c_tipo='C' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND cue_p_c_egr_codigo='" & cmbCotizacion.Text & "' " & _
                 " AND asi_numasiento='" & cmbCotizacion.Tag & "'"
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE asiento " & _
                 " SET asi_descripcion = CONCAT('FACTURA ELIMINADA " & vbNewLine & strUsuario & vbNewLine & " ',asi_descripcion)," & _
                 " asi_usumod='" & strUsuario & "',asi_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & cmbCotizacion.Tag & "'"
        clsSql.Ejecutar strSQL
        strSQL = " UPDATE pedido" & _
                 " set ped_estado=3,ped_observacion=CONCAT('FACTURA ELIMINADA " & vbNewLine & strUsuario & vbNewLine & " ',ped_observacion)," & _
                 " ped_usumod='" & strUsuario & "',ped_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE ped_tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND ped_egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        strSQL = " SELECT ped_codigo FROM pedido" & _
                 " WHERE ped_tip_egr_codigo='FAC' AND emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & cmbCliente.BoundText & "' AND ped_egr_codigo='" & cmbCotizacion.Text & "' "
        clsSql.Ejecutar strSQL
        LiberarIncentivo clsSql.adorec_Def("ped_codigo"), cmbCliente.BoundText
        MsgBox "Factura ELIMINADA.", vbInformation + vbOKOnly, "Eliminación"
        cmbCliente_Change
    Else
        MsgBox "No se puede eliminar esta factura." & vbNewLine & "La factura debe estar anulada", vbInformation + vbOKOnly, "Eliminación"
    End If
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
    strSQL = " SELECT egr_codigo " & _
             " FROM egreso INNER JOIN persona ON (egreso.emp_codigo = persona.emp_codigo) AND (egreso.per_codigo = persona.per_codigo) " & _
             " WHERE tip_egr_codigo='FAC' AND egreso.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' AND cat_p_tipo='C' " & _
             " ORDER BY egreso.egr_codigo "
    clsSql.Ejecutar strSQL
    cmbCotizacion = ""
    Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
    cmbCotizacion.ListField = "egr_codigo"
    cmbCotizacion.Tag = ""
    lblEstado.Caption = ""
    End If
End Sub

Private Sub cmdImpGuia_Click()
    frmReporte.strNumero = cmbCotizacion.BoundText
    frmReporte.strTipo = "FAC"
    frmReporte.strReporte = "rptGuiaRemision"
    frmReporte.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    Dim GuiaAutomatica As Boolean
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        frmReporte.strNumero = cmbCotizacion.BoundText
        'listo
        GuiaAutomatica = False
        frmReporte.strReporte = IIf(GuiaAutomatica = True, "rptFacturaGuia", "rptFacturaSola")
        frmReporte.Show

    Else
        MsgBox "No ha seleccionado una factura", vbInformation, "Factura"
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
    cmdEliminar.Enabled = False
    
    strSQL = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSQL)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")
    
    
    
'****** VENDEDORES
    'Coloca los datos de los vendedores en un listado
    strSQL = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as nombV " & _
             " FROM vendedor " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY nombV "
    clsSql.Ejecutar strSQL
    Set cmbVendedor.RowSource = clsSql.adorec_Def.DataSource
    cmbVendedor.ListField = "nombV"
    cmbVendedor.BoundColumn = "ven_codigo"
    
    strSQL = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombC " & _
             " FROM persona " & _
             " INNER JOIN tipo_pedido ON tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
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
