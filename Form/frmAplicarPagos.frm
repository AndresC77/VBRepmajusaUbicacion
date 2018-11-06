VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAplicarPagos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicar Pagos"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmAplicarPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   11475
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5910
      TabIndex        =   12
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4110
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11295
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optproveedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1110
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optcliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfgDocumento 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   11055
         _cx             =   19500
         _cy             =   2355
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarPagos.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   10695
         _cx             =   18865
         _cy             =   2778
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarPagos.frx":0452
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas por Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         Left            =   6600
         TabIndex        =   9
         Top             =   4110
         Width           =   600
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4320
         Picture         =   "frmAplicarPagos.frx":05A1
         ToolTipText     =   "Elimina una Fila"
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4560
         Picture         =   "frmAplicarPagos.frx":06D7
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2310
         TabIndex        =   6
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmAplicarPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private clsSql As New clsConsulta
Private clsPer As New clsConsulta
Private strSQL As String
Private t As String

Private Sub cmdAceptar_Click()
'Comprueba que todos los datos esten ingresados
    If FormatoD2(txtValor.Text) <> FormatoD2(vsfgDocumento.TextMatrix(vsfgDocumento.Row, 5)) Then
        MsgBox "No Completa el cobro", vbCritical, "Cobros"
    Else
        For i = 1 To VSFG1.Rows - 1
            If Abs(VSFG1.TextMatrix(i, 1)) = 1 Then
                k = VSFG1.TextMatrix(i, 10)
                If (VSFG1.TextMatrix(i, 10) <> "" Or VSFG1.TextMatrix(i, 10) <> "0") Then
                    'Calcula el máximo codigo de pago para la cuenta
                     strSQL = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                              " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                              "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                              "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                              " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'P'" & _
                              " GROUP BY pago.emp_codigo"
                    clsSql.Ejecutar strSQL
                    If clsSql.adorec_Def.EOF Then
                        maxpag = 1
                    Else
                        maxpag = clsSql.adorec_Def("pag") + 1
                    End If
                    
                    strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, pag_no_doc, pag_observacion,doc_pag_codigo,asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & Val(VSFG1.TextMatrix(i, 2)) & "', 'P', '" & Val(maxpag) & "', '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 7) & "', '" & FormatoD2(VSFG1.TextMatrix(i, 10)) & "', '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 4) & "', '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 6) & "', " & _
                             " '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 1) & "','" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 9) & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsSql.Ejecutar strSQL, "M"
                    If FormatoD2(VSFG1.TextMatrix(i, 9)) <= FormatoD2(VSFG1.TextMatrix(i, 10)) Then
                        strSQL = " UPDATE cuenta_p_c " & _
                                 " SET cue_p_c_fechapago='" & ffch & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                 " WHERE cue_p_c_tipo= 'P' AND cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' AND emp_codigo = '" & strEmpresa & "' "
                        clsSql.Ejecutar strSQL, "M"
                    End If
                End If
        
                
            End If
            
        Next i
    End If
    For i = 0 To 10000
    Next i
    dcmbBeneficiario_Change
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBeneficiario_Change()
    txtValor = 0
    t = "P"
    If Me.optcliente.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiario.MatchedWithList = True Then
       
        vsfgDocumento.Enabled = True
        'Consulta para el grid
        strSQL = " SELECT DISTINCT '0', com_egr_codigo, CONCAT(per_apellido,' ',per_nombre), " & _
                 " banco.ban_nombre, com_egr_ch_num,  com_egr_ch_valor,com_egr_descripcion,com_egr_ch_fecha,com_egr_fecha,comp_egreso.asi_numasiento " & _
                 " FROM (comp_egreso INNER JOIN banco ON comp_egreso.ban_codigo = banco.ban_codigo) " & _
                 " INNER JOIN persona ON comp_egreso.per_codigo=persona.per_codigo AND comp_egreso.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN cuenta_p_c ON comp_egreso.emp_codigo=cuenta_p_c.emp_codigo AND comp_egreso.per_codigo=cuenta_p_c.per_codigo AND cuenta_p_c.cue_p_c_pagado=0 AND cuenta_p_c.cue_p_c_tipo='P'" & _
                 " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND comp_egreso.com_egr_codigo=pago.doc_pag_codigo AND comp_egreso.asi_numasiento=pago.asi_numasiento " & _
                 " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND comp_egreso.per_codigo='" & dcmbBeneficiario.BoundText & "'  " & _
                 " AND pago.cue_p_c_codigo IS NULL " & _
                 " ORDER BY com_egr_codigo "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set vsfgDocumento.DataSource = clsSql.adorec_Def.DataSource
        Else
            vsfgDocumento.Clear 1
            vsfgDocumento.Rows = 1
        End If
        strSQL = " SELECT ' ','0', cuenta_p_c.cue_p_c_codigo, CONCAT(cue_p_c_fra_cuenta, '/' , cue_p_c_tot_cuenta ) as cue_p_c_fra_cuenta, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0), ' ' " & _
                 " FROM  (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo)" & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & "' AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'P' AND cue_p_c_pagado='0' " & _
                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cue_p_c_fra_cuenta,cue_p_c_tot_cuenta,cue_p_c_egr_codigo,cue_p_c_descripcion,cue_p_c_fechaemision,cue_p_c_fechapropuesta,cue_p_c_valor,com_ret_total "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Valor = clsSql.adorec_Def("cue_p_c_valor")
            Set VSFG1.DataSource = clsSql.adorec_Def.DataSource
             VSFG1.ColDataType(1) = flexDTBoolean
            'ponerBotones
        Else
            Valor = 0
            VSFG1.Clear 1
            VSFG1.Rows = 2
        End If
        For i = 1 To VSFG1.Rows - 1
          VSFG1.TextMatrix(i, 0) = i
        Next i
    End If
End Sub
'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa las clases para hacer distintas consultas
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    vsfgDocumento.Rows = 1
    
    'Seleccionamos el proveedor de la tabla persona (P), que esta por defecto
    
    optcliente.Value = True
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
        Persona = ""
        p = 0
    End If
End Sub

Private Sub OptCliente_Click()
  
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiario.Text = ""
    vsfgDocumento.Clear 1
    vsfgDocumento.Rows = 1
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub optproveedor_Click()
    
    p = 1
    Frame1.Caption = "Proveedor"
    dcmbBeneficiario.Text = ""
    vsfgDocumento.Clear 1
    vsfgDocumento.Rows = 1
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strComparar As String
    If Col = 10 Then
        'Verifica que solo se ingresen números en el campo Debe
        If Not IsNumeric(VSFG1.TextMatrix(Row, 10)) And VSFG1.TextMatrix(Row, 10) <> "" Then
            MsgBox "Ingrese solo números en el Valor de Pago.", vbInformation, "Pagos"
            VSFG1.TextMatrix(Row, 10) = 0
        End If
    End If
    If Row < VSFG1.Rows Then
        If Val(VSFG1.TextMatrix(Row, 10)) > Val(VSFG1.TextMatrix(Row, 9)) Then
            If MsgBox("El valor a pagar es mayor al Saldo." & vbNewLine & "Esta seguro de que el pago es mayor?", vbCritical + vbYesNo, "Pagos") = vbNo Then
                VSFG1.Select Row, 10
                VSFG1.TextMatrix(Row, 10) = 0
            End If
        End If
    End If
    pagos
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFG1.TextMatrix(Row, 1) = "0" Or VSFG1.TextMatrix(Row, 1) = "" Then
        If Col >= 10 Then
            Cancel = True
        End If
    End If
  
End Sub

Private Sub VSFG1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 2 Or NewCol = 3 Or NewCol = 4 Or NewCol = 5 Or NewCol = 6 Or NewCol = 7 Or NewCol = 8 Or NewCol = 9 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 11 Then
            a = 1
        End If
    If Col = 10 Then
        txtValor = 0
    End If
    If Col = 1 And Row > 0 Then
        If Abs(VSFG1.TextMatrix(Row, 1)) = 1 Then
            VSFG1.Select Row, 1, Row, 10
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HC0FFFF
            VSFG1.Select Row, 10
        ElseIf Abs(VSFG1.TextMatrix(Row, 1)) = 0 Then
            VSFG1.Select Row, 1, Row, 10
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HFFFFFF
            VSFG1.Select Row, 10
            VSFG1.TextMatrix(Row, 10) = ""
            If Row < VSFG1.Rows - 2 And VSFG1.TextMatrix(Row, 0) <> " " Then
                While VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row + 1, 0)
                    VSFG1.RemoveItem Row + 1
                    If Row = VSFG1.Rows - 1 Then Exit Sub
                Wend
            End If
        End If
    End If
End Sub

Private Sub vsfgDocumento_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    End If
End Sub


Private Sub vsfgDocumento_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 0 And Row > 0 Then
        If vsfgDocumento.TextMatrix(Row, 0) = "-1" Then
            For i = 1 To vsfgDocumento.Rows - 1
                If i <> Row And vsfgDocumento.TextMatrix(i, 0) = "-1" Then
                    vsfgDocumento.Select i, 0, i, 6
                    vsfgDocumento.FillStyle = flexFillRepeat
                    vsfgDocumento.CellBackColor = &HFFFFFF
                    vsfgDocumento.TextMatrix(i, 0) = "0"
                End If
            Next
            vsfgDocumento.Select Row, 0, Row, 6
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HC0FFFF
            vsfgDocumento.Select Row, 0
        ElseIf vsfgDocumento.TextMatrix(Row, 0) = "0" Then
            vsfgDocumento.Select Row, 0, Row, 6
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HFFFFFF
            vsfgDocumento.Select Row, 0
        End If
    End If
End Sub

Private Sub pagos()
    Dim aux As Long
    Dim i As Long
    Dim j As Long
    aux = 0
    For j = 2 To lonNFijas - 1
            VSFG.TextMatrix(j, 3) = 0
    Next j
    For i = 1 To VSFG1.Rows - 1
        If Abs(VSFG1.TextMatrix(i, 1)) = 1 Then
            If aux <> VSFG1.TextMatrix(i, 0) Then
                Suma = Suma + Val(VSFG1.TextMatrix(i, 10))
                aux = VSFG1.TextMatrix(i, 0)
            End If
        End If
    Next i
    txtValor = Format(Suma, "##0.00")
End Sub
