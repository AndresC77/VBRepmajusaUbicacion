VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPagoRetencion 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmPagoRetencion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8190
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Retenciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   7935
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7575
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   285
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            Value           =   42009.4904861111
         End
         Begin VB.TextBox txtAutorizacion 
            Height          =   285
            Left            =   6000
            TabIndex        =   3
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtDocumentoR 
            Height          =   285
            Left            =   2640
            TabIndex        =   2
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAServ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtSTcero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtBeneficiario 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtCX 
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Autorizacion:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4440
            TabIndex        =   28
            Top             =   990
            Width           =   1470
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de documento:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   27
            Top             =   990
            Width           =   1350
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4440
            TabIndex        =   26
            Top             =   1710
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Serv:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   2070
            Width           =   1380
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Prod:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   1710
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal 0%:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   1350
            Width           =   915
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4440
            TabIndex        =   22
            Top             =   2070
            Width           =   555
         End
         Begin VB.Label lblBeneficiario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beneficiario:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   630
            Width           =   900
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Emisión:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Documento:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   3360
            TabIndex        =   13
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de CXP:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5040
            TabIndex        =   12
            Top             =   630
            Width           =   855
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Top             =   4440
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         Top             =   2760
         Width           =   7680
         _cx             =   227357931
         _cy             =   227347162
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPagoRetencion.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4410
         Picture         =   "frmPagoRetencion.frx":03DC
         ToolTipText     =   "Elimina una Fila"
         Top             =   4440
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4170
         Picture         =   "frmPagoRetencion.frx":0512
         Top             =   4440
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4800
         TabIndex        =   16
         Top             =   4470
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2483
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4133
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmPagoRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso del comprobante de egresos comunes                         #
'#  frmComprobanteEgresoComun V1.0                                              #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el comprobante de egresos comunes                     #
'#  Permite ingresar los datos de egresos comunes y sus detalles                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  COMP_EGRESO: Esta tabla almacena los datos del comprobante                  #
'#  PERSONA: donde se guardan los datos de los benficiarios de los comprobantes #
'#  DET_COMP_EGRESO: Guarda los detalles del comprobante de Egreso              #
'#  RET_COMP_EGRESO: Guarda las retenciones que puede tener el comprobante      #
'#  CTA_BANCO: consulta los datos del numero de cuenta y el último cheque       #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private strFechaPago As String

Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private clsPag As New clsConsulta
Private clsPer As New clsConsulta
Private clsSql As New clsConsulta
Private clsEgr As New clsConsulta
Private clsCod As New clsConsulta
Private clsPgd As New clsConsulta
Private clsAsi As New clsConsulta
Private strSql As String
Private IVA As Double

Private Sub Form_Activate()
    strSql = " SELECT com_ret_autorizacion,com_ret_serie,com_ret_numero " & _
             " FROM comprobante_retencion  " & _
             " WHERE " & IIf(GeneraDocElec = 1, "com_ret_serie='001" & PtoEmiDocEle & "' AND ", "") & _
             " emp_codigo = '" & strEmpresa & "' AND cue_p_c_tipo='" & Me.Tag & "' " & _
             " ORDER BY com_ret_numero DESC"
             'LIMIT 0,1
    clsCta.Ejecutar strSql
    If Not clsCta.adorec_Def.EOF Then
        txtSerie.Text = IIf(GeneraDocElec = 1, "001" & PtoEmiDocEle & "", "")
        txtAutorizacion.Text = clsCta.adorec_Def("com_ret_autorizacion")
        txtDocumentoR.Text = FormatoD2(clsCta.adorec_Def("com_ret_numero")) + 1
        'txtDocumentoR.Text = InputBox("Numero Retencion", "Retencion")
    Else
        txtSerie.Text = IIf(GeneraDocElec = 1, "001" & PtoEmiDocEle & "", "")
        txtAutorizacion.Text = "0"
        txtDocumentoR.Text = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim RepReten As New frmReporte
    RepReten.strNumero = txtCX.Text
    RepReten.strAsiento = VSFG.Tag
    RepReten.Atencion = "Fecha de Pago: " & Format(strFechaPago, "yyyy-MM-dd")
    RepReten.strTipo = Me.Tag
    RepReten.strReporte = "rptRetencionDiario"
    RepReten.Show

'    frmReporte.strAsiento = VSFG.Tag
'    frmReporte.strReporte = "rptAsiento"
'    frmReporte.Show
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
    Set clsPag = Nothing
    Set clsPer = Nothing
    Set clsSql = Nothing
    Set clsEgr = Nothing
    Set clsCod = Nothing
    Set clsPgd = Nothing
    Set clsAsi = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
    
    'For i = 1 To (VSFG1.Rows - 1)
    '    VSFG1.TextMatrix(i, 0) = i
    'Next i
End Sub

Private Sub cmdAceptar_Click()
'Comprueba que todos los datos esten ingresados
    ffch = Format(dtpFecha.Value, "yyyy-mm-dd")
    If (IsDate(ffch) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Pagos"
        Exit Sub
    End If
    If Trim(txtSerie.Text) = "" Or Trim(txtDocumentoR.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
        MsgBox "Debe llenar todos los datos del Comprobante de Retención", vbInformation, "Pagos"
        Exit Sub
    End If

    'Ingreso de datos en la tabla pago
    n = VSFG.Rows - 1
    For i = 1 To n
        If VSFG.TextMatrix(i, 1) = "" And VSFG.TextMatrix(i, 2) = "" And VSFG.TextMatrix(i, 3) = "" And VSFG.TextMatrix(i, 4) = "0" Then
            VSFG.RemoveItem i
        Else
            
            'MsgBox "Los datos de Retenciones no estan completos", vbInformation, "Retenciones"
            'Exit Sub
        End If
    Next
    
    
    strSql = " SELECT cue_p_c_fechapropuesta " & _
                 " FROM cuenta_p_c " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND cue_p_c_tipo='" & Me.Tag & "' " & _
                 " AND cue_p_c_codigo='" & txtCX.Text & "' "
                 
    clsCta.Ejecutar strSql
    If clsCta.adorec_Def.RecordCount > 0 Then
        strFechaPago = clsCta.adorec_Def("cue_p_c_fechapropuesta")
    Else
        strFechaPago = ""
    End If
    
    strSql = " BEGIN TRAN "
    clsSql.Ejecutar strSql, "M"
    strSql = " SELECT com_ret_serie,com_ret_numero " & _
             " FROM comprobante_retencion  WITH (TABLOCKX)" & _
             " WHERE " & IIf(GeneraDocElec = 1, "com_ret_serie='001" & PtoEmiDocEle & "' AND ", "") & " emp_codigo = '" & strEmpresa & "' AND cue_p_c_tipo='" & Me.Tag & "' ORDER BY com_ret_numero DESC"
             'DESC LIMIT 0,1
    clsSql.Ejecutar strSql, "M"
    If Not clsSql.adorec_Def.EOF Then
        txtSerie.Text = IIf(GeneraDocElec = 1, "001" & PtoEmiDocEle & "", "")
        txtAutorizacion.Text = "0"
        txtDocumentoR.Text = FormatoD2(clsSql.adorec_Def("com_ret_numero")) + 1
        'txtDocumentoR.Text = InputBox("Numero Retencion", "Retencion")
    Else
        txtSerie.Text = IIf(GeneraDocElec = 1, "001" & PtoEmiDocEle & "", "")
        txtAutorizacion.Text = "0"
        txtDocumentoR.Text = 1
        'txtDocumentoR.Text = InputBox("Numero Retencion", "Retencion")
    End If
    strSql = " INSERT INTO comprobante_retencion (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,com_ret_fecha,com_ret_serie,com_ret_numero,com_ret_autorizacion,com_ret_total,com_ret_fechamod,com_ret_usumod) " & _
             " VALUES ('" & strEmpresa & "', '" & txtCX.Text & "','" & Me.Tag & "','" & Format(dtpFecha.Value, "yyyy-mm-dd") & "','" & txtSerie.Text & "','" & txtDocumentoR.Text & "','" & txtAutorizacion.Text & "', " & _
             TxtTotal.Text & ", CURRENT_TIMESTAMP, '" & strUsuario & "')"
    clsSql.Ejecutar strSql, "M"
    strSql = "COMMIT TRAN"
    clsSql.Ejecutar strSql, "M"
    With VSFG
        For i = 1 To .Rows - 1
            strSql = " INSERT INTO det_comp_ret (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_com_ret_fechamod,det_com_ret_usumod) " & _
                     " VALUES ('" & strEmpresa & "','" & txtCX.Text & "','" & Me.Tag & "', '" & .TextMatrix(i, 1) & "', " & _
                     FormatoD2(.TextMatrix(i, 2)) & "," & FormatoD2(.TextMatrix(i, 3)) & ", CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsSql.Ejecutar strSql, "M"
            If (VSFG.Tag <> "NO") Then
                strSql = " SELECT COALESCE(COUNT(*),0) as num " & _
                         " FROM det_asiento " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND asi_numasiento='" & VSFG.Tag & "' " & _
                         " AND cta_codigo='" & .TextMatrix(i, 5) & "'"
                clsSql.Ejecutar strSql
                If clsSql.adorec_Def("num") = 0 Then
                    strSql = " INSERT INTO det_asiento ( emp_codigo, asi_numasiento, cta_codigo, det_asi_debe, det_asi_haber, det_asi_fechamod, det_asi_usumod) " & _
                            " VALUES ('" & strEmpresa & "','" & VSFG.Tag & "','" & .TextMatrix(i, 5) & "', " & _
                            " '0', '" & FormatoD2(.TextMatrix(i, 4)) & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
                Else
                    strSql = " UPDATE det_asiento " & _
                             " SET det_asi_haber=det_asi_haber+" & FormatoD2(.TextMatrix(i, 4)) & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND asi_numasiento='" & VSFG.Tag & "' " & _
                             " AND cta_codigo='" & .TextMatrix(i, 5) & "'"
                End If
                clsSql.Ejecutar strSql, "M"
            End If
        Next i
    End With
    
    Dim obs As String
    obs = " - Retención: " & txtSerie.Text & "-" & txtDocumentoR.Text
    
    strSql = " UPDATE asiento " & _
             " SET asi_descripcion=CONCAT(asi_descripcion,'" & obs & "') " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & VSFG.Tag & "' "
    clsSql.Ejecutar strSql, "M"
    
    If (VSFG.Tag <> "NO") Then
        strSql = " UPDATE det_asiento " & _
                 " SET det_asi_haber=det_asi_haber-" & FormatoD2(TxtTotal.Text) & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & VSFG.Tag & "' " & _
                 " AND cta_codigo='" & txtDocumento.Tag & "' "
        clsSql.Ejecutar strSql, "M"
    End If
    DocElectronico "07", (txtCX.Text)
    MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
'    Dim RepReten As New frmReporte
'    RepReten.strNumero = txtCX.Text
'    RepReten.strAsiento = VSFG.Tag
'    RepReten.strTipo = Me.Tag
'    RepReten.strReporte = "rptRetencionDiario"
'    RepReten.Show

     'Actualiza la fecha
     Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
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
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsCod.Inicializar AdoConn, AdoConnMaster
    clsPgd.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    txtr = 0
    txtD = 0
    txtp = 0
    strFechaPago = ""
    strSql = " SELECT par_numero " & _
                 " FROM parametro " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND par_codigo='IVAC' "
    clsCta.Ejecutar strSql
    IVA = clsCta.adorec_Def("par_numero")
    strSql = " SELECT ret_codigo, ret_nombre,ret_porcentaje,ret_ctaconta" & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " AND ret_activo = 1 " & _
                 " AND ret_ctaconta!='' "
     clsCta.Ejecutar strSql

     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "ret_codigo, *ret_nombre", "ret_codigo")

End Sub

Private Sub txtValor_Change()
    VSFG.TextMatrix(1, 4) = txtValor
    txtValor = FormatoD2(txtValor)
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    TxtTotal.Text = 0
    For i = 1 To VSFG.Rows - 1
        TxtTotal.Text = FormatoD2(TxtTotal.Text) + FormatoD2(VSFG.TextMatrix(i, 4))
    Next i
    TxtTotal.Text = Format(TxtTotal.Text, "###0.00")
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

    ' only interesetd in left button
    If Button <> 1 Then Exit Sub

    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub

    If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub

    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        If r > 0 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pagos"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)

        'Recorro el FlexGrid para poner números a las filas

        If respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            PonerBotones
        Else
            VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes

    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            VSFG.AddItem ""
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        'hace la consulta para saber las cuentas contables que no tengan subcuentas
         strSql = " SELECT ret_codigo, ret_nombre,ret_porcentaje,ret_ctaconta,ret_gravara" & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " AND ret_activo = 1 " & _
                 " AND ret_ctaconta!='' "
         clsCta.Ejecutar strSql
    
         VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "ret_codigo, *ret_nombre", "ret_codigo")
    
    ' Asigna codigos de cuenta y nombres en el grid
        With VSFG
            If .TextMatrix(Row, Col) <> "" Then
                If Col = 1 Then
                     clsCta.Filtrar ("ret_codigo = '" & .TextMatrix(Row, 1) & "'")
                         .TextMatrix(Row, 3) = clsCta.adorec_Def("ret_porcentaje")
                         If UCase(clsCta.adorec_Def("ret_gravara")) = "IVA" Then
                            .TextMatrix(Row, 2) = FormatoD2(TxtIva.Text)
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "IVAPRODUCTOS" Then
                            .TextMatrix(Row, 2) = FormatoD2(FormatoD2(TxtIva.Text) * FormatoD2(txtSTIVAProd.Text) / (FormatoD2(txtSTIVAProd.Text) + FormatoD2(txtSTIVAServ.Text)))
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "IVASERVICIOS" Then
                            .TextMatrix(Row, 2) = FormatoD2(FormatoD2(TxtIva.Text) * FormatoD2(txtSTIVAServ.Text) / (FormatoD2(txtSTIVAProd.Text) + FormatoD2(txtSTIVAServ.Text)))
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "SUBTOTAL" Then
                            .TextMatrix(Row, 2) = FormatoD2(FormatoD2(txtSTcero.Text) + FormatoD2(txtSTIVAServ.Text) + FormatoD2(txtSTIVAProd.Text))
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "SUBTOTALPRODUCTOS" Then
                            .TextMatrix(Row, 2) = FormatoD2(txtSTIVAProd.Text)
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "SUBTOTALSERVICIOS" Then
                            .TextMatrix(Row, 2) = FormatoD2(txtSTIVAServ.Text)
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "IVA0%" Then
                            .TextMatrix(Row, 2) = FormatoD2(txtSTcero.Text)
                         ElseIf UCase(clsCta.adorec_Def("ret_gravara")) = "TOTAL" Then
                            .TextMatrix(Row, 2) = txtTotalDoc.Text
                         Else
                            .TextMatrix(Row, 2) = 0
                         End If
                         .TextMatrix(Row, 5) = clsCta.adorec_Def("ret_ctaconta")
                     clsCta.QuitarFiltro
                 End If
             End If
        End With
    End If
    VSFG.TextMatrix(Row, 4) = FormatoD2(Val(VSFG.TextMatrix(Row, 3)) * Val(VSFG.TextMatrix(Row, 2)) / 100)
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (VSFG.TextMatrix(VSFG.Row, 3) = "") Then
                VSFG.TextMatrix(VSFG.Row, 3) = 0
     ElseIf VSFG.TextMatrix(VSFG.Row, 4) = "" Then
                VSFG.TextMatrix(VSFG.Row, 4) = 0
     End If
End Sub
