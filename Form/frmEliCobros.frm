VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEliCobros 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anular Cobros"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmEliCobros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11475
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5835
         TabIndex        =   4
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Anular"
         Height          =   375
         Left            =   3795
         TabIndex        =   3
         Top             =   6240
         Width           =   1575
      End
      Begin VB.OptionButton optproveedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optcliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   11055
         _cx             =   1959349292
         _cy             =   1959332570
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEliCobros.frx":030A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1455
         Left            =   855
         TabIndex        =   7
         Top             =   4680
         Width           =   9120
         _cx             =   1959345879
         _cy             =   1959332358
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
         FormatString    =   $"frmEliCobros.frx":045C
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
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGCobros 
         Height          =   1935
         Left            =   555
         TabIndex        =   8
         Top             =   2640
         Width           =   6855
         _cx             =   1959341883
         _cy             =   1959333205
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEliCobros.frx":053E
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGRetencion 
         Height          =   1935
         Left            =   7635
         TabIndex        =   9
         Top             =   2640
         Width           =   3015
         _cx             =   1959335110
         _cy             =   1959333205
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEliCobros.frx":067E
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
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4320
         Picture         =   "frmEliCobros.frx":06FA
         ToolTipText     =   "Elimina una Fila"
         Top             =   720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4560
         Picture         =   "frmEliCobros.frx":0830
         Top             =   720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deudor:"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   255
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos de Pago"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmEliCobros"
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

Private clsSql As New clsConsulta
Private clsPer As New clsConsulta
Private strSQL As String
Private t As String

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    Dim i As Long
    Dim booPasar As Boolean
    Dim aux As String
    Dim Motivo As String
    booPasar = True
'    For i = 1 To VSFG.Rows - 1
'        VSFG.Select i, 1
'        If VSFG.CellBackColor <> &HC0FFFF Then
'            booPasar = False
'        End If
'    Next i
    If booPasar = True Then
        If MsgBox("Esta seguro de anular el cobro?", vbDefaultButton2 + vbQuestion + vbYesNo, "Cobros") = vbNo Then
            booPasar = False
        End If
    Else
        MsgBox "No puede anular hasta que este cuadrado de acuerdo al cobro"
    End If
    
    If booPasar = True Then
        For i = 1 To VSFG1.Rows - 1
            If Abs(VSFG1.TextMatrix(i, 1)) = 1 Then
                aux = VSFG1.TextMatrix(i, 2)
            End If
        Next i
        Motivo = ""
        While Motivo = ""
            Motivo = UCase(InputBox("Motivo de Anulacion", "Contabilidad"))
            If Motivo = "" Then
                If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Esta seguro de anular el cobro?", vbQuestion + vbYesNo, "Cobros") = vbNo Then
                    Exit Sub
                End If
            End If
        Wend
        Motivo = Motivo & vbNewLine & strUsuario & vbNewLine & HoyDia & " " & Format(Ahora, "HH:MM:SS")
        strSQL = " UPDATE det_doc_pago " & _
                 " SET det_doc_pag_debe=0,det_doc_pag_haber=0 " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND doc_pag_codigo='" & aux & "' "
        clsSql.Ejecutar strSQL, "M"
        strSQL = " UPDATE doc_pago " & _
                 " SET doc_pag_estado='ANULADO',doc_pag_valor=0,doc_pag_fechamod=CURRENT_TIMESTAMP,doc_pag_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND doc_pag_codigo='" & aux & "' "
        clsSql.Ejecutar strSQL, "M"
'        strSql = " UPDATE loc_facturacion " & _
'                 " SET doc_pag_codigo = NULL " & _
'                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
'                 " AND doc_pag_codigo='" & aux & "' "
'        clsSql.Ejecutar strSql
        
        For i = 1 To VSFGCobros.Rows - 1
            If Abs(VSFGCobros.TextMatrix(i, 1)) = 1 Then
                strSQL = " UPDATE pago " & _
                         " SET pag_monto=0,pag_observacion=CONCAT('ANULADO" & vbNewLine & Motivo & vbNewLine & "',pag_observacion) " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND pag_codigo='" & VSFGCobros.TextMatrix(i, 10) & "' " & _
                         " AND cue_p_c_tipo='C' AND cue_p_c_codigo='" & VSFGCobros.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSQL, "M"
                strSQL = " UPDATE cuenta_p_c " & _
                         " SET cue_p_c_pagado=0" & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
                         " AND cue_p_c_tipo='C' AND cue_p_c_codigo='" & VSFGCobros.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSQL, "M"
            End If
        Next i
        For i = 1 To VSFGRetencion.Rows - 1
            If Abs(VSFGRetencion.TextMatrix(i, 1)) = 1 Then
                strSQL = " DELETE FROM det_comp_ret " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND cue_p_c_tipo='C' AND cue_p_c_codigo='" & VSFGRetencion.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSQL, "M"
                strSQL = " DELETE FROM comprobante_retencion " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND cue_p_c_tipo='C' AND cue_p_c_codigo='" & VSFGRetencion.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSQL, "M"
                strSQL = " UPDATE cuenta_p_c " & _
                         " SET cue_p_c_pagado=0" & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
                         " AND cue_p_c_tipo='C' AND cue_p_c_codigo='" & VSFGRetencion.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSQL, "M"
            End If
        Next i
        MsgBox "Eliminación realizada con éxito"
        Unload Me
    End If
End Sub

Private Sub dcmbBeneficiario_Change()
    txtValor = 0
    t = "P"
    If Me.optcliente.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiario.MatchedWithList = True Then
        cmdEliminar.Enabled = True
        VSFG1.Enabled = True
        'Consulta para el grid sobre las cuentas por pagar del beneficiario seleccionado
        strSQL = " SELECT '0', doc_pag_codigo, IIF(doc_pago.tip_doc_pag_codigo is null or doc_pago.tip_doc_pag_codigo ='' ,'EFECTIVO',tipo_doc_pago.tip_doc_pag_nombre)as tip_doc_pag_nombre, " & _
                 " banco.ban_nombre, doc_pag_numero, doc_pag_fecha_doc, CONCAT(per_apellido,' ',per_nombre), doc_pag_valor, doc_pag_observacion, doc_pag_anticipo " & _
                 " FROM ((doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo" & _
                 " LEFT JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo) " & _
                 " LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo) " & _
                 " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' AND doc_pag_estado = 'GIRADO' " & _
                 " AND doc_pago.per_codigo='" & dcmbBeneficiario.BoundText & "' AND asi_numasiento='' " & _
                 " ORDER BY CONCAT(per_apellido,' ',per_nombre) "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set VSFG1.DataSource = clsSql.adorec_Def.DataSource
             VSFG1.ColDataType(1) = flexDTBoolean
            'ponerBotones
        Else
            VSFG1.Clear 1
            VSFG1.Rows = 2
        End If

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
    
    'Seleccionamos el proveedor de la tabla persona (P), que esta por defecto
    
    optcliente.Value = True
    strSQL = " SELECT cen_cos_codigo, cen_cos_nombre " & _
             " FROM centro_costo " & _
             " WHERE emp_codigo= '" & strEmpresa & "' " & _
             " ORDER BY cen_cos_nombre"
    clsPer.Ejecutar strSQL
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsPer.adorec_Def, "cen_cos_codigo,*cen_cos_nombre", "cen_cos_codigo")
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona " & _
             " INNER JOIN tipo_pedido " & _
             " ON tipo_pedido.emp_codigo=persona.emp_codigo " & _
             " AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
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
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona " & _
             " INNER JOIN tipo_pedido " & _
             " ON tipo_pedido.emp_codigo=persona.emp_codigo " & _
             " AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
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

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGCobros_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub
Private Sub VSFG1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 1 And Row > 0 Then
        
        If VSFG1.TextMatrix(Row, 1) = "-1" Then
            For i = 1 To VSFG1.Rows - 1
                If i <> Row And VSFG1.TextMatrix(i, 1) = "-1" Then
                    VSFG1.Select i, 1, i, 9
                    VSFG1.FillStyle = flexFillRepeat
                    VSFG1.CellBackColor = &HFFFFFF
                    VSFG1.TextMatrix(i, 1) = "0"
                End If
            Next
            VSFG1.Select Row, 1, Row, 9
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HC0FFFF
            VSFG1.Select Row, 1
            Llenar_Grid (Row)
            If Val(VSFG1.TextMatrix(Row, 10)) = 1 Then
                For i = 1 To VSFG.Rows - 1
                    VSFG.Select i, 1, i, 4
                    VSFG.FillStyle = flexFillRepeat
                    VSFG.CellBackColor = &HC0FFFF
                Next i
            End If
        ElseIf VSFG1.TextMatrix(Row, 1) = "0" Then
            VSFG1.Select Row, 1, Row, 9
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HFFFFFF
            VSFG1.Select Row, 1
            VSFG.Clear 1
            VSFG.Rows = 2
            VSFGCobros.Clear 1
            VSFGCobros.Rows = 1
            VSFGRetencion.Clear 1
            VSFGRetencion.Rows = 1
        End If
    End If
End Sub
Private Sub VSFGCobros_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 1 And Row > 0 Then
        
        If VSFGCobros.TextMatrix(Row, 1) = "-1" Then
            VSFGCobros.Select Row, 1, Row, 7
            VSFGCobros.FillStyle = flexFillRepeat
            VSFGCobros.CellBackColor = &HC0FFFF
            VSFGCobros.Select Row, 1
            MuestraRetentecion (Row)
        ElseIf VSFGCobros.TextMatrix(Row, 1) = "0" Then
            VSFGCobros.Select Row, 1, Row, 7
            VSFGCobros.FillStyle = flexFillRepeat
            VSFGCobros.CellBackColor = &HFFFFFF
            VSFGCobros.Select Row, 1
            QuitaRetentecion (Row)
        End If
        CuadraBanco
    End If
End Sub
Private Sub VSFGRetencion_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 1 And Row > 0 Then
        If VSFGRetencion.TextMatrix(Row, 1) = "-1" Then
            VSFGRetencion.Select Row, 1, Row, 3
            VSFGRetencion.FillStyle = flexFillRepeat
            VSFGRetencion.CellBackColor = &HC0FFFF
            VSFGRetencion.Select Row, 1
        ElseIf VSFGRetencion.TextMatrix(Row, 1) = "0" Then
            VSFGRetencion.Select Row, 1, Row, 3
            VSFGRetencion.FillStyle = flexFillRepeat
            VSFGRetencion.CellBackColor = &HFFFFFF
            VSFGRetencion.Select Row, 1
        End If
        CuadraRetencion
    End If
End Sub
Private Sub VSFGRetencion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub
Private Sub CuadraRetencion()
    Dim i As Long
    Dim dblRet As Double
    dblRet = 0
    For i = 1 To VSFGRetencion.Rows - 1
        If Abs(Val(VSFGRetencion.TextMatrix(i, 1))) = 1 Then
            dblRet = FormatoD2(dblRet + FormatoD2(VSFGRetencion.TextMatrix(i, 3)))
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) <> "" And FormatoD2(VSFG.TextMatrix(i, 3)) <> 0 Then
            dblRet = FormatoD2(dblRet - FormatoD2(VSFG.TextMatrix(i, 3)))
            VSFG.Select i, 1, i, 4
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
        End If
    Next i
    If FormatoD2(dblRet) = 0 Then
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 1) <> "" And FormatoD2(VSFG.TextMatrix(i, 3)) <> 0 Then
                VSFG.Select i, 1, i, 4
                VSFG.FillStyle = flexFillRepeat
                VSFG.CellBackColor = &HC0FFFF
            End If
        Next i
    End If
    VSFG.Select 1, 1
End Sub
Private Sub CuadraBanco()
    Dim i As Long
    Dim dblCobro As Double
    dblCobro = 0
    For i = 1 To VSFGCobros.Rows - 1
        If Abs(Val(VSFGCobros.TextMatrix(i, 1))) = 1 Then
            dblCobro = dblCobro + FormatoD2(VSFGCobros.TextMatrix(i, 7))
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        If (FormatoD2(VSFG.TextMatrix(i, 3)) = FormatoD2(dblCobro) And VSFG.TextMatrix(i, 1) = "*") Or FormatoD2(VSFG.TextMatrix(i, 3)) = 0 Then
            VSFG.Select i, 1, i, 4
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HC0FFFF
            
        Else
            VSFG.Select i, 1, i, 4
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
        End If
        VSFG.Select 1, 1
    Next i
End Sub
Private Sub MuestraRetentecion(num As Long)
    strSQL = " SELECT cue_p_c_codigo,com_ret_total" & _
             " FROM comprobante_retencion " & _
             " WHERE cue_p_c_codigo = '" & VSFGCobros.TextMatrix(num, 2) & _
             "' AND cue_p_c_tipo='C' AND emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFGRetencion.AddItem "" & vbTab & "0" & vbTab & clsSql.adorec_Def("cue_p_c_codigo") & vbTab & clsSql.adorec_Def("com_ret_total")
    End If
End Sub

Private Sub QuitaRetentecion(num As Long)
    Dim i As Long
    For i = 1 To VSFGRetencion.Rows - 1
        If VSFGCobros.TextMatrix(num, 2) = VSFGRetencion.TextMatrix(i, 2) Then
            VSFGRetencion.RemoveItem i
            i = VSFGRetencion.Rows
        End If
    Next i
End Sub

Private Sub Llenar_Grid(num As Long)
    strSQL = " SELECT det_doc_pago.cta_codigo as codigo, IIF(det_doc_pago.cta_codigo= '*' , 'CAJA', ctaconta.cta_nombre) as nombre, det_doc_pag_debe as debe, det_doc_pag_haber as haber,COALESCE(cen_cos_codigo,'') as cen_cos_codigo" & _
             " FROM det_doc_pago LEFT JOIN ctaconta ON det_doc_pago.emp_codigo =ctaconta.emp_codigo " & _
             " AND det_doc_pago.cta_codigo = ctaconta.cta_codigo " & _
             " WHERE doc_pag_codigo = '" & VSFG1.TextMatrix(num, 2) & _
             "' AND det_doc_pago.emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    VSFG.Clear 1
    VSFG.Rows = 2
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    VSFGCobros.Clear 1
    VSFGCobros.Rows = 1
    VSFGRetencion.Clear 1
    VSFGRetencion.Rows = 1
    If Val(VSFG1.TextMatrix(num, 10)) <> 1 Then
        strSQL = " SELECT '0', cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_fechaemision, pag_fecha, cue_p_c_valor, pag_monto,pag_no_doc,pag_observacion,pag_codigo " & _
                 " FROM (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo) " & _
                 " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & _
                 "' AND cuenta_p_c.emp_codigo = '" & strEmpresa & _
                 "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND doc_pag_codigo='" & VSFG1.TextMatrix(num, 2) & "' "
        clsSql.Ejecutar strSQL
        Set VSFGCobros.DataSource = clsSql.adorec_Def.DataSource
    End If
End Sub
