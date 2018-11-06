VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmModCobros 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Fecha de Cobros"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmModCobros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
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
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "ACTUALIZAR"
         Height          =   375
         Left            =   9120
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   4320
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   315
            Left            =   2160
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Value           =   42292.5835069444
         End
         Begin NEED2.dtpFecha dtpFechaCh 
            Height          =   315
            Left            =   2160
            TabIndex        =   14
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Value           =   42292.5835069444
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de documento:"
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
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Cobro:"
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
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2355
         TabIndex        =   4
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   315
         TabIndex        =   3
         Top             =   5640
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
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11055
         _cx             =   19500
         _cy             =   4471
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmModCobros.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGCobros 
         Height          =   1815
         Left            =   4680
         TabIndex        =   7
         Top             =   4320
         Width           =   6375
         _cx             =   2008099821
         _cy             =   2008091777
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
         FormatString    =   $"frmModCobros.frx":0495
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
      Begin MSDataListLib.DataCombo dcmbDocumento 
         Height          =   315
         Left            =   825
         TabIndex        =   15
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker Fecha1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   18
         Top             =   735
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60293123
         CurrentDate     =   37463
      End
      Begin MSComCtl2.DTPicker Fecha2 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   19
         Top             =   735
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60293123
         CurrentDate     =   37463
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL Pag:"
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
         Left            =   8400
         TabIndex        =   25
         Top             =   6150
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CH:"
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
         Left            =   6240
         TabIndex        =   23
         Top             =   3990
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   4320
         TabIndex        =   21
         Top             =   735
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   6600
         TabIndex        =   20
         Top             =   735
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc:"
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
         TabIndex        =   16
         Top             =   720
         Width           =   675
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   3960
         Picture         =   "frmModCobros.frx":05D6
         ToolTipText     =   "Elimina una Fila"
         Top             =   4200
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4200
         Picture         =   "frmModCobros.frx":070C
         Top             =   4200
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmModCobros"
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

Private Sub cmdAceptar_Click()
    Dim i As Long
    Dim booPasar As Boolean
    Dim aux As String
    Dim Asiento As String
    Dim Motivo As String
    Dim FechasDetalle As String
    booPasar = True
    If MsgBox("Esta seguro de cambiar las fechas del cobro?", vbDefaultButton2 + vbQuestion + vbYesNo, "Cobros") = vbNo Then
        booPasar = False
    End If
    If booPasar = True Then
        Motivo = ""
        While Motivo = ""
            Motivo = UCase(InputBox("Motivo del cambio de Fecha", "Contabilidad"))
            If Motivo = "" Then
                If MsgBox("Debe ingresar un motivo para realizar el cambio" & vbNewLine & "Esta seguro del cambio?", vbQuestion + vbYesNo, "Cobros") = vbNo Then
                    Exit Sub
                End If
            End If
        Wend
        
        For i = 1 To VSFG1.Rows - 1
            If Abs(Val(VSFG1.TextMatrix(i, 6))) = 1 Then
                aux = VSFG1.TextMatrix(i, 7)
                Asiento = VSFG1.TextMatrix(i, 12)
                FechasDetalle = "FECHA DE COBRO ORIGINAL: " & VSFG1.TextMatrix(i, 4) & vbNewLine & _
                                "FECHA DE COBRO NUEVA: " & dtpFecha.Value & vbNewLine & _
                                "FECHA DE DOC ORIGINAL: " & VSFG1.TextMatrix(i, 5) & vbNewLine & _
                                "FECHA DE DOC NUEVA: " & dtpFechaCh.Value
        
                strSQL = " UPDATE doc_pago " & _
                         " SET doc_pag_fecha_recepcion='" & dtpFecha.Value & "',doc_pag_fecha_efec='" & dtpFecha.Value & "',doc_pag_fecha_doc='" & dtpFechaCh.Value & "'," & _
                         " doc_pag_observacion=CONCAT('" & Motivo & vbNewLine & FechasDetalle & vbNewLine & "',doc_pag_observacion)," & _
                         " doc_pag_fechamod=CURRENT_TIMESTAMP,doc_pag_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND doc_pag_codigo='" & aux & "' "
                clsSql.Ejecutar strSQL, "M"
                strSQL = " UPDATE asiento " & _
                         " SET asi_fecha='" & dtpFecha.Value & "'," & _
                         " asi_descripcion=CONCAT('" & Motivo & vbNewLine & FechasDetalle & vbNewLine & "',asi_descripcion)" & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND asi_numasiento='" & Asiento & "' "
                clsSql.Ejecutar strSQL, "M"
                strSQL = " UPDATE pago " & _
                         " SET pag_observacion=CONCAT('" & Motivo & vbNewLine & FechasDetalle & vbNewLine & "',pag_observacion), " & _
                         " pag_fecha='" & dtpFecha.Value & "' " & _
                         " WHERE emp_codigo = '" & strEmpresa & "' " & _
                         " AND doc_pag_codigo='" & aux & "' " & _
                         " AND asi_numasiento='" & Asiento & "' " & _
                         " AND cue_p_c_tipo='C' "
                clsSql.Ejecutar strSQL, "M"
            End If
        Next i
        
        MsgBox "Cambio realizado con éxito"
        VSFG1.Clear 1
        VSFG1.Rows = 1
        
        VSFGCobros.Clear 1
        VSFGCobros.Rows = 1
        dcmbBeneficiario_Validate False
        'Unload Me
    End If
End Sub

Private Sub cmdActualizar_Click()
    Dim i As Long
    txtValor = 0
    t = "P"
    If Me.optcliente.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiario.MatchedWithList = True Then
        'cmdEliminar.Enabled = True
        VSFG1.Enabled = True
        'Consulta para el grid sobre las cuentas por pagar del beneficiario seleccionado
        If dcmbDocumento.MatchedWithList = True Then
            strSQL = " SELECT doc_pag_numero, iif(doc_pago.tip_doc_pag_codigo ='' ,'EFECTIVO',tipo_doc_pago.tip_doc_pag_nombre)as tip_doc_pag_nombre, " & _
                     " COALESCE(banco.ban_nombre,'') as ban_nombre, doc_pag_fecha_recepcion, doc_pag_fecha_doc,'0', doc_pag_codigo, CONCAT(per_apellido,' ',per_nombre), doc_pag_valor, doc_pag_observacion, doc_pag_anticipo,COALESCE(asi_numasiento2,asi_numasiento) as asien " & _
                     " FROM ((doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                     " INNER JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo) " & _
                     " LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo) " & _
                     " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' " & _
                     " AND (doc_pago.per_codigo='" & dcmbBeneficiario.BoundText & "' OR doc_pago.per_codigo_ch='" & dcmbBeneficiario.BoundText & "')  " & _
                     " AND doc_pago.tip_doc_pag_codigo='" & dcmbDocumento.BoundText & "'" & _
                     " AND doc_pago.doc_pag_fecha_doc between '" & Format(Fecha1.Value, "YYYY-mm-dd") & "' AND '" & Format(Fecha2.Value, "YYYY-mm-dd") & "'" & _
                     " ORDER BY doc_pag_codigo "
        Else
            strSQL = " SELECT doc_pag_numero, 'EFECTIVO' as tip_doc_pag_nombre, " & _
                     " 'RYB', doc_pag_fecha_recepcion, doc_pag_fecha_doc,'0', doc_pag_codigo, CONCAT(per_apellido,' ',per_nombre), doc_pag_valor, doc_pag_observacion, doc_pag_anticipo,COALESCE(asi_numasiento2,asi_numasiento) as asien " & _
                     " FROM doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                     " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' " & _
                     " AND (doc_pago.per_codigo='" & dcmbBeneficiario.BoundText & "' OR doc_pago.per_codigo_ch='" & dcmbBeneficiario.BoundText & "')  " & _
                     " AND (doc_pago.tip_doc_pag_codigo='' or doc_pago.tip_doc_pag_codigo IS NULL)" & _
                     " AND doc_pago.doc_pag_fecha_doc between '" & Format(Fecha1.Value, "YYYY-mm-dd") & "' AND '" & Format(Fecha2.Value, "YYYY-mm-dd") & "'" & _
                     " ORDER BY doc_pag_codigo "
        
        End If
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set VSFG1.DataSource = clsSql.adorec_Def.DataSource
             For i = 1 To VSFG1.Cols - 1
                VSFG1.TextMatrix(1, i) = clsSql.adorec_Def(i - 1)
             Next i
             If Me.dcmbDocumento.BoundText <> "" Then
                VSFG1.MergeCol(0) = True: VSFG1.MergeCol(1) = True: VSFG1.MergeCol(2) = True: VSFG1.MergeCol(3) = True: VSFG1.MergeCol(4) = True: VSFG1.MergeCol(5) = True: VSFG1.MergeCol(6) = True
             Else
                VSFG1.MergeCol(0) = True: VSFG1.MergeCol(1) = True: VSFG1.MergeCol(2) = True: VSFG1.MergeCol(3) = True: VSFG1.MergeCol(4) = True: VSFG1.MergeCol(5) = True
             End If
            'ponerBotones
        Else
            VSFG1.Clear 1
            VSFG1.Rows = 2
        End If
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBeneficiario_Validate(Cancel As Boolean)
    cmdActualizar_Click
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
    Fecha1.Value = HoyDia
    Fecha2.Value = HoyDia
    
    strSQL = " SELECT tip_doc_pag_codigo, tip_doc_pag_nombre " & _
             " FROM tipo_doc_pago "
    clsSql.Ejecutar strSQL
    
    Set dcmbDocumento.RowSource = clsSql.adorec_Def.DataSource
    dcmbDocumento.ListField = "tip_doc_pag_nombre"
    dcmbDocumento.BoundColumn = "tip_doc_pag_codigo"
    
    optcliente.Value = True
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
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
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    
    Dim i As Long
    Me.MousePointer = 11
    If Col = 6 Then
        VSFG1.GetMergedRange Row, Col, r1, c1, r2, c2
        For i = 1 To VSFG1.Rows - 1
            If i < r1 Or r2 < i Then
                VSFG1.Select i, 1, i, VSFG1.Cols - 1
                VSFG1.FillStyle = flexFillRepeat
                VSFG1.CellBackColor = &HFFFFFF
                VSFG1.TextMatrix(i, Col) = "0"
            Else
                VSFG1.TextMatrix(i, Col) = VSFG1.TextMatrix(Row, Col)
            End If
        Next i
        
        
        If Col = 6 And Row > 0 Then
            
            If VSFG1.TextMatrix(Row, Col) = "-1" Then
    '            For i = 1 To VSFG1.Rows - 1
    '                If i <> Row And VSFG1.TextMatrix(i, 1) = "-1" Then
    '                    VSFG1.Select i, 1, i, 9
    '                    VSFG1.FillStyle = flexFillRepeat
    '                    VSFG1.CellBackColor = &HFFFFFF
    '                    VSFG1.TextMatrix(i, 1) = "0"
    '                End If
    '            Next
                VSFG1.Select r1, 1, r2, VSFG1.Cols - 1
                VSFG1.FillStyle = flexFillRepeat
                VSFG1.CellBackColor = &HC0FFFF
                VSFG1.Select Row, Col
    '            Llenar_Grid (Row)
    '            If Val(VSFG1.TextMatrix(Row, 10)) = 1 Then
    '                For i = 1 To VSFG.Rows - 1
    '                    VSFG.Select i, 1, i, 4
    '                    VSFG.FillStyle = flexFillRepeat
    '                    VSFG.CellBackColor = &HC0FFFF
    '                Next i
    '            End If
            ElseIf VSFG1.TextMatrix(Row, Col) = "0" Then
                VSFG1.Select r1, 1, r2, VSFG1.Cols - 1
                VSFG1.FillStyle = flexFillRepeat
                VSFG1.CellBackColor = &HFFFFFF
                VSFG1.Select Row, Col
    '            VSFGCobros.Clear 1
    '            VSFGCobros.Rows = 1
            End If
            Llenar_Grid (Row)
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub VSFG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = 6 Then
        VSFG1.AutoSearch = flexSearchNone
    Else
        VSFG1.AutoSearch = flexSearchFromTop
    End If
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 6 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGCobros_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub
Private Sub VSFGCobros_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    Dim i As Long
'    If Col = 1 And Row > 0 Then
'
'        If VSFGCobros.TextMatrix(Row, 1) = "-1" Then
'            VSFGCobros.Select Row, 1, Row, 7
'            VSFGCobros.FillStyle = flexFillRepeat
'            VSFGCobros.CellBackColor = &HC0FFFF
'            VSFGCobros.Select Row, 1
'            MuestraRetentecion (Row)
'        ElseIf VSFGCobros.TextMatrix(Row, 1) = "0" Then
'            VSFGCobros.Select Row, 1, Row, 7
'            VSFGCobros.FillStyle = flexFillRepeat
'            VSFGCobros.CellBackColor = &HFFFFFF
'            VSFGCobros.Select Row, 1
'            QuitaRetentecion (Row)
'        End If
'        CuadraBanco
'    End If
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
            dblRet = dblRet + FormatoD2(VSFGRetencion.TextMatrix(i, 3))
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) <> "*" And FormatoD2(VSFG.TextMatrix(i, 3)) <> 0 Then
            dblRet = dblRet - FormatoD2(VSFG.TextMatrix(i, 3))
            VSFG.Select i, 1, i, 4
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
        End If
    Next i
    If FormatoD2(dblRet) = 0 Then
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 1) <> "*" And FormatoD2(VSFG.TextMatrix(i, 3)) <> 0 Then
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
    Dim i As Long
    Dim strDocPag As String
    Dim f1 As String
    Dim f2 As String
    Dim totCH As Double
    f1 = ""
    f2 = ""
    totCH = 0
    For i = 1 To VSFG1.Rows - 1
        If Abs(Val(VSFG1.TextMatrix(i, 6))) = 1 Then
        
            If f1 = "" And f2 = "" Then
                f1 = VSFG1.TextMatrix(num, 4)
                f2 = VSFG1.TextMatrix(num, 5)
            Else
                If f1 <> VSFG1.TextMatrix(num, 4) And f2 <> VSFG1.TextMatrix(num, 5) Then
                    VSFG1.TextMatrix(i, 6) = 0
                    VSFG1.Select i, 1, i, VSFG1.Cols - 1
                    VSFG1.FillStyle = flexFillRepeat
                    VSFG1.CellBackColor = &HFFFFFF
                    VSFG1.Select i, 6
                End If
            End If
            strDocPag = strDocPag & "'" & VSFG1.TextMatrix(i, 7) & "',"
            totCH = totCH + VSFG1.TextMatrix(i, 9)
            
        End If
    Next
    VSFGCobros.Clear 1
    VSFGCobros.Rows = 1
    If Len(strDocPag) > 2 Then
        strDocPag = Left(strDocPag, Len(strDocPag) - 1)
        If Val(VSFG1.TextMatrix(num, 10)) <> 1 Then
            strSQL = " SELECT '0', cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_fechaemision, pag_fecha, cue_p_c_valor, pag_monto,pag_no_doc,pag_observacion,pag_codigo " & _
                     " FROM (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo) " & _
                     " WHERE " & _
                     " cuenta_p_c.emp_codigo = '" & strEmpresa & _
                     "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND doc_pag_codigo in (" & strDocPag & ") "
            clsSql.Ejecutar strSQL
            Set VSFGCobros.DataSource = clsSql.adorec_Def.DataSource
            dtpFecha.Value = VSFG1.TextMatrix(num, 4)
            dtpFechaCh.Value = VSFG1.TextMatrix(num, 5)
        End If
    End If
    TxtTotal.Text = totCH
    totCH = 0
    For i = 1 To Me.VSFGCobros.Rows - 1
        totCH = totCH + VSFGCobros.TextMatrix(i, 7)
    Next i
    txtPago.Text = totCH
End Sub
