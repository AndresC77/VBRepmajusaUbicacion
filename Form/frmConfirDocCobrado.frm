VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConfirDocCobrado 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmacion de Cobros"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmConfirDocCobrado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11475
   Begin VB.CommandButton cmdProtesto 
      Caption         =   "Ch. &Protestado"
      Height          =   375
      Left            =   4095
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5805
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
   End
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
      Height          =   7215
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11295
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5617
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   6720
         Width           =   1935
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7537
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   1230
         TabIndex        =   16
         Top             =   3720
         Width           =   9015
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   285
            Left            =   6360
            TabIndex        =   26
            Top             =   240
            Width           =   2415
            _extentx        =   4260
            _extenty        =   503
         End
         Begin VB.TextBox txtDoc 
            Height          =   285
            Left            =   6360
            TabIndex        =   7
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtValor 
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
            Height          =   315
            Left            =   6360
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   960
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   1680
            TabIndex        =   4
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbTipo 
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha comprobante:"
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
            Left            =   4800
            TabIndex        =   25
            Top             =   270
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de nota:"
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
            Left            =   285
            TabIndex        =   24
            Top             =   975
            Width           =   930
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
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
            Left            =   285
            TabIndex        =   20
            Top             =   285
            Width           =   510
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Doc:"
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
            Left            =   4800
            TabIndex        =   19
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria:"
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
            Left            =   285
            TabIndex        =   18
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor del recargo:"
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
            Left            =   4800
            TabIndex        =   17
            Top             =   990
            Width           =   1305
         End
      End
      Begin VB.OptionButton optproveedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3510
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optcliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2310
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   6270
         TabIndex        =   2
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
         FormatString    =   $"frmConfirDocCobrado.frx":030A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGCobros 
         Height          =   1335
         Left            =   2310
         TabIndex        =   13
         Top             =   2280
         Width           =   6855
         _cx             =   12091
         _cy             =   2355
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConfirDocCobrado.frx":045A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1335
         Left            =   1170
         TabIndex        =   9
         Top             =   5280
         Width           =   9120
         _cx             =   16087
         _cy             =   2355
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
         FormatString    =   $"frmConfirDocCobrado.frx":0580
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
      Begin VB.Label lbltotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Height          =   255
         Left            =   4530
         TabIndex        =   23
         Top             =   6735
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4320
         Picture         =   "frmConfirDocCobrado.frx":065A
         ToolTipText     =   "Elimina una Fila"
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4560
         Picture         =   "frmConfirDocCobrado.frx":0790
         Top             =   600
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
         Left            =   5550
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   600
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmConfirDocCobrado"
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

Private Sub cmdProtesto_Click()
    Dim i As Long
    Dim strCod As String
    Dim strMaximo As String
    Dim ff As String
    Dim num As Long
    Dim dblVal As Double
    If dcmbBanco.MatchedWithList And dcmbCuenta.MatchedWithList And dcmbTipo.MatchedWithList And VSFG.Rows > 2 Then
        If FormatoD2(txtTotalHaber.Text) = FormatoD2(txtTotalDebe.Text) Then
            If MsgBox("Esta seguro del Cheque a protestar?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Cobro") = vbYes Then
                num = 0
                For i = 1 To vsfgDocumento.Rows - 1
                    If Abs(Val(vsfgDocumento.TextMatrix(i, 1))) = 1 Then
                        num = i
                        i = vsfgDocumento.Rows + 1
                    End If
                Next i
                If num <> 0 Then
                    'nota de debito y asiento
                    strSQL = " SELECT COALESCE(max(not_d_c_codigo),0) as num " & _
                             " FROM nota_d_c" & _
                             " WHERE emp_codigo = '" & strEmpresa & _
                             "' AND tip_not_d_c='D'" & _
                             " GROUP BY emp_codigo"
                    clsSql.Ejecutar strSQL
                    If clsSql.adorec_Def.EOF Then
                        strCod = 1
                    Else
                        strCod = Val(clsSql.adorec_Def("num")) + 1
                    End If
                    Dim clsAsiento As New clsContable
                    clsAsiento.Inicializar AdoConn, AdoConnMaster
                    clsAsiento.NuevoAsiento "D", Format(dtpFecha.Value, "yyyy-mm-dd"), 0, 0, FormatoD2(txtTotalDebe.Text), "DEBITO POR CH PROTESTADO No. " & vsfgDocumento.TextMatrix(num, 5) & vbNewLine & vsfgDocumento.TextMatrix(num, 10), True
                    strSQL = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha,not_d_c_descripcion, not_d_c_monto,asi_numasiento, not_d_c_fechamod, not_d_c_usumod) " & _
                             " VALUES ('D','" & strCod & "', '" & dcmbCuenta.Text & "', '" & dcmbBanco.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtDoc.Text & "','" & ff & "','DEBITO POR CH PROTESTADO No. " & vsfgDocumento.TextMatrix(num, 5) & vbNewLine & vsfgDocumento.TextMatrix(num, 10) & "','" & FormatoD2(VSFG.TextMatrix(1, 4)) + FormatoD2(txtValor.Text) & "','" & clsAsiento.NumAsiento & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsSql.Ejecutar (strSQL), "M"
                    For i = 1 To VSFG.Rows - 1
                        clsAsiento.NuevoDetAsiento VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 5), FormatoD2(VSFG.TextMatrix(i, 3)), FormatoD2(VSFG.TextMatrix(i, 4))
                    Next i
                    
                    'pago negativo
                    For i = 1 To VSFGCobros.Rows - 1
                        'Calcula el máximo codigo de pago para la cuenta
                         strSQL = " SELECT max(pag_codigo) as pag " & _
                                  " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                                  "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                                  "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                                  " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFGCobros.TextMatrix(i, 1) & "' AND cue_p_c_egr_codigo = '" & VSFGCobros.TextMatrix(i, 2) & "' AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'C'" & _
                                  " GROUP BY pago.emp_codigo"
                        clsSql.Ejecutar strSQL
                        If clsSql.adorec_Def.EOF Then
                            strCod = 1
                        Else
                            strCod = clsSql.adorec_Def("pag") + 1
                        End If
                        If i = 1 Then
                            dblVal = (FormatoD2(VSFGCobros.TextMatrix(i, 6)) + FormatoD2(txtValor.Text)) * -1
                        Else
                            dblVal = FormatoD2(VSFGCobros.TextMatrix(i, 6)) * -1
                        End If
                        strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, pag_no_doc, pag_observacion,doc_pag_codigo,asi_numasiento, pag_fechamod, pag_usumod) " & _
                                 " VALUES ('" & strEmpresa & "', '" & VSFGCobros.TextMatrix(i, 1) & "', 'C', '" & Val(strCod) & "', '" & Format(dtpFecha.Value, "yyyy-mm-dd") & "', '" & FormatoD2(dblVal) & "', '" & txtDoc.Text & "', 'DEBITO POR CH PROTESTADO No. " & vsfgDocumento.TextMatrix(num, 5) & "', " & _
                                 " '" & strCod & "','" & clsAsiento.NumAsiento & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                        clsSql.Ejecutar strSQL, "M"
                    'actualiza no pagado a la cuenta_p_c
                        strSQL = " UPDATE cuenta_p_c " & _
                                 " SET cue_p_c_fechapago='" & Format(dtpFecha.Value, "yyyy-mm-dd") & "', cue_p_c_pagado = 0 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                 " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & VSFGCobros.TextMatrix(i, 1) & "' AND cue_p_c_egr_codigo = '" & VSFGCobros.TextMatrix(i, 2) & "' AND emp_codigo = '" & strEmpresa & "' "
                        clsSql.Ejecutar strSQL, "M"
                    Next i
                    'actualiza estado documento
                    strSQL = " UPDATE doc_pago " & _
                             " SET doc_pag_pendiente=-1 , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                             " WHERE doc_pag_codigo= '" & vsfgDocumento.TextMatrix(num, 2) & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsSql.Ejecutar strSQL, "M"
                    frmReporte.strAsiento = clsAsiento.NumAsiento
                    frmReporte.strReporte = "rptAsiento"
                    frmReporte.Show
                    
                    Unload Me
                Else
                    MsgBox "Seleccione un Documente para anular", vbCritical, "Confirmar Cobro"
                End If
            End If
        Else
            MsgBox "Los datos del asiento no esta cuadrado", vbCritical, "Confirmar Cobro"
        End If
    Else
        MsgBox "Alguno de los datos del banco no esta ingresado correctamente", vbCritical, "Confirmar Cobro"
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBanco_Change()
    If dcmbBanco.MatchedWithList Then
        strSQL = " SELECT cta_ban_numero, cta_ban_ctaconta " & _
                 " FROM cta_banco " & _
                 " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
                 " AND emp_codigo = '" & strEmpresa & "' " & _
                 " ORDER BY cta_ban_numero "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set dcmbCuenta.RowSource = clsSql.adorec_Def.DataSource
            dcmbCuenta.ListField = "cta_ban_numero"
            dcmbCuenta.BoundColumn = "cta_ban_ctaconta"
        Else
            dcmbCuenta = ""
        End If
    End If
End Sub

Private Sub dcmbBeneficiario_Change()
    txtValor = 0
    t = "P"
    If Me.optcliente.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiario.MatchedWithList = True Then
        cmdProtesto.Enabled = True
        vsfgDocumento.Enabled = True
        'Consulta para el grid sobre las cuentas por pagar del beneficiario seleccionado
        strSQL = " SELECT '0', doc_pag_codigo, iif( (doc_pago.tip_doc_pag_codigo is not null) or doc_pago.tip_doc_pag_codigo ='' ,'EFECTIVO',tipo_doc_pago.tip_doc_pag_nombre)as tip_doc_pag_nombre, " & _
                 " banco.ban_nombre, doc_pag_numero, doc_pag_fecha_doc, per_codigo, doc_pag_valor,iif(doc_pag_pendiente=1,'PostFechado','-') as doc_pag_pendienteN, doc_pag_observacion  " & _
                 " FROM ((doc_pago LEFT JOIN tipo_doc_pago ON doc_pago.tip_doc_pag_codigo = tipo_doc_pago.tip_doc_pag_codigo) " & _
                 "                 LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo) " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' AND doc_pag_pendiente != -1 " & _
                 " AND per_codigo='" & dcmbBeneficiario.BoundText & "'  " & _
                 " ORDER BY per_codigo "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set vsfgDocumento.DataSource = clsSql.adorec_Def.DataSource
            vsfgDocumento.ColDataType(1) = flexDTBoolean
            PonerBotones
        Else
            vsfgDocumento.Clear 1
            vsfgDocumento.Rows = 1
        End If
    End If
End Sub

Private Sub dcmbCuenta_Change()
    If dcmbCuenta.MatchedWithList Then
        strSQL = " SELECT cta_codigo, cta_nombre " & _
                 " FROM ctaconta " & _
                 " WHERE cta_codigo='" & dcmbCuenta.BoundText & "'"
        clsSql.Ejecutar strSQL
        VSFG.TextMatrix(1, 1) = clsSql.adorec_Def("cta_codigo")
        VSFG.TextMatrix(1, 2) = clsSql.adorec_Def("cta_nombre")
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
    
    dtpFecha.Value = HoyDia
'     consulta para saber los  bancos existentes
    strSQL = " SELECT DISTINCT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON banco.ban_codigo=cta_banco.ban_codigo " & _
             " ORDER BY banco.ban_codigo"
    clsSql.Ejecutar strSQL

    If clsSql.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsSql.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        dcmbBanco = ""
    End If

'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsSql.Ejecutar strSQL

     VSFG.ColComboList(5) = VSFG.BuildComboList(clsSql.adorec_Def, "cen_cos_codigo,*cen_cos_nombre", "cen_cos_codigo")
'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsSql.Ejecutar strSQL

     VSFG.ColComboList(1) = VSFG.BuildComboList(clsSql.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsSql.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")
    

    strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'D'" & _
             " ORDER BY tip_not_codigo"
    clsSql.Ejecutar strSQL
    
    Set dcmbTipo.RowSource = clsSql.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
    
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

Private Sub txtValor_Validate(Cancel As Boolean)
    Dim num As Long
    Dim i As Long
    If IsNumeric(txtValor.Text) Then
        If vsfgDocumento.Rows > 1 Then
            For i = 1 To vsfgDocumento.Rows - 1
                If Abs(Val(vsfgDocumento.TextMatrix(i, 1))) = 1 Then
                    num = i
                    i = vsfgDocumento.Rows + 1
                End If
            Next i
            VSFG.TextMatrix(1, 4) = FormatoD2(vsfgDocumento.TextMatrix(num, 8)) + FormatoD2(txtValor.Text)
        Else
            VSFG.TextMatrix(1, 4) = FormatoD2(txtValor.Text)
        End If
        txtValor.Text = Format(txtValor.Text, "#0.00")
    Else
        MsgBox "El Valor no es un numero válido", vbCritical, "Confirmacion de Cobros"
        Cancel = True
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
            MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
            VSFG.TextMatrix(Row, 3) = ""
            VSFG.TextMatrix(Row, 4) = ""


        ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe

        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = intDato
        End If

        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 1 Then
        Cancel = True
    End If
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
        If r > 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Egresos Comunes"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)

        'Recorro el FlexGrid para poner números a las filas

        If respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            PonerBotones
            CalcuTotal
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

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
' filtra el nombre y codigo de cuenta para los combos del greed
If Col > 2 Then
    CalcuTotal
End If
If Row > 1 Then
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                     .TextMatrix(Row, 2) = .TextMatrix(Row, 1)
             End If

             If Col = 2 Then
                     .TextMatrix(Row, 1) = .TextMatrix(Row, 2)
             End If
         End If
    End With
End If
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

Private Sub VSFG_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            frmSelecCtaConta.Show
            Set frmSelecCtaConta.objEscribir = VSFG
        End If
    End If
End Sub

Private Sub vsfgDocumento_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDocumento_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 1 And Row > 0 Then
        
        If vsfgDocumento.TextMatrix(Row, 1) = "-1" Then
            For i = 1 To vsfgDocumento.Rows - 1
                If i <> Row And vsfgDocumento.TextMatrix(i, 1) = "-1" Then
                    vsfgDocumento.Select i, 1, i, 10
                    vsfgDocumento.FillStyle = flexFillRepeat
                    vsfgDocumento.CellBackColor = &HFFFFFF
                    vsfgDocumento.TextMatrix(i, 1) = "0"
                End If
            Next
            vsfgDocumento.Select Row, 1, Row, 10
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HC0FFFF
            vsfgDocumento.Select Row, 1
            Llenar_Grid (Row)
        ElseIf vsfgDocumento.TextMatrix(Row, 1) = "0" Then
            vsfgDocumento.Select Row, 1, Row, 10
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HFFFFFF
            vsfgDocumento.Select Row, 1
            VSFG.Clear 1
            VSFG.Rows = 2
            VSFGCobros.Clear 1
            VSFGCobros.Rows = 1
        End If
    End If
    CalcuTotal
End Sub


Private Sub CuadraBanco()
    Dim i As Long
    Dim dblCobro As Double
    dblCobro = 0
    For i = 1 To VSFGCobros.Rows - 1
        If Abs(Val(VSFGCobros.TextMatrix(i, 1))) = 1 Then
            dblCobro = dblCobro + VSFGCobros.TextMatrix(i, 7)
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        If (Val(VSFG.TextMatrix(i, 3)) = dblCobro And VSFG.TextMatrix(i, 1) = "*") Or Val(VSFG.TextMatrix(i, 3)) = 0 Then
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

Private Sub Llenar_Grid(num As Long)
    strSQL = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_fechaemision, pag_fecha, cue_p_c_valor, pag_monto,pag_no_doc,pag_observacion,pag_codigo " & _
             " FROM (cuenta_p_c INNER JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo AND pago.doc_pag_codigo='" & Me.vsfgDocumento.TextMatrix(num, 2) & "') " & _
             " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & _
             "' AND cuenta_p_c.emp_codigo = '" & strEmpresa & _
             "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND pag_no_doc='" & vsfgDocumento.TextMatrix(num, 5) & "' "
    clsSql.Ejecutar strSQL
    VSFGCobros.Clear 1
    VSFGCobros.Rows = 1
    Set VSFGCobros.DataSource = clsSql.adorec_Def.DataSource
    VSFG.TextMatrix(1, 4) = FormatoD2(vsfgDocumento.TextMatrix(num, 8)) + FormatoD2(txtValor.Text)
End Sub
Private Sub CalcuTotal()
   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    
    'Calcula total debe y haber
    
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    txtTotalHaber = Format(SumaHaber, "##0.00")
    
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
End Sub
