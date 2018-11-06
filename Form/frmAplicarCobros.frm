VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAplicarCobros 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicar Cobros"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmAplicarCobros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11475
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5910
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   7680
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
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   11055
         Begin VB.OptionButton optproveedorc 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Proveedor"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   960
            TabIndex        =   24
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optclientec 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Cliente"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dcmbBeneficiarioc 
            Height          =   315
            Left            =   2850
            TabIndex        =   25
            Top             =   120
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblBeneficiarioc 
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
            Left            =   2160
            TabIndex        =   26
            Top             =   135
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11055
         Begin VB.OptionButton optcliente 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Cliente"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optproveedor 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Proveedor"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   960
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcmbBeneficiario 
            Height          =   315
            Left            =   2850
            TabIndex        =   20
            Top             =   120
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            Left            =   2160
            TabIndex        =   21
            Top             =   135
            Width           =   525
         End
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   525
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   5400
         Width           =   7335
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
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   7080
         Width           =   1815
      End
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
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   7080
         Width           =   1935
      End
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
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfgDocumento 
         Height          =   1335
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   11055
         _cx             =   1998605356
         _cy             =   1998588211
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarCobros.frx":030A
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
         TabIndex        =   4
         Top             =   3360
         Width           =   10695
         _cx             =   1998604721
         _cy             =   1998588634
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
         FormatString    =   $"frmAplicarCobros.frx":048F
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
      Begin MSComCtl2.DTPicker dtpFecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   5040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   66715651
         CurrentDate     =   37463
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1095
         Left            =   840
         TabIndex        =   13
         Top             =   6000
         Width           =   9480
         _cx             =   1998602578
         _cy             =   1998587787
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarCobros.frx":05DE
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
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label lbltotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   7095
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Comprobante:"
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
         Top             =   5040
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas por Cobrar"
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
         TabIndex        =   6
         Top             =   3120
         Width           =   1425
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
         TabIndex        =   5
         Top             =   4950
         Width           =   600
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4320
         Picture         =   "frmAplicarCobros.frx":06B8
         ToolTipText     =   "Elimina una Fila"
         Top             =   720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4560
         Picture         =   "frmAplicarCobros.frx":07EE
         Top             =   720
         Visible         =   0   'False
         Width           =   225
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
         TabIndex        =   2
         Top             =   840
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmAplicarCobros"
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
    Dim i As Long
    Dim claAsi As New clsContable
    claAsi.Inicializar AdoConn, AdoConnMaster
    If FormatoD2(txtValor.Text) > FormatoD2(vsfgDocumento.TextMatrix(vsfgDocumento.Row, 6)) Then
        MsgBox "Esta aplicando mas que el saldo del Anticipo disponible", vbCritical, "Cobros"
    Else
        claAsi.NuevoAsiento "D", Format(dtpFecha.Value, "yyyy-mm-dd"), 0, 0, FormatoD2(txtTotalDebe.Text), txtDescripcion.Text, True
        For i = 1 To VSFG1.Rows - 1
            If Abs(VSFG1.TextMatrix(i, 1)) = 1 Then
                k = VSFG1.TextMatrix(i, 10)
                If (VSFG1.TextMatrix(i, 10) <> "" Or VSFG1.TextMatrix(i, 10) <> "0") Then
                    'Calcula el máximo codigo de pago para la cuenta
                     strSQL = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                              " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                              " AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                              " AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                              " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' " & _
                              " AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' " & _
                              " AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'C'"
                    clsSql.Ejecutar strSQL, "M"
                    If clsSql.adorec_Def.EOF Then
                        maxpag = 1
                    Else
                        maxpag = clsSql.adorec_Def("pag") + 1
                    End If
                    strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & Val(VSFG1.TextMatrix(i, 2)) & "', 'C', '" & Val(maxpag) & "', '" & dtpFecha & "', '" & FormatoD2(VSFG1.TextMatrix(i, 10)) & "', '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 4) & "', '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 7) & "', " & _
                             " '" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 1) & "','" & claAsi.NumAsiento & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsSql.Ejecutar strSQL, "M"
                    
                    If FormatoD2(VSFG1.TextMatrix(i, 9)) <= FormatoD2(VSFG1.TextMatrix(i, 10)) Then
                        strSQL = " UPDATE cuenta_p_c " & _
                                 " SET cue_p_c_fechapago='" & ffch & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                 " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' AND emp_codigo = '" & strEmpresa & "' "
                        clsSql.Ejecutar strSQL, "M"
                    End If
                End If
            End If
        Next i
        'doc_pag_saldo
        strSQL = "UPDATE doc_pago SET doc_pag_saldo=doc_pag_saldo+'" & FormatoD2(txtValor.Text) & "' WHERE doc_pag_codigo='" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 1) & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSQL, "M"
        strSQL = "UPDATE doc_pago SET doc_pag_anticipo=0 WHERE doc_pag_saldo>=doc_pag_valor AND doc_pag_codigo='" & vsfgDocumento.TextMatrix(vsfgDocumento.Row, 1) & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSQL, "M"
        For i = 1 To VSFG.Rows - 1
            claAsi.NuevoDetAsiento VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 5), FormatoD2(VSFG.TextMatrix(i, 3)), FormatoD2(VSFG.TextMatrix(i, 4))
        Next i
        Dim frmAsiApli As New frmReporte
        frmAsiApli.strAsiento = claAsi.NumAsiento
        frmAsiApli.strReporte = "rptAsiento"
        frmAsiApli.Show
    End If
    dcmbBeneficiario_Change
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBeneficiario_Change()
    Dim i As Long
    txtValor = 0
    t = "P"
    If Me.optcliente.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiario.MatchedWithList = True Then
       
        vsfgDocumento.Enabled = True
        'Consulta para el grid
        strSQL = " SELECT '' as sele, doc_pago.doc_pag_codigo, CONCAT(per_apellido,' ', per_nombre) as per, " & _
                 " COALESCE(banco.ban_nombre,'-') as ban_nombre, doc_pag_numero,  doc_pag_valor, doc_pag_valor-doc_pag_saldo,doc_pag_observacion,doc_pag_fecha_doc,doc_pag_fecha_recepcion,doc_pago.asi_numasiento,COALESCE(det_asiento.cta_codigo,cat_p_ctaconta_ant) " & _
                 " FROM doc_pago INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                 " INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                 " LEFT JOIN det_asiento ON doc_pago.emp_codigo=det_asiento.emp_codigo AND doc_pago.asi_numasiento=det_asiento.asi_numasiento " & _
                 " AND doc_pago.doc_pag_valor=det_asiento.det_asi_haber " & _
                 " LEFT JOIN banco ON doc_pago.ban_codigo = banco.ban_codigo " & _
                 " WHERE doc_pago.emp_codigo = '" & strEmpresa & "' AND doc_pag_anticipo = 1 " & _
                 " AND doc_pago.per_codigo='" & dcmbBeneficiario.BoundText & "' AND doc_pag_estado!='ANULADO' " & _
                 " ORDER BY doc_pago.doc_pag_codigo desc "
        clsSql.Ejecutar strSQL
        If clsSql.adorec_Def.EOF = False Then
            Set vsfgDocumento.DataSource = clsSql.adorec_Def.DataSource
        Else
            vsfgDocumento.Clear 1
            vsfgDocumento.Rows = 1
        End If
        strSQL = " SELECT IIF(cat_p_ctaconta IS NULL OR cat_p_ctaconta='',par_texto,cat_p_ctaconta) as par_texto " & _
                 " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                 " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                 " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & dcmbBeneficiario.BoundText & "'"
        clsSql.Ejecutar strSQL
        VSFG.Tag = clsSql.adorec_Def("par_texto")
    End If
    txtDescripcion.Text = ""
    dcmbBeneficiarioc = dcmbBeneficiario
End Sub

Private Sub dcmbBeneficiarioc_Change()
    Dim i As Long
    txtValor = 0
    t = "P"
    If Me.optclientec.Value = True Then
        t = "C"
    End If
    
    If dcmbBeneficiarioc.MatchedWithList = True Then
       
        strSQL = " SELECT ' ','0', cuenta_p_c.cue_p_c_codigo, CONCAT(cue_p_c_fra_cuenta, '/' , cue_p_c_tot_cuenta ) as cue_p_c_fra_cuenta, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0), ' ' " & _
                 " FROM (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo)" & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE per_codigo = '" & dcmbBeneficiarioc.BoundText & "' AND cue_p_c_valor!=0 AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " GROUP BY cuenta_p_c.cue_p_c_codigo, cue_p_c_fra_cuenta, cue_p_c_tot_cuenta, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,com_ret_total "
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
    VSFG.Rows = 1
    txtDescripcion.Text = ""
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
    optclientec.Value = True
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
        Set dcmbBeneficiarioc.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiarioc.ListField = "nombre"
        dcmbBeneficiarioc.BoundColumn = "per_codigo"
        Persona = ""
        p = 0
    End If
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub OptCliente_Click()
  
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiario.Text = ""
    vsfgDocumento.Clear 1
    vsfgDocumento.Rows = 1
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub OptClientec_Click()
  
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiarioc.Text = ""
    VSFG1.Clear 1
    VSFG1.Rows = 1
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiarioc.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiarioc.ListField = "nombre"
        dcmbBeneficiarioc.BoundColumn = "per_codigo"
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

Private Sub optproveedorc_Click()
    
    p = 1
    Frame1.Caption = "Proveedor"
    dcmbBeneficiarioc.Text = ""
    VSFG1.Clear 1
    VSFG1.Rows = 1
    strSQL = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiarioc.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiarioc.ListField = "nombre"
        dcmbBeneficiarioc.BoundColumn = "per_codigo"
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
    CrearAsiento
End Sub

Private Sub CrearAsiento()
    Dim i As Long
    VSFG.Rows = 1
    txtDescripcion.Text = ""
    For i = 1 To vsfgDocumento.Rows - 1
        If Abs(FormatoD0(vsfgDocumento.TextMatrix(i, 0))) = 1 Then
            strSQL = " SELECT cta_nombre " & _
                     " FROM ctaconta " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cta_codigo='" & vsfgDocumento.TextMatrix(i, 11) & "' "
            clsSql.Ejecutar strSQL
            If clsSql.adorec_Def.RecordCount > 0 Then
                VSFG.AddItem "" & vbTab & vsfgDocumento.TextMatrix(i, 11) & vbTab & clsSql.adorec_Def("cta_nombre") & vbTab & txtValor.Text
            End If
            txtDescripcion.Text = "APLICACION DE COBRO NO. " & vsfgDocumento.TextMatrix(i, 1) & vbNewLine & _
                                  "CLIENTE: " & dcmbBeneficiario.Text & vbNewLine & _
                                  "VALOR TOTAL: " & vsfgDocumento.TextMatrix(i, 5) & vbNewLine & _
                                  "SALDO: " & vsfgDocumento.TextMatrix(i, 6) & vbNewLine & _
                                  "APLICADO: " & txtValor.Text & vbNewLine & _
                                  "DESCRIPCION: " & vsfgDocumento.TextMatrix(i, 8) & vbNewLine & vbNewLine & _
                                  "CLIENTE FAC: " & dcmbBeneficiarioc.Text
            Exit For
        End If
    Next i
    For i = 1 To VSFG1.Rows - 1
        If Abs(FormatoD0(VSFG1.TextMatrix(i, 1))) = 1 Then
            txtDescripcion.Text = txtDescripcion.Text & vbNewLine & "FACTURA: " & VSFG1.TextMatrix(i, 4) & " DESCRIPCION: " & VSFG1.TextMatrix(i, 5) & " VALOR APLICADO: " & VSFG1.TextMatrix(i, 10)
        End If
    Next i
    strSQL = " SELECT cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cta_codigo='" & VSFG.Tag & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG.AddItem "" & vbTab & VSFG.Tag & vbTab & clsSql.adorec_Def("cta_nombre") & vbTab & vbTab & txtValor.Text
    End If
    CalcuTotal
End Sub
Private Sub CalcuTotal()

   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double

    'Calcula total debe
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")

    'Calcula total haber
    For i = 1 To VSFG.Rows - 1
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")

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
                    vsfgDocumento.Select i, 0, i, 7
                    vsfgDocumento.FillStyle = flexFillRepeat
                    vsfgDocumento.CellBackColor = &HFFFFFF
                    vsfgDocumento.TextMatrix(i, 0) = "0"
                End If
            Next
            vsfgDocumento.Select Row, 0, Row, 7
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HC0FFFF
            vsfgDocumento.Select Row, 0
        ElseIf vsfgDocumento.TextMatrix(Row, 0) = "0" Then
            vsfgDocumento.Select Row, 0, Row, 7
            vsfgDocumento.FillStyle = flexFillRepeat
            vsfgDocumento.CellBackColor = &HFFFFFF
            vsfgDocumento.Select Row, 0
        End If
        CrearAsiento
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
