VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReprocesoRet 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproceso Retenciones"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReprocesoRet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6750
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfgKardex 
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      _cx             =   11456
      _cy             =   9551
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4095
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   2535
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDesde 
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
      Left            =   2550
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   70320131
      CurrentDate     =   37463
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
   Begin VB.Label LblDetalle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   1830
      TabIndex        =   6
      Top             =   180
      Width           =   525
   End
End
Attribute VB_Name = "frmReprocesoRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim strIni As String
    Dim strSQL As String
    Dim clsConAUX As New clsConsulta
    strIni = Format(dtpDesde.Value, "yyyy-mm-dd")
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT asiento.asi_numasiento,asiento.asi_fecha,comprobante_retencion.com_ret_fecha,cue_p_c_fechaemision,not_d_c_fecha," & _
             " EOMONTH(cue_p_c_fechaemision) as ffm, " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',CHAR(10),'ANTERIRO: ',LEFT(asi_fecha,10),char(10),asi_descripcion) as asi_desc, " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',CHAR(10),'ANTERIRO: ',LEFT(pago.pag_fecha,10),char(10),pag_observacion) as pag_desc," & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',CHAR(10),'ANTERIRO: ',LEFT(not_d_c_fecha,10),char(10),not_d_c_descripcion) as not_desc," & _
             " cuenta_p_c.cue_p_c_codigo,not_d_c_codigo,asiento.emp_codigo" & _
             " FROM comprobante_retencion, pago, Asiento, cuenta_p_c, nota_d_c" & _
             " WHERE comprobante_retencion.emp_codigo = pago.emp_codigo" & _
             " and comprobante_retencion.cue_p_c_codigo=pago.cue_p_c_codigo" & _
             " and comprobante_retencion.cue_p_c_tipo=pago.cue_p_c_tipo" & _
             " and pago.pag_monto=0 and pag_observacion like'%RETENCI%N%'" & _
             " and pag_observacion not like'%ANULAD%'" & _
             " and pago.emp_codigo=asiento.emp_codigo" & _
             " and pago.asi_numasiento=asiento.asi_numasiento" & _
             " and pago.emp_codigo=cuenta_p_c.emp_codigo" & _
             " and pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo" & _
             " and pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo" & _
             " and comprobante_retencion.cue_p_c_tipo='C'" & _
             " and LEFT(com_ret_fecha,7)!=LEFT(cue_p_c_fechaemision,7) AND asi_fecha>='" & strIni & "'" & _
             " and asiento.emp_codigo=nota_d_c.emp_codigo" & _
             " and asiento.asi_numasiento=nota_d_c.asi_numasiento" & _
             " and LEFT(asi_fecha,7)!=LEFT(cue_p_c_fechaemision,7)"
    clsConAUX.Ejecutar strSQL
    Set vsfgKardex.DataSource = clsConAUX.adorec_Def.DataSource
    If MsgBox("desea Continuar con el reproceso", vbYesNo) = vbYes Then
        While Not clsConAUX.adorec_Def.EOF
            strSQL = " UPDATE asiento " & _
                     " SET asiento.asi_fecha='" & clsConAUX.adorec_Def("ffm") & "'," & _
                     " asiento.asi_descripcion='" & clsConAUX.adorec_Def("asi_desc") & "'" & _
                     " WHERE asiento.emp_codigo='" & clsConAUX.adorec_Def("emp_codigo") & "'" & _
                     " and asiento.asi_numasiento='" & clsConAUX.adorec_Def("asi_numasiento") & "'"
            clsCon_Def.Ejecutar strSQL
            strSQL = " UPDATE comprobante_retencion " & _
                     " SET comprobante_retencion.com_ret_fecha='" & clsConAUX.adorec_Def("ffm") & "'" & _
                     " WHERE comprobante_retencion.emp_codigo ='" & clsConAUX.adorec_Def("emp_codigo") & "'" & _
                     " and cue_p_c_tipo='C' and cue_p_c_codigo='" & clsConAUX.adorec_Def("cue_p_c_codigo") & "'"
            clsCon_Def.Ejecutar strSQL
            strSQL = " UPDATE nota_d_c" & _
                     " SET not_d_c_fecha='" & clsConAUX.adorec_Def("ffm") & "'," & _
                     " not_d_c_descripcion='" & clsConAUX.adorec_Def("not_desc") & "'" & _
                     " WHERE nota_d_c.emp_codigo='" & clsConAUX.adorec_Def("emp_codigo") & "'" & _
                     " and not_d_c_codigo='" & clsConAUX.adorec_Def("not_d_c_codigo") & "'"
            clsCon_Def.Ejecutar strSQL
            strSQL = " UPDATE pago" & _
                     " SET pago.pag_fecha='" & clsConAUX.adorec_Def("ffm") & "'," & _
                     " pago.pag_observacion='" & clsConAUX.adorec_Def("pag_desc") & "'" & _
                     " WHERE pago.emp_codigo='" & clsConAUX.adorec_Def("emp_codigo") & "'" & _
                     " and pago.asi_numasiento='" & clsConAUX.adorec_Def("asi_numasiento") & "'" & _
                     " and pago.cue_p_c_codigo='" & clsConAUX.adorec_Def("cue_p_c_codigo") & "'" & _
                     " and pago.cue_p_c_tipo='C'"
            clsCon_Def.Ejecutar strSQL
            clsConAUX.adorec_Def.MoveNext
        Wend
    End If
    MsgBox "FIN" & vbNewLine & strIni & vbNewLine & strFin
    'Recostear Format(dtpDesde.Value, "yyyy-mm-dd"), Format(dtpHasta.Value, "yyyy-mm-dd")
End Sub

Private Sub Command1_Click()
    Dim strIni As String
    Dim strSQL As String
    strIni = Format(dtpDesde.Value, "yyyy-mm-dd")
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT asiento.asi_numasiento,asiento.asi_fecha,comprobante_retencion.com_ret_fecha,cue_p_c_fechaemision,not_d_c_fecha," & _
             " EOMONTH(cue_p_c_fechaemision), " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(asi_fecha,10),char(10),asi_descripcion), " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(pago.pag_fecha,10),char(10),pag_observacion)," & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(not_d_c_fecha,10),char(10),not_d_c_descripcion)" & _
             " FROM comprobante_retencion, pago, Asiento, cuenta_p_c, nota_d_c" & _
             " WHERE comprobante_retencion.emp_codigo = pago.emp_codigo" & _
             " and comprobante_retencion.cue_p_c_codigo=pago.cue_p_c_codigo" & _
             " and comprobante_retencion.cue_p_c_tipo=pago.cue_p_c_tipo" & _
             " and pago.pag_monto=0 and pag_observacion like'%RETENCI%N%'" & _
             " and pag_observacion not like'%ANULAD%'" & _
             " and pago.emp_codigo=asiento.emp_codigo" & _
             " and pago.asi_numasiento=asiento.asi_numasiento" & _
             " and pago.emp_codigo=cuenta_p_c.emp_codigo" & _
             " and pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo" & _
             " and pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo" & _
             " and comprobante_retencion.cue_p_c_tipo='C'" & _
             " and com_ret_fecha>='" & strIni & "'" & _
             " and asiento.emp_codigo=nota_d_c.emp_codigo" & _
             " and asiento.asi_numasiento=nota_d_c.asi_numasiento" & _
             " and LEFT(asi_fecha,7)!=LEFT(cue_p_c_fechaemision,7)"
    strSQL = " SELECT asiento.asi_numasiento,asiento.asi_fecha,comprobante_retencion.com_ret_fecha,cue_p_c_fechaemision,not_d_c_fecha," & _
             " EOMONTH(cue_p_c_fechaemision), " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(asi_fecha,10),char(10),asi_descripcion), " & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(pago.pag_fecha,10),char(10),pag_observacion)," & _
             " CONCAT('CAMBIO DE FECHA RETENCION.',char(10),'ANTERIRO: ',LEFT(not_d_c_fecha,10),char(10),not_d_c_descripcion),cue_p_c_egr_codigo" & _
             " FROM comprobante_retencion, pago, Asiento, cuenta_p_c, nota_d_c" & _
             " WHERE comprobante_retencion.emp_codigo = pago.emp_codigo" & _
             " and comprobante_retencion.cue_p_c_codigo=pago.cue_p_c_codigo" & _
             " and comprobante_retencion.cue_p_c_tipo=pago.cue_p_c_tipo" & _
             " and pago.pag_monto=0 and pag_observacion like'%RETENCI%N%'" & _
             " and pag_observacion not like'%ANULAD%'" & _
             " and pago.emp_codigo=asiento.emp_codigo" & _
             " and pago.asi_numasiento=asiento.asi_numasiento" & _
             " and pago.emp_codigo=cuenta_p_c.emp_codigo" & _
             " and pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo" & _
             " and pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo" & _
             " and comprobante_retencion.cue_p_c_tipo='C'" & _
             " and asi_fecha>='" & strIni & "'" & _
             " and LEFT(com_ret_fecha,7)!=LEFT(cue_p_c_fechaemision,7)" & _
             " and asiento.emp_codigo=nota_d_c.emp_codigo" & _
             " and asiento.asi_numasiento=nota_d_c.asi_numasiento" & _
             " and LEFT(asi_fecha,7)!=LEFT(cue_p_c_fechaemision,7)"
    clsCon_Def.Ejecutar strSQL
    Set vsfgKardex.DataSource = clsCon_Def.adorec_Def.DataSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpDesde.Value = HoyDia
    Set ucrtVSFG.VSFGControl = vsfgKardex
    ucrtVSFG.Inicializar
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
