VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCartera 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartera"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3290
      TabIndex        =   0
      Top             =   4800
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4080
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7980
      _cx             =   14076
      _cy             =   7197
      Appearance      =   1
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
      FormatString    =   $"frmCartera.frx":030A
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Public Persona As String
Public Sub Carga()
'***********************************************
        
        strSql = " CREATE TEMPORARY TABLE Abo ( " & _
                 " emp_codigo char(3) NOT NULL default ''," & _
                 " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
                 " cue_p_c_tipo char(1) NOT NULL default ''," & _
                 " abono decimal(14,2) default NULL," & _
                 " abonoNC decimal(14,2) default NULL," & _
                 " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO Abo " & _
                 " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(sum(if(pag_observacion!='NOTA DE CREDITO',pag_monto,0)),0.000) as abono,COALESCE(sum(if(pag_observacion='NOTA DE CREDITO',pag_monto,0)),0.000) as abonoNC " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = pago.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='C' " & _
                 " AND cuenta_p_c.per_codigo like '" & Persona & "' " & _
                 " GROUP BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
                 " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TEMPORARY TABLE RetFech ( " & _
                 " emp_codigo char(3) NOT NULL default ''," & _
                 " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
                 " cue_p_c_tipo char(1) NOT NULL default ''," & _
                 " reten decimal(14,2) default NULL," & _
                 " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO RetFech " & _
                 " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(comprobante_retencion.com_ret_total,0.000) as reten " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='C' " & _
                 " AND  cuenta_p_c.per_codigo like '" & Persona & "' " & _
                 " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TEMPORARY TABLE Ret ( " & _
                 " emp_codigo char(3) NOT NULL default ''," & _
                 " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
                 " cue_p_c_tipo char(1) NOT NULL default ''," & _
                 " reten decimal(14,2) default NULL," & _
                 " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO Ret " & _
                 " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(RetFech.reten,0.000) as reten " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                 " LEFT JOIN RetFech ON cuenta_p_c.cue_p_c_codigo = RetFech.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = RetFech.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = RetFech.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='C' " & _
                 " AND  cuenta_p_c.per_codigo like '" & Persona & "' " & _
                 " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TEMPORARY TABLE Cob ( " & _
                 " emp_codigo char(3) NOT NULL default ''," & _
                 " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
                 " cue_p_c_tipo char(1) NOT NULL default ''," & _
                 " abono decimal(14,2) default NULL," & _
                 " abonoNC decimal(14,2) default NULL," & _
                 " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo)) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO Cob " & _
                 " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(Abo.abono,0.000) as abono,COALESCE(Abo.abonoNC,0.000) as abonoNC " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
                 " LEFT JOIN Abo ON cuenta_p_c.cue_p_c_codigo = Abo.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = Abo.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = Abo.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo like '" & Persona & "' " & _
                 " ORDER BY cue_p_c_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " DROP TABLE Abo "
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TEMPORARY TABLE Cuentas " & _
                 " SELECT cue_p_c_egr_codigo,cue_p_c_fechaemision as emision, cue_p_c_fechapropuesta as vencimiento, " & _
                 " COALESCE(cue_p_c_valor,0.000) as valor, COALESCE(Cob.abono,0.000) as abono, COALESCE(Cob.abonoNC,0.000) as abonoNC," & _
                 " COALESCE(Ret.reten,0.000) as reten, COALESCE(cue_p_c_st_cero,0.000) as flete,ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Cob.abonoNC,0.000) - COALESCE(Ret.reten,0.000),2) as saldo " & _
                 " FROM cuenta_p_c INNER JOIN Ret ON cuenta_p_c.cue_p_c_codigo = Ret.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = Ret.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = Ret.emp_codigo " & _
                 " INNER JOIN Cob ON cuenta_p_c.cue_p_c_codigo = Cob.cue_p_c_codigo  " & _
                 " AND cuenta_p_c.cue_p_c_tipo = Cob.cue_p_c_tipo " & _
                 " AND cuenta_p_c.emp_codigo = Cob.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo like '" & Persona & "' " & _
                 " AND ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Cob.abonoNC,0.000) - COALESCE(Ret.reten,0.000),2)>0 " & _
                 " ORDER BY cue_p_c_egr_codigo, emision "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO Cuentas " & _
                 " SELECT CONCAT('NC ',ing_codigo),ing_fecha as emision, ing_fecha as vencimiento, " & _
                 " COALESCE(-1 * ing_total,0.000) as valor, '0.000' as abono, COALESCE(-1 * ing_saldo,0.000) as abonoNC, " & _
                 " '0.000' as reten, COALESCE(ing_subtotal_o,0) as flete, -1*ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2) as saldo " & _
                 " FROM ingreso  " & _
                 " WHERE ingreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND ingreso.tip_ing_codigo='DCL' AND  ingreso.per_codigo like '" & Persona & "' " & _
                 " AND ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2)>0 " & _
                 " ORDER BY ing_codigo, emision "
        clsCon_Def.Ejecutar strSql
        strSql = " SELECT * FROM Cuentas ORDER BY cue_p_c_egr_codigo, emision "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
        End If
        strSql = " DROP TABLE Ret "
        clsCon_Def.Ejecutar strSql
        strSql = " DROP TABLE RetFech "
        clsCon_Def.Ejecutar strSql
        strSql = " DROP TABLE Cob "
        clsCon_Def.Ejecutar strSql
        strSql = " DROP TABLE Cuentas "
        clsCon_Def.Ejecutar strSql
'***********************************************
    ucrtVSFG.PonerNum
    VSFG.Subtotal flexSTClear
    VSFG.Subtotal flexSTSum, -1, 4, , vbRed, vbWhite, True, "Totales"
    VSFG.Subtotal flexSTSum, -1, 5
    VSFG.Subtotal flexSTSum, -1, 6
    VSFG.Subtotal flexSTSum, -1, 7
    VSFG.Subtotal flexSTSum, -1, 8
    VSFG.Subtotal flexSTSum, -1, 9
    VSFG.ShowCell VSFG.Rows - 1, VSFG.Cols - 1
    frmV_PedBod.txtDisponible.Text = FormatoD2(frmV_PedBod.txtCredito.Text) - FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, VSFG.Cols - 1))
    
End Sub

Private Sub VSFG_DblClick()
    Dim i As Long
    Set DAT = New frmDatos
    If VSFG.Row >= 1 Then
        DAT.Show
        DAT.VSFG.Rows = VSFG.Cols
        For i = 1 To VSFG.Cols - 1
            DAT.VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(0, i)
            DAT.VSFG.Cell(flexcpText, i, 1) = VSFG.Cell(flexcpTextDisplay, VSFG.Row, i)
            If VSFG.ColComboList(i) <> "" Then
                DAT.VSFG.TextMatrix(i, 2) = VSFG.ColComboList(i)
                DAT.VSFG.Cell(flexcpText, i, 3) = VSFG.Cell(flexcpText, VSFG.Row, i)
            End If
        Next i
        DAT.VSFG.Cell(flexcpBackColor, 1, 1, DAT.VSFG.Rows - 1, 1) = VSFG.Cell(flexcpBackColor, VSFG.Row, VSFG.Col)
        DAT.VSFG.RowHidden(DAT.VSFG.Rows - 1) = True
        Set DAT.VSFGOrigen = VSFG
        DAT.VSFGOrigen.Tag = VSFG.Row
        DAT.Caption = Tipo
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
    End If
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

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = 0
    Me.Top = mdiPrincipal.Height - Me.Height - 1300
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
