VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProductoPromo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmProductoPomo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7635
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2038
      TabIndex        =   1
      Top             =   3120
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3943
      TabIndex        =   0
      Top             =   3120
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7380
      _cx             =   13017
      _cy             =   4233
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProductoPomo.frx":030A
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
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66519041
         CurrentDate     =   39449
      End
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmProductoPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSQL As String
Private Tipo As String
Private Tipo2 As String
Public CodPrd As String
Private Sub IniDato()
    Tipo = " Precio "
    Tipo2 = "el Precio "
    Me.Caption = Tipo
End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
  
    strSQL = " SELECT prd_pro_porcentaje,prd_pro_fechaini,prd_pro_fechafin," & _
             " prd_pro_fechamod, prd_pro_usumod, '0' as modi" & _
             " FROM producto_promo " & _
             " INNER JOIN producto " & _
             " ON producto_promo.emp_codigo=producto.emp_codigo " & _
             " AND producto.prd_codigo=producto_promo.prd_codigo " & _
             " WHERE producto_promo.emp_codigo LIKE '" & strEmpresa & "' " & _
             " AND producto_promo.prd_codigo='" & CodPrd & "' " & _
             " ORDER BY prd_pro_fechaini,prd_pro_fechafin "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum
    
    VSFG.ColComboList(1) = "Dummy"
    VSFG.ColComboList(2) = "Dummy"
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long, control As Integer
    control = 0
    If VSFG.Rows > 1 Then
        VSFG.Select 1, VSFG.Cols - 1
        VSFG.Sort = flexSortGenericDescending
        
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
                strSQL = " UPDATE producto_promo SET " & _
                         " prd_pro_fechaini='" & VSFG.TextMatrix(i, 1) & "', " & _
                         " prd_pro_fechafin='" & VSFG.TextMatrix(i, 2) & "', " & _
                         " prd_pro_porcentaje='" & FormatoD2(VSFG.TextMatrix(i, 3)) & "', " & _
                         " prd_pro_fechamod=CURRENT_TIMESTAMP, " & _
                         " prd_pro_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND prd_codigo='" & CodPrd & "' " & _
                         " AND prd_pro_fechaini='" & VSFG.TextMatrix(i, 4) & "' " & _
                         " AND prd_pro_fechafin='" & VSFG.TextMatrix(i, 5) & "' "
                clsCon_Def.Ejecutar (strSQL), "M"
            ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            
                If Trim(VSFG.TextMatrix(i, 1)) = "" Then
                    MsgBox "No puede ingresar Promoción, falta la Fecha inicio", vbInformation, "Ingreso"
                    control = 1
                ElseIf Trim(VSFG.TextMatrix(i, 2)) = "" Then
                    MsgBox "No puede ingresar Promoción, falta la Fecha fin", vbInformation, "Ingreso"
                    control = 1
                ElseIf Format(VSFG.TextMatrix(i, 1), "yyyy-MM-dd") > Format(VSFG.TextMatrix(i, 2), "yyyy-MM-dd") Then
                    MsgBox "No puede ingresar promoción, la Fecha fin debe ser mayor a Fecha inicio", vbInformation, "Ingreso"
                    control = 1
                Else
                    strSQL = " SELECT count(*) as num " & _
                            " FROM producto_promo " & _
                            " WHERE emp_codigo='" & strEmpresa & "' " & _
                            " AND prd_codigo='" & CodPrd & "' " & _
                            " AND ('" & VSFG.TextMatrix(i, 1) & "' BETWEEN prd_pro_fechaini AND prd_pro_fechafin " & _
                            " OR '" & VSFG.TextMatrix(i, 2) & "' BETWEEN prd_pro_fechaini AND prd_pro_fechafin)"
                    clsCon_Def.Ejecutar strSQL
                    If FormatoD0(clsCon_Def.adorec_Def("num")) <> 0 Then
                        MsgBox "Tiene conflictos con las fechas", vbInformation, "Promociones"
                        control = 1
                    Else
                        strSQL = " INSERT INTO producto_promo " & _
                                  " (emp_codigo,prd_codigo,prd_pro_fechaini,prd_pro_fechafin,prd_pro_porcentaje,prd_pro_fechamod,prd_pro_usumod) " & _
                                  " VALUES ('" & strEmpresa & "','" & CodPrd & "','" & VSFG.TextMatrix(i, 1) & "','" & VSFG.TextMatrix(i, 2) & "','" & FormatoD2(VSFG.TextMatrix(i, 3)) & "'," & _
                                  " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                        clsCon_Def.Ejecutar (strSQL), "M"
                    End If
                End If
            ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
                strSQL = " DELETE FROM producto_promo " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND prd_codigo='" & CodPrd & "' " & _
                         " AND prd_pro_fechaini='" & VSFG.TextMatrix(i, 4) & "' " & _
                         " AND prd_pro_fechafin='" & VSFG.TextMatrix(i, 5) & "' "
                clsCon_Def.Ejecutar (strSQL), "M"
            ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
                Exit For
            End If
        Next i
        If control = 0 Then
            Carga
        End If
        
    End If
End Sub


Private Sub dtpFecha_Change()
    If Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "2" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "2"
    ElseIf Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "3" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "3"
    End If
    
    If VSFG.Col = 1 Or VSFG.Col = 2 Then
        'If FormatoFecha(VSFG.TextMatrix(VSFG.Row, 3)) <> "" And FormatoFecha(VSFG.TextMatrix(VSFG.Row, 5)) <> "" _
        'And FormatoFecha(VSFG.TextMatrix(VSFG.Row, 3)) > FormatoFecha(VSFG.TextMatrix(VSFG.Row, 5)) Then
            'MsgBox "La Fecha de Finalización debe ser mayor a la Fecha de Inicio", vbCritical, "Fecha de Finalización"
            'VSFG.TextMatrix(VSFG.Row, 5) = ""
        'Else
            VSFG.Text = Format(dtpFecha.Value, "yyyy-MM-dd")
        'End If
    End If
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            VSFG = dtpFecha.Tag
            dtpFecha.Visible = False
        Case vbKeyReturn
            dtpFecha.Visible = False
    End Select
End Sub

Private Sub dtpFecha_LostFocus()
    dtpFecha.Visible = False
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
        If Col = 1 Or Col = 2 Then
            If VSFG.TextMatrix(Row, Col) = "" Then
                VSFG.TextMatrix(Row, Col) = Format(Date, "yyyy-MM-dd")
            End If
        ElseIf Col = 3 Then
            If VSFG.TextMatrix(Row, Col) = "" Then
                VSFG.TextMatrix(Row, Col) = "0"
            End If
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col >= VSFG.Cols - 3 Then
            Cancel = True
        End If
        If Col = 1 Or Col = 2 Then
            If VSFG.TextMatrix(Row, Col) = "" Then
                VSFG.TextMatrix(Row, Col) = Format(Date, "yyyy-MM-dd")
            End If
        ElseIf Col = 3 Then
            If VSFG.TextMatrix(Row, Col) = "" Then
                VSFG.TextMatrix(Row, Col) = "0"
            End If
        End If
    End If
    
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


Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.Value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.Value = 1 Then
        txtCodigo.Enabled = True
    Else
        txtCodigo.Enabled = False
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
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar
    IniDato
    dtpFecha.Format = dtpCustom
    dtpFecha.CustomFormat = "yyyy-MM-dd"
    dtpFecha.Visible = False
    
    Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
    If Col = 3 Then
        If VSFG.TextMatrix(Row, 3) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, 3)) Then
            MsgBox "Ingrese un valor válido", vbInformation, "Descuento"
            VSFG.TextMatrix(Row, 3) = "0"
        End If
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

Private Sub VSFG_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpFecha.Visible Then Cancel = True
End Sub

Private Sub VSFG_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If dtpFecha.Visible Then Cancel = True
End Sub

Private Sub VSFG_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Or Col = 2 Then
        If VSFG.ColDataType(Col) = flexDTDate Then
            Cancel = True
            dtpFecha.Move VSFG.CellLeft, VSFG.CellTop, VSFG.CellWidth, VSFG.CellHeight
            dtpFecha.Value = VSFG
            dtpFecha.Tag = VSFG
            
            dtpFecha.Visible = True
            dtpFecha.SetFocus
            
            SendKeys vbKeyF4
        End If
    End If
    
End Sub

