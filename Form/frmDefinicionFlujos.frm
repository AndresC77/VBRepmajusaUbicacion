VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDefinicionFlujos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Flujos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "frmDefinicionFlujos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12075
   Begin VB.CommandButton cmbAceptar2 
      Caption         =   "&Aceptar Detalle"
      Height          =   360
      Left            =   8400
      TabIndex        =   6
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar / Recargar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar Flujo"
      Height          =   360
      Left            =   10200
      TabIndex        =   1
      Top             =   3000
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   10200
      TabIndex        =   0
      Top             =   6480
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   1920
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   11820
      _cx             =   20849
      _cy             =   3387
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDefinicionFlujos.frx":030A
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
      TabIndex        =   3
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   2880
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   11820
      _cx             =   20849
      _cy             =   5080
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
      FormatString    =   $"frmDefinicionFlujos.frx":03E9
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
   Begin NEED2.uctrVSFG ucrtVSFG2 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmDefinicionFlujos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Private Sub IniDato()
    Tipo = "Flujo"
    Tipo2 = "Flujo"
    Me.Caption = Tipo
End Sub

Private Sub cmbAceptar2_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFG2.Select 1, VSFG2.Cols - 1
    VSFG2.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG2.Rows - 1
        'update
        If VSFG2.TextMatrix(i, VSFG2.Cols - 1) = 3 Then
            strSql = " UPDATE det_flujo " & _
                 " SET det_flu_nombre='" & UCase(VSFG2.TextMatrix(i, 2)) & "'," & _
                 " det_flu_descripcion='" & UCase(VSFG2.TextMatrix(i, 3)) & "'," & _
                 " det_flu_tiempo='" & FormatoD0(VSFG2.TextMatrix(i, 4)) & "'," & _
                 " res_flu_codigo='" & UCase(VSFG2.TextMatrix(i, 5)) & "'," & _
                 " det_flu_fechamod=CURRENT_TIMESTAMP," & _
                 " det_flu_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND det_flu_codigo='" & VSFG2.TextMatrix(i, 1) & "'" & _
                 " AND flu_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG2.TextMatrix(i, VSFG2.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG2.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el código", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG2.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG2.TextMatrix(i, 4) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta tiempo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG2.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta responsable", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT det_flu_codigo" & _
                    " FROM det_flujo " & _
                    " WHERE emp_codigo='" & strEmpresa & "'" & _
                    " AND flu_codigo='" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
                    " AND det_flu_codigo='" & VSFG2.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    strSql = " INSERT INTO det_flujo(emp_codigo, flu_codigo, det_flu_codigo,det_flu_nombre, " & _
                             " det_flu_descripcion,det_flu_tiempo,res_flu_codigo, " & _
                             " det_flu_fechamod,det_flu_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(VSFG.Row, 1)) & "','" & UCase(VSFG2.TextMatrix(i, 1)) & "','" & UCase(VSFG2.TextMatrix(i, 2)) & "'," & _
                             " '" & UCase(VSFG2.TextMatrix(i, 3)) & "', '" & FormatoD0(VSFG2.TextMatrix(i, 4)) & "', '" & UCase(VSFG2.TextMatrix(i, 5)) & "', " & _
                            " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "El código de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG2.TextMatrix(i, VSFG2.Cols - 1) = 1 Then
        
            strSql = " SELECT count(flu_codigo) as existe " & _
                    " FROM det_historia_flujo " & _
                    " WHERE flu_codigo = '" & VSFG.TextMatrix(VSFG.Row, 1) & "' " & _
                    " AND det_flu_codigo = '" & VSFG.TextMatrix(i, 1) & "' " & _
                    " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql)
    
            ' Si existe no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSql = " DELETE " & _
                    " FROM det_flujo " & _
                    " WHERE flu_codigo = '" & VSFG.TextMatrix(VSFG.Row, 1) & "' " & _
                    " AND det_flu_codigo = '" & VSFG.TextMatrix(i, 1) & "' " & _
                    " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar (strSql), "M"
                
            End If
        ElseIf VSFG2.TextMatrix(i, VSFG2.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        CargaDetalle
    End If

End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    strSql = " SELECT flu_codigo,flu_nombre,flu_descripcion,flu_fechamod, flu_usumod, '0' as modi " & _
             " FROM flujo " & _
             " WHERE emp_codigo LIKE '" & strEmpresa & "'" & _
             " ORDER BY flu_nombre "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum
    
End Sub
Private Sub CargaDetalle()
    strSql = " SELECT det_flu_codigo,det_flu_nombre,det_flu_descripcion,det_flu_tiempo,res_flu_codigo,det_flu_fechamod, det_flu_usumod, '0' as modi " & _
             " FROM det_flujo " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " AND flu_codigo = '" & VSFG.TextMatrix(VSFG.Row, 1) & "'" & _
             " ORDER BY det_flu_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG2.PonerNum
    
    strSql = " SELECT res_flu_codigo,res_flu_nombre " & _
             " FROM responsable_flujo " & _
             " ORDER BY res_flu_nombre"
    clsCon_Def.Ejecutar strSql
    
    VSFG2.ColComboList(5) = VSFG2.BuildComboList(clsCon_Def.adorec_Def, "res_flu_codigo,*res_flu_nombre", "res_flu_codigo")
    
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            strSql = " UPDATE flujo " & _
                 " SET flu_nombre='" & UCase(VSFG.TextMatrix(i, 2)) & "'," & _
                 " flu_descripcion='" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                 " flu_fechamod=CURRENT_TIMESTAMP," & _
                 " flu_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND flu_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                 
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 1) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el código", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT flu_codigo" & _
                    " FROM flujo " & _
                    " WHERE emp_codigo='" & strEmpresa & "'" & _
                    " AND flu_codigo='" & VSFG.TextMatrix(i, 1) & "'"
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    strSql = " INSERT INTO flujo(emp_codigo, flu_codigo, " & _
                             " flu_nombre, flu_descripcion, " & _
                             " flu_fechamod, flu_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(i, 1)) & "'," & _
                             " '" & UCase(VSFG.TextMatrix(i, 2)) & "', '" & UCase(VSFG.TextMatrix(i, 3)) & "', " & _
                            " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                Else
                    MsgBox "El código de" & Tipo2 & " ya existe", vbInformation, "Ingreso"
                End If
             End If
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
        
            strSql = " SELECT count(flu_codigo) as existe " & _
                    " FROM det_flujo " & _
                    " WHERE flu_codigo = '" & VSFG.TextMatrix(i, 1) & "' " & _
                    " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql)
    
            ' Si existe no puedo eliminar
            If clsCon_Def.adorec_Def("existe") > 0 Then
                MsgBox "No Puede eliminar " & Tipo2, vbInformation, "Eliminación"
            Else
                strSql = " DELETE " & _
                    " FROM flujo " & _
                    " WHERE flu_codigo = '" & VSFG.TextMatrix(i, 1) & "' " & _
                    " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar (strSql), "M"
                
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga
    End If
    
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > 0 Then
        If NewRow <> OldRow Then
            CargaDetalle
        End If
    End If
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow And OldRow > 0 And VSFG.Rows > 1 Then
        If VSFG.TextMatrix(OldRow, VSFG.Cols - 1) <> 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG2_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow And OldRow > 0 And VSFG2.Rows > 1 Then
        If VSFG2.TextMatrix(OldRow, VSFG2.Cols - 1) <> 0 Then
            'Cancel = True
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

Private Sub VSFG2_DblClick()
    Dim i As Long
    Set DAT = New frmDatos
    If VSFG2.Row >= 1 Then
        DAT.Show
        DAT.VSFG.Rows = VSFG2.Cols
        For i = 1 To VSFG2.Cols - 1
            DAT.VSFG.TextMatrix(i, 0) = VSFG2.TextMatrix(0, i)
            DAT.VSFG.Cell(flexcpText, i, 1) = VSFG2.Cell(flexcpTextDisplay, VSFG2.Row, i)
            If VSFG2.ColComboList(i) <> "" Then
                DAT.VSFG.TextMatrix(i, 2) = VSFG2.ColComboList(i)
                DAT.VSFG.Cell(flexcpText, i, 3) = VSFG2.Cell(flexcpText, VSFG2.Row, i)
            End If
        Next i
        DAT.VSFG.Cell(flexcpBackColor, 1, 1, DAT.VSFG.Rows - 1, 1) = VSFG2.Cell(flexcpBackColor, VSFG2.Row, VSFG2.Col)
        DAT.VSFG.RowHidden(DAT.VSFG.Rows - 1) = True
        Set DAT.VSFGOrigen = VSFG2
        DAT.VSFGOrigen.Tag = VSFG2.Row
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

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 0 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 2 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -2 Then
        If Col >= VSFG2.Cols - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 3 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG2.Cols - 3 Then
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
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    Set ucrtVSFG2.VSFGControl = VSFG2
    IniDato
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
End Sub

Private Sub VSFG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -2 Then
        VSFG2.TextMatrix(Row, VSFG2.Cols - 1) = 2
    ElseIf Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -3 Then
        VSFG2.TextMatrix(Row, VSFG2.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG2_KeyPress(KeyAscii As Integer)
    ucrtVSFG2.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub

Private Sub VSFG2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG2.MouseRow > 0 Then
        ucrtVSFG2.VerMenu
    End If
End Sub
