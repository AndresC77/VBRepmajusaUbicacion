VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSeguimientoFlujo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Flujos"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "frmSeguimientoFlujo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12075
   Begin VB.CommandButton cmdEnviarCorreo 
      Caption         =   "&Enviar Correos a Responsables"
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmbAceptar1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   10200
      TabIndex        =   7
      Top             =   4200
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   10200
      TabIndex        =   0
      Top             =   6840
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGFlujos 
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11820
      _cx             =   20849
      _cy             =   2328
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSeguimientoFlujo.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   1920
      Left            =   120
      TabIndex        =   3
      Top             =   2280
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
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSeguimientoFlujo.frx":03E9
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
      TabIndex        =   4
      Top             =   1920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGDet 
      Height          =   2160
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   11820
      _cx             =   20849
      _cy             =   3810
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
      Rows            =   1
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSeguimientoFlujo.frx":0507
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
   Begin NEED2.uctrVSFG ucrtVSFG3 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmSeguimientoFlujo"
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

Private Sub cmbAceptar1_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    VSFG2.Select 1, VSFG2.Cols - 1
    VSFG2.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    
    For i = 1 To VSFG2.Rows - 1
        'update
        If VSFG2.TextMatrix(i, VSFG2.Cols - 1) = 3 Then
            strSql = " UPDATE historia_flujo " & _
                 " SET his_flu_nombre='" & UCase(VSFG2.TextMatrix(i, 2)) & "'," & _
                 " his_flu_descripcion='" & UCase(VSFG2.TextMatrix(i, 3)) & "'," & _
                 " his_flu_fechamod=CURRENT_TIMESTAMP," & _
                 " his_flu_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND his_flu_codigo='" & VSFG2.TextMatrix(i, 1) & "'" & _
                 " AND flu_codigo='" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'"
            clsCon_Def.Ejecutar strSql, "M"
        'insert
        ElseIf VSFG2.TextMatrix(i, VSFG2.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG2.TextMatrix(i, 2) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta nombre", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG2.TextMatrix(i, 3) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta descripcion", vbInformation, "Ingreso"
                control = 1
            Else
                Dim nf As Long
                Dim FechaCalc As String
                strSql = " SELECT COALESCE(MAX(his_flu_codigo),0)+1 as n" & _
                    " FROM historia_flujo " & _
                    " WHERE emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar strSql
                'controla que no se repita el código
                If clsCon_Def.adorec_Def.RecordCount = 0 Then
                    nf = 1
                Else
                    nf = clsCon_Def.adorec_Def("n")
                End If
                
                strSql = " INSERT INTO historia_flujo(emp_codigo, his_flu_codigo, flu_codigo, his_flu_nombre," & _
                         " his_flu_descripcion, his_flu_fechamod, his_flu_usumod) " & _
                         " VALUES ('" & strEmpresa & "','" & nf & "','" & UCase(VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1)) & "','" & UCase(VSFG2.TextMatrix(i, 2)) & "'," & _
                         " '" & UCase(VSFG2.TextMatrix(i, 3)) & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
                
                strSql = " SELECT det_flu_codigo," & _
                         " LEFT(CURRENT_TIMESTAMP,10) as fecha,det_flu_tiempo " & _
                         " FROM det_flujo " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND flu_codigo='" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'" & _
                         " ORDER BY det_flu_codigo"
                clsCon_Def.Ejecutar strSql, "M"
                FechaCalc = clsCon_Def.adorec_Def("fecha")
                While Not clsCon_Def.adorec_Def.EOF
                    strSql = " INSERT INTO det_historia_flujo" & _
                             " (emp_codigo, his_flu_codigo, flu_codigo, det_flu_codigo, " & _
                             " det_his_flu_fecha_prevista_inicio, det_his_flu_fecha_prevista_fin, " & _
                             " det_his_flu_fecha_real_inicio, det_his_flu_fecha_real_fin, " & _
                             " det_his_flu_observacion, det_his_flu_fechamod, det_his_flu_usumod) " & _
                             " VALUES('" & strEmpresa & "','" & nf & "','" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "','" & clsCon_Def.adorec_Def("det_flu_codigo") & "'," & _
                             " '" & FechaCalc & "','" & SumaDiasHabiles(FechaCalc, clsCon_Def.adorec_Def("det_flu_tiempo")) & "'," & _
                             " '',''," & _
                             " '',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    clsAux.Ejecutar strSql, "M"
                    FechaCalc = SumaDiasHabiles(FechaCalc, clsCon_Def.adorec_Def("det_flu_tiempo") + 1)
                    clsCon_Def.adorec_Def.MoveNext
                Wend
                
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
        Carga
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
    Set VSFGFlujos.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG.PonerNum
    
End Sub
Private Sub CargaDetalle()
    strSql = " SELECT historia_flujo.his_flu_codigo,his_flu_nombre,his_flu_descripcion,MIN(det_his_flu_fecha_prevista_inicio),MAX(det_his_flu_fecha_prevista_fin),his_flu_fechamod, his_flu_usumod, '0' as modi " & _
             " FROM historia_flujo INNER JOIN det_historia_flujo " & _
             " ON historia_flujo.emp_codigo=det_historia_flujo.emp_codigo " & _
             " AND historia_flujo.flu_codigo=det_historia_flujo.flu_codigo " & _
             " AND historia_flujo.his_flu_codigo=det_historia_flujo.his_flu_codigo " & _
             " WHERE historia_flujo.emp_codigo = '" & strEmpresa & "'" & _
             " AND historia_flujo.flu_codigo = '" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'" & _
             " GROUP BY historia_flujo.his_flu_codigo " & _
             " ORDER BY historia_flujo.his_flu_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG2.PonerNum
    CargaDetalleHistorial
End Sub
Private Sub CargaDetalleHistorial()
    Dim i As Long
    strSql = " SELECT det_flujo.det_flu_codigo,det_flu_nombre,det_flu_descripcion," & _
             " det_flu_tiempo,res_flu_nombre,LEFT(det_his_flu_fecha_prevista_inicio,10)," & _
             " LEFT(det_his_flu_fecha_prevista_fin,10),if(det_his_flu_fecha_real_inicio='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_inicio,10))," & _
             " IF(det_his_flu_fecha_real_fin='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_fin,10))," & _
             " TO_DAYS(IF(det_his_flu_fecha_real_fin='0000-00-00 00:00:00','" & HoyDia & "',LEFT(det_his_flu_fecha_real_fin,10)))-TO_DAYS(LEFT(det_his_flu_fecha_prevista_fin,10)) as retraso,det_his_flu_observacion," & _
             " responsable_flujo.usu_codigo,det_his_flu_fechamod, det_his_flu_usumod, '0' as modi " & _
             " FROM det_historia_flujo INNER JOIN det_flujo " & _
             " ON det_historia_flujo.emp_codigo=det_flujo.emp_codigo " & _
             " AND det_historia_flujo.flu_codigo=det_flujo.flu_codigo " & _
             " AND det_historia_flujo.det_flu_codigo=det_flujo.det_flu_codigo " & _
             " INNER JOIN responsable_flujo ON det_flujo.emp_codigo=responsable_flujo.emp_codigo" & _
             " AND det_flujo.res_flu_codigo=responsable_flujo.res_flu_codigo" & _
             " WHERE det_historia_flujo.emp_codigo = '" & strEmpresa & "'" & _
             " AND det_historia_flujo.flu_codigo = '" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'" & _
             " AND det_historia_flujo.his_flu_codigo = '" & VSFG2.TextMatrix(VSFG2.Row, 1) & "'" & _
             " GROUP BY det_historia_flujo.det_flu_codigo " & _
             " ORDER BY det_flu_codigo "
    clsCon_Def.Ejecutar strSql
    Set VSFGDet.DataSource = clsCon_Def.adorec_Def.DataSource
    ucrtVSFG3.PonerNum
    
    For i = 1 To VSFGDet.Rows - 1
        If VSFGDet.TextMatrix(i, 10) > 0 Then
            VSFGDet.Cell(flexcpBackColor, i, 1, i, VSFGDet.Cols - 1) = vbRed
        ElseIf VSFGDet.TextMatrix(i, 10) = 0 Or VSFGDet.TextMatrix(i, 10) = 1 Then
            VSFGDet.Cell(flexcpBackColor, i, 1, i, VSFGDet.Cols - 1) = vbGreen
        Else
            VSFGDet.Cell(flexcpBackColor, i, 1, i, VSFGDet.Cols - 1) = vbWhite
        End If
    Next i
    
    
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

Private Sub cmdEnviarCorreo_Click()
    Dim clsAuxCorreo As New clsConsulta
    Dim strAsunto As String
    Dim strCuerpo As String
    Dim strCopia As String
    strCopia = "acevallos@enlacedigital.com.ec"
    clsAuxCorreo.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT det_flujo.det_flu_codigo,det_flu_nombre,det_flu_descripcion," & _
             " det_flu_tiempo,res_flu_nombre,LEFT(det_his_flu_fecha_prevista_inicio,10) as det_his_flu_fecha_prevista_inicio," & _
             " LEFT(det_his_flu_fecha_prevista_fin,10) as det_his_flu_fecha_prevista_fin,if(det_his_flu_fecha_real_inicio='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_inicio,10))," & _
             " if(det_his_flu_fecha_real_fin='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_fin,10)),det_his_flu_observacion," & _
             " responsable_flujo.usu_codigo,res_flu_email,res_flu_nombre,det_his_flu_fechamod, det_his_flu_usumod, '0' as modi " & _
             " FROM det_historia_flujo INNER JOIN det_flujo " & _
             " ON det_historia_flujo.emp_codigo=det_flujo.emp_codigo " & _
             " AND det_historia_flujo.flu_codigo=det_flujo.flu_codigo " & _
             " AND det_historia_flujo.det_flu_codigo=det_flujo.det_flu_codigo " & _
             " INNER JOIN responsable_flujo ON det_flujo.emp_codigo=responsable_flujo.emp_codigo" & _
             " AND det_flujo.res_flu_codigo=responsable_flujo.res_flu_codigo" & _
             " WHERE det_historia_flujo.emp_codigo = '" & strEmpresa & "'" & _
             " AND det_historia_flujo.flu_codigo = '" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'" & _
             " AND det_historia_flujo.his_flu_codigo = '" & VSFG2.TextMatrix(VSFG2.Row, 1) & "'" & _
             " GROUP BY det_historia_flujo.det_flu_codigo " & _
             " ORDER BY det_flu_codigo "
    clsAuxCorreo.Ejecutar strSql
    strAsunto = "Seguimiento de Flujo N." & VSFG2.TextMatrix(VSFG2.Row, 1) & " - " & VSFG2.TextMatrix(VSFG2.Row, 2)
    While Not clsAuxCorreo.adorec_Def.EOF
        strCuerpo = "Estomad@." & vbNewLine & vbNewLine & _
                    "Se le ha asignado la responsabilidad de la siguiente tarea en el módulo de flujos:" & vbNewLine & _
                    "Tarea: " & clsAuxCorreo.adorec_Def("det_flu_nombre") & vbNewLine & _
                    "Descripción: " & clsAuxCorreo.adorec_Def("det_flu_descripcion") & vbNewLine & _
                    "Tiempo máximo de la tarea: " & clsAuxCorreo.adorec_Def("det_flu_tiempo") & vbNewLine & _
                    "Fecha estimada de inicio de la tarea: " & clsAuxCorreo.adorec_Def("det_his_flu_fecha_prevista_inicio") & vbNewLine & _
                    "Fecha estimada de finalización de la tarea: " & clsAuxCorreo.adorec_Def("det_his_flu_fecha_prevista_fin")
        EnviarMail "Flujos", "jefedeproductojsn@rbimportadores.com", clsAuxCorreo.adorec_Def("res_flu_nombre"), clsAuxCorreo.adorec_Def("res_flu_email"), strCopia, strAsunto, strCuerpo
        clsAuxCorreo.adorec_Def.MoveNext
    Wend
End Sub

Private Sub VSFG2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow <> 0 Then
        CargaDetalleHistorial
    End If
    If NewRow = 1 And OldRow = 1 And VSFGDet.Rows = 1 Then
        CargaDetalleHistorial
    End If
End Sub

Private Sub VSFG2_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow And OldRow > 0 And VSFG2.Rows > 1 Then
        'If VSFG2.TextMatrix(OldRow, VSFG2.Cols - 1) <> 0 Then
            'Cancel = True
        'End If
    End If
End Sub

Private Sub VSFGDet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If Row = 1 And Col = 6 Then
        For i = 1 To VSFGDet.Rows - 1
            VSFGDet.TextMatrix(i, 7) = SumaDiasHabiles(VSFGDet.TextMatrix(i, 6), VSFGDet.TextMatrix(i, 4))
            strSql = " UPDATE det_historia_flujo " & _
                     " SET det_his_flu_fecha_prevista_inicio='" & VSFGDet.TextMatrix(i, 6) & "'," & _
                     " det_his_flu_fecha_real_inicio=if(1=" & i & ",'" & VSFGDet.TextMatrix(i, 6) & "',det_his_flu_fecha_real_inicio)," & _
                     " det_his_flu_fecha_prevista_fin='" & VSFGDet.TextMatrix(i, 7) & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND his_flu_codigo='" & VSFG2.TextMatrix(VSFG2.Row, 1) & "'" & _
                     " AND flu_codigo='" & VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1) & "'" & _
                     " AND det_flu_codigo='" & VSFGDet.TextMatrix(i, 1) & "'"
            clsAux.Ejecutar strSql, "M"
            If VSFGDet.Rows - 1 >= i + 1 Then
                VSFGDet.TextMatrix(i + 1, 6) = SumaDiasHabiles(VSFGDet.TextMatrix(i, 6), VSFGDet.TextMatrix(i, 4) + 1)
            End If
        Next i
    End If
End Sub

Private Sub VSFGDet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Row = 1 And Col = 6) Then
        Cancel = True
    End If
End Sub

Private Sub VSFGDet_DblClick()
    If strUsuario = VSFGDet.TextMatrix(VSFGDet.Row, 12) Then
        frmHacerSeguimientoFlujo.txtTituloFlujo.Text = VSFGFlujos.TextMatrix(VSFGFlujos.Row, 2)
        frmHacerSeguimientoFlujo.txtTituloFlujo.Tag = VSFGFlujos.TextMatrix(VSFGFlujos.Row, 1)
        frmHacerSeguimientoFlujo.txtNombreFlujo.Text = VSFG2.TextMatrix(VSFG2.Row, 2)
        frmHacerSeguimientoFlujo.txtNombreFlujo.Tag = VSFG2.TextMatrix(VSFG2.Row, 1)
        frmHacerSeguimientoFlujo.txtDescripcionFlujo.Text = VSFG2.TextMatrix(VSFG2.Row, 3)
        frmHacerSeguimientoFlujo.txtFechaInicioFlujo.Text = VSFG2.TextMatrix(VSFG2.Row, 4)
        frmHacerSeguimientoFlujo.txtFechaFinFlujo.Text = VSFG2.TextMatrix(VSFG2.Row, 5)
        frmHacerSeguimientoFlujo.txtNombreTarea.Tag = VSFGDet.TextMatrix(VSFGDet.Row, 1)
        frmHacerSeguimientoFlujo.txtNombreTarea.Text = VSFGDet.TextMatrix(VSFGDet.Row, 2)
        frmHacerSeguimientoFlujo.txtDescripcionTarea.Text = VSFGDet.TextMatrix(VSFGDet.Row, 3)
        frmHacerSeguimientoFlujo.txtNumeroDias.Text = VSFGDet.TextMatrix(VSFGDet.Row, 4)
        frmHacerSeguimientoFlujo.txtFechaInicioTareaProgramada.Text = VSFGDet.TextMatrix(VSFGDet.Row, 6)
        frmHacerSeguimientoFlujo.txtFechaFinTareaProgramada.Text = VSFGDet.TextMatrix(VSFGDet.Row, 7)
        frmHacerSeguimientoFlujo.txtFechaInicioTareaReal.Text = VSFGDet.TextMatrix(VSFGDet.Row, 8)
        frmHacerSeguimientoFlujo.txtFechaFinTareaReal.Text = VSFGDet.TextMatrix(VSFGDet.Row, 9)
        frmHacerSeguimientoFlujo.txtObservaciones.Text = VSFGDet.TextMatrix(VSFGDet.Row, 11)
        frmHacerSeguimientoFlujo.Show
    End If
End Sub

Private Sub VSFGFlujos_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow <> 0 Then
        CargaDetalle
    End If
End Sub

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 0 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 2 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -2 Then
        If Col >= VSFG2.Cols - 5 Then
            Cancel = True
        End If
    ElseIf Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = 3 Or Val(VSFG2.TextMatrix(Row, VSFG2.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG2.Cols - 5 Then
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
    Set ucrtVSFG.VSFGControl = VSFGFlujos
    ucrtVSFG.Inicializar False, False, False, True, True, True, False, False, True
    Set ucrtVSFG2.VSFGControl = VSFG2
    ucrtVSFG2.Inicializar True, False, True, True, True, True, False, False, True
    Set ucrtVSFG3.VSFGControl = VSFGDet
    ucrtVSFG3.Inicializar False, False, True, True, True, True, False, False, True
    IniDato
    Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
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
