VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmJuntarPedidos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Juntar Pedidos"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJuntarPedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9270
   Begin VB.CommandButton cmdJuntar 
      Caption         =   "Juntar"
      Height          =   375
      Left            =   3068
      TabIndex        =   13
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtCantEnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   6360
      Width           =   930
   End
   Begin VB.TextBox txtCantPed 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Top             =   6360
      Width           =   930
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Pedidos:"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   9015
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "ACT"
         Height          =   1815
         Left            =   8640
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkSeleccionar 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Seleccionar todos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   8475
         _cx             =   1986280037
         _cy             =   1986268289
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
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmJuntarPedidos.frx":030A
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
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   600
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4748
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   8955
      _cx             =   1986280884
      _cy             =   1986269136
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmJuntarPedidos.frx":0580
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
End
Attribute VB_Name = "frmJuntarPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'
Private strSql As String
Private clsSql As New clsConsulta

Private Sub chkSeleccionar_Click()

    Dim tip As Integer
    tip = 0
    If CBool(chkSeleccionar.Value) = True Then
        tip = 1
    End If
    For i = 1 To VSFGPeds.Rows - 1
        If VSFGPeds.TextMatrix(i, 8) = "Facturado" Or VSFGPeds.TextMatrix(i, 8) = "De Baja" Then
            VSFGPeds.TextMatrix(i, 0) = 0
        Else
            VSFGPeds.TextMatrix(i, 0) = tip
        End If
    Next i
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.MatchedWithList = True Then
        CargaPedidos
    End If
End Sub

Private Sub CargaPedidos()

    strSql = " SELECT tip_ped_ptofac " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    clsSql.Ejecutar strSql

    strSql = " SELECT '' as seleccionar,ped_codigo, ped_fechamod, CONCAT(persona.per_apellido,' ',persona.per_nombre) as nombC, " & _
            " CONCAT(ven_apellido,' ',ven_nombre) as nombV,CONCAT(GZ.per_apellido,' ',GZ.per_nombre) as nombG,CONCAT(DI.per_apellido,' ',DI.per_nombre) as nombD, " & _
            " ped_observacion, est_descripcion, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
            " pedido.ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
            " persona.per_sec_publico,persona.per_siniva,persona.per_fac_flete,persona.per_dcto,pedido.tar_cre_codigo " & _
            " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
            " INNER JOIN persona ON pedido.per_codigo = persona.per_codigo " & _
            " AND pedido.emp_codigo = persona.emp_codigo " & _
            " INNER JOIN tipo_factura ON pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo LEFT JOIN vendedor ON vendedor.ven_codigo = persona.ven_codigo " & _
            " AND vendedor.emp_codigo = persona.emp_codigo " & _
            " LEFT JOIN persona as GZ ON persona.emp_codigo = GZ.emp_codigo " & _
            " AND persona.per_codigo_ref = GZ.per_codigo " & _
            " LEFT JOIN persona as DI ON persona.emp_codigo = DI.emp_codigo " & _
            " AND persona.per_codigo_ref2 = DI.per_codigo " & _
            " LEFT JOIN tarjeta_credito ON pedido.emp_codigo = tarjeta_credito.emp_codigo AND pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo " & _
            " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado=1 AND pedido.per_codigo='" & cmbCliente.BoundText & "'" & _
            " AND ped_codigo LIKE CONCAT('" & strSucursal & clsSql.adorec_Def(0) & "'+0,'%') " & _
            " ORDER BY ped_estado,ped_codigo "
    clsSql.Ejecutar strSql
    Set VSFGPeds.DataSource = clsSql.adorec_Def.DataSource
    chkSeleccionar.Value = 0
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        LimpiarTodo
    Else
        Exit Sub
    End If
    
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
    VSFGPeds.Clear 1
    VSFGPeds.Rows = 1
End Sub

Private Sub cmdActualizar_Click()
    CargaPedidos
End Sub

Private Sub cmdJuntar_Click()
    Dim pedFin As String
    Dim strPed As String
    Dim i As Long
    For i = 1 To VSFGPeds.Rows - 1
        If Abs(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = 1 Then
            pedFin = VSFGPeds.TextMatrix(i, 1)
            strPed = strPed & pedFin & ","
        End If
    Next i
    strPed = Left(strPed, Len(strPed) - 1)
    MsgBox "Todos los pedidos se juntaran en el pedido:" & vbNewLine & pedFin, vbInformation, "Pedidos"
    strSql = " UPDATE pedido " & _
             " SET ped_observacion='PEDIDOS JUNTADOS:" & strPed & "', " & _
             " ped_fechamod=CURRENT_TIMESTAMP , " & _
             " ped_usumod='" & strUsuario & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_codigo=" & pedFin & " "
    clsSql.Ejecutar strSql, "M"
    strSql = " UPDATE pedido " & _
             " SET ped_observacion='PEDIDO JUNTADO EN:" & pedFin & "', " & _
             " ped_estado=3 " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_codigo IN (" & strPed & ") and ped_codigo!='" & pedFin & "' "
    clsSql.Ejecutar strSql, "M"
    
    strSql = "DELETE FROM det_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_codigo='" & pedFin & "'"
    clsSql.Ejecutar strSql, "M"
    
    For i = 1 To VSFG.Rows - 1
        strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                 " det_ped_cant_entregada, det_ped_cant_confirmada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
                 " VALUES ('" & strEmpresa & "'," & pedFin & ",'" & VSFG.TextMatrix(i, 1) & "','" & VSFG.TextMatrix(i, 0) & "'," & VSFG.TextMatrix(i, 3) & ", " & _
                 VSFG.TextMatrix(i, 4) & "," & VSFG.TextMatrix(i, 5) & "," & VSFG.TextMatrix(i, 6) & "," & VSFG.TextMatrix(i, 7) & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
        clsSql.Ejecutar (strSql), "M"
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    cargarTipoPedido
    
End Sub

Private Sub VSFGPeds_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGPeds_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Marca toda la fila con otra tonalidad si el pedido puede ser vendido
    If Col = 0 And Row <> 0 Then
        If Abs(VSFGPeds.TextMatrix(Row, 0)) = 1 Then
            VSFGPeds.Select Row, 0, Row, VSFGPeds.Cols - 1
            VSFGPeds.FillStyle = flexFillRepeat
            VSFGPeds.CellBackColor = &HC0C0FF
        Else
            VSFGPeds.Select Row, 0, Row, VSFGPeds.Cols - 1
            VSFGPeds.FillStyle = flexFillRepeat
            VSFGPeds.CellBackColor = &HFFFFFF
        End If
        CargaDetallePedido
    End If
End Sub

Private Sub CargaDetallePedido()
    Dim i As Long
    Dim ped As String
    
    For i = 1 To VSFGPeds.Rows - 1
        If Abs(FormatoD0(VSFGPeds.TextMatrix(i, 0))) = 1 Then
            ped = ped & VSFGPeds.TextMatrix(i, 1) & ","
        End If
    Next i
    VSFG.Clear 1
    VSFG.Rows = 1
    If ped <> "" Then
        ped = Left(ped, Len(ped) - 1)
        strSql = " SELECT dep_codigo, det_pedido.prd_codigo, prd_nombre, SUM(det_ped_cant_pedida), SUM(det_ped_cant_entregada), SUM(det_ped_cant_confirmada), MAX(det_ped_precio),  " & _
                " SUM(det_ped_dcto)" & _
                " FROM ((pedido INNER JOIN det_pedido ON (pedido.ped_codigo = det_pedido.ped_codigo) " & _
                " AND (pedido.emp_codigo = det_pedido.emp_codigo)) INNER JOIN producto " & _
                " ON (det_pedido.emp_codigo = producto.emp_codigo) AND (det_pedido.prd_codigo = producto.prd_codigo)) " & _
                " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                " det_pedido.ped_codigo IN (" & ped & ") GROUP BY dep_codigo, det_pedido.prd_codigo, prd_nombre " & _
                " ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo "
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    End If
    txtCantEnt.Text = 0
    txtCantPed.Text = 0
    
    For i = 1 To VSFG.Rows - 1
        txtCantEnt.Text = FormatoD0(txtCantEnt.Text) + FormatoD0(VSFG.TextMatrix(i, 5))
        txtCantPed.Text = FormatoD0(txtCantPed.Text) + FormatoD0(VSFG.TextMatrix(i, 3))
    Next i
End Sub
