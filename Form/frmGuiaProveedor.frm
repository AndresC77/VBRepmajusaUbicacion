VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGuiaProveedor 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantener Guías de Remisión Simples"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmGuiaProveedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   12285
   Begin VB.CommandButton CmdDevolver 
      Caption         =   "Devolver"
      Height          =   375
      Left            =   3855
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6975
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   5415
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "DATOS DE GUIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   285
         Left            =   9480
         TabIndex        =   8
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         Value           =   42706.593125
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGGuia 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   11895
         _cx             =   20981
         _cy             =   5953
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmGuiaProveedor.frx":030A
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   2055
         TabIndex        =   4
         Top             =   360
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Factura:"
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
         Left            =   8160
         TabIndex        =   7
         Top             =   420
         Width           =   1320
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un Proveedor:"
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
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   1860
      End
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   495
      Picture         =   "frmGuiaProveedor.frx":04FD
      ToolTipText     =   "Elimina una Fila"
      Top             =   4800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   735
      Picture         =   "frmGuiaProveedor.frx":0633
      Top             =   4800
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmGuiaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################
'#  Forma para ver y facturar un proyecto de venta, en base a los ingresos y
'#  egresos que se han realizado en el proyecto.
'#  frmV_FacProVenta V1.0
'#  Copyright (C) 2002
'#
'#  Opciones que permite:
'#  *   En una lista se despliegan los datos del los distintos proyectos de
'#      trabajo de una emprea como el cliente y el vendedor que lo atiende y el
'#      estado del mismo.
'#  *   De igual manera es necesario seleccionar el tipo de facturación que se
'#      va a aplicar al proyecto y la fecha en que se lo factura.
'#  *   Es necesario también seleccionar la forma de pago.
'#  *   El usuario puede seleccionar los posibles recargos que puede generar
'#      la facturación de proyecto.
'#
'#  Procesos internos que maneja:
'#  *   La lista que muestra los distintos proyectos, se refresca automáticamente
'#      cada 20 segundos para buscar un nuevo proyecto de trabajo.
'#  *   Al dar un click en la lista de proyectos, automáticamente se cargan los
'#      detalles de los movimientos del mismo en un segundo grid.
'#  *   Una vez que el proyecto ha sido facturado su estado pasa a vendido. Al
'#      igual que su cotización relacionada.
'#  *   Se pueden ver solo los proyectos que no están facturados y los que ya
'#      se han facturado el día de hoy.
'#  *   Una vez que se va a facturar el proyecto se generan automáticamente las
'#      respectivas retenciones que puede tener el cliente del mismo.
'#
'#  Tablas que maneja:
'#
'#  persona:
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el
'#      proyecto que se está facturando.
'#  *   También se extrae el nombre del vendedor asignado al pedido.
'#  persona_ret:
'#  *   De esta tala se extraen las diferentes retenciones que puede tener un
'#      cliente determinado para luego aplicarlas a esta factura.
'#  retencion:
'#  *   De aquí se extraen los valores y descripciones de las retenciones, que
'#      se aplicarán posteriormente.
'#  det_egreso:
'#  *   En esta tabla se guardan los detalles del nuevo documento de egreso de
'#      productos.
'#  ocargo:
'#  *   De esta tabla se extraen los diferentes recargos que se puede manejar
'#      al realizar un nuevo egreso de productos de bodega, como pueden ser:
'#      transporte, fletes, etc.
'#  det_egreso_c:
'#  *   En esta tabla se guardan los diferentes recargos que puede tener esta
'#      nueva compra o egreso de productos.
'#  det_egreso_ret:
'#  *   En esta tabla se guardan los valores de las retenciones aplicadas a este
'#      ingreso de productos a bodega.
'#
'################################################################################

Private clsClie As New clsConsulta
Private clsSql As New clsConsulta
Private clsFPago As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsPrds As New clsConsulta
Private clsBods As New clsConsulta
Private clsLstPrds As New clsConsulta

Private IVA As Double
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsClie = Nothing
    Set clsSql = Nothing
    Set clsFPago = Nothing
    Set clsRecargos = Nothing
    Set clsPrds = Nothing
    Set clsBods = Nothing
    Set clsLstPrds = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With VSFGReca
        For i = 1 To (.Rows - 1)
            .TextMatrix(i, 0) = i
            If conBot = True Then
                'Coloca los botones de elimniar fila en el grid
                .Cell(flexcpPicture, i, 0) = imgBtnUp
                .Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub
Private Sub PonerBotonesFac(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With VSFG
        For i = 1 To (.Rows - 1)
            '.TextMatrix(i, 0) = i
            If conBot = True And Val(.TextMatrix(i, 7)) <> 1 Then
                'Coloca los botones de elimniar fila en el grid
                .Cell(flexcpPicture, i, 0) = imgBtnUp
                .Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub
Private Sub CalcuReca()
    'Calcula el total del pedido
    Dim Suma As Double
    For i = 1 To VSFGReca.Rows - 1
        Suma = Suma + FormatoD2(VSFGReca.TextMatrix(i, 3))
    Next i
    TxtRecargo = Format(Suma, "####0.00")
    TxtTotal = Format(Suma + Val(TxtIva) + Val(TxtSubTotal), "####0.00")
End Sub

Private Sub CalcuTotal()
    'Calcula es total del pedido
    Dim Suma As Double, Columna As Long
    'Busca cual es la columna del total
    For i = 0 To VSFG.Cols - 1
        If VSFG.TextMatrix(0, i) = "Total" Then
            Columna = i
            Exit For
        End If
    Next i
    For i = 1 To VSFG.Rows - 1
        Suma = Suma + FormatoD2(VSFG.TextMatrix(i, Columna))
    Next i
    'Coloca los totales parciales de la factura
    TxtSubTotal = Format(Suma, "####0.00")
    TxtIva = Format(Suma * IVA / 100, "####0.00")
    TxtTotal = Format(Suma + Val(TxtIva) + Val(TxtRecargo) - Val(TxtDesc), "####0.00")
End Sub


Private Sub cmbCliente_Change()
    CmdLimpiar = True
    'Cargar datos de Guias
    If cmbCliente.MatchedWithList = True Then
        CargarGuias
    Else
        VSFGGuia.Clear 1
        VSFGGuia.Rows = 1
    End If
    
End Sub

Private Sub CmbFpago_Change()
'    CmdLimpiar = True
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Sub cmdDevolver_Click()
    If (Devolver = True) Then Unload Me
End Sub

Private Sub cmdLimpiar_Click()
    'Muestra el formulario como si se hubiera cargado por primera vez
    'CmdConfirmar.Enabled = False
    'CmdDeBaja.Enabled = False
    TxtSubTotal = ""
    TxtTotal = ""
    TxtRecargo = ""
    TxtIva = ""
    TxtDesc = ""
    fila = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
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
    clsClie.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
'****** PROVEEDOR
    'Obtiene todos los clientes de una empresa con su respectiva lista de precios y vendedor asociado
    strSQL = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')') as nombC " & _
             " FROM persona " & _
             " Where persona.emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY nombC "
    clsClie.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsClie.adorec_Def.DataSource
    If Not clsClie.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
        cmbCliente = clsClie.adorec_Def("nombC")
    Else
        cmbCliente = "No hay proveedores en la empresa: " & strEmpresa
    End If
    'Selecciona el primer elemento del combo de cotizaciones
    dtpFecha.Value = HoyDia
End Sub

Private Sub TxtDesc_Change()
    'TxtDesc = Replace(TxtDesc, ",", ".")
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    'Valida que solo se ingresen números en el campo de descuento
    If KeyAscii < 44 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDesc_LostFocus()
    'Calcula el total de la factura
    CalcuTotal
End Sub

Private Sub TxtIva_Change()
    'TxtIva = Replace(TxtIva, ",", ".")
End Sub

Private Sub TxtRecargo_Change()
    'TxtRecargo = Replace(TxtRecargo, ",", ".")
End Sub

Private Sub TxtSubTotal_Change()
    'TxtSubTotal = Replace(TxtSubTotal, ",", ".")
End Sub

Private Sub TxtSubTotal_KeyPress(KeyAscii As Integer)
    'Valida que solo se ingresen números en el campo de subtotal
    If KeyAscii < 44 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSubTotal_LostFocus()
    'Calcula el total de la factura
    CalcuTotal
End Sub

Private Sub txtTotal_Change()
    TxtTotal = Format(TxtTotal, "###0.00")
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Verifica que solo se ingresen números tanto en la cantidad como en el precio
    If Col = 4 Or Col = 5 Then
        'Verifica que solo se ingresen números en el campo cantidad
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en la cantidad.", vbInformation, "Cantidad"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
        'Verifica que solo se ingresen números en el campo precio
        If Not IsNumeric(VSFG.TextMatrix(Row, 5)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el precio.", vbInformation, "Precio"
            VSFG.TextMatrix(Row, 5) = intDato
        End If
        'Actualiza el total del producto pedido
        VSFG.TextMatrix(Row, 6) = Val(VSFG.TextMatrix(Row, 5)) * Val(VSFG.TextMatrix(Row, 4))
        CalcuTotal
    End If
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid en caso de ser necesario
    If OldCol = 5 And OldRow = VSFG.Rows - 1 And NewCol = 6 And VSFG.TextMatrix(OldRow, 2) <> "" Then
        VSFG.AddItem ""
        VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, 7)) = 1 And Col <> 5 Then
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
    If r <= 0 Then Exit Sub
    If Val(VSFG.TextMatrix(r, 7)) = 1 Then Exit Sub
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         Dim h As Long
         VSFG.RemoveItem (r)
         For h = 1 To VSFGGuia.Rows - 1
            If VSFGGuia.TextMatrix(h, 12) > r Then
                VSFGGuia.TextMatrix(h, 12) = VSFGGuia.TextMatrix(h, 12) - 1
            End If
         Next h
         PonerBotonesFac
         CalcuTotal
    Else
        VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    'No permite entrar en las celdas de las columnas siguientes
    If NewCol = 3 Or NewCol = 6 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Coloca la descripción del producto en caso que se haga un pedido manual y el usuario haya seleccionado un código de producto
    If Col = 1 Or Col = 2 Then
        If VSFG.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero una bodega", vbInformation, "Bodega"
            VSFG.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
        'Verifica que no se seleccione más de una vez el mismo producto en la misma bodega
'        For i = 1 To VSFG.Rows - 1
'            If VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(i, 2) And VSFG.TextMatrix(Row, 1) = VSFG.TextMatrix(i, 1) And Row <> i Then
'                MsgBox "Ese producto ya fue seleccionado en la bodega " & VSFG.TextMatrix(i, 2) & ", solo cambie la candidad del mismo.", vbInformation, "Producto"
'                VSFG.RemoveItem Row
'                PonerBotones
'                VSFG.Row = i
'                VSFG.Col = 2
'                Exit Sub
'            End If
'        Next i
        'Coloca los datos de un producto seleccionado
        If VSFG.TextMatrix(Row, 2) <> "" Then
            'Busca el producto seleccionado y coloca sus datos respectivos
            clsLstPrds.adorec_Def.MoveFirst
            clsLstPrds.Filtrar "dep_codigo='" & VSFG.TextMatrix(Row, 1) & "' AND prd_codigo='" & VSFG.TextMatrix(Row, 2) & "'"
            If Not clsLstPrds.adorec_Def.EOF Then
                VSFG.TextMatrix(Row, 3) = clsLstPrds.adorec_Def("prd_nombre")
                'Coloca el costo del producto en una columna oculta
                VSFG.TextMatrix(Row, 9) = clsLstPrds.adorec_Def("prd_costo")
                VSFG.TextMatrix(Row, 7) = 0
                VSFG.TextMatrix(Row, 4) = 1
                VSFG.TextMatrix(Row, 6) = VSFG.TextMatrix(Row, 5) * VSFG.TextMatrix(Row, 4)
                VSFG.TextMatrix(Row, 8) = clsLstPrds.adorec_Def("exi_cantidad")
            End If
            clsLstPrds.QuitarFiltro
            CalcuTotal
        End If
    End If
    If Col = 5 Then
        CalcuTotal
    End If
End Sub

Private Sub VSFGGuia_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Then Exit Sub
    If NewCol = 0 Or ((NewCol = 9 Or NewCol = 10) And Abs(VSFGGuia.TextMatrix(NewRow, 0)) = 1) Then
        VSFGGuia.Editable = flexEDKbdMouse
    Else
        VSFGGuia.Editable = flexEDNone
    End If
End Sub

Private Sub VSFGGuia_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Col = 9 Then
        If VSFGGuia.TextMatrix(Row, Col) = "" Then
            VSFGGuia.TextMatrix(Row, Col) = 0
        End If
    ElseIf Col = 10 Then
        If VSFGGuia.TextMatrix(Row, Col) = "" Then
            VSFGGuia.TextMatrix(Row, Col) = 0
        End If
    End If
    If VSFGGuia.Tag <> "A" Then
        If Col = 0 And Row > 0 Then
            If Abs(VSFGGuia.TextMatrix(Row, 0)) = 1 Then
                VSFGGuia.Select Row, 0, Row, 13
                VSFGGuia.FillStyle = flexFillRepeat
                VSFGGuia.CellBackColor = &HC0FFFF
                VSFGGuia.Select Row, 0
            ElseIf Abs(VSFGGuia.TextMatrix(Row, 0)) = 0 Then
                VSFGGuia.Select Row, 0, Row, 13
                VSFGGuia.FillStyle = flexFillRepeat
                VSFGGuia.CellBackColor = &HFFFFFF
                
                If FormatoD2(VSFGGuia.TextMatrix(Row, 11)) > 0 Then
                    For i = 1 To VSFGGuia.Rows - 1
                        If VSFGGuia.TextMatrix(i, 12) > VSFGGuia.TextMatrix(Row, 12) Then
                            VSFGGuia.TextMatrix(i, 12) = VSFGGuia.TextMatrix(i, 12) - 1
                        End If
                    Next i
                    VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                End If
                VSFGGuia.TextMatrix(Row, 12) = 0
                VSFGGuia.TextMatrix(Row, 11) = 0
                VSFGGuia.TextMatrix(Row, 10) = 0
                VSFGGuia.TextMatrix(Row, 9) = 0
                VSFGGuia.Select Row, 0
            End If
        ElseIf (Col = 9 Or Col = 10) And Row > 0 And VSFGGuia.TextMatrix(Row, 9) <> "" And VSFGGuia.TextMatrix(Row, 10) <> "" Then
            If CDbl(VSFGGuia.TextMatrix(Row, 9)) + CDbl(VSFGGuia.TextMatrix(Row, 10)) > CDbl(VSFGGuia.TextMatrix(Row, 7)) Or CDbl(VSFGGuia.TextMatrix(Row, Col)) < 0 Then
                MsgBox "La cantidad debe mayor a 0 y menor a " & VSFGGuia.TextMatrix(Row, 7) - VSFGGuia.TextMatrix(Row, IIf(Col = 10, 9, 10)), vbCritical, "ERROR"
                VSFGGuia.TextMatrix(Row, Col) = 0
                If CDbl(VSFGGuia.TextMatrix(Row, 11)) > 0 Then
                    For i = 1 To VSFGGuia.Rows - 1
                        If VSFGGuia.TextMatrix(i, 12) > VSFGGuia.TextMatrix(Row, 12) Then
                            VSFGGuia.TextMatrix(i, 12) = VSFGGuia.TextMatrix(i, 12) - 1
                        End If
                    Next i
                    VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                    VSFGGuia.TextMatrix(Row, 12) = 0
                End If
            End If
            VSFGGuia.TextMatrix(Row, 11) = VSFGGuia.TextMatrix(Row, 8) * VSFGGuia.TextMatrix(Row, 10)
            If Col = 10 And VSFGGuia.TextMatrix(Row, 12) <> "" Then
                If Val(VSFGGuia.TextMatrix(Row, 12)) = 0 And FormatoD2(VSFGGuia.TextMatrix(Row, 10)) > 0 Then
                    If VSFG.Rows > 1 Then
                        If VSFG.TextMatrix(VSFG.Rows - 1, 2) = "" Then
                            VSFG.RemoveItem VSFG.Rows - 1
                        End If
                    End If
                    VSFG.AddItem vbTab & VSFGGuia.TextMatrix(Row, 3) & vbTab & VSFGGuia.TextMatrix(Row, 4) & vbTab & VSFGGuia.TextMatrix(Row, 5) & vbTab & VSFGGuia.TextMatrix(Row, 10) & vbTab & VSFGGuia.TextMatrix(Row, 8) & vbTab & VSFGGuia.TextMatrix(Row, 11) & vbTab & 1 & vbTab & vbTab & VSFGGuia.TextMatrix(Row, 14)
                    VSFGGuia.TextMatrix(Row, 12) = VSFG.Rows - 1
                ElseIf Val(VSFGGuia.TextMatrix(Row, 12)) > 0 Then
                    If CDbl(VSFGGuia.TextMatrix(Row, 10)) = 0 Then
                        For i = 1 To VSFGGuia.Rows - 1
                            If CDbl(VSFGGuia.TextMatrix(i, 12)) > CDbl(VSFGGuia.TextMatrix(Row, 12)) Then
                                VSFGGuia.TextMatrix(i, 12) = CDbl(VSFGGuia.TextMatrix(i, 12)) - 1
                            End If
                        Next i
                        VSFG.RemoveItem VSFGGuia.TextMatrix(Row, 12)
                        VSFGGuia.TextMatrix(Row, 12) = 0
                    Else
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 4) = VSFGGuia.TextMatrix(Row, 10)
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 5) = VSFGGuia.TextMatrix(Row, 8)
                        VSFG.TextMatrix(VSFGGuia.TextMatrix(Row, 12), 6) = VSFGGuia.TextMatrix(Row, 8) * VSFGGuia.TextMatrix(Row, 10)
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 And VSFGReca.TextMatrix(OldRow, 1) <> "" Then
        VSFGReca.AddItem ""
        PonerBotones
    End If
End Sub

Private Sub VSFGReca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 0 del recargo
    If Col = 2 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    With VSFGReca
        ' only interesetd in left button
        If Button <> 1 Then Exit Sub
        
        ' get cell that was clicked
        Dim r&, c&
        r = .MouseRow
        c = .MouseCol
        
        ' make sure the click was on the sheet
        If r < 0 Or c < 0 Then Exit Sub
        
        If (c <> 0 Or r = (.Rows - 1)) Then Exit Sub
         
        ' make sure the click was on a cell with a button
        If .Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        
        ' click was on a button: do the work
         .Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
            
        'Recorro el FlexGrid para poner números a las filas
            
        If respuesta = vbYes Then
             Dim i As Integer
              .RemoveItem (r)
             PonerBotones
             CalcuReca
        Else
             .Cell(flexcpPicture, r, c) = imgBtnUp
        End If
        
        ' cancel default processing
        ' note: this is not strictly necessary in this case, because
        '       the dialog box already stole the focus etc, but let's be safe.
        Cancel = True
    End With
End Sub

Private Sub VSFGReca_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Busca y coloca el valor del recargo seleccionado
    If Row > 0 And VSFGReca.TextMatrix(Row, 1) <> "" And Col <> 3 Then
        clsRecargos.Filtrar "oca_codigo='" & VSFGReca.TextMatrix(Row, 1) & "'"
        VSFGReca.TextMatrix(Row, 2) = clsRecargos.adorec_Def("oca_nombre")
        VSFGReca.TextMatrix(Row, 3) = clsRecargos.adorec_Def("oca_precio")
        clsRecargos.QuitarFiltro
        'Verifica que no se haya escogido antes el mismo recargo, en ese caso suma sus valores
        For i = 1 To VSFGReca.Rows - 1
            If VSFGReca.TextMatrix(Row, 1) = VSFGReca.TextMatrix(i, 1) And Row <> i Then
                VSFGReca.TextMatrix(i, 3) = Val(VSFGReca.TextMatrix(i, 3)) + (VSFGReca.TextMatrix(Row, 3))
                VSFGReca.RemoveItem Row
                PonerBotones
                Exit For
            End If
        Next i
    End If
        CalcuReca
    
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'Verifica que solo se ingresen números en el grid de recargos en caso de ser necesario
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CargarGuias()
    Dim strSQL As String
    
    strSQL = " SELECT '0',ingreso.ing_codigo,ing_factura,det_ingreso.dep_codigo,det_ingreso.prd_codigo,producto.prd_nombre,det_ing_cantidad,det_ing_cantidad - COALESCE(sum(det_egr_cantidad),0),det_ing_precio,'0' as devo,'0' as fact,'0.00' as tota,'0' as linea,ing_observacion,(producto.prd_costo/(1 - IIF(det_ingreso.prd_codigo='BEL1583A',0.0715,0.15))) as prd_costo,det_ing_costo " & _
             " FROM (((ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.ing_anulado=0) " & _
             " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo) " & _
             " LEFT JOIN egreso ON ingreso.ing_codigo=egreso.egr_factura AND ingreso.emp_codigo=egreso.emp_codigo AND egreso.tip_egr_codigo='EGR' AND egreso.egr_anulado=0) " & _
             " LEFT JOIN  det_egreso ON egreso.egr_codigo=det_egreso.egr_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.emp_codigo=det_egreso.emp_codigo AND det_ingreso.prd_codigo=det_egreso.prd_codigo AND det_ingreso.dep_codigo=det_egreso.dep_codigo " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' AND ingreso.per_codigo='" & cmbCliente.BoundText & "' " & _
             " AND ingreso.tip_ing_codigo='IGR' " & _
             " GROUP BY ingreso.ing_codigo,ing_factura,det_ingreso.dep_codigo,det_ingreso.prd_codigo,producto.prd_nombre,det_ing_cantidad,det_ing_precio,ing_observacion,producto.prd_costo,det_ing_costo " & _
             " HAVING det_ing_cantidad - COALESCE(sum(det_egr_cantidad),0)>0 "
    clsSql.Ejecutar strSQL
    'VSFGGuia.AllowUserResizing = flexResizeNone
    VSFGGuia.Tag = "A"
    Set VSFGGuia.DataSource = clsSql.adorec_Def.DataSource
    VSFGGuia.Tag = ""
    'VSFGGuia.AllowUserResizing = flexResizeColumns
End Sub
Private Function Devolver() As Boolean
    Dim i As Integer
    Dim guia_actual As Double
    Dim ultima_guia As Double
    Dim numero_ingreso As Double
    Dim operacion As Boolean
    Dim MAguia As String
    Dim ex As Boolean
    Dim clsEgreso As New clsInventario
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    ex = False
    operacion = False
    If VSFGGuia.Tag <> "A" And MsgBox("Desea hacer la devolución de todas las guias seleccionadas?", vbYesNo + vbQuestion, "Devolución de Guias") = vbYes Then
        For i = 1 To VSFGGuia.Rows - 1
            'Verifica que la casilla esté seleccionada
            If Abs(VSFGGuia.TextMatrix(i, 0)) = 1 Then
                'Verifica si hay un valor en el campo Devolución
                If (Val(VSFGGuia.TextMatrix(i, 9)) > 0 And VSFGGuia.TextMatrix(i, 9) <> "") Then
                    guia_actual = FormatoD0(VSFGGuia.TextMatrix(i, 1))
                    operacion = True
                    If (guia_actual = ultima_guia) Then
                        'Añadir DET_INGRESO en último INGRESO
                        'Inserta el detalle de ingreso al proyecto
                        clsEgreso.NuevoDetEgr VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD0(VSFGGuia.TextMatrix(i, 9)), FormatoD4(VSFGGuia.TextMatrix(i, 8)), FormatoD4(VSFGGuia.TextMatrix(i, 15)), 0
                    Else
                        If ex = True Then
                            drptDevGuia.Tag = numero_ingreso
                            drptDevGuia.PrintReport True
                            drptDevGuia.Hide
                            Unload drptDevGuia
                        End If
                        ex = True
                        'Crear nuevo INGRESO con su DET_INGRESO
                        clsEgreso.NuevoEgr True, "EGR", True, strSucursal, strPtoFactura, , , cmbCliente.BoundText, Format(dtpFecha.Value, "yyyy-mm-dd"), (guia_actual), , "DECOLUCION GUIA: " & guia_actual
                        clsEgreso.NuevoDetEgr VSFGGuia.TextMatrix(i, 4), VSFGGuia.TextMatrix(i, 3), FormatoD0(VSFGGuia.TextMatrix(i, 9)), FormatoD4(VSFGGuia.TextMatrix(i, 8)), FormatoD4(VSFGGuia.TextMatrix(i, 15)), 0
                        strSQL = " UPDATE ingreso " & _
                                 " SET ing_observacion = CONCAT('" & UCase(MAguia) & "',' - ',ing_observacion),ing_fechamod=CURRENT_TIMESTAMP, ing_usumod='" & strUsuario & "' " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IGR' AND ing_codigo='" & guia_actual & "' "
                        clsSql.Ejecutar (strSQL), "M"
                    End If
                    ultima_guia = guia_actual
                End If
            End If
        Next i
        If operacion = True Then
            drptDevGuia.Tag = clsEgreso.strDoc
            drptDevGuia.PrintReport True
            drptDevGuia.Hide
            Unload drptDevGuia
            MsgBox "Devolución realizada con éxito", vbInformation
        End If
    End If
    Devolver = operacion
    Set clsEgreso = Nothing
End Function
