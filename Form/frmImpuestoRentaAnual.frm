VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImpuestoRentaAnual 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impuesto Renta Anual"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   Icon            =   "frmImpuestoRentaAnual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10260
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar diferencias"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2783
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   375
      Left            =   4583
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impuesto Renta Anual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F1905&
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9975
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3975
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   9495
         _cx             =   16748
         _cy             =   7011
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
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImpuestoRentaAnual.frx":030A
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
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Año 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         CustomFormat    =   "yyyyXX"
         Format          =   106168323
         UpDown          =   -1  'True
         CurrentDate     =   38054
      End
      Begin MSDataListLib.DataCombo dcmbEmpleado 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   380
         Width           =   855
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleados"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   380
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6023
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "frmImpuestoRentaAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSql As String
Private clsSql As New clsConsulta
Private clsSql1 As New clsConsulta
Private clsSocio As New clsConsulta
Private adorec_Socio As ADODB.Recordset
Private adorec_SocioCodigo As ADODB.Recordset
Private SociosCargados As Boolean
Private HacerChange As Boolean
Private Hacer As Boolean
Private Factor As String
Private FactorInteres As String
Private CuentaContable As String
Dim i As Integer
Dim j As Integer
Public Ingreso As Boolean
Private PrimeraVez As Boolean
Private CodImp() As String
Private CodImpuesto As String
Private Fecha1 As Variant
Private Fecha2 As Variant
Private Condicion As String

Private Sub MostrarQuincena()
    'Me.lblQuincena.Caption = QuincenaText(Me.Año.Year & Me.cmbQuincena.List(Me.cmbQuincena.ListIndex))
    Me.cmdBuscar.Enabled = True
    'Me.cmdEditar.Visible = False
    'Me.Frame2.Visible = False
End Sub

Private Sub Año_Change()
    CambiarFecha
End Sub

Private Sub NumerarVSFG()
    For i = 2 To VSFG.Rows - 1
        If VSFG.IsSubtotal(i) = False Then VSFG.TextMatrix(i, 0) = i - 1
    Next i
End Sub
Private Sub SumarVSFG()
    Me.VSFG.SubtotalPosition = flexSTBelow
    VSFG.Subtotal flexSTSum, -1, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
    VSFG.Subtotal flexSTSum, -1, 6, "#,##0.00"
    VSFG.Subtotal flexSTSum, -1, 7, "#,##0.00"
    VSFG.Subtotal flexSTSum, -1, 8, "#,##0.00"
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
           
    Fecha1 = Me.Año.Year & "-01-01"
    Fecha2 = Me.Año.Year & "-12-31"
    'MostrarAsientos
    Me.cmdGenerar.Enabled = False
    Me.cmdBuscar.Enabled = True
    VSFG.Rows = 2
    UnirCeldas
End Sub

Private Sub cmdBuscar_Click()
    BuscarDescuentos
End Sub

Private Sub cmdExportar_Click()
    SeleccionarFlexGrid2 Me.VSFG
    CopiarFlexGrid2 Me.VSFG
    MsgBox "Se ha copiado la tabla de al portapapeles.", vbInformation, "Información"
End Sub

Private Sub cmdGenerar_Click()
    If MsgBox("¿Está seguro de generar las diferencias por Impuesto a la Renta? " & vbNewLine & "Se grabarán en el mes de DICIEMBRE del año " & Me.Año.Year & ".", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Si no están contabilizados borrar
    strSql = " DELETE FROM descuento WHERE emp_codigo='" & strEmpresa & "'" & _
             " des_pagado=0 AND des_fecha='" & Fecha2 & "' AND ("
             
    For i = 0 To UBound(CodImp)
        strSql = strSql & " tip_des_codigo='" & CodImp(i) & "' "
        If i <> UBound(CodImp) Then
            strSql = strSql & " OR "
        End If
    Next i
    strSql = strSql & ") "
    clsSql.Ejecutar strSql, "M"
    
    Dim valor As Double
    'Generar de nuevo
    For i = 2 To VSFG.Rows - 2
        valor = VSFG.TextMatrix(i, 15)
        If valor > 0 Then
            'codigo quemado
            GrabarDescuento clsSql, CodImp(0), VSFG.TextMatrix(i, 1), CStr(Fecha2), valor
        ElseIf valor < 0 Then
            valor = valor * -1
            'codigo quemado
            GrabarDescuento clsSql, CodImp(1), VSFG.TextMatrix(i, 1), CStr(Fecha2), valor
        End If
    Next i
    Screen.MousePointer = vbDefault
    cmdGenerar.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub UnirCeldas()
    VSFG.MergeCells = flexMergeFixedOnly
    VSFG.MergeCol(0) = True
    VSFG.MergeCol(1) = True
    VSFG.MergeCol(2) = True
    VSFG.MergeCol(3) = True
    VSFG.MergeCol(6) = True
    VSFG.MergeCol(12) = True
End Sub

Private Sub PonerTotales()
    Dim ValorTemporal As Double
    For i = 2 To VSFG.Rows - 1
        If VSFG.IsSubtotal(i) = False Then
            If Not IsNumeric(VSFG.TextMatrix(i, 4)) Then VSFG.TextMatrix(i, 4) = "0"
            If Not IsNumeric(VSFG.TextMatrix(i, 6)) Then VSFG.TextMatrix(i, 6) = "0"
            If Not IsNumeric(VSFG.TextMatrix(i, 9)) Then VSFG.TextMatrix(i, 9) = "0"
            If Not IsNumeric(VSFG.TextMatrix(i, 5)) Then VSFG.TextMatrix(i, 5) = "0"
            If Not IsNumeric(VSFG.TextMatrix(i, 7)) Then VSFG.TextMatrix(i, 7) = "0"
            If Not IsNumeric(VSFG.TextMatrix(i, 10)) Then VSFG.TextMatrix(i, 10) = "0"
            
            VSFG.TextMatrix(i, 11) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(i, 6)) + Val(VSFG.TextMatrix(i, 9))
            VSFG.TextMatrix(i, 12) = Val(VSFG.TextMatrix(i, 5)) + Val(VSFG.TextMatrix(i, 7)) + Val(VSFG.TextMatrix(i, 10))
            VSFG.TextMatrix(i, 13) = ImpuestoRentaAño(Val(VSFG.TextMatrix(i, 11)))
            VSFG.TextMatrix(i, 14) = Val(VSFG.TextMatrix(i, 12)) - Val(VSFG.TextMatrix(i, 13))
            VSFG.TextMatrix(i, 15) = Val(VSFG.TextMatrix(i, 14))
            If VSFG.TextMatrix(i, 14) > 0 Then
                'Si hay que pagarle al empleado máximo devolverle lo que aportó en la empresa
                ValorTemporal = Val(VSFG.TextMatrix(i, 7)) + Val(VSFG.TextMatrix(i, 10))
                If Val(VSFG.TextMatrix(i, 14)) > ValorTemporal Then
                    VSFG.TextMatrix(i, 15) = ValorTemporal
                End If
            End If
        End If
    Next i
    
    VSFG.SubtotalPosition = flexSTBelow
    VSFG.Subtotal flexSTSum, -1, 4, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
    VSFG.Subtotal flexSTSum, -1, 5
    VSFG.Subtotal flexSTSum, -1, 7
    VSFG.Subtotal flexSTSum, -1, 8
    VSFG.Subtotal flexSTSum, -1, 9
    VSFG.Subtotal flexSTSum, -1, 10
    VSFG.Subtotal flexSTSum, -1, 11
    VSFG.Subtotal flexSTSum, -1, 12
End Sub

Private Sub BuscarDescuentos()
    Dim HayRegistros As Boolean
    Screen.MousePointer = vbHourglass
    strSql = " SELECT descuento.epl_codigo, concat(epl_apellidos,' ',epl_nombres) AS nombre, 0, IFNULL(imp_ren_renta1,0), IFNULL(imp_ren_impuesto1,0), IFNULL(imp_ren_renta,0), IFNULL(imp_ren_impuesto,0),  0, 0, SUM(des_valor),0,0,0,0,0 " & _
             " FROM descuento INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
             " LEFT JOIN impuesto_renta ON descuento.epl_codigo=impuesto_renta.epl_codigo AND descuento.emp_codigo=impuesto_renta.emp_codigo AND impuesto_renta.imp_ren_año='" & Me.Año.Year & "'" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_codigo='" & CodImpuesto & "' AND descuento.epl_codigo LIKE '" & Me.dcmbEmpleado.BoundText & "' " & _
             " GROUP BY descuento.epl_codigo, epl_apellidos, epl_nombres, imp_ren_renta1, imp_ren_impuesto1, imp_ren_renta, imp_ren_impuesto ORDER BY epl_apellidos, epl_nombres"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then HayRegistros = True
    HacerChange = False
    Set Me.VSFG.DataSource = clsSql.adorec_Def.DataSource
    UnirCeldas
    NumerarVSFG
    Dim i As Long
    For i = 2 To VSFG.Rows - 1
        PeriodoAño VSFG.TextMatrix(i, 1), i
        VSFG.TextMatrix(i, 9) = RentaAño(VSFG.TextMatrix(i, 1), Fecha1, Fecha2)
    Next i
    PonerTotales
    Me.cmdBuscar.Enabled = False
    'Verificar si ya existen los registros de diferencia en el año
    strSql = " SELECT des_codigo FROM descuento WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND des_pagado=1 AND des_fecha='" & Fecha2 & "' AND ("
             
    For i = 0 To UBound(CodImp)
        strSql = strSql & " tip_des_codigo='" & CodImp(i) & "' "
        If i <> UBound(CodImp) Then
            strSql = strSql & " OR "
        End If
    Next i
    strSql = strSql & ") "
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmdGenerar.Enabled = False
    Else
        If HayRegistros = True Then
            cmdGenerar.Enabled = True
        Else
            cmdGenerar.Enabled = False
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub PeriodoAño(Empleado As String, Row As Long)
    strSql = " SELECT MIN(des_fecha) AS min, MAX(des_fecha) AS max " & _
             " FROM descuento" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_codigo='" & CodImpuesto & "' AND descuento.epl_codigo='" & Empleado & "'" & _
             " GROUP BY emp_codigo"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
    
        VSFG.TextMatrix(Row, 8) = Format(clsSql.adorec_Def("min"), "mmm") & " - " & Format(clsSql.adorec_Def("max"), "mmm")
        If Format(clsSql.adorec_Def("min"), "mm") <> "01" Then
            VSFG.TextMatrix(Row, 3) = Format(Fecha1, "mmm") & " - " & Format(DateAdd("m", -1, clsSql.adorec_Def("min")), "mmm")
        Else
            VSFG.TextMatrix(Row, 3) = ""
        End If
    Else
        VSFG.TextMatrix(Row, 8) = ""
        VSFG.TextMatrix(Row, 3) = ""
    End If
End Sub

Private Sub dcmbEmpleado_Change()
    Me.cmdGenerar.Enabled = False
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)

    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster
    clsSocio.Inicializar AdoConn, AdoConnMaster
    Año = Date
    CambiarFecha
    Dim cont As Integer
    cont = 0
    strSql = " SELECT tip_des_codigo " & _
            " FROM tipo_descuento " & _
            " WHERE upper(tip_des_nombre) LIKE 'IMPUESTO%RENTA' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        CodImpuesto = clsSql.adorec_Def(0)
    End If
    
    strSql = " SELECT tip_des_codigo " & _
            " FROM tipo_descuento " & _
            " WHERE upper(tip_des_nombre) LIKE '%DIFERENCIA IMPUESTO RENTA%' " & _
            " ORDER BY tip_des_ingreso DESC,tip_des_codigo "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        ReDim CodImp(Abs(clsSql.adorec_Def.RecordCount - 1))
        
        While clsSql.adorec_Def.EOF = False
            CodImp(cont) = clsSql.adorec_Def(0)
            cont = cont + 1
            clsSql.adorec_Def.MoveNext
        Wend
    End If
    strSql = " SELECT '%' AS epl_codigo, ' ---Todos los empleados---                          ' AS nombre UNION SELECT epl_codigo, CONCAT(epl_apellidos,' ',epl_nombres) AS nombre " & _
             " FROM empleado WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY nombre"
    clsSql.Ejecutar strSql
    Set Me.dcmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    dcmbEmpleado.ListField = "nombre"
    dcmbEmpleado.BoundColumn = "epl_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbEmpleado.BoundText = clsSql.adorec_Def("epl_codigo")
    Else
        MsgBox "Sin empleados ingresados no se puede usar esta pantalla.", vbCritical, "Información"
        Unload Me
        Exit Sub
    End If
    
    PrimeraVez = True
    Hacer = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Public Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    PonerTotales
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFG.IsSubtotal(Row) = True Then
        Cancel = True
        Exit Sub
    End If
    If Col <> 7 And Col <> 6 And Col <> 5 And Col <> 4 Then
        Cancel = True
    Else
        If Trim(VSFG.TextMatrix(Row, 3)) = "" Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row < 2 Then Exit Sub
    If VSFG.IsSubtotal(Row) = True Then Exit Sub
    If Col = 3 Or Col = 8 Then
        VSFG.Cell(flexcpForeColor, Row, Col) = RGB(0, 120, 0)
        VSFG.Cell(flexcpBackColor, Row, Col) = RGB(245, 245, 245)
    End If
    If Col = 11 Or Col = 12 Then
        VSFG.Cell(flexcpForeColor, Row, Col) = RGB(120, 0, 0)
        VSFG.Cell(flexcpBackColor, Row, Col) = RGB(245, 245, 245)
    End If
    If Col = 13 Then
        VSFG.Cell(flexcpForeColor, Row, Col) = RGB(0, 0, 120)
    End If
    If Col = 3 Then
        If Trim(VSFG.TextMatrix(Row, Col)) = "" Then
            VSFG.Cell(flexcpForeColor, Row, 4, Row, 5) = RGB(180, 180, 180)
        End If
    End If
End Sub

Private Sub VSFG_EnterCell()
    If VSFG.Col = 15 Then
        VSFG.ToolTipText = "Un valor positivo en Diferencia a pagar indica un valor a favor del empleado"
    Else
        VSFG.ToolTipText = ""
    End If
End Sub

Private Sub VSFG_KeyUp(KeyCode As Integer, Shift As Integer)
    If vbCtrlMask > 0 Then
        If KeyCode = 17 Then
            CopiarFlexGrid2 VSFG
            MsgBox "Se copió la selección al portapapeles.", vbInformation, "Información"
        End If
    End If
End Sub

Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Screen.MousePointer = vbHourglass
    Dim strUpdate As String
    Dim Valor1 As Double
    Dim Valor2 As Double
    Dim Valor3 As Double
    Dim Valor4 As Double
    'VSFG.EditText = Format(VSFG.EditText, "###0.00")
    If Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Then
        If Col = 4 Then
            If IsNumeric(VSFG.EditText) Then
                strUpdate = " imp_ren_renta1='" & VSFG.EditText & "'"
                Valor1 = VSFG.EditText
                VSFG.TextMatrix(Row, Col) = VSFG.EditText
            Else
                strUpdate = " imp_ren_renta1='0'"
                Valor1 = 0
                VSFG.TextMatrix(Row, Col) = "0"
            End If
            Valor2 = 0
            Valor3 = 0
            Valor4 = 0
        ElseIf Col = 5 Then
            If IsNumeric(VSFG.EditText) Then
                strUpdate = " imp_ren_impuesto1='" & VSFG.EditText & "'"
                Valor2 = VSFG.EditText
                VSFG.TextMatrix(Row, Col) = VSFG.EditText
            Else
                strUpdate = " imp_ren_impuesto1='0'"
                Valor2 = 0
                VSFG.TextMatrix(Row, Col) = "0"
            End If
            Valor1 = 0
            Valor3 = 0
            Valor4 = 0
        ElseIf Col = 6 Then
            If IsNumeric(VSFG.EditText) Then
                strUpdate = " imp_ren_renta='" & VSFG.EditText & "'"
                Valor3 = VSFG.EditText
                VSFG.TextMatrix(Row, Col) = VSFG.EditText
            Else
                strUpdate = " imp_ren_renta='0'"
                Valor3 = 0
                VSFG.TextMatrix(Row, Col) = "0"
            End If
            Valor1 = 0
            Valor2 = 0
            Valor4 = 0
        ElseIf Col = 7 Then
            If IsNumeric(VSFG.EditText) Then
                strUpdate = " imp_ren_impuesto='" & VSFG.EditText & "'"
                Valor4 = VSFG.EditText
                VSFG.TextMatrix(Row, Col) = VSFG.EditText
            Else
                strUpdate = " imp_ren_impuesto='0'"
                Valor4 = 0
                VSFG.TextMatrix(Row, Col) = "0"
            End If
            Valor1 = 0
            Valor2 = 0
            Valor3 = 0
        End If
        strSql = " SELECT imp_ren_año FROM impuesto_renta " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'" & _
                 " AND imp_ren_año='" & Me.Año.Year & "'"
        clsSql.Ejecutar (strSql)
        If clsSql.adorec_Def.RecordCount > 0 Then
            
            strSql = " UPDATE impuesto_renta SET " & strUpdate & ", imp_ren_fechamod=CURRENT_TIMESTAMP, imp_ren_usumod='" & strUsuario & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & VSFG.TextMatrix(Row, 1) & "'" & _
                     " AND imp_ren_año='" & Me.Año.Year & "'"
            clsSql.Ejecutar strSql, "M"
        Else
            strSql = " INSERT INTO impuesto_renta (imp_ren_año, emp_codigo, epl_codigo, imp_ren_renta1, imp_ren_impuesto1, imp_ren_renta, " & _
                     " imp_ren_impuesto, imp_ren_fechamod, imp_ren_usumod) VALUES (" & _
                     "'" & Me.Año.Year & "','" & strEmpresa & "','" & VSFG.TextMatrix(Row, 1) & "','" & Valor1 & "'," & _
                     "'" & Valor2 & "', '" & Valor3 & "', '" & Valor4 & "', CURRENT_TIMESTAMP,'" & strUsuario & "')"
            clsSql.Ejecutar strSql, "M"
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
