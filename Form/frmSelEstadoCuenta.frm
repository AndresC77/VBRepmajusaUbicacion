VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelEstadoCuenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roles de Pago"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9975
   Icon            =   "frmSelEstadoCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdRecibir 
      Caption         =   "&Contabilizar Rol"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4335
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6015
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir todo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2655
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Parámetros de Búsqueda"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   9735
      Begin VB.ComboBox cmbMesI 
         Height          =   315
         ItemData        =   "frmSelEstadoCuenta.frx":030A
         Left            =   240
         List            =   "frmSelEstadoCuenta.frx":0335
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1425
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcmbSocios 
         Height          =   315
         Left            =   2520
         TabIndex        =   0
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
      Begin MSComCtl2.DTPicker Año 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
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
         CurrentDate     =   38419
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes y Año"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   380
         Width           =   2175
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleados"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   375
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7455
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Roles de Pago"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   9735
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3615
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   9255
         _cx             =   16325
         _cy             =   6376
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
         Rows            =   1
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelEstadoCuenta.frx":039E
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
   End
End
Attribute VB_Name = "frmSelEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSql As String
Private clsSql As New clsConsulta
Private clsSql1 As New clsConsulta
Private HacerChange As Boolean
Private FactorCapital As String
Private FactorInteres As String
Private strSqlImpresion As String
Dim i As Integer
Dim j As Integer
Private HacerFecha As Boolean
Private Fecha1 As Variant
Private Fecha2 As Variant
Public Liquidacion As Boolean
Private CuentaNomina As String

Private Sub RestaurarBotones()
    Me.cmdExportar.Enabled = False
    Me.cmdImprimir.Enabled = False
    Me.cmdRecibir.Enabled = False
    Me.cmdVistaPrevia.Enabled = False
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub Año_Change()
    CambiarFecha
End Sub

Private Sub CambiarFecha()
    If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
           
    Fecha1 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-1"
    Fecha2 = ""
    DiaFinal = 31
    While (IsDate(Fecha2) = False)
        Fecha2 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
    RestaurarBotones
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub cmdBuscar_Click()
    BuscarEstadoCuenta
End Sub

Private Sub cmdExportar_Click()
    SeleccionarFlexGrid2 Me.VSFG
    CopiarFlexGrid2 Me.VSFG
    MsgBox "Se ha copiado el Estado de Cuenta de " & Me.dcmbSocios & " al portapapeles.", vbInformation, "Información"
End Sub

Private Sub cmdImprimir_Click()
    Dim NumeroSocios As Long
    Dim fila As Long
    Dim Mensaje As String
    Dim rptRol As New frmReporte
    
    'Contar número de estados de cuenta a imprimir
    'Sin tomar en cuenta el gran total
    For i = 0 To VSFG.Rows - 2
        If VSFG.TextMatrix(i, 0) <> VSFG.TextMatrix(i + 1, 0) Then
            If i = 0 Or (VSFG.IsSubtotal(i) = True And VSFG.IsSubtotal(i + 1) = False) Then
                NumeroSocios = NumeroSocios + 1
            End If
        End If
    Next i
    If NumeroSocios > 1 Then
        Mensaje = "los roles de pago de " & NumeroSocios & " empleados"
    ElseIf NumeroSocios = 1 Then
        Mensaje = "el rol de pago de 1 empleado"
    Else
        Exit Sub
    End If
    If MsgBox("¿Está seguro de imprimir " & Mensaje & "?", vbQuestion + vbYesNo, "Información") = vbNo Then Exit Sub
    fila = 1
    NumeroSocios = 0
    For i = 0 To VSFG.Rows - 2
        If VSFG.TextMatrix(i, 0) <> VSFG.TextMatrix(i + 1, 0) Then
            If i = 0 Or (VSFG.IsSubtotal(i) = True And VSFG.IsSubtotal(i + 1) = False) Then
                NumeroSocios = NumeroSocios + 1
                fila = i + 1

'                a.Mes = UCase(Me.Frame2.Caption)
'                Set a.VSFG = Me.VSFG
'                'drptRol.DataReport_Activate
'                a.Show
'                If NumeroSocios = 4 Then Exit Sub
            
            
'                drptRol.FilaInicial = fila
'                drptRol.Mes = UCase(Me.Frame2.Caption)
'                Set drptRol.VSFG = Me.VSFG
'                drptRol.DataReport_Activate
'                drptRol.Show

                rptRol.strReporte = "rptRolPagos"
                rptRol.strAsiento = Fecha1 & "," & Fecha2 & "," & cmbMesI.List(cmbMesI.ListIndex) & " " & CStr(Year(Año))
                rptRol.strNumero = VSFG.TextMatrix(fila, 1)
                rptRol.strTipo = CuentaNomina
                rptRol.Show
                
                If NumeroSocios = 1 Then
'                    drptRol.PrintReport True ', rptRangeFromTo, 1, 1
'                    rptRol.PrintForm
                    rptRol.VSPrint.PrintDoc
                    'rptRol.VSPrint.PrintDoc
                Else
'                    drptRol.PrintReport False ', rptRangeFromTo, 1, 1
                    rptRol.VSPrint.PrintDoc
                End If
                'MsgBox "Ver"
                'Unload rptRol
            End If
        End If
    Next i
End Sub

Private Sub cmdRecibir_Click()
    Screen.MousePointer = vbHourglass
    frmContabilizarRol.Fecha = Fecha2
    If Me.dcmbSocios.BoundText = "%" Then
        frmContabilizarRol.txtDescripcion = "REGISTRO CONTABLE DEL " & UCase(Me.Frame2.Caption) & " - TODOS LOS EMPLEADOS"
    Else
        frmContabilizarRol.txtDescripcion = "REGISTRO CONTABLE DEL " & UCase(Me.Frame2.Caption) & " - EMPLEADO: " & dcmbSocios
        If Liquidacion = True Then
            frmContabilizarRol.txtDescripcion = frmContabilizarRol.txtDescripcion & " - LIQUIDACIÓN DEL EMPLEADO"
        End If
    End If
    
    'BUSCA SI ESTA PARA LIQUIDAR
    strSql = " SELECT epl_fec_salida FROM empleado " & _
            " WHERE emp_codigo='" & strEmpresa & "' " & _
            " AND epl_codigo='" & Me.dcmbSocios.BoundText & "' "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        If Trim(clsSql.adorec_Def(0)) <> "" Then
            Liquidacion = True
        Else
            Liquidacion = False
        End If
    End If
    frmContabilizarRol.CodigoEmpleado = Me.dcmbSocios.BoundText
    frmContabilizarRol.Liquidacion = Liquidacion
    frmContabilizarRol.CuentaNomina = CuentaNomina
    frmContabilizarRol.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Public Sub BuscarEstadoCuenta()
    Screen.MousePointer = vbHourglass
    
    Me.Frame2.Caption = "Rol de Pagos " & StrConv(Me.cmbMesI.List(Me.cmbMesI.ListIndex), vbProperCase) & " " & Año.Year
     strSql = " DROP TABLE IF EXISTS EstadoCuentaVB "
    clsSql.Ejecutar (strSql)
    'Crear tabla temporal
    strSql = " CREATE TABLE EstadoCuentaVB " & _
            " (persona varchar(255),per_codigo varchar(10), tipo varchar(50), producto varchar(100), " & _
            " ingresos numeric(17,2), egresos numeric(17,2), TotalRecibir numeric(17,2), otros numeric(17,2), cuenta1 varchar(24), cuenta2 varchar(24), des_valor numeric(17,2), des_codigo numeric(18,0), sel1 int, sel2 int, orden int, orden2 int)"
    clsSql.Ejecutar (strSql)
    
    'Ingresos
    strSql = " INSERT INTO EstadoCuentaVB " & _
             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'INGRESOS' AS tipo, tip_des_nombre AS producto, " & _
                " des_valor AS ingresos, 0 AS egresos, des_valor AS TotalRecibir, 0 AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 0 AS orden, tip_des_orden AS orden2 " & _
                " FROM descuento " & _
                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND tipo_descuento.cta_codigo='" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=1 AND des_pagado=0"
    clsSql.Ejecutar (strSql)
    
    'Egresos
    strSql = " INSERT INTO EstadoCuentaVB " & _
             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'EGRESOS' AS tipo, tip_des_nombre AS producto, " & _
                " 0 AS ingresos, des_valor AS egresos, des_valor*-1 AS TotalRecibir, 0 AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 1 AS orden, tip_des_orden AS orden2 " & _
                " FROM descuento " & _
                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND det_tip_descuento.cta_codigo='" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=0 AND des_pagado=0"
    clsSql.Ejecutar (strSql)
    
    'Otros Ingresos
    strSql = " INSERT INTO EstadoCuentaVB " & _
             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'OTROS' AS tipo, tip_des_nombre AS producto, " & _
                " 0 AS ingresos, 0 AS egresos, 0 AS TotalRecibir, des_valor AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 2 AS orden, tip_des_orden AS orden2 " & _
                " FROM descuento " & _
                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND tipo_descuento.cta_codigo<>'" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=1 AND des_pagado=0"
    clsSql.Ejecutar (strSql)
    
    'Otros Egresos
    strSql = " INSERT INTO EstadoCuentaVB " & _
             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'OTROS' AS tipo, tip_des_nombre AS producto, " & _
                " 0 AS ingresos, 0 AS egresos, 0 AS TotalRecibir, des_valor AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 2 AS orden, tip_des_orden AS orden2 " & _
                " FROM descuento " & _
                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND det_tip_descuento.cta_codigo<>'" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=0 AND des_pagado=0"
    clsSql.Ejecutar (strSql)
    
    strSql = " SELECT * from EstadoCuentaVB ORDER BY persona, orden, orden2"
    clsSql.Ejecutar (strSql)
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    If clsSql.adorec_Def.RecordCount = 0 Then
        Me.cmdExportar.Enabled = False
        Me.cmdVistaPrevia.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.cmdRecibir.Enabled = False
    Else
        If dcmbSocios.BoundText <> "%" Then
            Me.cmdExportar.Enabled = True
        Else
            Me.cmdExportar.Enabled = False
        End If
        Me.cmdVistaPrevia.Enabled = True
        Me.cmdImprimir.Enabled = True
        If Me.dcmbSocios.BoundText <> "%" Then
            Me.cmdRecibir.Enabled = True
        Else
            Me.cmdRecibir.Enabled = False
        End If
    End If
    
    strSql = " DROP TABLE EstadoCuentaVB "
    clsSql.Ejecutar (strSql)
    
    VSFG.MergeCells = flexMergeRestrictRows
    VSFG.MergeCol(0) = True
    VSFG.MergeCol(1) = True
    VSFG.MergeCol(2) = True
    
    VSFG.SubtotalPosition = flexSTBelow
    'Total por tipo (ingreso, egreso, otros)
    VSFG.Subtotal flexSTSum, 2, 4, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 2, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 2, 6, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 2, 7, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    
    'Total por empleado
    VSFG.Subtotal flexSTSum, 0, 4, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 0, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 0, 6, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    VSFG.Subtotal flexSTSum, 0, 7, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0)
    
    If Me.dcmbSocios.BoundText = "%" Then
        'Total de los totales
        VSFG.Subtotal flexSTSum, -1, 4, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
        VSFG.Subtotal flexSTSum, -1, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
        VSFG.Subtotal flexSTSum, -1, 6, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
        VSFG.Subtotal flexSTSum, -1, 7, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
    
    End If
    Me.cmdBuscar.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVistaPrevia_Click()
    Dim rptRol As New frmReporte
    rptRol.strReporte = "rptRolPagos"
    rptRol.strAsiento = Fecha1 & "," & Fecha2 & "," & cmbMesI.List(cmbMesI.ListIndex) & " " & CStr(Year(Año))
    rptRol.strNumero = VSFG.TextMatrix(VSFG.Row, 1)
    rptRol.strTipo = CuentaNomina
    rptRol.Show
'    drptRol.FilaInicial = 1
'    drptRol.Mes = UCase(Me.Frame2.Caption)
'    Set drptRol.VSFG = Me.VSFG
'    drptRol.Show
End Sub

Private Sub dcmbSocios_Change()
    RestaurarBotones
End Sub

Private Sub Form_Activate()
    Me.cmdBuscar.Enabled = True
    If Liquidacion = True Then
        Me.Caption = "Liquidación de empleado"
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)

    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster
    Liquidacion = False
    HacerFecha = False
    Año = Date
    HacerFecha = True
    'Selecciona el mes actual
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(Date)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    'Buscar cuenta de nómina
    strSql = " SELECT par_con_cta_codigo FROM parametro_contable" & _
            " WHERE emp_codigo='" & strEmpresa & "' AND par_con_tipo='RRHH' AND par_con_codigo='1' "
    clsSql.Ejecutar (strSql)
    
    CuentaNomina = clsSql.adorec_Def(0)
    
    strSql = " SELECT '%' AS epl_codigo, ' ---Todos los Empleados---                    ' AS nombre UNION" & _
             " SELECT  epl_codigo, concat(epl_apellidos,' ',epl_nombres) AS nombre " & _
             " FROM empleado WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY nombre"
    clsSql.Ejecutar strSql
    Set Me.dcmbSocios.RowSource = clsSql.adorec_Def.DataSource
    dcmbSocios.ListField = "nombre"
    dcmbSocios.BoundColumn = "epl_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbSocios.BoundText = clsSql.adorec_Def("epl_codigo")
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtRol1_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtRol1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Temp As String
    If Len(txtRol1) = 5 Then
        Temp = Me.dcmbSocios.BoundText
        Me.dcmbSocios.BoundText = txtRol1
        If Me.dcmbSocios.MatchedWithList = False Then
            Me.dcmbSocios.BoundText = Temp
            txtRol1 = ""
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmContabilizarRol
End Sub



Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 18 Then
        If VSFG.TextMatrix(Row, 18) = "0" Then
            VSFG.TextMatrix(Row, 19) = 0
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 13 Then
        Cancel = True
    Else
'        'Que no permita poner el check de Liquidar si el check de sel está apagado
'        If VSFG.TextMatrix(Row, 18) = "0" And Col = 19 Then
'            Cancel = True
'        End If
'        'Que no permita liquidar si ya no queda saldo pendiente
'        If FormatoD(VSFG.TextMatrix(Row, 11)) = 0 And Col = 19 Then
'            Cancel = True
'        End If
    End If
    If VSFG.IsSubtotal(Row) = True Then
        Cancel = True
    End If
    
    If dcmbSocios.BoundText = "%" Then
        Cancel = True
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        If Col = 6 Then
            VSFG.Cell(flexcpBackColor, Row, Col) = RGB(230, 230, 230)
        End If
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 3 Then
'        CopiarFlexGrid VSFG
'        MsgBox "Se copió la selección al portapapeles.", vbInformation, "Información"
'    End If
End Sub

