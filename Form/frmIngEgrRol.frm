VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIngEgrRol 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos y Egresos Rol"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   Icon            =   "frmIngEgrRol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13275
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Pago Provisiones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   5055
      Begin VSFlex8Ctl.VSFlexGrid VSFGProvision 
         Height          =   1215
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4575
         _cx             =   8070
         _cy             =   2143
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIngEgrRol.frx":030A
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
   End
   Begin VB.Frame fraDetalle 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Rol de Pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   12975
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4320
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   12540
         _cx             =   22119
         _cy             =   7620
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
         Rows            =   6
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIngEgrRol.frx":0399
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
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   5292
      TabIndex        =   4
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   7212
      TabIndex        =   3
      Top             =   7080
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5625
      Begin VB.ComboBox cmbMesI 
         Height          =   315
         ItemData        =   "frmIngEgrRol.frx":05D7
         Left            =   240
         List            =   "frmIngEgrRol.frx":0602
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1425
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker Año 
         Height          =   315
         Left            =   1665
         TabIndex        =   6
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
         CurrentDate     =   38054
      End
      Begin VB.Label lblDias 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDCCBB&
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         ForeColor       =   &H002F1905&
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lbld 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDCCBB&
         BackStyle       =   0  'Transparent
         Caption         =   "Días:"
         ForeColor       =   &H002F1905&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   390
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes y Año"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   375
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmIngEgrRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSql As String
Private clsSql As New clsConsulta
Private strSql1 As String
Private clsSql1 As New clsConsulta
Private clsSqlAux As New clsConsulta
Private i As Long
Private j As Long
Private Fecha1 As Variant
Private Fecha2 As Variant
Private toalIngresos As Double
Private totalEgresos As Double
Private totalRecibir As Double
Private SMes As Double
Private SBas As Double
Private SIESS As Double
Private IRMes As Double
Private Factor As String
Private FactorInteres As String
Private strSqlImpresion As String
Private CuentaContable As String
Private PrimeraVez As Boolean
Private colInicial As Long
Private rowInicial As Long
Private Cambio As Boolean
Private Contabilizado As Boolean
Private Const datos As Long = 9
Private Const EMP As Long = 2
Private Const INICIO As Long = 19
Public EMP1 As Long
Public DATOS1 As Long
Public INICIO1 As Long
Private colIngreso As Long
Private colEgreso As Long
Private colOtros As Long
Dim empleados As Long





Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
           
    Fecha1 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-1"
    Fecha2 = ""
    DiaFinal = 31
    While (IsDate(Fecha2) = False)
        Fecha2 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        lblDias = DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
    Fecha1 = Format(Fecha1, "yyyy-mm-dd")
    Fecha2 = Format(Fecha2, "yyyy-mm-dd")
    'MostrarAsientos
    'Me.cmdBuscar.Enabled = True
    'Me.cmdEditar.Visible = False
    'Me.Frame2.Visible = False
    'PonerEtiquetas
End Sub


Private Sub CargaEmpleados()
    Dim Row As Long
    
    Dim CadenaEval As String
    
   
    
    Row = 1
    strSql = " SELECT epl_codigo, concat(epl_apellidos,' ',epl_nombres) As nombre, epl_sueldo, epl_fec_ingreso, epl_fec_salida, are_lab_nombre " & _
             " FROM empleado e" & _
             " INNER JOIN area_laboral a ON a.are_lab_codigo = e.are_lab_codigo " & _
             " AND a.emp_codigo = e.emp_codigo " & _
             " WHERE e.emp_codigo = '" & strEmpresa & "' AND epl_baja=0 " & _
             " ORDER BY are_lab_nombre,epl_apellidos,epl_nombres"
    'are_lab_codigo LIKE '" & Trim(Me.dcmbArea.BoundText) & "' " & Condicion & " AND
    clsSql.Ejecutar strSql
    'clsSqlAux.Ejecutar strSql
    empleados = 0
    Row = 0
    
    If clsSql.adorec_Def.RecordCount <> 0 Then
        empleados = FormatoD0(clsSql.adorec_Def.RecordCount)
    While clsSql.adorec_Def.EOF = False
        Row = Row + 1
        If Row > VSFG.Rows Then
            VSFG.AddItem "", Row
        End If
        VSFG.TextMatrix(Row, EMP) = clsSql.adorec_Def("epl_codigo")
        VSFG.TextMatrix(Row, EMP + 1) = clsSql.adorec_Def("are_lab_nombre")
        VSFG.TextMatrix(Row, EMP + 2) = clsSql.adorec_Def("nombre")
        VSFG.TextMatrix(Row, INICIO - 3) = Format(clsSql.adorec_Def("epl_sueldo"), "###0.00")
        VSFG.TextMatrix(Row, EMP + 3) = clsSql.adorec_Def("epl_fec_ingreso")
        If IsNull(clsSql.adorec_Def("epl_fec_salida")) = False Then
            VSFG.TextMatrix(Row, EMP + 4) = clsSql.adorec_Def("epl_fec_salida")
        Else
            VSFG.TextMatrix(Row, EMP + 4) = ""
        End If
        
        strSql = "SELECT des_codigo, descuento.epl_codigo, concat(epl_apellidos,' ',epl_nombres) AS nombre, des_valor, 0, 0, epl_sueldo, "
'    If Me.Check1(1).Value = 1 Then
'        strSql = strSql & "det_asiento.asi_numasiento+' '+LEFT(CONVERT(VARCHAR,asi_fecha,20),10)+' '+CAST(det_asi_debe AS VARCHAR)"
'    Else
        strSql = strSql & "''"
    'End If
    strSql = strSql & ", IFNULL(des_pagado,0) as des_pagado, IFNULL(des_valor1,0), IFNULL(des_valor2,0), epl_fec_ingreso,'','', epl_fec_salida " & _
             " FROM descuento INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo"
    ''Si tiene un asiento relacionado buscar fecha y valor
    'If Me.Check1(1).Value = 1 Then
     '   strSql = strSql & " LEFT JOIN asiento ON descuento.asi_numasiento=asiento.asi_numasiento AND descuento.emp_codigo=asiento.emp_codigo" & _
     '           " LEFT JOIN det_asiento ON det_asiento.asi_numasiento=asiento.asi_numasiento AND det_asiento.emp_codigo=asiento.emp_codigo AND det_asi_haber=0 AND det_asiento.cta_codigo = '" & CuentaContable & "'"
   ' End If
    strSql = strSql & " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND empleado.epl_codigo='" & VSFG.TextMatrix(Row, EMP) & "' " & _
             " ORDER BY epl_apellidos, epl_nombres"
    clsSqlAux.Ejecutar strSql
    
    If clsSqlAux.adorec_Def.RecordCount = 0 Then
        VSFG.TextMatrix(Row, INICIO - 2) = "0"
    Else
        VSFG.TextMatrix(Row, INICIO - 2) = clsSqlAux.adorec_Def("des_pagado")
    End If
        
        
        VSFG.TextMatrix(Row, EMP + 5) = DiasFinDeMes(CStr(Fecha2), VSFG.TextMatrix(Row, EMP + 3), VSFG.TextMatrix(Row, EMP + 4))
        VSFG.TextMatrix(Row, EMP + 6) = DiasFondo(CStr(Fecha2), VSFG.TextMatrix(Row, EMP + 3), VSFG.TextMatrix(Row, EMP + 4))

        'Ingresar Variables
        
        VSFG.TextMatrix(Row, datos + 1) = Format(SueldoMes(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2), "###0.00")
        VSFG.TextMatrix(Row, datos + 2) = Format(SueldoAño(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2), "###0.00")
        VSFG.TextMatrix(Row, datos + 3) = Format(RentaMes(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2), "###0.00")
        VSFG.TextMatrix(Row, datos + 4) = Format(SueldoIESS(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2), "###0.00")
        VSFG.TextMatrix(Row, datos) = Format(Val(VSFG.TextMatrix(Row, INICIO - 3)) * CInt(VSFG.TextMatrix(Row, EMP + 5)) / CInt(lblDias), "###0.00")
'                End If
'                'Si es fondo de cesantía calcular con la columna 15
'                If Me.dcmbTipo.BoundText = "1003" Then
'                    ElCapital = FormatoD(ElCapital * CInt(VSFG.TextMatrix(Row, 15)) / CInt(lblDias))
'                End If
 
        VSFG.TextMatrix(Row, 1) = "0"
       
        clsSql.adorec_Def.MoveNext
        
        If clsSql.adorec_Def.EOF = False Then
            'Row = Row + 1 'VSFG.Rows
            If clsSql.adorec_Def("epl_fec_salida") <> "" And Format(clsSql.adorec_Def("epl_fec_salida"), "yyyy-MM") < Format(Fecha2, "yyyy-MM") Then clsSql.adorec_Def.MoveNext
        End If
    Wend
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, EMP) <> "" Then
            'empleados = clsSql.adorec_Def.RecordCount
        Else
            VSFG.RowHidden(i) = True
        End If
    Next i
    PonerNum
    VSFG.AddItem "", VSFG.Rows
    
    
    SumarTotal
    End If
End Sub


Private Function CargaIngresosOpcion2() As Boolean
 
    Dim Row As Long
    Dim CUENTA As String
    
    strSql = " SELECT par_con_cta_codigo FROM parametro_contable " & _
              " WHERE emp_codigo='" & strEmpresa & "' AND par_con_codigo=1 AND par_con_tipo='RRHH'"
    clsSql.Ejecutar strSql
    CUENTA = clsSql.adorec_Def(0)
    
     strSql = " SELECT tip_des_codigo, tip_des_nombre, tip_des_factor " & _
            " ,tip_des_sueldo_mes,tip_des_impuesto_renta,tip_des_iess " & _
            " FROM tipo_descuento " & _
            " WHERE tipo_descuento.emp_codigo ='" & strEmpresa & "' " & _
            " AND tipo_descuento.cta_codigo='" & CUENTA & "' " & _
            " AND tipo_descuento.tip_des_ingreso=1 " & _
            " GROUP BY tipo_descuento.tip_des_nombre " & _
            " ORDER BY tip_des_orden "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount <> 0 Then
        While clsSql.adorec_Def.EOF = False
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_codigo")
            VSFG.TextMatrix(1, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_factor")
            VSFG.TextMatrix(3, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_sueldo_mes")
            VSFG.TextMatrix(4, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_impuesto_renta")
            VSFG.TextMatrix(5, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_iess")
            strSql1 = " SELECT IFNULL(B.tip_des_codigo,0) AS cod_provision, IFNULL(B.tip_des_nombre,'') AS provision " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, VSFG.Cols - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
            clsSql1.Ejecutar (strSql1)
            If clsSql1.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(2, VSFG.Cols - 1) = clsSql1.adorec_Def("cod_provision")
            End If
            VSFG.ColHidden(VSFG.Cols - 1) = True
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_nombre")
            VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
            VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
            VSFG.AutoSize (VSFG.Cols - 1)
            clsSql.adorec_Def.MoveNext
        Wend
    Else
        MsgBox "No ha parametrizado las cuentas contables de Tipos de Ingresos", vbInformation, "Ingresos"
        CargaIngresosOpcion2 = False
        Exit Function
    End If
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total Ingresos"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
    CargaIngresosOpcion2 = True
End Function

Private Function CargaEgresosOpcion2() As Boolean
 
    Dim Row As Long
    Dim CUENTA As String
    
    strSql = " SELECT par_con_cta_codigo FROM parametro_contable " & _
              " WHERE emp_codigo='" & strEmpresa & "' AND par_con_codigo=1 AND par_con_tipo='RRHH'"
    clsSql.Ejecutar strSql
    CUENTA = clsSql.adorec_Def(0)
    

    
    strSql = " SELECT tipo_descuento.tip_des_codigo,tip_des_nombre,tipo_descuento.tip_des_factor " & _
            " ,tip_des_sueldo_mes,tip_des_impuesto_renta,tip_des_iess " & _
            " FROM tipo_descuento " & _
            " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo " & _
            " AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo " & _
            " WHERE tipo_descuento.emp_codigo ='" & strEmpresa & "' " & _
            " AND det_tip_descuento.cta_codigo='" & CUENTA & "' " & _
            " AND tipo_descuento.tip_des_ingreso=0 " & _
            " GROUP BY tipo_descuento.tip_des_nombre " & _
            " ORDER BY tip_des_orden "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount <> 0 Then
        While clsSql.adorec_Def.EOF = False
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_codigo")
            VSFG.TextMatrix(1, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_factor")
            VSFG.TextMatrix(3, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_sueldo_mes")
            VSFG.TextMatrix(4, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_impuesto_renta")
            VSFG.TextMatrix(5, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_iess")
            strSql1 = " SELECT IFNULL(B.tip_des_codigo,0) AS cod_provision, IFNULL(B.tip_des_nombre,'') AS provision " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, VSFG.Cols - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
            clsSql1.Ejecutar (strSql1)
            If clsSql1.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(2, VSFG.Cols - 1) = clsSql1.adorec_Def("cod_provision")
            End If
            VSFG.ColHidden(VSFG.Cols - 1) = True
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_nombre")
            VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
            VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
            VSFG.AutoSize (VSFG.Cols - 1)
            clsSql.adorec_Def.MoveNext
        Wend
    Else
        MsgBox "No ha parametrizado las cuentas contables de Tipos de Egresos", vbInformation, "Egresos"
        CargaEgresosOpcion2 = False
        Exit Function
    End If
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total Egresos"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total a Recibir"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
    CargaEgresosOpcion2 = True
End Function

Private Function CargaOtrosOpcion2() As Boolean
 
    Dim Row As Long
    Dim CUENTA As String
    
    strSql = " SELECT par_con_cta_codigo FROM parametro_contable " & _
              " WHERE emp_codigo='" & strEmpresa & "' AND par_con_codigo=1 AND par_con_tipo='RRHH'"
    clsSql.Ejecutar strSql
    CUENTA = clsSql.adorec_Def(0)
    
    'primero provisiones y luego los otros importes
    
    
'''     'Otros Ingresos
'''    strSql = " INSERT INTO EstadoCuentaVB " & _
'''             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'OTROS' AS tipo, tip_des_nombre AS producto, " & _
'''                " 0 AS ingresos, 0 AS egresos, 0 AS TotalRecibir, des_valor AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 2 AS orden, tip_des_orden AS orden2 " & _
'''                " FROM descuento " & _
'''                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
'''                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
'''                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
'''                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
'''                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND tipo_descuento.cta_codigo<>'" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=1 AND des_pagado=0"
'''    clsSql.Ejecutar (strSql)
'''
'''    'Otros Egresos
'''    strSql = " INSERT INTO EstadoCuentaVB " & _
'''             " SELECT concat(epl_apellidos,' ',epl_nombres) AS persona, descuento.epl_codigo, 'OTROS' AS tipo, tip_des_nombre AS producto, " & _
'''                " 0 AS ingresos, 0 AS egresos, 0 AS TotalRecibir, des_valor AS otros, det_tip_descuento.cta_codigo AS cuenta1, tipo_descuento.cta_codigo AS cuenta2, des_valor, des_codigo, 1 AS sel1, 0 AS sel2, 2 AS orden, tip_des_orden AS orden2 " & _
'''                " FROM descuento " & _
'''                " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
'''                " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
'''                " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
'''                " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & dcmbSocios.BoundText & "'" & _
'''                " AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND det_tip_descuento.cta_codigo<>'" & CuentaNomina & "' AND tipo_descuento.tip_des_ingreso=0 AND des_pagado=0"
'''    clsSql.Ejecutar (strSql)
    
    
    
    strSql = " SELECT tip_des_codigo,tip_des_nombre,tip_des_factor " & _
            " ,tip_des_sueldo_mes,tip_des_impuesto_renta,tip_des_iess " & _
            " FROM tipo_descuento " & _
            " WHERE tipo_descuento.emp_codigo ='" & strEmpresa & "' " & _
            " AND tipo_descuento.cta_codigo<>'" & CUENTA & "' " & _
            " AND tipo_descuento.tip_des_ingreso=1 " & _
            " GROUP BY tipo_descuento.tip_des_nombre "
    strSql = strSql & " UNION "
    strSql = strSql & " SELECT tipo_descuento.tip_des_codigo,tip_des_nombre,tipo_descuento.tip_des_factor " & _
            " ,tip_des_sueldo_mes,tip_des_impuesto_renta,tip_des_iess " & _
            " FROM tipo_descuento " & _
            " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo " & _
            " AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo " & _
            " WHERE tipo_descuento.emp_codigo ='" & strEmpresa & "' " & _
            " AND det_tip_descuento.cta_codigo<>'" & CUENTA & "' " & _
            " AND tipo_descuento.tip_des_ingreso=0 " & _
            " GROUP BY tipo_descuento.tip_des_nombre "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount <> 0 Then
        While clsSql.adorec_Def.EOF = False
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_codigo")
            VSFG.TextMatrix(1, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_factor")
            VSFG.TextMatrix(3, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_sueldo_mes")
            VSFG.TextMatrix(4, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_impuesto_renta")
            VSFG.TextMatrix(5, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_iess")
            strSql1 = " SELECT IFNULL(B.tip_des_codigo,0) AS cod_provision, IFNULL(B.tip_des_nombre,'') AS provision " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, VSFG.Cols - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
            clsSql1.Ejecutar (strSql1)
            If clsSql1.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(2, VSFG.Cols - 1) = clsSql1.adorec_Def("cod_provision")
            End If
            VSFG.ColHidden(VSFG.Cols - 1) = True
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_nombre")
            VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
            VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
            VSFG.AutoSize (VSFG.Cols - 1)
            clsSql.adorec_Def.MoveNext
        Wend
    End If
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total Otros"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
    CargaOtrosOpcion2 = True
End Function


Private Sub CargaIngresos()
 
    Dim Row As Long
    strSql = " SELECT tip_des_codigo,tip_des_nombre,tip_des_factor " & _
            " FROM tipo_descuento " & _
            " WHERE tip_des_ingreso=1 AND emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount <> 0 Then
        While clsSql.adorec_Def.EOF = False
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_codigo")
            VSFG.TextMatrix(1, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_factor")
            strSql1 = " SELECT IFNULL(B.tip_des_codigo,0) AS cod_provision, IFNULL(B.tip_des_nombre,'') AS provision " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, VSFG.Cols - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
            clsSql1.Ejecutar (strSql1)
            If clsSql1.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(2, VSFG.Cols - 1) = clsSql1.adorec_Def("cod_provision")
            End If
            VSFG.ColHidden(VSFG.Cols - 1) = True
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_nombre")
            VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
            VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
            VSFG.AutoSize (VSFG.Cols - 1)
            clsSql.adorec_Def.MoveNext
        Wend
    End If
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total Ingresos"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
End Sub

Private Sub CargaEgresos()
 
    Dim Row As Long
    strSql = " SELECT tip_des_codigo,tip_des_nombre,tip_des_factor " & _
            " FROM tipo_descuento " & _
            " WHERE tip_des_ingreso=0 AND emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount <> 0 Then
        While clsSql.adorec_Def.EOF = False
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_codigo")
            VSFG.TextMatrix(1, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_factor")
            strSql1 = " SELECT IFNULL(B.tip_des_codigo,0) AS cod_provision, IFNULL(B.tip_des_nombre,'') AS provision " & _
             " FROM tipo_descuento " & _
             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
             " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, VSFG.Cols - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
            clsSql1.Ejecutar (strSql1)
            If clsSql1.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(2, VSFG.Cols - 1) = clsSql1.adorec_Def("cod_provision")
            End If
            VSFG.ColHidden(VSFG.Cols - 1) = True
            VSFG.Cols = VSFG.Cols + 1
            VSFG.TextMatrix(0, VSFG.Cols - 1) = clsSql.adorec_Def("tip_des_nombre")
            VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
            VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
            VSFG.AutoSize (VSFG.Cols - 1)
            clsSql.adorec_Def.MoveNext
        Wend
    End If
    VSFG.Cols = VSFG.Cols + 1
    VSFG.TextMatrix(0, VSFG.Cols - 1) = "Total Egresos"
    VSFG.FixedAlignment(VSFG.Cols - 1) = flexAlignCenterCenter
    VSFG.Cell(flexcpAlignment, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = 7
    VSFG.Cell(flexcpBackColor, 1, VSFG.Cols - 1, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(200, 200, 250)
    VSFG.AutoSize (VSFG.Cols - 1)
End Sub


Private Sub CargaValores()
    Dim SueldoBas As Double
    Dim SueldoMes As Double
    Dim SueldoIESS As Double
    Dim ImpRentaMes As Double
    Dim formula As String
    Dim operacion As String
    Dim valor As Double
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        SueldoBas = FormatoD2(VSFG.TextMatrix(i, datos))
        SueldoMes = FormatoD2(VSFG.TextMatrix(i, datos + 1))
        SueldoIESS = FormatoD2(VSFG.TextMatrix(i, datos + 4))
        ImpRentaMes = FormatoD2(ImpuestoRentaMes(FormatoD2(VSFG.TextMatrix(i, datos + 3)))) 'CDbl(VSFG.TextMatrix(i, 11))
        For j = (INICIO - 1) To VSFG.Cols - 2
            If Trim(VSFG.TextMatrix(1, j)) <> "" Then
                'Existe formula o valor
                'VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                formula = Trim(VSFG.TextMatrix(1, j))
                If InStr(1, UCase(formula), "SUELDOBAS") <> 0 Then
                    operacion = Replace(UCase(formula), "SUELDOBAS", SueldoBas)
                    strSql = " SELECT " & operacion
                    clsSql.Ejecutar strSql
                    VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                    VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                ElseIf InStr(1, UCase(formula), "SUELDOMES") <> 0 Then
                    operacion = Replace(UCase(formula), "SUELDOMES", SueldoMes)
                    strSql = " SELECT " & operacion
                    clsSql.Ejecutar strSql
                    VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                    VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                ElseIf InStr(1, UCase(formula), "SUELDOIESS") <> 0 Then
                    operacion = Replace(UCase(formula), "SUELDOIESS", SueldoIESS)
                    strSql = " SELECT " & operacion
                    clsSql.Ejecutar strSql
                    VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                    VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                ElseIf InStr(1, UCase(formula), "IMPRENTAMES") <> 0 Then
                    operacion = Replace(UCase(formula), "IMPRENTAMES", ImpRentaMes)
                    strSql = " SELECT " & operacion
                    clsSql.Ejecutar strSql
                    VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                    VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                Else
                    If IsNumeric(VSFG.TextMatrix(0, j)) Then
                        strSql = " SELECT " & formula
                        clsSql.Ejecutar strSql
                        VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                        VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                    End If
                    
                    
                    
'                    operacion = VSFG.TextMatrix(1, j)
'                    VSFG.TextMatrix(i, j + 1) = Format(operacion, "###0.00")
                End If
    
            ElseIf Trim(VSFG.TextMatrix(2, j)) <> "" And Trim(VSFG.TextMatrix(2, j)) <> "0" Then
                VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
                VSFG.TextMatrix(i, j + 1) = Format(SumarProvisionesPendientes(Trim(VSFG.TextMatrix(2, j)), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(0, j)), "###0.00")
            End If
            
           
            
        Next j
        End If
    Next i
    CalculoTotal
    CalculoIngresos
    CalculoEgresos
    CalculoRecibir
    CalculoOtros
    CalculoTodo
End Sub

Private Sub CalculoTotal()
    Dim Total As Double
    For i = INICIO To VSFG.Cols - 1
        Total = 0
        If VSFG.TextMatrix(VSFG.Rows - 1, i) = "*" Then
            For j = 1 To VSFG.Rows - 2
                If j <= empleados Then
                If Trim(VSFG.TextMatrix(j, i)) <> "" Then
                    Total = Total + CDbl(VSFG.TextMatrix(j, i))
                End If
                End If
            Next j
            VSFG.TextMatrix(VSFG.Rows - 1, i) = Format(Total, "###0.00")
        End If
        
    Next i
    
End Sub


Private Sub CalculoIngresos()
    Dim Col As Long
    Dim Total As Double
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            Col = i
            Exit For
        End If
    Next i
    colIngreso = Col
    For i = 1 To VSFG.Rows - 2
        Total = 0
        If i <= empleados Then
        For j = INICIO To Col Step 2
            
            If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                Total = Total + CDbl(VSFG.TextMatrix(i, j))
            End If
        Next j
        End If
        VSFG.TextMatrix(i, Col) = Format(Total, "###0.00")
    Next i
    
    
End Sub


Private Sub CalculoEgresos()
    Dim cole As Long
    Dim coli As Long
    Dim Total As Double
    coli = 0: cole = 0
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            coli = i
            Exit For
        End If
    Next i
    
    If coli = 0 Then coli = INICIO
    For i = coli To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL EGRESOS" Then
            cole = i
            Exit For
        End If
    Next i
    colEgreso = cole
    For i = 1 To VSFG.Rows - 2
        Total = 0
        If i <= empleados Then
            For j = (coli + 2) To cole Step 2
                If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                    Total = Total + CDbl(VSFG.TextMatrix(i, j))
                End If
            Next j
        End If
        VSFG.TextMatrix(i, cole) = Format(Total, "###0.00")
    Next i
    'VSFG.Cell(flexcpForeColor, VSFG.Rows - 1, 17, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(155, 0, 0)
End Sub


Private Sub CalculoOtros()
    Dim cole As Long
    Dim coli As Long
    Dim colo As Long
    Dim Total As Double
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            coli = i
            Exit For
        End If
    Next i
    For i = coli To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL EGRESOS" Then
            cole = i + 1
            Exit For
        End If
    Next i
    For i = cole To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL OTROS" Then
            colo = i
            Exit For
        End If
    Next i
    colOtros = colo
    For i = 1 To VSFG.Rows - 2
        Total = 0
        If i <= empleados Then
        For j = (cole + 2) To colo Step 2
            If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                Total = Total + CDbl(VSFG.TextMatrix(i, j))
            End If
        Next j
        End If
        VSFG.TextMatrix(i, colo) = Format(Total, "###0.00")
    Next i
    VSFG.Cell(flexcpForeColor, VSFG.Rows - 1, INICIO, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(155, 0, 0)
End Sub

Private Sub CalculoRecibir()
    Dim cole As Long
    Dim coli As Long
    Dim Total As Double
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            coli = i
            Exit For
        End If
    Next i
    For i = coli To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL EGRESOS" Then
            cole = i
            Exit For
        End If
    Next i
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        totalRecibir = CDbl(VSFG.TextMatrix(i, coli)) - CDbl(VSFG.TextMatrix(i, cole))
        VSFG.TextMatrix(i, cole + 1) = Format(totalRecibir, "###0.00")
        End If
    Next i
    'VSFG.Cell(flexcpForeColor, VSFG.Rows - 1, 17, VSFG.Rows - 1, VSFG.Cols - 1) = RGB(155, 0, 0)
End Sub

Public Sub SumarTotal()
    Dim SueldoBasico As Double
    Dim SueldoMes As Double
    Dim SueldoAnio As Double
    Dim ImpRentaMes As Double
    Dim SueldoIESS As Double
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        SueldoBasico = SueldoBasico + FormatoD2(VSFG.TextMatrix(i, datos))
        SueldoMes = SueldoMes + FormatoD2(VSFG.TextMatrix(i, datos + 1))
        SueldoAnio = SueldoAnio + FormatoD2(VSFG.TextMatrix(i, datos + 2))
        ImpRentaMes = ImpRentaMes + FormatoD2(VSFG.TextMatrix(i, datos + 3))
        SueldoIESS = SueldoIESS + FormatoD2(VSFG.TextMatrix(i, datos + 4))
        End If
    Next i
    
    If PrimeraVez Then
        SMes = SueldoMes
        SBas = SueldoBasico
        SIESS = SueldoIESS
        IRMes = ImpRentaMes
        PrimeraVez = False
    End If
    
        VSFG.TextMatrix(VSFG.Rows - 1, datos) = Format(SueldoBasico, "###0.00")
        VSFG.TextMatrix(VSFG.Rows - 1, datos + 1) = Format(SueldoMes, "###0.00")
        VSFG.TextMatrix(VSFG.Rows - 1, datos + 2) = Format(SueldoAnio, "###0.00")
        VSFG.TextMatrix(VSFG.Rows - 1, datos + 3) = Format(ImpRentaMes, "###0.00")
        VSFG.TextMatrix(VSFG.Rows - 1, datos + 4) = Format(SueldoIESS, "###0.00")

    VSFG.TextMatrix(VSFG.Rows - 1, 0) = "Totales"
    VSFG.Cell(flexcpForeColor, VSFG.Rows - 1, datos, VSFG.Rows - 1, datos + 4) = RGB(155, 0, 0)
    
End Sub

Public Sub PonerNum()
    For i = 1 To VSFG.Rows - 1
        If i <= empleados Then
        VSFG.TextMatrix(i, 0) = i
        End If
    Next i
    VSFG.Cell(flexcpAlignment, 1, 0, VSFG.Rows - 1, 0) = 4
End Sub



Private Sub Año_Change()
    CambiarFecha
End Sub





Private Sub cmbAceptar_Click()
    If MsgBox("¿Desea generar el rol de pagos?", vbQuestion + vbYesNo, "Aceptar") = vbNo Then Exit Sub
    Me.MousePointer = 11
    Dim Mensaje As String
    Dim cole As Long
    Dim coli As Long
    Dim colo As Long
    Dim Descuento As Long
    
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            coli = i
            Exit For
        End If
    Next i
    
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        If VSFG.Cell(flexcpBackColor, i, 1, i, 1) = &HC0FFC0 Then
            EliminarDescuento clsSql, VSFG.TextMatrix(i, EMP), CStr(Fecha2)
        End If
        For j = INICIO To coli Step 2
            If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                If CBool(VSFG.TextMatrix(i, 1)) Then
                    GrabarDescuento clsSql, VSFG.TextMatrix(0, j - 1), VSFG.TextMatrix(i, EMP), CStr(Fecha2), Format(VSFG.TextMatrix(i, j), "###0.00")
                End If
            End If
        Next j
        End If
    Next i
    
    
    For i = coli To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL EGRESOS" Then
            cole = i
            Exit For
        End If
    Next i
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        For j = (coli + 2) To cole Step 2
            If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                If CBool(VSFG.TextMatrix(i, 1)) Then
                    GrabarDescuento clsSql, VSFG.TextMatrix(0, j - 1), VSFG.TextMatrix(i, EMP), CStr(Fecha2), Format(VSFG.TextMatrix(i, j), "###0.00")
                End If
            End If
        Next j
        End If
    Next i
    
     For i = (cole + 1) To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL OTROS" Then
            colo = i
            Exit For
        End If
    Next i
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        For j = (cole + 3) To colo Step 2
            If Trim(VSFG.TextMatrix(i, j)) <> "" Then
                If CBool(VSFG.TextMatrix(i, 1)) Then
                    GrabarDescuento clsSql, VSFG.TextMatrix(0, j - 1), VSFG.TextMatrix(i, EMP), CStr(Fecha2), Format(VSFG.TextMatrix(i, j), "###0.00")
                End If
            End If
        Next j
        End If
    Next i
  
    Dim contador As Long
    contador = 0
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        If CBool(VSFG.TextMatrix(i, 1)) Then
            contador = contador + 1
        End If
        End If
    Next i
    Me.MousePointer = 0
    Mensaje = "Se añadieron " & contador & " registros."
    MsgBox Mensaje, vbInformation, "Información"
    
'    For i = VSFG.Rows - 1 To 1 Step -1
'        VSFG.RemoveItem i
'    Next i
    cmdMostrar_Click
    
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub CmdCerrar_Click()
    Unload Me
End Sub



Private Sub cmdMostrar_Click()
    CambiarFecha
    If Fecha1 < Format(Date, "yyyy-MM-01") Then
        Contabilizado = True
        'Cuando ya el rol esta realizado
    End If
    VSFG.Cols = colInicial
    VSFG.Clear 1
    VSFG.Rows = 6
'    For i = VSFG.Rows - 1 To 1 Step -1
'        VSFG.RemoveItem i
'    Next i
    cargar
End Sub

Private Sub Form_Load()
 'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    rowInicial = VSFG.Rows
    colInicial = VSFG.Cols
    Año = Date
    Contabilizado = False
    PrimeraVez = True
    Cambio = False
    EMP1 = EMP
    INICIO1 = INICIO
    DATOS1 = datos
    'Selecciona el mes actual
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(Date)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    cargar

End Sub

Private Sub cargar()
    CargaProvision
    CargaEmpleados
    '*******************
    If CargaIngresosOpcion2 = False Then
        VSFG.Rows = 1
        fraDetalle.Enabled = False
        Frame1.Enabled = False
        Frame2.Enabled = False
        cmbAceptar.Enabled = False
        Exit Sub
    Else
        If CargaEgresosOpcion2 = False Then
            VSFG.Rows = 1
            fraDetalle.Enabled = False
            Frame1.Enabled = False
            Frame2.Enabled = False
            cmbAceptar.Enabled = False
            Exit Sub
        Else
            fraDetalle.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            cmbAceptar.Enabled = True
            
            CargaOtrosOpcion2
            '*******************
            CargaValores
            CargaVariables
        End If
    End If
End Sub


Private Sub CargaProvision()
    strSql = " SELECT '0',t1.tip_des_codigo,t2.tip_des_codigo,t1.tip_des_nombre " & _
             " from tipo_descuento t1 " & _
             " inner join tipo_descuento t2 " & _
             " on t1.emp_codigo=t2.emp_codigo " & _
             " and t2.tip_des_provision=t1.tip_des_codigo " & _
             " WHERE t1.emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY t1.tip_des_nombre "
    clsSql.Ejecutar strSql
    Set VSFGProvision.DataSource = clsSql.adorec_Def.DataSource
End Sub



Private Sub CargaVariables()
    Dim SueldoBas As Double
    Dim SueldoMes As Double
    Dim SueldoIESS As Double
    Dim ImpRentaMes As Double
    Dim formula As String
    Dim operacion As String
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        SMes = 0
        SIESS = 0
        IRMes = 0
        SueldoBas = 0
        SueldoMes = 0
        SueldoIESS = 0
        ImpRentaMes = 0
        
        '*******SUELDO IESS****************
        For j = (INICIO - 1) To VSFG.Cols - 2
            'Caso en el que es REVISAR
            If IsNumeric(Val(VSFG.TextMatrix(0, j))) Then
                'SUELDO IESS
                If VSFG.TextMatrix(5, j) = "1" Then
                    If Trim(VSFG.TextMatrix(i, j + 1)) <> "" Then
                        If j < colIngreso Then
                            SIESS = SIESS + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colIngreso And j < colEgreso Then
                            SIESS = SIESS - CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colEgreso Then
                            SIESS = SIESS + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                    End If
                End If
            End If
        Next j
        End If
        VSFG.TextMatrix(i, datos + 4) = Format(SIESS, "###0.00")
        
    Next i
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
     SueldoIESS = CDbl(VSFG.TextMatrix(i, datos + 4))
        For j = (INICIO - 1) To VSFG.Cols - 2
        If Trim(VSFG.TextMatrix(1, j)) <> "" Then
            formula = Trim(VSFG.TextMatrix(1, j))
            If InStr(1, UCase(formula), "SUELDOIESS") <> 0 Then
                operacion = Replace(UCase(formula), "SUELDOIESS", SueldoIESS)
                strSql = " SELECT " & operacion
                clsSql.Ejecutar strSql
                VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            End If
        ElseIf Trim(VSFG.TextMatrix(2, j)) <> "" And Trim(VSFG.TextMatrix(2, j)) <> "0" Then
            VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            'VSFG.TextMatrix(i, j + 1) = Format(SumarProvisionesPendientes(Trim(VSFG.TextMatrix(2, j)), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(0, j)), "###0.00")
        End If
        Next j
        End If
    Next i
    
    '*********SUELDO MES****************
     For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        SMes = 0
        SueldoMes = 0
        For j = (INICIO - 1) To VSFG.Cols - 2
            'Caso en el que es REVISAR
            If IsNumeric(Val(VSFG.TextMatrix(0, j))) Then
                'SUELDO MES
                If VSFG.TextMatrix(3, j) = "1" Then
                    If Trim(VSFG.TextMatrix(i, j + 1)) <> "" Then
                        If j < colIngreso Then
                            SMes = SMes + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colIngreso And j < colEgreso Then
                            SMes = SMes - CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colEgreso Then
                            SMes = SMes + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                    End If
                End If
            End If
        Next j
        End If
        VSFG.TextMatrix(i, datos + 1) = Format(SMes, "###0.00")
    Next i
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
     SueldoMes = CDbl(VSFG.TextMatrix(i, datos + 1))
        For j = (INICIO - 1) To VSFG.Cols - 2
        If Trim(VSFG.TextMatrix(1, j)) <> "" Then
            formula = Trim(VSFG.TextMatrix(1, j))
            If InStr(1, UCase(formula), "SUELDOMES") <> 0 Then
                operacion = Replace(UCase(formula), "SUELDOMES", SueldoMes)
                strSql = " SELECT " & operacion
                clsSql.Ejecutar strSql
                VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            End If
        ElseIf Trim(VSFG.TextMatrix(2, j)) <> "" And Trim(VSFG.TextMatrix(2, j)) <> "0" Then
            VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            'VSFG.TextMatrix(i, j + 1) = Format(SumarProvisionesPendientes(Trim(VSFG.TextMatrix(2, j)), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(0, j)), "###0.00")
        End If
        Next j
        End If
    Next i
    
    '****************IMPUESTO A LA RENTA**************
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        IRMes = 0
        ImpRentaMes = 0
        For j = (INICIO - 1) To VSFG.Cols - 2
            'Caso en el que es REVISAR
            If IsNumeric(Val(VSFG.TextMatrix(0, j))) Then
                'SUELDO IESS
                If VSFG.TextMatrix(4, j) = "1" Then
                    If Trim(VSFG.TextMatrix(i, j + 1)) <> "" Then
                        If j < colIngreso Then
                            IRMes = IRMes + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colIngreso And j < colEgreso Then
                            IRMes = IRMes - CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                        If j > colEgreso Then
                            IRMes = IRMes + CDbl(VSFG.TextMatrix(i, j + 1))
                        End If
                    End If
                End If
            End If
        Next j
        End If
        VSFG.TextMatrix(i, datos + 3) = Format(IRMes, "###0.00")
    Next i
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
     ImpRentaMes = CDbl(VSFG.TextMatrix(i, datos + 3))
        For j = (INICIO - 1) To VSFG.Cols - 2
        If Trim(VSFG.TextMatrix(1, j)) <> "" Then
            formula = Trim(VSFG.TextMatrix(1, j))
            If InStr(1, UCase(formula), "IMPRENTAMES") <> 0 Then
                ImpRentaMes = ImpuestoRentaMes(ImpRentaMes)
                operacion = Replace(UCase(formula), "IMPRENTAMES", ImpRentaMes)
                strSql = " SELECT " & operacion
                clsSql.Ejecutar strSql
                VSFG.TextMatrix(i, j + 1) = Format(clsSql.adorec_Def(0), "###0.00")
                VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            End If
        ElseIf Trim(VSFG.TextMatrix(2, j)) <> "" And Trim(VSFG.TextMatrix(2, j)) <> "0" Then
            VSFG.TextMatrix(VSFG.Rows - 1, j + 1) = "*"
            VSFG.TextMatrix(i, j + 1) = Format(SumarProvisionesPendientes(Trim(VSFG.TextMatrix(2, j)), VSFG.TextMatrix(i, EMP), VSFG.TextMatrix(0, j)), "###0.00")
        End If
        Next j
        End If
    Next i
    calculoProvision
    CalculoTotal
    CalculoIngresos
    CalculoEgresos
    CalculoRecibir
    CalculoOtros
    CalculoTodo
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        If VSFG.TextMatrix(i, INICIO - 2) <> "0" Then
            VSFG.Cell(flexcpBackColor, i, 1, i, VSFG.Cols - 1) = &HC0C0FF
        End If
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long
    PrimeraVez = False
    On Error Resume Next
    
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    Set clsSql1 = Nothing
    Set clsSqlAux = Nothing
End Sub

Private Sub CalculoTodo()
    Dim coli As Long
    Dim cole As Long
    Dim colo As Long
    Dim Total As Double
    
    
    coli = 0: cole = 0: colo = 0
    
    
    
    For i = INICIO To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL INGRESOS" Then
            coli = i
            Exit For
        End If
    Next i
    If coli = 0 Then coli = INICIO
    
    
    For i = INICIO To coli Step 2
        Total = 0
        For j = 1 To VSFG.Rows - 2
            If j <= empleados Then
            If Trim(VSFG.TextMatrix(j, i)) <> "" Then
                If Not CBool(VSFG.TextMatrix(j, INICIO - 2)) Then
                    Total = Total + CDbl(VSFG.TextMatrix(j, i))
                End If
            End If
            End If
        Next j
        VSFG.TextMatrix(VSFG.Rows - 1, i) = Format(Total, "###0.00")
    Next i
    
    For i = coli To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL EGRESOS" Then
            cole = i
            Exit For
        End If
    Next i
    
    For i = (coli + 2) To cole Step 2
        Total = 0
        For j = 1 To VSFG.Rows - 2
            If j <= empleados Then
            If Trim(VSFG.TextMatrix(j, i)) <> "" Then
                If Not CBool(VSFG.TextMatrix(j, INICIO - 2)) Then
                    Total = Total + CDbl(VSFG.TextMatrix(j, i))
                End If
            End If
            End If
        Next j
        VSFG.TextMatrix(VSFG.Rows - 1, i) = Format(Total, "###0.00")
    Next i
    
    
    
    CalculoRecibir
    
    
    
    For i = (cole + 1) To VSFG.Cols - 1
        If UCase(VSFG.TextMatrix(0, i)) = "TOTAL OTROS" Then
            colo = i
            Exit For
        End If
    Next i
    
    For i = (cole + 3) To colo Step 2
        Total = 0
        For j = 1 To VSFG.Rows - 2
            If j <= empleados Then
            If Trim(VSFG.TextMatrix(j, i)) <> "" Then
                If Not CBool(VSFG.TextMatrix(j, INICIO - 2)) Then
                    Total = Total + CDbl(VSFG.TextMatrix(j, i))
                End If
            End If
            End If
        Next j
        VSFG.TextMatrix(VSFG.Rows - 1, i) = Format(Total, "###0.00")
    Next i



    For i = INICIO To VSFG.Cols - 1
        Total = 0
        If i = coli Or i = cole Or i = (cole + 1) Or i = colo Then
            For j = 1 To VSFG.Rows - 2
                If j <= empleados Then
                If Not CBool(VSFG.TextMatrix(j, INICIO - 2)) Then
                    Total = Total + CDbl(VSFG.TextMatrix(j, i))
                End If
                End If
            Next j
            VSFG.TextMatrix(VSFG.Rows - 1, i) = Format(Total, "###0.00")
        End If
        
    Next i
End Sub


Public Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        
    If Col = 1 Then
        If CBool(VSFG.TextMatrix(Row, 1)) Then
             strSql = " SELECT d.epl_codigo " & _
             " FROM descuento d" & _
             " WHERE d.epl_codigo='" & VSFG.TextMatrix(Row, EMP) & "' AND d.emp_codigo='" & strEmpresa & "' AND d.des_fecha='" & Fecha2 & "' "
            clsSql.Ejecutar (strSql)
            If clsSql.adorec_Def.RecordCount > 0 Then
                VSFG.Cell(flexcpBackColor, Row, 1, Row, VSFG.Cols - 1) = &HC0FFC0
            Else
                VSFG.Cell(flexcpBackColor, Row, 1, Row, VSFG.Cols - 1) = &H80FFFF
            End If
        Else
            VSFG.Cell(flexcpBackColor, Row, 1, Row, VSFG.Cols - 1) = vbDefault
        End If
    End If
    'VSFG.Cell(flexcpBackColor, 1, colIngreso, VSFG.Rows - 1, colIngreso) = RGB(200, 200, 250)
    
    If Trim(VSFG.TextMatrix(Row, Col)) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, Col)) Then
        MsgBox "Ingrese un número válido", vbInformation, "Valor incorrecto"
        VSFG.TextMatrix(Row, Col) = ""
    Else
        'If Trim(VSFG.TextMatrix(Row, Col)) <> "" Then
            VSFG.TextMatrix(Row, Col) = Format(VSFG.TextMatrix(Row, Col), "###0.00")
        'End If
        strSql = " SELECT tipo_descuento.tip_des_sueldo_mes, tipo_descuento.tip_des_impuesto_renta, tipo_descuento.tip_des_iess " & _
         " FROM tipo_descuento " & _
         " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
         " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
         " WHERE tipo_descuento.tip_des_codigo='" & VSFG.TextMatrix(0, Col - 1) & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar (strSql)
        
        
        strSql = " SELECT epl_codigo, concat(epl_apellidos,' ',epl_nombres) As nombre, epl_sueldo, epl_fec_ingreso, epl_fec_salida, are_lab_nombre " & _
             " FROM empleado e" & _
             " INNER JOIN area_laboral a ON a.are_lab_codigo = e.are_lab_codigo " & _
             " AND a.emp_codigo = e.emp_codigo " & _
             " WHERE e.emp_codigo = '" & strEmpresa & "' AND epl_baja=0 " & _
             " AND epl_codigo='" & VSFG.TextMatrix(Row, EMP) & "'" & _
             " ORDER BY are_lab_nombre,epl_apellidos,epl_nombres"
    
        'clsSql.Ejecutar strSql
        clsSqlAux.Ejecutar strSql
    
    
    
    
        CargaVariables
        Exit Sub
    
    End If
        CalculoIngresos
        CalculoEgresos
        CalculoRecibir
        CalculoOtros
        CalculoTodo
        CargaValores
    'End If
End Sub








Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim Columna As Long
    If Val(VSFG.TextMatrix(Row, INICIO - 2)) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    If Row <> VSFG.Rows - 1 Then
        If VSFG.Cols > (INICIO - 1) Then
            If Col > INICIO Then
                If VSFG.TextMatrix(1, Col - 1) <> "" Then
                    Cancel = True
                End If
            ElseIf Col <> 1 Then
                Cancel = True
            End If
        End If
    Else
        Cancel = True
    End If
End Sub



Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = EMP + 3 Or Col = EMP + 4 Then
        FechaEsteMes Row
    End If
End Sub

Private Sub VSFG_Click()
    If VSFG.Col = 1 Then
        CargaVariables
        calculoProvision
        CargaVariables
    End If
End Sub

Private Sub VSFG_DblClick()
    If VSFG.Row > 0 And InStr(1, UCase(VSFG.TextMatrix(0, VSFG.Col)), "HORAS") <> 0 Then
        frmHorasExtras.VSFG.TextMatrix(1, 1) = Format(VSFG.TextMatrix(VSFG.Row, datos + 5), "###0.00")
        frmHorasExtras.VSFG.TextMatrix(2, 1) = Format(VSFG.TextMatrix(VSFG.Row, datos + 6), "###0.00")
        frmHorasExtras.SueldoBasico = Format(VSFG.TextMatrix(VSFG.Row, datos), "###0.00")
        frmHorasExtras.Columna = CLng(VSFG.Col)
        frmHorasExtras.Show
    End If
End Sub


Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Val(VSFG.TextMatrix(VSFG.Row, (INICIO - 2))) <> 0 Then Exit Sub
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
       ' ucrtVSFG.VerMenu
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    If Val(VSFG.TextMatrix(VSFG.Row, (INICIO - 2))) <> 0 Then Exit Sub
    'ucrtVSFG.Editar KeyAscii
   If VSFG.Row > 0 And InStr(1, UCase(VSFG.TextMatrix(0, VSFG.Col)), "HORAS") <> 0 Then
        VSFG_DblClick
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub VSFG_EnterCell()
    If VSFG.Col = EMP + 3 Or VSFG.Col = EMP + 5 Or VSFG.Col = EMP + 6 Then
        VSFG.ToolTipText = "Fecha en Rojo: empleado entró en este mes. Fecha en Verde: empleado cumple un año."
    Else
        VSFG.ToolTipText = ""
    End If
End Sub


Private Sub FechaEsteMes(Row As Long)
    Dim dia As Integer
    Dim Mes As Integer
    Dim Año As Integer
    
    Dim DiaI As Integer
    Dim MesI As Integer
    Dim AñoI As Integer
    
    Dim dias As Integer
    Dim MesS As Integer
    Dim AñoS As Integer
    
    Dim IngresóMes As Boolean
    Dim SalióMes As Boolean
    Dim CumpleañosMes As Boolean
    Dim FechaIngreso As String
    Dim FechaSalida As String
    
    dia = CInt(Mid(Fecha2, 9, 2))
    Mes = CInt(Mid(Fecha2, 6, 2))
    Año = CInt(Left(Fecha2, 4))
    
    FechaIngreso = VSFG.TextMatrix(Row, (EMP + 3))
    FechaSalida = VSFG.TextMatrix(Row, (EMP + 4))
    
    DiaI = FormatoD0(Mid(FechaIngreso, 9, 2))
    MesI = FormatoD0(Mid(FechaIngreso, 6, 2))
    AñoI = FormatoD0(Left(FechaIngreso, 4))
    
    'Si el empleado entró este mes a trabajar tiene menos días
    If Mes = MesI And Año = AñoI Then
        IngresóMes = True
    End If

    If Trim(FechaSalida) <> "" Then
        dias = FormatoD0(Mid(FechaSalida, 9, 2))
        MesS = FormatoD0(Mid(FechaSalida, 6, 2))
        AñoS = FormatoD0(Left(FechaSalida, 4))
        'Si el empleado salió este mes de trabajar tiene menos días
        If Mes = MesS And Año = AñoS Then
            SalióMes = True
        End If
    End If
    'Ingresó este mes
    If IngresóMes = True And SalióMes = False Then
        VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(190, 0, 0)
        VSFG.Cell(flexcpForeColor, Row, EMP + 4) = RGB(0, 0, 0)
    'Salió este mes
    ElseIf IngresóMes = False And SalióMes = True Then
        VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(0, 0, 0)
        VSFG.Cell(flexcpForeColor, Row, EMP + 4) = RGB(190, 0, 0)
    'Ingresó y salió este mes
    ElseIf IngresóMes = True And SalióMes = True Then
        VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(190, 0, 0)
        VSFG.Cell(flexcpForeColor, Row, EMP + 4) = RGB(190, 0, 0)
    Else
        'Si ya tiene un año en la empresa
        If AñoI = Año + 1 And MesI = Mes Then
            VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(0, 120, 0)
        Else
            VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(0, 0, 0)
        End If
    End If
    
    'Si el empleado entró este mes a trabajar tiene menos días
    If Mes = MesI And Año = AñoI + 1 Then
        'Pone en verde si es el cumpleaños
        VSFG.Cell(flexcpForeColor, Row, EMP + 6) = RGB(0, 120, 0)
        VSFG.Cell(flexcpForeColor, Row, EMP + 3) = RGB(0, 120, 0)
    End If
End Sub



Private Sub VSFGProvision_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If CBool(VSFGProvision.TextMatrix(Row, 0)) Then
        VSFGProvision.Cell(flexcpBackColor, Row, 0, Row, VSFGProvision.Cols - 1) = &H80FFFF
    Else
        VSFGProvision.Cell(flexcpBackColor, Row, 0, Row, VSFGProvision.Cols - 1) = vbDefault
    End If
End Sub

Private Sub VSFGProvision_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    End If
End Sub


Private Sub VSFGProvision_Click()
    If VSFGProvision.Col = 0 Then
        CargaVariables
        calculoProvision
        CargaVariables
    End If
End Sub

Private Sub calculoProvision()
    Dim codEmpleado As String
    Dim codDescuento As String
    Dim codProvision As String
    Dim k As Long
    Dim l As Long
    
        
'    'If VSFGProvision.Col = 0 Then
'        For i = 1 To VSFGProvision.Rows - 1
'            'SI PROVISION SE HA HECHO CLICK
'            If CBool(VSFGProvision.TextMatrix(i, 0)) = True Then
'                codProvision = VSFGProvision.TextMatrix(i, 2)
'                codDescuento = VSFGProvision.TextMatrix(i, 1)
'                For j = 1 To VSFG.Rows - 2
'                    'SI SE HA SELECCIONADO PARA CONTABILIZAR
'                    If CBool(VSFG.TextMatrix(j, 1)) Then
'                        For k = INICIO To VSFG.Cols - 1
'                            If VSFG.TextMatrix(0, k) = VSFGProvision.TextMatrix(i, 1) Then
'                                codEmpleado = VSFG.TextMatrix(j, EMP)
'                                VSFG.TextMatrix(j, k + 1) = Format(SumarProvisionesPendientes(codProvision, codEmpleado, codDescuento), "###0.00")
'                            End If
'                        Next k
'                    End If
'                Next j
'
'            End If
'        Next i
    'End If
    'CargaVariables
    Dim value As Double
    
    For i = 1 To VSFG.Rows - 2
        If i <= empleados Then
        'SI SE HA SELECCIONADO PARA CONTABILIZAR
        If CBool(FormatoD0(VSFG.TextMatrix(i, 1))) = True Then
            codEmpleado = VSFG.TextMatrix(i, EMP)
            For j = 1 To VSFGProvision.Rows - 1
                'SI PROVISION SE HA HECHO CLICK
                If CBool(VSFGProvision.TextMatrix(j, 0)) = True Then
                    For k = INICIO To VSFG.Cols - 1
                        If VSFG.TextMatrix(0, k) = VSFGProvision.TextMatrix(j, 1) Then
                            codProvision = VSFGProvision.TextMatrix(j, 2)
                            codDescuento = VSFGProvision.TextMatrix(j, 1)
                            For l = INICIO To VSFG.Cols - 1
                                If VSFG.TextMatrix(0, l) = VSFGProvision.TextMatrix(j, 2) Then
                                    value = CDbl(VSFG.TextMatrix(i, l + 1))
                                End If
                            Next l
                            VSFG.TextMatrix(i, k + 1) = Format(SumarProvisionesPendientes(codProvision, codEmpleado, codDescuento) + value, "###0.00")
                        End If
                    Next k
                End If
            Next j

        End If
        End If
    Next i
End Sub

