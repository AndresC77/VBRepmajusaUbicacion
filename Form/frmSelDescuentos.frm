VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSelDescuentos 
   BackColor       =   &H00DDCCBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos / Egresos Rol"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   Icon            =   "frmSelDescuentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10260
   Begin VB.OptionButton Option2 
      BackColor       =   &H00DDCCBB&
      Caption         =   "Egresos Rol"
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
      Height          =   375
      Left            =   5483
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00DDCCBB&
      Caption         =   "Ingresos Rol"
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
      Height          =   375
      Left            =   3083
      TabIndex        =   22
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDCCBB&
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
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   9975
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Define variable Sueldo IESS"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   32
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtProvision 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Define variable Impuesto Renta"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   28
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Define variable Sueldo Mes"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   27
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Se calcula en función de provisión"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   26
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmbMesI 
         Height          =   315
         ItemData        =   "frmSelDescuentos.frx":030A
         Left            =   240
         List            =   "frmSelDescuentos.frx":0335
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   1425
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Préstamo o anticipo"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Sólo para grupos de empleados"
         ForeColor       =   &H002F1905&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   578
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Año 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
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
         Format          =   47120387
         UpDown          =   -1  'True
         CurrentDate     =   38054
      End
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   315
         Left            =   2520
         TabIndex        =   12
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
      Begin VB.Label lblDias 
         BackColor       =   &H00DDCCBB&
         Caption         =   "31"
         ForeColor       =   &H002F1905&
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Días:"
         ForeColor       =   &H002F1905&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCuenta 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Cuenta Contable:"
         ForeColor       =   &H002F1905&
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label lblFactor 
         BackColor       =   &H00DDCCBB&
         Caption         =   "Factor Cálculo:"
         ForeColor       =   &H002F1905&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes y Año:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipos de Ingresos:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   3720
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
      Left            =   4440
      TabIndex        =   9
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDCCBB&
      Caption         =   "Ingresos / Egresos Rol"
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
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   9975
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   3840
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "&Borrar descuentos"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmdAñadir2 
         Caption         =   "&Añadir Empleados"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   3825
         Width           =   1455
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFG 
         Height          =   2655
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   9495
         _cx             =   16748
         _cy             =   4683
         _ConvInfo       =   1
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelDescuentos.frx":039E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
      End
      Begin MSDataListLib.DataCombo dcmbArea 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   3840
         Width           =   3360
         _ExtentX        =   5927
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
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H002F1905&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre de ingreso/egreso"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   9495
      End
      Begin VB.Label lblPorcentaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00663300&
         Height          =   195
         Left            =   4920
         TabIndex        =   15
         Top             =   3375
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   6600
         Picture         =   "frmSelDescuentos.frx":05D0
         ToolTipText     =   "Elimina una Fila"
         Top             =   240
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   6840
         Picture         =   "frmSelDescuentos.frx":0706
         Top             =   240
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área Laboral:"
         ForeColor       =   &H002F1905&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   3600
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmSelDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private strSql As String
'Private clsSql As New clsConsulta
'Private clsSql1 As New clsConsulta
'Private clsSocio As New clsConsulta
'Private adorec_Socio As ADODB.Recordset
'Private adorec_SocioCodigo As ADODB.Recordset
'Private SociosCargados As Boolean
'Private HacerChange As Boolean
'Private Hacer As Boolean
'Private Factor As String
'Private FactorInteres As String
'Private strSqlImpresion As String
'Private CuentaContable As String
'Dim i As Integer
'Dim j As Integer
'Public Ingreso As Boolean
'Private PrimeraVez As Boolean
'
'Private Fecha1 As Variant
'Private Fecha2 As Variant
'Private Condicion As String
'
'Dim Dia1 As Integer
'Dim Dia2 As Integer
'Dim Mes1 As Integer
'Dim mes2 As Integer
'Dim Año1 As Integer
'Dim Año2 As Integer
'Private HacerReglaDe3 As Boolean
'
'Private Sub MostrarQuincena()
'    'Me.lblQuincena.Caption = QuincenaText(Me.Año.Year & Me.cmbQuincena.List(Me.cmbQuincena.ListIndex))
'    Me.cmdBuscar.Enabled = True
'    Me.cmdEditar.Visible = False
'    Me.Frame2.Visible = False
'End Sub
'
'Private Sub Año_Change()
'    CambiarFecha
'End Sub
'
'Private Sub Check1_Click(Index As Integer)
'    If Hacer = True Then Exit Sub
'    If Check1(Index).Value = 1 Then
'        Hacer = True
'        Check1(Index).Value = 0
'        Hacer = False
'    ElseIf Check1(Index).Value = 0 Then
'        Hacer = True
'        Check1(Index).Value = 1
'        Hacer = False
'    End If
'End Sub
'
'Private Sub NumerarVSFG()
'    For i = 1 To VSFG.Rows - 1
'        If VSFG.IsSubtotal(i) = False Then VSFG.TextMatrix(i, 0) = i
'        If Trim(VSFG.TextMatrix(i, 2)) <> "" And VSFG.IsSubtotal(i) = False Then
'            VSFG.Cell(flexcpPicture, i, 1) = Me.imgBtnUp
'            VSFG.Cell(flexcpPictureAlignment, i, 1) = flexPicAlignRightCenter
'        Else
'            VSFG.Cell(flexcpPicture, i, 1) = Nothing
'        End If
'    Next i
'End Sub
'Private Sub SumarVSFG()
'    Me.VSFG.SubtotalPosition = flexSTBelow
'    VSFG.SubTotal flexSTSum, -1, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
'    VSFG.SubTotal flexSTSum, -1, 6, "#,##0.00"
'    VSFG.SubTotal flexSTSum, -1, 7, "#,##0.00"
'    VSFG.SubTotal flexSTSum, -1, 8, "#,##0.00"
'End Sub
'
'Private Sub cmbMesI_Click()
'    CambiarFecha
'End Sub
'
'Private Sub CambiarFecha()
'    'If HacerFecha = False Then Exit Sub
'    Dim DiaFinal As Integer
'
'    Fecha1 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-1"
'    Fecha2 = ""
'    DiaFinal = 31
'    While (IsDate(Fecha2) = False)
'        Fecha2 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
'        lblDias = DiaFinal
'        DiaFinal = DiaFinal - 1
'    Wend
'    Fecha1 = Format(Fecha1, "yyyy-mm-dd")
'    Fecha2 = Format(Fecha2, "yyyy-mm-dd")
'    'MostrarAsientos
'    Me.cmdBuscar.Enabled = True
'    Me.cmdEditar.Visible = False
'    Me.Frame2.Visible = False
'    PonerEtiquetas
'End Sub
'
'Private Sub PonerEtiquetas()
'    If Me.Ingreso = True Then
'        Me.lblNombre.Caption = "Tipos de Ingresos:"
'        Me.Caption = "Ingresos Rol - " & StrConv(Me.cmbMesI.List(Me.cmbMesI.ListIndex), vbProperCase) & " " & Me.Año.Year
'        Me.Frame2 = "Ingreso Rol de Pagos - " & StrConv(Me.cmbMesI.List(Me.cmbMesI.ListIndex), vbProperCase) & " " & Me.Año.Year
'        Me.cmdBorrar.Caption = "Borrar Ingresos"
'    Else
'        Me.lblNombre.Caption = "Tipos de Egresos:"
'        Me.Caption = "Egresos Rol - " & StrConv(Me.cmbMesI.List(Me.cmbMesI.ListIndex), vbProperCase) & " " & Me.Año.Year
'        Me.Frame2 = "Egreso Rol de Pagos - " & StrConv(Me.cmbMesI.List(Me.cmbMesI.ListIndex), vbProperCase) & " " & Me.Año.Year
'        Me.cmdBorrar.Caption = "Borrar Egresos"
'    End If
'End Sub
'
'Private Sub cmdAñadir2_Click()
'    If Me.dcmbArea.BoundText = "%" Then
'        If MsgBox("¿Está seguro de AÑADIR todos los empleados en todos los estados?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub
'    Else
'        If MsgBox("¿Está seguro de AÑADIR todos los empleados en estado " & Me.dcmbArea & "?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub
'    End If
'    Dim Encontro As Boolean
'    Dim Registros As Long
'    Dim Mensaje As String
'    Dim CadenaEval As String
'    Dim ElCapital As Double
'    Dim ElInteres As Double
'    Dim HacerBusqueda As Boolean
'    'Dim NumeroSocios
'    Dim Row As Long
'    Screen.MousePointer = vbHourglass
'    strSql = " SELECT epl_codigo, epl_apellidos+' '+epl_nombres As nombre, epl_sueldo, epl_fec_ingreso, epl_fec_salida FROM empleado " & _
'             " WHERE are_lab_codigo LIKE '" & Trim(Me.dcmbArea.BoundText) & "' " & Condicion & " AND epl_baja=0 ORDER BY epl_apellidos,epl_nombres"
'    clsSql.Ejecutar (strSql)
'
'    Me.lblPorcentaje.Visible = True
'    Me.lblPorcentaje.Caption = "0%"
'    If clsSql.adorec_Def.RecordCount > 0 Then
'        ProgressBar1.Max = clsSql.adorec_Def.RecordCount
'    Else
'        ProgressBar1.Max = 1
'    End If
'    ProgressBar1.Value = ProgressBar1.Min
'    ProgressBar1.Visible = True
'    HacerChange = False
'    If VSFG.Rows = 2 Then
'        HacerBusqueda = False
'    Else
'        HacerBusqueda = True
'    End If
'    While clsSql.adorec_Def.EOF = False
'        Encontro = False
'        If HacerBusqueda = True Then
'            For i = 1 To VSFG.Rows - 1
'                If VSFG.TextMatrix(i, 3) = clsSql.adorec_Def("epl_codigo") Then
'                    Encontro = True
'                    Exit For
'                End If
'            Next i
'        End If
'        If Encontro = False Then
'            'Añadir
'            Row = VSFG.Rows - 2
'            VSFG.AddItem "", Row
'
'            VSFG.TextMatrix(Row, 3) = clsSql.adorec_Def("epl_codigo")
'            VSFG.TextMatrix(Row, 4) = clsSql.adorec_Def("nombre")
'            VSFG.TextMatrix(Row, 8) = clsSql.adorec_Def("epl_sueldo")
'            VSFG.TextMatrix(Row, 13) = clsSql.adorec_Def("epl_fec_ingreso")
'            If IsNull(clsSql.adorec_Def("epl_fec_salida")) = False Then
'                VSFG.TextMatrix(Row, 16) = clsSql.adorec_Def("epl_fec_salida")
'            Else
'                VSFG.TextMatrix(Row, 16) = ""
'            End If
'            VSFG.TextMatrix(Row, 14) = DiasFinDeMes(CStr(Fecha2), VSFG.TextMatrix(Row, 13), VSFG.TextMatrix(Row, 16))
'            VSFG.TextMatrix(Row, 15) = DiasFondo(CStr(Fecha2), VSFG.TextMatrix(Row, 13), VSFG.TextMatrix(Row, 16))
'
'
'            HacerReglaDe3 = True
'            'Capital
'            CadenaEval = Replace(Factor, "SueldoBas", clsSql.adorec_Def("epl_sueldo"))
'            'Si es que se calcula en función de provisiones
'            If Check1(4).Value = 1 And CadenaEval <> "0" Then
'                HacerReglaDe3 = False
'                CadenaEval = SumarProvisionesPendientes(Me.txtProvision.Tag, clsSql.adorec_Def("epl_codigo"), Me.dcmbTipo.BoundText)
'            Else
'                If VSFG.ColHidden(6) = False Then
'                    HacerReglaDe3 = False
'                    If VSFG.TextMatrix(0, 6) = "Sueldo Mes" Then
'                        VSFG.TextMatrix(Row, 6) = FormatoD(SueldoMes(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2))
'                        CadenaEval = Replace(CadenaEval, "SueldoMes", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                    ElseIf VSFG.TextMatrix(0, 6) = "Sueldo Año" Then
'                        VSFG.TextMatrix(Row, 6) = FormatoD(SueldoAño(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2))
'                        CadenaEval = Replace(CadenaEval, "SueldoAño", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                    ElseIf VSFG.TextMatrix(0, 6) = "Renta Mes" Then
'                        VSFG.TextMatrix(Row, 6) = FormatoD(RentaMes(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2))
'                        CadenaEval = Replace(CadenaEval, "ImpRentaMes", Format(ImpuestoRentaMes(VSFG.TextMatrix(Row, 6)), "#0.00"))
'                    ElseIf VSFG.TextMatrix(0, 6) = "Sueldo IESS" Then
'                        VSFG.TextMatrix(Row, 6) = FormatoD(SueldoIESS(clsSql.adorec_Def("epl_codigo"), Fecha1, Fecha2))
'                        CadenaEval = Replace(CadenaEval, "SueldoIESS", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                    End If
'                End If
'            End If
'            If Trim(CadenaEval) <> "" Then
'                'Se evaluará la expresión con la base de datos
'                strSql = " SELECT " & CadenaEval
'                clsSql1.Ejecutar (strSql)
'                ElCapital = Formato(clsSql1.adorec_Def(0))
'                'Sacar el proporcional según días del mes trabajados
'                If HacerReglaDe3 = True Then
'                    ElCapital = FormatoD(ElCapital * CInt(VSFG.TextMatrix(Row, 14)) / CInt(lblDias))
'                End If
'                'Si es fondo de cesantía calcular con la columna 15
'                If Me.dcmbTipo.BoundText = "1003" Then
'                    ElCapital = FormatoD(ElCapital * CInt(VSFG.TextMatrix(Row, 15)) / CInt(lblDias))
'                End If
'            Else
'                ElCapital = 0
'            End If
'            VSFG.TextMatrix(Row, 5) = ElCapital
'            VSFG.TextMatrix(Row, 2) = GrabarDescuento(clsSql1, Me.dcmbTipo.BoundText, VSFG.TextMatrix(Row, 3), CStr(Fecha2), FormatoD(VSFG.TextMatrix(Row, 5)))
'            Registros = Registros + 1
'        End If
'        ProgressBar1.Value = ProgressBar1.Value + 1
'
'        Me.lblPorcentaje.Caption = Val(Format(ProgressBar1.Value * 100 / ProgressBar1.Max, "#0")) & "%"
'        Me.lblPorcentaje.Refresh
'        clsSql.adorec_Def.MoveNext
'    Wend
'
'    NumerarVSFG
'    SumarVSFG
'    HacerChange = True
'    ProgressBar1.Visible = False
'    lblPorcentaje.Visible = False
'    Screen.MousePointer = vbDefault
'    Select Case Registros
'    Case 0
'        Mensaje = "No se añadió ningún registro."
'    Case 1
'        Mensaje = "Se añadió 1 registro."
'    Case Else
'        Mensaje = "Se añadieron " & Registros & " registros."
'    End Select
'    MsgBox Mensaje, vbInformation, "Información"
'End Sub
'
'Private Sub cmdBorrar_Click()
'    If MsgBox("¿Está seguro de BORRAR todos los " & Me.Frame2.Caption & vbNewLine & "de " & Me.dcmbTipo & " para el mes de " & Me.cmbMesI.List(Me.cmbMesI.ListIndex) & " de " & Me.Año.Year & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Pregunta") = vbNo Then Exit Sub
'    ProgressBar1.Max = VSFG.Rows - 1
'    ProgressBar1.Value = ProgressBar1.Min
'    ProgressBar1.Visible = True
'    lblPorcentaje.Visible = True
'    lblPorcentaje.Caption = "0%"
'    'Papayas gracias a la super función jejejejeje
'    Screen.MousePointer = vbHourglass
'    For i = 1 To VSFG.Rows - 1
'        If Trim(VSFG.TextMatrix(i, 2)) <> "" And Val(VSFG.TextMatrix(i, 10)) = 0 And VSFG.IsSubtotal(i) = False Then
'            EliminarDescuento clsSql1, VSFG.TextMatrix(i, 2)
'        End If
'        ProgressBar1.Value = ProgressBar1.Value + 1
'        Me.lblPorcentaje.Caption = Val(Format(ProgressBar1.Value * 100 / ProgressBar1.Max, "#0")) & "%"
'        Me.lblPorcentaje.Refresh
'    Next i
'    'VSFG.Rows = 2
'    'VSFG.Clear 1
'    'NumerarVSFG
'    'SumarVSFG
'    BuscarDescuentos
'    EditarDescuentos
'    ProgressBar1.Visible = False
'    lblPorcentaje.Visible = False
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub cmdBuscar_Click()
'    BuscarDescuentos
'End Sub
'
'Private Sub cmdEditar_Click()
'    'If VerificarFechaContable(Me.Año) = False Then Exit Sub
'    Unload frmSelDescuentos2
'    EditarDescuentos
'End Sub
'
'Private Sub cmdExportar_Click()
'    SeleccionarFlexGrid2 Me.VSFG
'    CopiarFlexGrid2 Me.VSFG
'    MsgBox "Se ha copiado la tabla de " & Me.Frame2 & " al portapapeles.", vbInformation, "Información"
'End Sub
'
'Private Sub cmdImprimir_Click()
'
'    'AA = 23
'    'drptDescuentos.TieneInteres = Me.Check1(2).Value
'    drptDescuentos.Ingreso = Me.Ingreso
'    drptDescuentos.TieneInteres = 0
'    drptDescuentos.TipoDescuento = Me.dcmbTipo
'    drptDescuentos.Quincena = Me.cmbMesI.List(Me.cmbMesI.ListIndex) & " de " & Me.Año.Year
'    drptDescuentos.Total1 = VSFG.TextMatrix(VSFG.Rows - 1, 5)
'    drptDescuentos.Total2 = VSFG.TextMatrix(VSFG.Rows - 1, 6)
'    drptDescuentos.Total3 = VSFG.TextMatrix(VSFG.Rows - 1, 7)
'    drptDescuentos.strSqlConsulta = strSqlImpresion
'    drptDescuentos.Show
'End Sub
'
'Private Sub CmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub BuscarDescuentos()
'    Screen.MousePointer = vbHourglass
'
'    'Me.Frame2.Caption = "Descuentos " & StrConv(Me.lblQuincena.Caption, vbProperCase)
'
'    strSql = " SELECT des_codigo, descuento.epl_codigo, epl_apellidos+ ' '+epl_nombres AS nombre, des_valor, 0, 0, epl_sueldo, "
'    If Me.Check1(1).Value = 1 Then
'        strSql = strSql & "det_asiento.asi_numasiento+' '+LEFT(CONVERT(VARCHAR,asi_fecha,20),10)+' '+CAST(det_asi_debe AS VARCHAR)"
'    Else
'        strSql = strSql & "''"
'    End If
'    strSql = strSql & ", des_pagado, ISNULL(des_valor1,0), ISNULL(des_valor2,0), epl_fec_ingreso,'','', epl_fec_salida " & _
'             " FROM descuento INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo"
'    'Si tiene un asiento relacionado buscar fecha y valor
'    If Me.Check1(1).Value = 1 Then
'        strSql = strSql & " LEFT JOIN asiento ON descuento.asi_numasiento=asiento.asi_numasiento AND descuento.emp_codigo=asiento.emp_codigo" & _
'                " LEFT JOIN det_asiento ON det_asiento.asi_numasiento=asiento.asi_numasiento AND det_asiento.emp_codigo=asiento.emp_codigo AND det_asi_haber=0 AND det_asiento.cta_codigo = '" & CuentaContable & "'"
'    End If
'    strSql = strSql & " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
'             " AND tip_des_codigo='" & Me.dcmbTipo.BoundText & "' ORDER BY epl_apellidos, epl_nombres"
'    clsSql.Ejecutar (strSql)
'    strSqlImpresion = strSql
'    HacerChange = False
'    VSFG.Editable = flexEDNone
'    Set Me.VSFG.DataSource = clsSql.adorec_Def.DataSource
'    NumerarVSFG
'
'    'Verificar si hay variable de SueldoMes, SueldoAño, SueldoIESS o ImpRentaMes
'    If InStr(1, Factor, "SueldoMes", vbTextCompare) <> 0 Then
'        VSFG.ColHidden(6) = False
'        VSFG.TextMatrix(0, 6) = "Sueldo Mes"
'        For i = 1 To VSFG.Rows - 1
'            VSFG.TextMatrix(i, 6) = FormatoD(SueldoMes(VSFG.TextMatrix(i, 3), Fecha1, Fecha2))
'        Next i
'    End If
'    If InStr(1, Factor, "SueldoAño", vbTextCompare) <> 0 Then
'        VSFG.ColHidden(6) = False
'        VSFG.TextMatrix(0, 6) = "Sueldo Año"
'        For i = 1 To VSFG.Rows - 1
'            VSFG.TextMatrix(i, 6) = FormatoD(SueldoAño(VSFG.TextMatrix(i, 3), Fecha1, Fecha2))
'        Next i
'    End If
'    If InStr(1, Factor, "ImpRentaMes", vbTextCompare) <> 0 Then
'        VSFG.ColHidden(6) = False
'        VSFG.TextMatrix(0, 6) = "Renta Mes"
'        For i = 1 To VSFG.Rows - 1
'            VSFG.TextMatrix(i, 6) = FormatoD(RentaMes(VSFG.TextMatrix(i, 3), Fecha1, Fecha2))
'        Next i
'    End If
'    If InStr(1, Factor, "SueldoIESS", vbTextCompare) <> 0 Then
'        VSFG.ColHidden(6) = False
'        VSFG.TextMatrix(0, 6) = "Sueldo IESS"
'        For i = 1 To VSFG.Rows - 1
'            VSFG.TextMatrix(i, 6) = FormatoD(SueldoIESS(VSFG.TextMatrix(i, 3), Fecha1, Fecha2))
'        Next i
'    End If
'
'    If Me.dcmbTipo.BoundText = "1003" Then
'        VSFG.ColHidden(15) = False
'    End If
'
'    For i = 1 To VSFG.Rows - 1
'        VSFG.TextMatrix(i, 14) = DiasFinDeMes(CStr(Fecha2), VSFG.TextMatrix(i, 13), VSFG.TextMatrix(i, 16))
'        VSFG.TextMatrix(i, 15) = DiasFondo(CStr(Fecha2), VSFG.TextMatrix(i, 13), VSFG.TextMatrix(i, 16))
'    Next i
'
'    If Me.Check1(1).Value = 1 Then
'        Me.VSFG.ColHidden(9) = False
'    Else
'        Me.VSFG.ColHidden(9) = True
'    End If
'
'    SumarVSFG
'
'    Me.cmdBuscar.Enabled = False
'    Me.Frame2.Visible = True
'    Me.cmdEditar.Visible = True
'    If Check1(1).Value = 1 Then
'        Me.cmdEditar.Enabled = False
'    Else
'        Me.cmdEditar.Enabled = True
'    End If
'    Me.cmdAñadir2.Enabled = False
'    Me.cmdBorrar.Enabled = False
'    Me.cmdExportar.Enabled = True
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub EditarDescuentos()
'    Screen.MousePointer = vbHourglass
'
'    'Añadir item para añadir
'    If VSFG.Rows = 1 Then
'        Me.VSFG.AddItem VSFG.Rows, VSFG.Rows
'    Else
'        Me.VSFG.AddItem VSFG.Rows - 1, VSFG.Rows - 1
'    End If
'    Dim Primera As Boolean
'
'    Condicion = ""
'    Primera = True
'    strSql = " SELECT are_lab_codigo " & _
'             " FROM det_tip_descuento " & _
'             " WHERE det_tip_descuento.emp_codigo='" & strEmpresa & "' AND tip_des_codigo='" & Me.dcmbTipo.BoundText & "'"
'    clsSql.Ejecutar (strSql)
'    While clsSql.adorec_Def.EOF = False
'        If Primera = True Then
'            Condicion = "AND (are_lab_codigo='" & clsSql.adorec_Def(0) & "' "
'            Primera = False
'        Else
'            Condicion = Condicion & "OR are_lab_codigo='" & clsSql.adorec_Def(0) & "' "
'        End If
'        clsSql.adorec_Def.MoveNext
'    Wend
'    If Condicion <> "" Then Condicion = Condicion & ")"
'
'    If Me.Check1(0).Value = 1 Then
'        VSFG.Editable = flexEDNone
'    Else
'        'CargarSocios
'        strSql = " SELECT epl_codigo,epl_apellidos+' '+epl_nombres AS nombre, epl_sueldo, epl_fec_ingreso, epl_fec_salida " & _
'                 " FROM empleado " & _
'                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_baja=0 " & Condicion & _
'                 " ORDER BY epl_apellidos, epl_nombres"
'        clsSocio.Ejecutar strSql
'        Set adorec_Socio = clsSocio.adorec_Def.Clone
'        VSFG.Editable = flexEDKbdMouse
'        'VSFG.ColComboList(3) = VSFG.BuildComboList(adorec_SocioCodigo, "*epl_codigo, nombre", "epl_codigo")
'        VSFG.ColComboList(4) = VSFG.BuildComboList(adorec_Socio, "epl_codigo, *nombre", "epl_codigo")
'    End If
'
'    strSql = " SELECT '%' AS codigo, ' --Todas Las Áreas Laborales--' AS nombre UNION" & _
'             " SELECT are_lab_codigo AS codigo, are_lab_nombre AS nombre FROM area_laboral " & _
'             " WHERE emp_codigo='" & strEmpresa & "' " & Condicion & " ORDER BY codigo"
'    clsSql.Ejecutar (strSql)
'    If clsSql.adorec_Def.RecordCount > 0 Then
'        Set dcmbArea.RowSource = clsSql.adorec_Def.DataSource
'        dcmbArea.ListField = "nombre"
'        dcmbArea.BoundColumn = "codigo"
'        dcmbArea.BoundText = clsSql.adorec_Def(0)
'    End If
'    SumarVSFG
'    HacerChange = True
'    Me.cmdEditar.Enabled = False
'    Me.cmdAñadir2.Enabled = True
'    Me.cmdBorrar.Enabled = True
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub dcmbTipo_Change()
'    Screen.MousePointer = vbHourglass
'    Me.cmdBuscar.Enabled = True
'    Me.Frame2.Visible = False
'    Me.cmdEditar.Visible = False
'
'    'Buscar parámetros del tipo de Descuento
'
'
'    strSql = " SELECT ISNULL((cta1.cta_codigo+' - '+cta1.cta_nombre),'') AS cta_nombre, ISNULL(cta1.cta_codigo,'') AS cta_codigo," & _
'             " ISNULL(tipo_descuento.tip_des_factor,'') AS tip_des_factor, tipo_descuento.tip_des_solo_grupos, tipo_descuento.tip_des_prestamo, tipo_descuento.tip_des_sueldo_mes, tipo_descuento.tip_des_impuesto_renta, tipo_descuento.tip_des_iess, tipo_descuento.tip_des_orden, ISNULL(B.tip_des_codigo,0) AS cod_provision, ISNULL(B.tip_des_nombre,'') AS provision " & _
'             " FROM tipo_descuento " & _
'             " LEFT JOIN ctaconta cta1 ON  tipo_descuento.cta_codigo=cta1.cta_codigo AND tipo_descuento.emp_codigo=cta1.emp_codigo" & _
'             " LEFT JOIN tipo_descuento B ON tipo_descuento.tip_des_codigo=B.tip_des_provision AND tipo_descuento.emp_codigo=B.emp_codigo" & _
'             " WHERE tipo_descuento.tip_des_codigo='" & dcmbTipo.BoundText & "' AND tipo_descuento.emp_codigo='" & strEmpresa & "'"
'    clsSql.Ejecutar (strSql)
'    CuentaContable = clsSql.adorec_Def("cta_codigo")
'    lblCuenta.Caption = "Cuenta Contable HABER: " & clsSql.adorec_Def("cta_nombre")
'    Factor = clsSql.adorec_Def("tip_des_factor")
'    lblFactor.Caption = "Factor Cálculo: " & Factor
'    'FactorInteres = clsSql.adorec_Def("tip_des_factor_interes")
'    Hacer = True
'    If clsSql.adorec_Def("tip_des_solo_grupos") = True Then
'        Me.Check1(0).Value = 1
'    Else
'        Me.Check1(0).Value = 0
'    End If
'    If clsSql.adorec_Def("tip_des_prestamo") = True Then
'        Me.Check1(1).Value = 1
'    Else
'        Me.Check1(1).Value = 0
'    End If
'    If clsSql.adorec_Def("tip_des_sueldo_mes") = True Then
'        Me.Check1(2).Value = 1
'    Else
'        Me.Check1(2).Value = 0
'    End If
'    If clsSql.adorec_Def("tip_des_impuesto_renta") = True Then
'        Me.Check1(3).Value = 1
'    Else
'        Me.Check1(3).Value = 0
'    End If
'    If Trim(clsSql.adorec_Def("provision")) <> "" = True Then
'        Me.Check1(4).Value = 1
'        Me.txtProvision.Enabled = True
'        Me.txtProvision = clsSql.adorec_Def("provision")
'        Me.txtProvision.Tag = clsSql.adorec_Def("cod_provision")
'    Else
'        Me.Check1(4).Value = 0
'        Me.txtProvision.Enabled = False
''        Me.txtProvision = clsSql.adorec_Def("provision")
''        Me.txtProvision.Tag = clsSql.adorec_Def("cod_provision")
'    End If
'    If clsSql.adorec_Def("tip_des_iess") = True Then
'        Me.Check1(5).Value = 1
'    Else
'        Me.Check1(5).Value = 0
'    End If
'
'    Me.lblTipo.Caption = Me.dcmbTipo
'    Hacer = False
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub Form_Activate()
'    If PrimeraVez = True Then
'        strSql = " SELECT tip_des_codigo, tip_des_nombre FROM tipo_descuento" & _
'                 " WHERE emp_codigo='" & strEmpresa & "' AND tip_des_ingreso=" & Abs(CInt(Ingreso)) & "" & _
'                 " ORDER BY tip_des_orden"
'        clsSql.Ejecutar (strSql)
'        Set Me.dcmbTipo.RowSource = clsSql.adorec_Def.DataSource
'        dcmbTipo.ListField = "tip_des_nombre"
'        dcmbTipo.BoundColumn = "tip_des_codigo"
'        If clsSql.adorec_Def.RecordCount > 0 Then
'            dcmbTipo.BoundText = clsSql.adorec_Def("tip_des_codigo")
'        End If
'        PonerEtiquetas
'        PrimeraVez = False
'    End If
'End Sub
'
'Private Sub Form_Load()
'    'Centra esta forma dentro de la forma MDI
'    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
'    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)
'
'    clsSql.Inicializar AdoConn
'    clsSql1.Inicializar AdoConn
'    clsSocio.Inicializar AdoConn
'    Año = Date
'    'Selecciona el mes actual
'    For i = 0 To 11
'        If (cmbMesI.ItemData(i) = Month(Date)) Then
'            cmbMesI.ListIndex = i
'            Exit For
'        End If
'    Next i
'
'    PrimeraVez = True
'    SociosCargados = False
'    Hacer = False
'End Sub
'
'Private Sub CargarEstado()
'    'Lleno combo quincena
'    For i = 1 To 24
'        cmbQuincena.AddItem Format(i, "00")
'    Next i
'    'Poner quincena actual
'    'cmbQuincena.ListIndex = Val(Right(Quincena(Date), 2)) - 1
'
'
'
'
''    strSql = " SELECT cli_est_codigo,cli_est_nombre " & _
''             " FROM cliente_estado " & _
''             " ORDER BY cli_est_nombre"
''    clsSql.Ejecutar strSql
''    Set Me.dcmbEstadoSocio.RowSource = clsSql.adorec_Def.DataSource
''    dcmbEstadoSocio.ListField = "cli_est_nombre"
''    dcmbEstadoSocio.BoundColumn = "cli_est_codigo"
''    If clsSql.adorec_Def.RecordCount > 0 Then
''        dcmbEstadoSocio.BoundText = "1 "
''    End If
'End Sub
'
'Private Sub CargarSocios()
'    If SociosCargados = True Then Exit Sub
'
'
''    strSql = " SELECT cli_codigo,cli_nombre, cli_sueldo_bas " & _
''             " FROM cliente " & _
''             " ORDER BY cli_codigo"
''    clsSql.Ejecutar strSql
''    Set adorec_SocioCodigo = clsSql.adorec_Def.Clone
'    SociosCargados = True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = 0
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub Option1_Click()
'    Ingreso = True
'    PrimeraVez = True
'    Call Form_Activate
'End Sub
'
'Private Sub Option2_Click()
'    Ingreso = False
'    PrimeraVez = True
'    Call Form_Activate
'End Sub
'
'Public Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    If Col = 5 Then
'        VSFG.TextMatrix(Row, Col) = FormatoD(VSFG.TextMatrix(Row, Col))
'        'Para llamar al change
'        'VSFG.TextMatrix(Row, 7) = "  "
'    End If
'
'    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 9 Then
'        If Trim(VSFG.TextMatrix(Row, 2)) = "" Then
'            If Trim(VSFG.TextMatrix(VSFG.Row, 3)) = "" And Trim(VSFG.TextMatrix(VSFG.Row, 4)) = "" Then Exit Sub
'            'Grabar
'            VSFG.TextMatrix(Row, 2) = GrabarDescuento(clsSql1, Me.dcmbTipo.BoundText, VSFG.TextMatrix(Row, 3), CStr(Fecha2), FormatoD(VSFG.TextMatrix(Row, 5)))
'            Me.VSFG.AddItem "", VSFG.Rows - 1
'            NumerarVSFG
'        Else
'            ActualizarDescuento VSFG.TextMatrix(Row, 2), VSFG.TextMatrix(Row, 3), FormatoD(VSFG.TextMatrix(Row, 5)), 0, VSFG.TextMatrix(Row, 11), VSFG.TextMatrix(Row, 12), Trim(Left(VSFG.TextMatrix(Row, 9), 14))
'        End If
'        If Col = 5 Or Col = 6 Then
'            SumarVSFG
'        End If
'    End If
'End Sub
'
'Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Val(VSFG.TextMatrix(Row, 10)) <> 0 Then
'        Cancel = True
'        Exit Sub
'    End If
'    If VSFG.IsSubtotal(Row) = True Then
'        Cancel = True
'        Exit Sub
'    End If
'    If Col <> 3 And Col <> 4 And Col <> 5 Then 'And Col <> 9 Then
'        Cancel = True
'    End If
'    If Col <> 3 Then
'        VSFG.ComboList = ""
'    End If
'    'Bloquear cuando sean horas extras
'    If Trim(VSFG.TextMatrix(Row, 2)) = "" Then
'        If Col = 5 Then
'            Cancel = True
'        End If
'    End If
'    If Col = 5 And Me.dcmbTipo.BoundText = "2" Then
'        Cancel = True
'    End If
''    If Col = 9 Then
''        If Trim(VSFG.TextMatrix(Row, 3)) <> "" Then
''            strSql = " SELECT det_asiento.asi_numasiento, det_asiento.asi_numasiento+' '+LEFT(CONVERT(VARCHAR,asi_fecha,20),10)+' '+CAST(det_asi_debe AS VARCHAR) AS Info FROM det_asiento" & _
''                     " INNER JOIN asiento ON det_asiento.asi_numasiento=asiento.asi_numasiento AND det_asiento.emp_codigo=asiento.emp_codigo" & _
''                     " WHERE det_asiento.emp_codigo='" & strEmpresa & "' AND det_asi_haber=0 AND cta_codigo = '" & CuentaContable & "'" & _
''                     " AND asi_fecha BETWEEN '" & DateAdd("m", -6, Fecha1) & "' AND '" & Fecha2 & "'" & _
''                     " ORDER BY asi_fecha DESC, det_asiento.asi_numasiento"
''            clsSql1.Ejecutar (strSql)
''            'Buscar asientos
''            VSFG.ComboList = VSFG.BuildComboList(clsSql1.adorec_Def, "asi_numasiento, *Info", "asi_numasiento")
''        Else
''            Cancel = True
''        End If
''    End If
'End Sub
'
'Private Sub FechaEsteMes(Row As Long)
'    Dim dia As Integer
'    Dim Mes As Integer
'    Dim Año As Integer
'
'    Dim DiaI As Integer
'    Dim MesI As Integer
'    Dim AñoI As Integer
'
'    Dim dias As Integer
'    Dim MesS As Integer
'    Dim AñoS As Integer
'
'    Dim IngresóMes As Boolean
'    Dim SalióMes As Boolean
'    Dim CumpleañosMes As Boolean
'    Dim FechaIngreso As String
'    Dim FechaSalida As String
'
'    dia = CInt(Mid(Fecha2, 9, 2))
'    Mes = CInt(Mid(Fecha2, 6, 2))
'    Año = CInt(Left(Fecha2, 4))
'
'    FechaIngreso = VSFG.TextMatrix(Row, 13)
'    FechaSalida = VSFG.TextMatrix(Row, 16)
'
'    DiaI = CInt(Mid(FechaIngreso, 9, 2))
'    MesI = CInt(Mid(FechaIngreso, 6, 2))
'    AñoI = CInt(Left(FechaIngreso, 4))
'
'    'Si el empleado entró este mes a trabajar tiene menos días
'    If Mes = MesI And Año = AñoI Then
'        IngresóMes = True
'    End If
'
'    If Trim(FechaSalida) <> "" Then
'        dias = CInt(Mid(FechaSalida, 9, 2))
'        MesS = CInt(Mid(FechaSalida, 6, 2))
'        AñoS = CInt(Left(FechaSalida, 4))
'        'Si el empleado salió este mes de trabajar tiene menos días
'        If Mes = MesS And Año = AñoS Then
'            SalióMes = True
'        End If
'    End If
'    'Ingresó este mes
'    If IngresóMes = True And SalióMes = False Then
'        VSFG.Cell(flexcpForeColor, Row, 13) = RGB(190, 0, 0)
'        VSFG.Cell(flexcpForeColor, Row, 16) = RGB(0, 0, 0)
'    'Salió este mes
'    ElseIf IngresóMes = False And SalióMes = True Then
'        VSFG.Cell(flexcpForeColor, Row, 13) = RGB(0, 0, 0)
'        VSFG.Cell(flexcpForeColor, Row, 16) = RGB(190, 0, 0)
'    'Ingresó y salió este mes
'    ElseIf IngresóMes = True And SalióMes = True Then
'        VSFG.Cell(flexcpForeColor, Row, 13) = RGB(190, 0, 0)
'        VSFG.Cell(flexcpForeColor, Row, 16) = RGB(190, 0, 0)
'    Else
'        'Si ya tiene un año en la empresa
'        If AñoI = Año + 1 And MesI = Mes Then
'            VSFG.Cell(flexcpForeColor, Row, 13) = RGB(0, 120, 0)
'        Else
'            VSFG.Cell(flexcpForeColor, Row, 13) = RGB(0, 0, 0)
'        End If
'    End If
'
'    'Si el empleado entró este mes a trabajar tiene menos días
'    If Mes = MesI And Año = AñoI + 1 Then
'        'Pone en verde si es el cumpleaños
'        VSFG.Cell(flexcpForeColor, Row, 15) = RGB(0, 120, 0)
'        VSFG.Cell(flexcpForeColor, Row, 13) = RGB(0, 120, 0)
'    End If
'End Sub
'
'Private Sub VSFG_BeforeScrollTip(ByVal Row As Long)
'    VSFG.ScrollTipText = "Empleado: " & VSFG.TextMatrix(Row, 4)
'End Sub
'
'Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    If Row = 0 Then Exit Sub
'    If Col = 5 Then
'        VSFG.Cell(flexcpForeColor, Row, 5) = &H80&
'    End If
'    If Col = 16 Or Col = 14 Then
'        FechaEsteMes Row
'    End If
'    If Col = 14 Then
'        If VSFG.TextMatrix(Row, 14) <> Me.lblDias.Caption Then
'            VSFG.Cell(flexcpForeColor, Row, 14) = RGB(190, 0, 0)
'        End If
'    End If
'    If HacerChange = False Then Exit Sub
'    With VSFG
'        If .TextMatrix(Row, Col) <> "" Then
'            If Col = 3 Then
'                clsSocio.Filtrar ("epl_codigo = '" & .TextMatrix(Row, 3) & "'")
'                    .TextMatrix(Row, 4) = clsSocio.adorec_Def("nombre")
'                    .TextMatrix(Row, 8) = clsSocio.adorec_Def("epl_sueldo")
'                    .TextMatrix(Row, 13) = clsSocio.adorec_Def("epl_fec_ingreso")
'                    If IsNull(clsSocio.adorec_Def("epl_fec_salida")) = False Then
'                        VSFG.TextMatrix(Row, 16) = clsSocio.adorec_Def("epl_fec_salida")
'                    Else
'                        VSFG.TextMatrix(Row, 16) = ""
'                    End If
'                    .TextMatrix(Row, 14) = DiasFinDeMes(CStr(Fecha2), .TextMatrix(Row, 13), .TextMatrix(Row, 16))
'                    .TextMatrix(Row, 15) = DiasFondo(CStr(Fecha2), .TextMatrix(Row, 13), .TextMatrix(Row, 16))
'                   FechaEsteMes Row
'                clsSocio.QuitarFiltro
'            End If
'
'            If Col = 4 Then
'                clsSocio.Filtrar ("epl_codigo = '" & .TextMatrix(Row, 4) & "'")
'                    .TextMatrix(Row, 3) = clsSocio.adorec_Def("epl_codigo")
'                    .TextMatrix(Row, 8) = clsSocio.adorec_Def("epl_sueldo")
'                    .TextMatrix(Row, 13) = clsSocio.adorec_Def("epl_fec_ingreso")
'                    If IsNull(clsSocio.adorec_Def("epl_fec_salida")) = False Then
'                        VSFG.TextMatrix(Row, 16) = clsSocio.adorec_Def("epl_fec_salida")
'                    Else
'                        VSFG.TextMatrix(Row, 16) = ""
'                    End If
'                    .TextMatrix(Row, 14) = DiasFinDeMes(CStr(Fecha2), .TextMatrix(Row, 13), .TextMatrix(Row, 16))
'                    .TextMatrix(Row, 15) = DiasFondo(CStr(Fecha2), .TextMatrix(Row, 13), .TextMatrix(Row, 16))
'
'                    FechaEsteMes Row
'                clsSocio.QuitarFiltro
'            End If
'
'            If Col = 3 Or Col = 4 Or Col = 5 Then
'                'If Col <> 5 And FormatoD(VSFG.TextMatrix(Row, 5)) = 0 Then
'                If Col <> 5 Then
'                    Dim CadenaEval As String
'                    HacerReglaDe3 = True
'                    'Capital
'                    CadenaEval = Replace(Factor, "SueldoBas", .TextMatrix(Row, 8))
'                    'Si es que se calcula en función de provisiones
'                    If Check1(4).Value = 1 And CadenaEval <> "0" Then
'                        HacerReglaDe3 = False
'                        CadenaEval = SumarProvisionesPendientes(Me.txtProvision.Tag, .TextMatrix(Row, 3), Me.dcmbTipo.BoundText)
'                    Else
'                        If VSFG.ColHidden(6) = False Then
'                            HacerReglaDe3 = False
'                            If VSFG.TextMatrix(0, 6) = "Sueldo Mes" Then
'                                VSFG.TextMatrix(Row, 6) = FormatoD(SueldoMes(.TextMatrix(Row, 3), Fecha1, Fecha2))
'                                CadenaEval = Replace(CadenaEval, "SueldoMes", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                            ElseIf VSFG.TextMatrix(0, 6) = "Sueldo Año" Then
'                                VSFG.TextMatrix(Row, 6) = FormatoD(SueldoAño(.TextMatrix(Row, 3), Fecha1, Fecha2))
'                                CadenaEval = Replace(CadenaEval, "SueldoAño", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                            ElseIf VSFG.TextMatrix(0, 6) = "Renta Mes" Then
'                                VSFG.TextMatrix(Row, 6) = FormatoD(RentaMes(.TextMatrix(Row, 3), Fecha1, Fecha2))
'                                CadenaEval = Replace(CadenaEval, "ImpRentaMes", Format(ImpuestoRentaMes(VSFG.TextMatrix(Row, 6)), "#0.00"))
'                            ElseIf VSFG.TextMatrix(0, 6) = "Sueldo IESS" Then
'                                VSFG.TextMatrix(Row, 6) = FormatoD(SueldoIESS(.TextMatrix(Row, 3), Fecha1, Fecha2))
'                                CadenaEval = Replace(CadenaEval, "SueldoIESS", Format(VSFG.TextMatrix(Row, 6), "#0.00"))
'                            End If
'                        End If
'                    End If
'
'                    'CadenaEval = Replace(Factor, "SueldoBas", .TextMatrix(Row, 8))
'                    If Trim(CadenaEval) <> "" Then
'                        'Se evaluará la expresión con la base de datos
'                        strSql = " SELECT " & CadenaEval
'                        clsSql1.Ejecutar (strSql)
'                        ElCapital = Formato(clsSql1.adorec_Def(0))
'                        'Sacar el proporcional según días del mes trabajados
'                        If HacerReglaDe3 = True Then
'                            ElCapital = FormatoD(ElCapital * CInt(VSFG.TextMatrix(Row, 14)) / CInt(lblDias))
'                        End If
'                        'Si es fondo de cesantía calcular con la columna 15
'                        If Me.dcmbTipo.BoundText = "1003" Then
'                            ElCapital = FormatoD(ElCapital * CInt(VSFG.TextMatrix(Row, 15)) / CInt(lblDias))
'                        End If
'                    Else
'                        ElCapital = 0
'                    End If
'                    VSFG.TextMatrix(Row, 5) = ElCapital
'                    VSFG.Cell(flexcpForeColor, Row, 5) = &H80&
'                End If
'                ElCapital = VSFG.TextMatrix(Row, 5)
''                'Interés
''                CadenaEval = Replace(FactorInteres, "SueldoBas", .TextMatrix(Row, 8))
''                CadenaEval = Replace(CadenaEval, "Capital", Formato(ElCapital))
''                If Trim(CadenaEval) <> "" Then
''                    'Se evaluará la expresión con la base de datos
''                    strSql = " SELECT " & CadenaEval
''                    clsSql1.Ejecutar (strSql)
''                    ElInteres = Formato(clsSql1.adorec_Def(0))
''                Else
''                    ElInteres = 0
''                End If
''                VSFG.TextMatrix(Row, 6) = ElInteres
''                VSFG.TextMatrix(Row, 7) = FormatoD(VSFG.TextMatrix(Row, 6)) + FormatoD(VSFG.TextMatrix(Row, 5))
''                VSFG.Cell(flexcpForeColor, Row, 7) = &H80&
'            End If
'         End If
'    End With
'End Sub
'
'Private Sub VSFG_DblClick()
'    If VSFG.Row > 0 And VSFG.Col = 5 Then
'        If Me.dcmbTipo.BoundText = 2 And VSFG.Editable <> flexEDNone And Trim(VSFG.TextMatrix(VSFG.Row, 2)) <> "" And VSFG.IsSubtotal(VSFG.Row) = False And Val(VSFG.TextMatrix(VSFG.Row, 10)) = 0 Then
'            frmHorasExtras.VSFG.TextMatrix(1, 1) = FormatoD(VSFG.TextMatrix(VSFG.Row, 11))
'            frmHorasExtras.VSFG.TextMatrix(2, 1) = FormatoD(VSFG.TextMatrix(VSFG.Row, 12))
'            frmHorasExtras.SueldoBasico = FormatoD(VSFG.TextMatrix(VSFG.Row, 8))
'            frmHorasExtras.Show
'        End If
'    End If
'End Sub
'
'Private Sub VSFG_EnterCell()
'    If VSFG.Col = 13 Or VSFG.Col = 14 Or VSFG.Col = 15 Then
'        VSFG.ToolTipText = "Fecha en Rojo: empleado entró en este mes. " & vbNewLine & _
'            "Fecha en Verde: empleado cumple un año."
'    Else
'        VSFG.ToolTipText = ""
'    End If
'End Sub
'
'Private Sub VSFG_KeyPress(KeyAscii As Integer)
'    'Que muestre la pantalla de horas cuando se da clic
'    If VSFG.Col = 5 And Me.dcmbTipo.BoundText = "2" Then
'        VSFG_DblClick
'    End If
'End Sub
'
'Private Sub VSFG_KeyUp(KeyCode As Integer, Shift As Integer)
'    If vbCtrlMask > 0 Then
'        If KeyCode = 17 Then
'            CopiarFlexGrid VSFG
'            MsgBox "Se copió la selección al portapapeles.", vbInformation, "Información"
'        End If
'    End If
'End Sub
'
'Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 1 Then Exit Sub
'    If Me.cmdEditar.Enabled = True Then Exit Sub
'
'    Row = VSFG.MouseRow
'    Col = VSFG.MouseCol
'    If Row < 0 Then Exit Sub
'    If Val(VSFG.TextMatrix(Row, 10)) <> 0 Then Exit Sub
'    If Col = 1 Then
'        If VSFG.Cell(flexcpPicture, Row, Col) Is Nothing Then Exit Sub
'        VSFG.Cell(flexcpPicture, Row, Col) = Me.imgBtnDn
'        If MsgBox("¿Está seguro de eliminar este descuento?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then GoTo Final
'        'Borrar
'        EliminarDescuento clsSql1, VSFG.TextMatrix(Row, 2)
'        VSFG.RemoveItem Row
'        NumerarVSFG
'        SumarVSFG
'        'VSFG.Cell(flexcpPicture, Row, Col) = Me.imgBtnUp
'    End If
'    Exit Sub
'Final:
'    VSFG.Cell(flexcpPicture, Row, Col) = Me.imgBtnUp
'End Sub
'
'Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    'Para evitar que ingresen dos veces el mismo socio
'    If Col = 3 Or Col = 4 Then
'        If Trim(VSFG.EditText) <> "" Then
'            Screen.MousePointer = vbHourglass
'            For i = 1 To VSFG.Rows - 1
'                If VSFG.Cell(flexcpTextDisplay, i, Col) = VSFG.EditText And i <> Row Then
'                    'Encontro = True
'                    Cancel = True
'                    Screen.MousePointer = vbDefault
'                    MsgBox "Ya se ingresó antes el socio " & VSFG.Cell(flexcpTextDisplay, i, 4) & " (" & VSFG.TextMatrix(i, 3) & ")" & vbNewLine & "No se lo puede ingresar de nuevo.", vbInformation, "Inforación"
'                    Exit For
'                End If
'            Next i
'            Screen.MousePointer = vbDefault
'        End If
'    End If
'End Sub
