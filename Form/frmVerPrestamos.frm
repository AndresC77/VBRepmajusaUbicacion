VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmVerPrestamos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Préstamos y Anticipos"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerPrestamos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9750
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4133
      TabIndex        =   6
      Top             =   6480
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbMesI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmVerPrestamos.frx":030A
         Left            =   4080
         List            =   "frmVerPrestamos.frx":0335
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   315
         Left            =   240
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
         Left            =   5520
         TabIndex        =   2
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
         Format          =   49872899
         UpDown          =   -1  'True
         CurrentDate     =   38054
      End
      Begin MSDataListLib.DataCombo dcmbEmpleado 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   5520
         _ExtentX        =   9737
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
         Caption         =   "Empleado"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1100
         Width           =   5535
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipos de Ingresos"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   380
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes y Año"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   380
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Préstamos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   9495
      Begin MSDataListLib.DataCombo dcmbPrestamo 
         Height          =   1350
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   2381
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
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
      Begin VSFlex7Ctl.VSFlexGrid VSFG 
         Height          =   1815
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   6255
         _cx             =   11033
         _cy             =   3201
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerPrestamos.frx":039E
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
         TabBehavior     =   0
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
      End
      Begin VB.Label lblVSFG 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 Cuotas"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2060
         Width           =   6255
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleado / Asiento / Fecha / Monto"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   380
         Width           =   9015
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   5693
      TabIndex        =   7
      Top             =   6480
      Width           =   1485
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   400
      Left            =   2573
      TabIndex        =   5
      Top             =   6480
      Width           =   1485
   End
End
Attribute VB_Name = "frmVerPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSql As String
Private clsSql As New clsConsulta
Private Fecha1 As Variant
Private Fecha2 As Variant
Private CuentaContable As String

Private Sub Año_Change()
    CambiarFecha
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
           
    Fecha1 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-1"
    Fecha2 = ""
    DiaFinal = 31
    While (IsDate(Fecha2) = False)
        Fecha2 = Me.Año.Year & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
    Restaurar
End Sub

Private Sub Restaurar()
    Frame1.Visible = False
    cmdBuscar.Enabled = True
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub cmdBuscar_Click()
    BuscarPrestamos
End Sub

Public Sub BuscarPrestamos()
    Set Me.dcmbPrestamo.RowSource = Nothing
    dcmbPrestamo = ""
    dcmbPrestamo.BoundText = ""
    
    strSql = " SELECT descuento.asi_numasiento, concat(epl_apellidos,' ',epl_nombres,' ',descuento.asi_numasiento,' ',LEFT(CONVERT(VARCHAR,asi_fecha,20),10),' ',ISNULL(CAST(det_asi_debe AS VARCHAR),CASE ISNULL(com_egr_ch_num,-1) WHEN -1 THEN ' NOTA BANCOS' WHEN 0 THEN ' TRANSFERENCIA' ELSE ' CHEQUE ',CAST(com_egr_ch_num AS VARCHAR) END,'')) AS Info " & _
             " FROM descuento INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
             " INNER JOIN asiento ON descuento.asi_numasiento=asiento.asi_numasiento AND descuento.emp_codigo=asiento.emp_codigo" & _
             " LEFT JOIN det_asiento ON det_asiento.asi_numasiento=asiento.asi_numasiento AND det_asiento.emp_codigo=asiento.emp_codigo AND det_asi_haber=0 AND det_asiento.cta_codigo = '" & CuentaContable & "'" & _
             " LEFT JOIN comp_egreso ON asiento.asi_numasiento=comp_egreso.com_egr_numasiento AND det_asiento.emp_codigo=comp_egreso.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND descuento.epl_codigo LIKE '" & Me.dcmbEmpleado.BoundText & "'" & _
             " AND tip_des_codigo='" & Me.dcmbTipo.BoundText & "' GROUP BY descuento.asi_numasiento, epl_apellidos, epl_nombres,asi_fecha,det_asi_debe, com_egr_ch_num ORDER BY descuento.asi_numasiento DESC"
    clsSql.Ejecutar (strSql)
    Set Me.dcmbPrestamo.RowSource = clsSql.adorec_Def.DataSource
    dcmbPrestamo.ListField = "Info"
    dcmbPrestamo.BoundColumn = "asi_numasiento"
    If clsSql.adorec_Def.RecordCount > 0 Then
        If clsSql.adorec_Def.RecordCount = 1 Then
            Me.lblPrestamo = "1 Registro"
        Else
            Me.lblPrestamo = clsSql.adorec_Def.RecordCount & " Registros"
        End If
        dcmbPrestamo.BoundText = clsSql.adorec_Def("asi_numasiento")
        Me.cmdEliminar.Enabled = True
    Else
        Me.lblPrestamo = "Ningún Registro"
        Me.cmdEliminar.Enabled = False
    End If
    Me.lblPrestamo = Me.lblPrestamo & " - Empleado / Asiento / Fecha / Monto"
    
    Frame1.Visible = True
    cmdBuscar.Enabled = False
End Sub

Private Sub cmdEliminar_Click()
    'Verificar fecha de cierre contable
    'If VerificarFechaContable(Me.Año) = False Then Exit Sub
    
    'Verificar que el asiento no esté mayorizado
    strSql = " SELECT asi_numasiento FROM asiento WHERE asi_numasiento='" & Me.dcmbPrestamo.BoundText & "' AND emp_codigo='" & strEmpresa & "' AND asi_mayorizado=1"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        MsgBox "No puede eliminar este asiento. El asiento está mayorizado.", vbCritical, "Información"
        Exit Sub
    End If
    
    'Buscar si ya se ha contabilizado algun registro de préstamo
    strSql = " SELECT asi_numasiento FROM descuento WHERE asi_numasiento='" & Me.dcmbPrestamo.BoundText & "' AND emp_codigo='" & strEmpresa & "' AND des_pagado=1"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        MsgBox "No puede eliminar este asiento. El préstamo relacionado ya tiene pagos.", vbCritical, "Información"
        Exit Sub
    End If
    'Borrar registros de préstamos si existen
    strSql = " SELECT COUNT(asi_numasiento) FROM descuento WHERE asi_numasiento='" & dcmbPrestamo.BoundText & "' AND emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        If clsSql.adorec_Def(0) = 1 Then
            NumeroDocumentos = vbNewLine & "Se eliminará 1 registro de préstamo o anticipo"
        ElseIf clsSql.adorec_Def(0) = 0 Then
            NumeroDocumentos = ""
        Else
            NumeroDocumentos = vbNewLine & "Se eliminarán " & clsSql.adorec_Def(0) & " registros de préstamos o anticipos"
        End If
    Else
        NumeroDocumentos = ""
    End If
    If MsgBox("¿Está seguro de eliminar el asiento de préstamos número " & dcmbPrestamo.BoundText & "?" & NumeroDocumentos, vbQuestion + vbYesNo + vbDefaultButton2, "Pregunta - Eliminar") = vbNo Then Exit Sub
    If NumeroDocumentos <> "" Then
        strSql = " DELETE FROM descuento WHERE asi_numasiento='" & dcmbPrestamo.BoundText & "' AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar (strSql)
    End If
    strSql = " DELETE FROM det_asiento WHERE asi_numasiento ='" & dcmbPrestamo.BoundText & "' and emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    
    strSql = " DELETE FROM asiento WHERE asi_numasiento ='" & dcmbPrestamo.BoundText & "' and emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    MsgBox "Asiento " & dcmbPrestamo.BoundText & " eliminado.", vbInformation, "Eliminar"
    BuscarPrestamos
End Sub

Private Sub cmdnuevo_Click()
    frmPrestamos.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbEmpleado_Change()
    Restaurar
End Sub

Private Sub dcmbPrestamo_Change()
    VSFG.Rows = 1
    lblVSFG.Caption = "0 Cuotas"
    VSFG.Clear 1
    If Trim(dcmbPrestamo) = "" Then
        Me.cmdEliminar.Enabled = False
        Exit Sub
    End If
    strSql = " SELECT '', des_fecha, des_valor, des_pagado" & _
             " FROM descuento " & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha >= '" & Fecha1 & "' AND asi_numasiento='" & Me.dcmbPrestamo.BoundText & "'" & _
             " AND tip_des_codigo='" & Me.dcmbTipo.BoundText & "' ORDER BY des_fecha"
    clsSql.Ejecutar (strSql)
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    If VSFG.Rows = 2 Then
        lblVSFG.Caption = "1 Cuota"
    Else
        lblVSFG.Caption = CStr(VSFG.Rows - 1) & " Cuotas"
    End If
    Dim i As Integer
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
        VSFG.TextMatrix(i, 1) = FechaLetras(VSFG.TextMatrix(i, 2))
    Next i
    VSFG.SubtotalPosition = flexSTBelow
    VSFG.SubTotal flexSTSum, -1, 3, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
    Me.cmdEliminar.Enabled = True
End Sub

Private Function FechaLetras(Fecha As String) As String
    Dim Año1 As String
    Dim Mes1 As String
    Año1 = Left(Fecha, 4)
    Mes1 = Mid(Fecha, 6, 2)
    FechaLetras = Me.cmbMesI.List(CInt(Mes1) - 1) & " " & Año1
End Function

Private Sub dcmbTipo_Change()
    Restaurar
    strSql = " SELECT cta_codigo  " & _
             " FROM tipo_descuento " & _
             " WHERE tip_des_codigo='" & Me.dcmbTipo.BoundText & "' AND emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar (strSql)
    CuentaContable = clsSql.adorec_Def("cta_codigo")
End Sub

Private Sub Form_Activate()
    Me.cmdBuscar.Enabled = True
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    Año = Date
    'Selecciona el mes actual
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(Date)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    strSql = " SELECT tip_des_codigo, tip_des_nombre FROM tipo_descuento" & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_des_prestamo=1" & _
             " ORDER BY tip_des_orden"
    clsSql.Ejecutar (strSql)
    Set Me.dcmbTipo.RowSource = clsSql.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_des_nombre"
    dcmbTipo.BoundColumn = "tip_des_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbTipo.BoundText = clsSql.adorec_Def("tip_des_codigo")
    End If
    
    strSql = " SELECT '%' AS codigo, ' --Todos los empleados--' AS nombre UNION " & _
             " SELECT epl_codigo AS codigo, concat(epl_apellidos,' ',epl_nombres) AS nombre FROM empleado" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY nombre"
    clsSql.Ejecutar (strSql)
    Set Me.dcmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    dcmbEmpleado.ListField = "nombre"
    dcmbEmpleado.BoundColumn = "codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbEmpleado.BoundText = clsSql.adorec_Def("codigo")
    End If
End Sub
