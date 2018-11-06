VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerDepreciacionAF 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depreciación de Activos Fijos"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerDepreciacionAF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10575
   Begin VB.Frame Frame5 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Asiento de Depreciación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   10335
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   10095
         _cx             =   17806
         _cy             =   2990
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerDepreciacionAF.frx":030A
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.TextBox TxtTotalDebe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2040
         Width           =   1545
      End
      Begin VB.TextBox TxtTotalHaber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2040
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suma total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5280
         TabIndex        =   9
         Top             =   2085
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Depreciación de Activos Fijos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10335
      Begin VSFlex8Ctl.VSFlexGrid VSFDep 
         Height          =   2655
         Left            =   225
         TabIndex        =   15
         Top             =   1320
         Width           =   9975
         _cx             =   17595
         _cy             =   4683
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerDepreciacionAF.frx":03EE
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton cmdRevalorizar 
         Caption         =   "&Revalorizar"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   600
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
         ItemData        =   "frmVerDepreciacionAF.frx":06F7
         Left            =   240
         List            =   "frmVerDepreciacionAF.frx":0722
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   1425
      End
      Begin VB.ComboBox cmbAñoI 
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
         ItemData        =   "frmVerDepreciacionAF.frx":078B
         Left            =   1680
         List            =   "frmVerDepreciacionAF.frx":078D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   780
      End
      Begin VB.CommandButton cmbCalcular 
         Caption         =   "&Depreciar"
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   578
         Width           =   1455
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes a depreciar"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Detalle"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   9965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Depreciación del Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5190
         TabIndex        =   5
         Top             =   4140
         Visible         =   0   'False
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmVerDepreciacionAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de Depreciacion de Activos Fijos con los que trabajará la mpresa.     #
'#  frmDepreciacion_af V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la Depreciacion de los Activo Fijos.                           #
'#  Permitirá Depreciar en la base de datos Activo Fijos                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     Activo_Fijo: En esta tabla se almacenan los nuevos Activo Fijos y se     #
'#               modifican los datos de estos.                                  #
'#     Tipo_activo_fijo : En esta tabla se sacan el # de cuenta que se pueden   #
'#               asignar a los diferentes Activo Fijos.                         #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#               clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Act As New clsConsulta
Private clsCon_Tip_Asi As New clsConsulta
Private clsCon_Fech As New clsConsulta
Private clsCon_Update As New clsConsulta
Private clsCon_Max As New clsConsulta
Private clsCon_Insert As New clsConsulta
Private clsCon_Insert_Det As New clsConsulta
Private clsCon_Inner As New clsConsulta
Dim strSql As String
Dim t As Double
Dim dd As Double
Dim dp As Double
Dim ban As Integer
Dim fi As Variant
Dim ff As Variant
'La cuenta que juega contra las revalorizaciones
Private CuentaPatrimonio As String
Private NombreCuentaPatrimonio As String

Private Fecha1 As Variant
Private Fecha2 As Variant
Private HacerFecha As Boolean
Private estado As String

Private Sub cmbAceptar_Click()
    'If VerificarFechaContable(Fecha2) = False Then Exit Sub
    Dim strSql As String
    Dim fa As String
    Dim fh As String
    Dim numAsi As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim h As Long
    Dim mat1() As Variant
    Dim mat2() As Variant
    Dim mat3() As Variant
    Dim maxMat1 As Long
    Dim maxMat2 As Long
    Dim maxMat3 As Long
    
    ' Si todos los campos estan llenos
    fa = Fecha2 'último día del mes
    fh = Format(fa, "yyyymm")
    If TxtTotal = "" Then
        MsgBox "No ha realizado la depreciación." & vbCrLf & "Verifíquelo, por favor.", vbInformation, "Información"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    clsCon_Max.Inicializar AdoConn, AdoConnMaster
    clsCon_Insert.Inicializar AdoConn, AdoConnMaster
    clsCon_Update.Inicializar AdoConn, AdoConnMaster
    clsCon_Inner.Inicializar AdoConn, AdoConnMaster

    On Error GoTo errhandler
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    
'    'Grabar asiento
'    clsAsiento.NuevoAsiento "", 0, CStr(Fecha), 0, 0, Format(TxtTotal1Debe, "#0.00"), UCase(txtDescripcion)
'    strMaximo = clsAsiento.NumAsiento
'    'strMaximo = AsientoNuevo(CStr(Fecha), "RRH", 0, 0, Format(TxtTotal1Debe, "##0.00"), UCase(txtDescripcion.Text))
'    With VSFG
'        For i = 1 To .Rows - 1
'            clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), Val(Format(.TextMatrix(i, 3), "##0.00")), Val(Format(.TextMatrix(i, 4), "##0.00"))
'    '        NuevoDetAsiento strMaximo, .TextMatrix(i, 1), Val(Format(.TextMatrix(i, 3), "##0.00")), Val(Format(.TextMatrix(i, 4), "##0.00")), .TextMatrix(i, 5)
'        Next i
'    End With
    
    If estado = "DEPRECIAR" Then
        clsAsiento.NuevoAsiento "D", CStr(fa), 0, 0, CDbl(txtTotalDebe.Text), UCase("DEPRECIACION DEL PERIODO DESDE " & fi & " HASTA " & ff)
    Else
        clsAsiento.NuevoAsiento "D", CStr(fa), 0, 0, CDbl(txtTotalDebe.Text), UCase("REVALORIZACIÓN DEL PERIODO DESDE " & fi & " HASTA " & ff)
    End If
    numAsi = clsAsiento.NumAsiento
            
    'Actualiza los valores de depreciacion de los activos fijos del grid en la tabla activos fijos
    For i = 1 To (VSFDep.Rows - 2) '-2 porque hay fila de totales
        If estado = "DEPRECIAR" Then
            strSql = " UPDATE activo_fijo SET act_fij_depreciado = (act_fij_depreciado + '" & Format(Val(VSFDep.TextMatrix(i, 7)), "##0.00") & "'), act_fij_fecha_dep='" & Fecha2 & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND act_fij_codigo = '" & VSFDep.TextMatrix(i, 1) & "' "
            clsCon_Update.Ejecutar strSql, "M"
            NuevoDetalleActivo VSFDep.TextMatrix(i, 1), "DEP", numAsi, (Me.cmbMesI.ListIndex + 1), Me.cmbAñoI.List(Me.cmbAñoI.ListIndex), VSFDep.TextMatrix(i, 7)
            ' Si la columna 14 es diferente de cero grabar
            If FormatoD2(VSFDep.TextMatrix(i, 14)) <> 0 Then
                strSql = " UPDATE activo_fijo SET act_fij_depreciado2 = (act_fij_depreciado2 + '" & Format(Val(VSFDep.TextMatrix(i, 14)), "##0.00") & "') " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND act_fij_codigo = '" & VSFDep.TextMatrix(i, 1) & "' "
                clsCon_Update.Ejecutar strSql, "M"
                NuevoDetalleActivo VSFDep.TextMatrix(i, 1), "DE2", numAsi, (Me.cmbMesI.ListIndex + 1), Me.cmbAñoI.List(Me.cmbAñoI.ListIndex), VSFDep.TextMatrix(i, 14)
            End If
        Else
            strSql = " UPDATE activo_fijo SET act_fij_revalorizado = (act_fij_revalorizado + '" & Format(Val(VSFDep.TextMatrix(i, 12)), "##0.00") & "'), act_fij_fecha_dep='" & Fecha2 & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND act_fij_codigo = '" & VSFDep.TextMatrix(i, 1) & "' "
            clsCon_Update.Ejecutar strSql, "M"
            NuevoDetalleActivo VSFDep.TextMatrix(i, 1), "REV", numAsi, (Me.cmbMesI.ListIndex + 1), Me.cmbAñoI.List(Me.cmbAñoI.ListIndex), VSFDep.TextMatrix(i, 12)
        End If
    Next i
        
    For h = 1 To VSFG.Rows - 1
        clsAsiento.NuevoDetAsiento VSFG.TextMatrix(h, 0), VSFG.TextMatrix(h, 4), FormatoD2(VSFG.TextMatrix(h, 2)), FormatoD2(VSFG.TextMatrix(h, 3))
    Next h
    Screen.MousePointer = vbDefault
    MsgBox " Los datos han sido ingresados ", vbInformation, "Depreciación Activos Fijos"
            
    Unload Me
    Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbAñoI_Click()
    CambiarFecha
    VSFDep.Clear 1
    VSFDep.Rows = 1
    VSFG.Clear 1
    VSFG.Rows = 1
    txtTotalHaber.Text = "0.00"
    txtTotalDebe.Text = "0.00"
     TxtTotal.Text = ""
End Sub

Private Sub cmbCalcular_Click()
    Screen.MousePointer = vbHourglass
    Me.VSFG.Rows = 1
    Me.txtTotalDebe = "0.00"
    Me.txtTotalHaber = "0.00"
    Me.TxtTotal = "0.00"
    BuscarActivos True
    fi = Format(Fecha1, "yyyy-mm-dd")
    ff = Format(Fecha2, "yyyy-mm-dd")
    If ban = 1 Then
        Call Calculo_Depreciacion
        Call Calculo_Por_Depreciar
        Call Calculo_E_DepPorPeriodo
        TxtTotal.Text = Format(dp, "##0.00")
        If FormatoD2(TxtTotal) > 0 Then
            cmbAceptar.Enabled = True
        Else
            cmbAceptar.Enabled = False
        End If
    Else
        cmbAceptar.Enabled = False
    End If
    PonerTotales
    PonerTotalesCuentas
    estado = "DEPRECIAR"
    Frame5.Caption = "Asiento de Depreciación"
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerTotales()
    VSFDep.SubtotalPosition = flexSTBelow
    VSFDep.Subtotal flexSTSum, -1, 5, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Totales"
    VSFDep.Subtotal flexSTSum, -1, 6, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Totales"
    VSFDep.Subtotal flexSTSum, -1, 7, "#,##0.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Totales"
    
    VSFDep.Subtotal flexSTSum, -1, 11, "#,##0.00"
    VSFDep.Subtotal flexSTSum, -1, 12, "#,##0.00"
    VSFDep.Subtotal flexSTSum, -1, 13, "#,##0.00"
    VSFDep.Subtotal flexSTSum, -1, 14, "#,##0.00"
    VSFDep.Subtotal flexSTSum, -1, 15, "#,##0.00"
    
End Sub

Private Sub BuscarActivos(Depreciacion As Boolean)
    Dim strWhere As String
    If Depreciacion = True Then
        strWhere = " AND det_act_fij_tipo<>'REV'"
        'cmbCalcular.Enabled = False
    Else
        strWhere = " AND det_act_fij_tipo='REV'"
        'cmdRevalorizar.Enabled = False
    End If
    'Consulta los Activos Fijos que estan disponibles para el grid
    strSql = " SELECT activo_fijo.act_fij_codigo,act_fij_nombre,act_fij_Vida_Util,act_fij_fecha_adq,act_fij_valor,act_fij_depreciado , 0 as valor1, 0 as valor2,tip_act_ctaconta2,c2.cta_nombre,  " & _
             " IFNULL(act_fij_revalorizado,0),0,IFNULL(act_fij_depreciado2,0),0,0, tip_act_ctaconta3,c3.cta_nombre, tip_act_ctaconta4,c4.cta_nombre, ''" & _
             " FROM activo_fijo INNER JOIN tipo_activo ON activo_fijo.tip_act_codigo =tipo_activo.tip_act_codigo AND activo_fijo.emp_codigo =tipo_activo.emp_codigo " & _
             " INNER JOIN ctaconta c2 ON tipo_activo.emp_codigo=c2.emp_codigo AND tipo_activo.tip_act_ctaconta2=c2.cta_codigo" & _
             " INNER JOIN ctaconta c3 ON tipo_activo.emp_codigo=c3.emp_codigo AND tipo_activo.tip_act_ctaconta3=c3.cta_codigo" & _
             " INNER JOIN ctaconta c4 ON tipo_activo.emp_codigo=c4.emp_codigo AND tipo_activo.tip_act_ctaconta4=c4.cta_codigo" & _
             " LEFT JOIN det_activo_fijo ON activo_fijo.act_fij_codigo=det_activo_fijo.act_fij_codigo AND activo_fijo.emp_codigo=det_activo_fijo.emp_codigo AND det_activo_fijo.det_act_fij_mes='" & (Me.cmbMesI.ListIndex + 1) & "' AND det_activo_fijo.det_act_fij_año='" & Me.cmbAñoI.List(Me.cmbAñoI.ListIndex) & "'" & strWhere & _
             " WHERE activo_fijo.emp_codigo='" & strEmpresa & "' AND det_activo_fijo.act_fij_codigo IS NULL" & _
             " AND act_fij_baja = 0 AND act_fij_fecha_adq<='" & Fecha2 & "'" & _
             " ORDER BY activo_fijo.act_fij_codigo "
    clsCon_Act.Ejecutar (strSql)
    If (clsCon_Act.adorec_Def.RecordCount > 0) Then
        t = 1
        Set VSFDep.DataSource = clsCon_Act.adorec_Def.DataSource
        Call PonerNumeros
        t = 0
        ban = 1
    Else
        VSFDep.Rows = 1
        'MsgBox "Todos los Activos Fijos están Depreciados o No Existen Activos Fijos!", vbExclamation, "Depreciacion Activos Fijos"
        ban = 0
        cmbAceptar.Enabled = False
    End If
    'Sacar el total para revalorizar, sumando las revalorizaciones y restando depreciaciones
    For i = 1 To VSFDep.Rows - 1
        'Fila 15 = 5-6+11-13
        VSFDep.TextMatrix(i, 15) = FormatoD2(VSFDep.TextMatrix(i, 5)) - FormatoD2(VSFDep.TextMatrix(i, 6)) + FormatoD2(VSFDep.TextMatrix(i, 11)) - FormatoD2(VSFDep.TextMatrix(i, 13))
        'Calcular la fecha máxima de vida del activo
        VSFDep.TextMatrix(i, 20) = DateAdd("yyyy", VSFDep.TextMatrix(i, 3), VSFDep.TextMatrix(i, 4))
    Next i
    
End Sub

Private Sub CambiarFecha()
    If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    'Me.Label5.Caption = "Asientos costo promedio del mes de " & cmbMesI.List(cmbMesI.ListIndex)
    'Me.Label6.Caption = "Asientos costo última compra del mes de " & cmbMesI.List(cmbMesI.ListIndex)
        
    Fecha1 = cmbAñoI & "-" & cmbMesI.ListIndex + 1 & "-1"
    Fecha2 = ""
    DiaFinal = 31
    While (IsDate(Fecha2) = False)
        Fecha2 = cmbAñoI & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
    Me.cmbCalcular.Enabled = True
    Me.cmdRevalorizar.Enabled = True
    'MostrarAsientos
    'cmdCargar.Enabled = True
    'cmdGuardarA(0).Enabled = False
    'cmdGuardarA(1).Enabled = False
    'cmdVerificar.Enabled = False
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
    VSFDep.Clear 1
    VSFDep.Rows = 1
    VSFG.Clear 1
    VSFG.Rows = 1
    txtTotalHaber.Text = "0.00"
    txtTotalDebe.Text = "0.00"
    TxtTotal.Text = ""
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdRevalorizar_Click()
    Dim strPorcentaje As String
    Dim Porcentaje As Double
    strPorcentaje = InputBox("Ingrese el porcentaje para revalorizar los activos fijos:", "Porcentaje")
    If strPorcentaje = "" Then Exit Sub 'botón cancelar
    Porcentaje = FormatoD2(strPorcentaje)
    If Porcentaje = 0 Then
        MsgBox "No se puede revalorizar por 0%", vbInformation, "Información"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de revalorizar en un " & Porcentaje & "% todos los activos fijos?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub
    
    Me.VSFG.Rows = 1
    Me.txtTotalDebe = "0.00"
    Me.txtTotalHaber = "0.00"
    Me.TxtTotal = "0.00"
    BuscarActivos False
    fi = Fecha1
    ff = Fecha2
    
    'Hacer revalorización
    For i = 1 To VSFDep.Rows - 1
        'Si es ya no se pasó de la fecha de fin de vida del activo
        If DateDiff("d", fi, VSFDep.TextMatrix(i, 20)) > 0 Then
            VSFDep.TextMatrix(i, 12) = FormatoD2(FormatoD2(VSFDep.TextMatrix(i, 15)) * FormatoD2(Porcentaje) / 100)
            'Añadir cuentas
            AñadirCuenta VSFDep.TextMatrix(i, 16), VSFDep.TextMatrix(i, 17), FormatoD2(VSFDep.TextMatrix(i, 12)), 0
            'Cuenta de patrimonio
            AñadirCuenta CuentaPatrimonio, NombreCuentaPatrimonio, 0, FormatoD2(VSFDep.TextMatrix(i, 12))
        End If
    Next i
    
    estado = "REVALORIZAR"
    Frame5.Caption = "Asiento de Revalorización"
    PonerTotales
    PonerTotalesCuentas
    cmbAceptar.Enabled = True
    'MsgBox "¿Está seguro de revalorizar en un " & FormatoD(Porcentaje) & "% los activos fijos?", vbInformation, "Información"
End Sub

Private Sub AñadirCuenta(CUENTA As String, Nombre As String, DEBE As Double, HABER As Double, Optional CentroCostoCodigo As String = "", Optional CentroCostoNombre As String = "")
    If FormatoD2(DEBE) = 0 And FormatoD2(HABER) = 0 Then Exit Sub
    For j = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(j, 0) = CUENTA And VSFG.TextMatrix(j, 4) = CentroCostoCodigo Then Exit For
    Next j
    If j < VSFG.Rows Then
        VSFG.TextMatrix(j, 2) = FormatoD2(VSFG.TextMatrix(j, 2)) + DEBE
        VSFG.TextMatrix(j, 3) = FormatoD2(VSFG.TextMatrix(j, 3)) + HABER
    Else
        VSFG.AddItem CUENTA & vbTab & Nombre & vbTab & DEBE & vbTab & HABER & vbTab & CentroCostoCodigo & vbTab & CentroCostoNombre
    End If
End Sub

Private Sub PonerTotalesCuentas()
    Dim ElDebe As Double
    Dim ElHaber As Double
    ElDebe = 0
    ElHaber = 0
    For i = 1 To VSFG.Rows - 1
        ElDebe = ElDebe + VSFG.TextMatrix(i, 2)
        ElHaber = ElHaber + VSFG.TextMatrix(i, 3)
    Next i
    txtTotalDebe = Format(ElDebe, "#,##0.00")
    txtTotalHaber = Format(ElHaber, "#,##0.00")
    VSFG.Col = 0
    VSFG.Sort = flexSortGenericAscending
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    'Call Centrar_Forma
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    For i = 2003 To 2060
        cmbAñoI.AddItem i
    Next i
    HacerFecha = False
    cmbAñoI.Text = CStr(Year(Date))
    HacerFecha = True
    'Selecciona el mes actual
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(Date)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    clsCon_Act.Inicializar AdoConn, AdoConnMaster
    clsCon_Tip_Asi.Inicializar AdoConn, AdoConnMaster
    cmbAceptar.Enabled = False
    
    strSql = " SELECT par_con_cta_codigo, cta_nombre FROM parametro_contable" & _
            " INNER JOIN ctaconta ON parametro_contable.par_con_cta_codigo=ctaconta.cta_codigo AND parametro_contable.emp_codigo=ctaconta.emp_codigo" & _
            " WHERE parametro_contable.emp_codigo='" & strEmpresa & "' AND par_con_tipo='ACTIVOS FIJOS' AND par_con_codigo='1' "
    clsCon_Act.Ejecutar (strSql)
    CuentaPatrimonio = clsCon_Act.adorec_Def(0)
    NombreCuentaPatrimonio = clsCon_Act.adorec_Def(1)
End Sub

Private Sub PonerNumeros(Optional conBot As Boolean = True)
    For i = 1 To (VSFDep.Rows - 1)
        VSFDep.TextMatrix(i, 0) = i
    Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Calculo_Depreciacion()
   'Calcula la depreciacion de las Activos en el grid
    Dim clsCon_Aux As New clsConsulta
    Dim Subtotal As Double
    Dim dep As Double
    Dim dep2 As Double
    Dim i As Long
    Dim j As Long
    Dim Meses As Long
    Dim GasDep As Double
    Dim GasDepP As Double
    
    
    VSFG.Clear 1
    VSFG.Rows = 1
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    dep = 0
    txtTotalDebe.Text = "0.00"
    txtTotalHaber.Text = "0.00"
'    fi = Format(cmbAñoInicio.Text + "-" + cmbMesInicio + "-" + cmbDiaInicio, "yyyy-mm-dd")
'    ff = Format(cmbAñoFin.Text + "-" + cmbMesFin + "-" + cmbDiaFin, "yyyy-mm-dd")
    For i = 1 To (VSFDep.Rows - 1)
        ' fd = CDate(cmbAñoAsiento.Text + "-" + cmbMesAsiento + "-" + cmbDiaAsiento)
        fd = Format(VSFDep.TextMatrix(i, 4), "yyyy-mm-dd")
        
        'Si la fecha de adquisición es en este periodo
        If fi <= fd And fd <= ff Then
            dd = DateDiff("d", fd, ff) + 1
        ElseIf fd < fi And fi <= ff Then
            dd = DateDiff("d", fi, ff) + 1
        ElseIf fi < ff And ff < fd Then
            dd = 0
        End If
        
        dep = (FormatoD2(VSFDep.TextMatrix(i, 5)) / (FormatoD2(VSFDep.TextMatrix(i, 3)) * 365))
        'VSFDep.TextMatrix(i, 7) = (dep * dd)
        If VSFDep.TextMatrix(i, 5) - VSFDep.TextMatrix(i, 6) < FormatoD2(dep * dd) Then
            VSFDep.TextMatrix(i, 7) = FormatoD2(VSFDep.TextMatrix(i, 5) - VSFDep.TextMatrix(i, 6))
        Else
            VSFDep.TextMatrix(i, 7) = FormatoD2(dep * dd)
        End If
'        If Trim(VSFDep.TextMatrix(i, 1)) = "EQC 209" Then
'            a = 2562
'        End If
        'Depreciación de la revalorización
        'Se pone en cero la columna de Dep Rev Periodo
        VSFDep.TextMatrix(i, 14) = 0
        'Buscar todos los registros de revalorización para este activo fijo
        strSql = " SELECT det_act_fij_valor, det_act_fij_mes, det_act_fij_año " & _
                 " FROM det_activo_fijo WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND act_fij_codigo='" & VSFDep.TextMatrix(i, 1) & "'" & _
                 " AND det_act_fij_tipo='REV'"
        clsCon_Aux.Ejecutar strSql
        While Not clsCon_Aux.adorec_Def.EOF

            fd = Format(clsCon_Aux.adorec_Def("det_act_fij_año") & "-" & clsCon_Aux.adorec_Def("det_act_fij_mes") & "-01", "yyyy-mm-dd")

            'Si la fecha de revalorización es en este periodo
            'dd = DateDiff("d", fi, ff) + 1
            'Calcular cuanto falta desde la revalorización para el final
            Meses = DateDiff("m", fd, VSFDep.TextMatrix(i, 20)) + 1
            If DateDiff("d", fi, VSFDep.TextMatrix(i, 20)) > 0 Then
                If Meses > 0 And FormatoD2(clsCon_Aux.adorec_Def("det_act_fij_valor")) > 0 Then
                    dep2 = (FormatoD2(clsCon_Aux.adorec_Def("det_act_fij_valor")) / Meses)
                    VSFDep.TextMatrix(i, 14) = FormatoD2(VSFDep.TextMatrix(i, 14)) + FormatoD2(dep2) 'Esta depreciación ya está por mes
                End If
            End If
            
            clsCon_Aux.adorec_Def.MoveNext
        Wend
                 
        'dep2 = (FormatoD(VSFDep.TextMatrix(i, 11)) / (FormatoD(VSFDep.TextMatrix(i, 3)) * 365))
        
        
        
        'Sacar asiento contable depreciación
        AñadirCuenta VSFDep.TextMatrix(i, 9), VSFDep.TextMatrix(i, 10), 0, FormatoD2(VSFDep.TextMatrix(i, 7))
        
        'Sacar asiento contable depreciación revalorización
        AñadirCuenta VSFDep.TextMatrix(i, 18), VSFDep.TextMatrix(i, 19), 0, FormatoD2(VSFDep.TextMatrix(i, 14))
        
        
'        For j = 1 To VSFG.Rows - 1
'            If VSFG.TextMatrix(j, 0) = VSFDep.TextMatrix(i, 9) Then Exit For
'        Next j
'        If j < VSFG.Rows Then
'            'VSFG.TextMatrix(j, 0) = VSFDep.TextMatrix(i, 9)
'            'VSFG.TextMatrix(j, 1) = VSFDep.TextMatrix(i, 10)
'            VSFG.TextMatrix(j, 3) = FormatoD(VSFG.TextMatrix(j, 3)) + FormatoD(VSFDep.TextMatrix(i, 7))
'        Else
'            VSFG.AddItem VSFDep.TextMatrix(i, 9) & vbTab & VSFDep.TextMatrix(i, 10) & vbTab & 0 & vbTab & FormatoD(VSFDep.TextMatrix(i, 7))
'        End If
        'TxtTotalHaber.Text = FormatoD(TxtTotalHaber.Text) + FormatoD(VSFDep.TextMatrix(i, 7))
        
        
        strSql = " SELECT det_gasto_act_are.det_gas_act_are_ctaconta,cta_nombre,depreciacion_activo.cen_cos_codigo,COALESCE(cen_cos_nombre,'') as cen_cos_nombre,sum(dep_act_porcentaje) as porcentaje " & _
                 " FROM activo_fijo INNER JOIN depreciacion_activo ON activo_fijo.emp_codigo=depreciacion_activo.emp_codigo AND activo_fijo.act_fij_codigo=depreciacion_activo.act_fij_codigo " & _
                 " INNER JOIN det_gasto_act_are ON activo_fijo.emp_codigo=det_gasto_act_are.emp_codigo AND activo_fijo.tip_act_codigo=det_gasto_act_are.tip_act_codigo AND depreciacion_activo.are_codigo=det_gasto_act_are.are_codigo " & _
                 " INNER JOIN ctaconta ON det_gasto_act_are.emp_codigo=ctaconta.emp_codigo AND det_gasto_act_are.det_gas_act_are_ctaconta=ctaconta.cta_codigo " & _
                 " LEFT JOIN centro_costo ON depreciacion_activo.emp_codigo=centro_costo.emp_codigo AND depreciacion_activo.cen_cos_codigo=centro_costo.cen_cos_codigo " & _
                 " WHERE activo_fijo.act_fij_codigo='" & VSFDep.TextMatrix(i, 1) & "' AND activo_fijo.emp_codigo='" & strEmpresa & "' " & _
                 " GROUP BY det_gasto_act_are.det_gas_act_are_ctaconta,cta_nombre,depreciacion_activo.cen_cos_codigo "
        clsCon_Aux.Ejecutar strSql
        GasDep = 0
        GasDepP = 0
        While Not clsCon_Aux.adorec_Def.EOF
            AñadirCuenta clsCon_Aux.adorec_Def("det_gas_act_are_ctaconta"), clsCon_Aux.adorec_Def("cta_nombre"), FormatoD2(FormatoD2(clsCon_Aux.adorec_Def("porcentaje")) * FormatoD2(VSFDep.TextMatrix(i, 7) - GasDep) / FormatoD2(100 - GasDepP)), 0, clsCon_Aux.adorec_Def("cen_cos_codigo"), clsCon_Aux.adorec_Def("cen_cos_nombre")
            AñadirCuenta clsCon_Aux.adorec_Def("det_gas_act_are_ctaconta"), clsCon_Aux.adorec_Def("cta_nombre"), FormatoD2(FormatoD2(clsCon_Aux.adorec_Def("porcentaje")) * VSFDep.TextMatrix(i, 14) / 100), 0, clsCon_Aux.adorec_Def("cen_cos_codigo"), clsCon_Aux.adorec_Def("cen_cos_nombre")
            GasDep = GasDep + FormatoD2(FormatoD2(clsCon_Aux.adorec_Def("porcentaje")) * FormatoD2(VSFDep.TextMatrix(i, 7) - GasDep) / FormatoD2(100 - GasDepP))
            GasDepP = GasDepP + FormatoD2(clsCon_Aux.adorec_Def("porcentaje"))
'            For j = 1 To VSFG.Rows - 1
'                If VSFG.TextMatrix(j, 0) = clsCon_Aux.adorec_Def("det_gas_act_are_ctaconta") Then Exit For
'            Next j
'            If j < VSFG.Rows Then
'                VSFG.TextMatrix(j, 0) = clsCon_Aux.adorec_Def("det_gas_act_are_ctaconta")
'                VSFG.TextMatrix(j, 1) = clsCon_Aux.adorec_Def("cta_nombre")
'                VSFG.TextMatrix(j, 2) = FormatoD(VSFG.TextMatrix(j, 2)) + FormatoD(FormatoD(clsCon_Aux.adorec_Def("porcentaje")) * VSFDep.TextMatrix(i, 7) / 100)
'            Else
'                VSFG.AddItem clsCon_Aux.adorec_Def("det_gas_act_are_ctaconta") & vbTab & clsCon_Aux.adorec_Def("cta_nombre") & vbTab & FormatoD(FormatoD(clsCon_Aux.adorec_Def("porcentaje")) * VSFDep.TextMatrix(i, 7) / 100) & vbTab & 0
'            End If
            'TxtTotalDebe.Text = FormatoD(TxtTotalDebe.Text) + FormatoD(FormatoD(clsCon_Aux.adorec_Def("porcentaje")) * VSFDep.TextMatrix(i, 7) / 100)
            clsCon_Aux.adorec_Def.MoveNext
        Wend
        PonerTotalesCuentas
    Next i
End Sub
Private Sub Calculo_Por_Depreciar()
   'Calcula la diferencia que falta por depreciar de las Activos en el grid
   Dim por As Double
    por = 0
    For i = 1 To (VSFDep.Rows - 1)
        'por=valor depreciacuin del activo - depreciacion periodo
        por = (FormatoD2(VSFDep.TextMatrix(i, 5)) - FormatoD2(VSFDep.TextMatrix(i, 6)) - FormatoD2(VSFDep.TextMatrix(i, 7)))
        VSFDep.TextMatrix(i, 8) = por
    Next i
End Sub

Private Sub Calculo_E_DepPorPeriodo()
   'Calcula la sumatoria de depreciaion por periodo del grid
    dp = 0
    For i = 1 To (VSFDep.Rows - 1)
        dp = dp + (FormatoD2(VSFDep.TextMatrix(i, 7)))
    Next i
End Sub

Private Sub Borrar()
    'función que recorre el flexGrid y limpia los campos
    For i = 1 To VSFDep.Rows - 1
        VSFDep.TextMatrix(i, 7) = ""
        VSFDep.TextMatrix(i, 8) = ""
    Next
    TxtTotal = ""
End Sub

Private Sub VSFDep_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        If Col = 5 Or Col = 15 Then
            Me.VSFDep.Cell(flexcpForeColor, Row, Col) = RGB(120, 0, 0)
        End If
        If Col = 7 Or Col = 12 Or Col = 14 Then
            Me.VSFDep.Cell(flexcpBackColor, Row, Col) = RGB(230, 230, 230)
        End If
    End If
End Sub

