VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmContabilizarRol 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilizar Rol"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmContabilizarRol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8910
   Begin VB.CommandButton Command1 
      Caption         =   "Sacar Cheque"
      Height          =   375
      Left            =   1748
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8655
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3015
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8175
         _cx             =   14420
         _cy             =   5318
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmContabilizarRol.frx":030A
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
      Begin VB.TextBox TxtTotal1Debe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3480
         Width           =   1545
      End
      Begin VB.TextBox TxtTotal1Haber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3480
         Width           =   1545
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Tag             =   "7"
         Top             =   3840
         Width           =   8175
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
         Left            =   3930
         TabIndex        =   7
         Top             =   3525
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar Asiento"
      Height          =   375
      Left            =   3668
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5588
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmContabilizarRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsSql1 As New clsConsulta
Private strSql As String
Private PrimeraVez As Boolean
Public Fecha As Variant
Public Liquidacion As Boolean
Public CodigoEmpleado As String
Private Valor As Double
Public CuentaNomina As String

Private Sub PonerCuenta(CUENTA As String, DEBE As Double, HABER As Double, Empleado As String)
    If DEBE = 0 And HABER = 0 Then Exit Sub
    DEBE = FormatoD2(DEBE)
    HABER = FormatoD2(HABER)
    
    'Centro de costos según línea de empleado
'    If Left(Cuenta, 1) = "4" Or Left(Cuenta, 1) = "5" Then
'        Dim NumeroLineas As Integer
'        Dim linea As Integer
'        Dim debeCC As Double
'        Dim haberCC As Double
'        Dim TotalDebeCC As Double
'        Dim TotalHaberCC As Double
'
'        strSql = " SELECT lin_codigo FROM empleado_linea WHERE emp_codigo='" & strEmpresa & "'" & _
'                 " AND epl_codigo='" & Empleado & "'"
'        clsSql.Ejecutar (strSql)
'        If clsSql.adorec_Def.RecordCount > 0 Then
'            'Poner valores divididos en lineas
'            NumeroLineas = clsSql.adorec_Def.RecordCount
'            linea = 0
'            While clsSql.adorec_Def.EOF = False
'                linea = linea + 1
'                debeCC = FormatoD(Debe / NumeroLineas)
'                haberCC = FormatoD(Haber / NumeroLineas)
'                'Para poner la diferencia de la última división en la última línea
'                TotalDebeCC = TotalDebeCC + debeCC
'                TotalHaberCC = TotalHaberCC + haberCC
'                If linea = NumeroLineas Then
'                    If TotalDebeCC <> Debe Then
'                        debeCC = debeCC - TotalDebeCC + Debe
'                    End If
'                    If TotalHaberCC <> Haber Then
'                        haberCC = haberCC - TotalHaberCC + Haber
'                    End If
'                End If
'                PonerCuenta2 Cuenta, debeCC, haberCC, "L" & clsSql.adorec_Def(0)
'                clsSql.adorec_Def.MoveNext
'            Wend
'            Exit Sub
'        End If
'    End If
    'Poner valor total sin división de líneas
    PonerCuenta2 CUENTA, DEBE, HABER, ""
End Sub

Public Sub PonerCuenta2(CUENTA As String, DEBE As Double, HABER As Double, CentroCosto As String)
    If DEBE = 0 And HABER = 0 Then Exit Sub
    
    Dim Posicion As Long
    DEBE = FormatoD2(DEBE)
    HABER = FormatoD2(HABER)
    Posicion = 0
    
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) = CUENTA And VSFG.TextMatrix(i, 5) = CentroCosto Then
            Posicion = i
            Exit For
        End If
    Next i
    If Posicion = 0 Then
        strSql = " SELECT cta_nombre FROM ctaconta WHERE emp_codigo = '" & strEmpresa & "' AND cta_codigo ='" & CUENTA & "'"
        clsSql1.Ejecutar strSql
        If clsSql1.adorec_Def.RecordCount > 0 Then
            VSFG.AddItem "" & vbTab & CUENTA & vbTab & clsSql1.adorec_Def(0) & vbTab & DEBE & vbTab & HABER & vbTab & CentroCosto
        Else
            Unload Me
        End If
    Else
        VSFG.TextMatrix(Posicion, 3) = Val(VSFG.TextMatrix(Posicion, 3)) + DEBE
        VSFG.TextMatrix(Posicion, 4) = Val(VSFG.TextMatrix(Posicion, 4)) + HABER
        'Poner el valor del cheque
        If CUENTA = Me.CuentaNomina Then
            Valor = VSFG.TextMatrix(Posicion, 4)
        End If
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    'If VerificarFechaContable(Me.Fecha) = False Then Exit Sub
    
    If TxtTotal1Debe = 0 Or TxtTotal1Haber = 0 Then
        MsgBox "El debe o el haber no pueden tener valor cero.", vbInformation, "Información"
        VSFG.SetFocus
        Exit Sub
    End If
    'verifica que el debe y el haber esten cuadrados
    If TxtTotal1Debe <> TxtTotal1Haber Then
        MsgBox "No está cuadrado el debe y el haber.", vbInformation, "Información"
        VSFG.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim strMaximo As String
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    'Grabar asiento
    clsAsiento.NuevoAsiento "D", CStr(Fecha), 0, 0, Format(TxtTotal1Debe, "#0.00"), UCase(txtDescripcion)
    strMaximo = clsAsiento.NumAsiento
    'strMaximo = AsientoNuevo(CStr(Fecha), "RRH", 0, 0, Format(TxtTotal1Debe, "##0.00"), UCase(txtDescripcion.Text))
    With VSFG
        For i = 1 To .Rows - 1
            clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), "", CDbl(Format(.TextMatrix(i, 3), "##0.00")), CDbl(Format(.TextMatrix(i, 4), "##0.00"))
    '        NuevoDetAsiento strMaximo, .TextMatrix(i, 1), Val(Format(.TextMatrix(i, 3), "##0.00")), Val(Format(.TextMatrix(i, 4), "##0.00")), .TextMatrix(i, 5)
        Next i
    End With
    
    
    'Marcar Cuentas como Pagadas y grabar número de asiento.
    For i = 1 To frmSelEstadoCuenta.VSFG.Rows - 1
        If frmSelEstadoCuenta.VSFG.IsSubtotal(i) = False Then
            strSql = " UPDATE descuento SET asi_numasiento2='" & strMaximo & "', des_pagado=1" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND des_codigo='" & frmSelEstadoCuenta.VSFG.TextMatrix(i, 11) & "'"
            clsSql.Ejecutar strSql, "M"
        End If
    Next i
    
    'Si es liquidación da de baja el empleado y pone el número de asiento
    If Me.Liquidacion = True Then
        strSql = " UPDATE empleado SET epl_baja=1, asi_numasiento='" & strMaximo & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & CodigoEmpleado & "'"
        clsSql.Ejecutar strSql, "M"
    End If
    'Mandar a recargar el estado de cuenta
    frmSelEstadoCuenta.BuscarEstadoCuenta
    Screen.MousePointer = vbDefault
    MsgBox "Asiento de Rol de Pagos " & strMaximo & " generado", vbInformation, "Información"
    'drptCompConta.Tag = strMaximo
    'drptCompConta.Show
    Dim rptAsiento As New frmReporte
    rptAsiento.strReporte = "rptAsiento"
    rptAsiento.strAsiento = strMaximo
    rptAsiento.Show
    '************SacarCheque
    Unload Me
End Sub

Private Sub SacarCheque()
    'Sacar cheque automáticamente
    Me.MousePointer = vbArrowHourglass
    frmComprobanteEgresoComun.Show
    'frmComprobanteEgresoComun.optEmpleado.Value = True
    frmComprobanteEgresoComun.dcmbBeneficiario.BoundText = CodigoEmpleado
    frmComprobanteEgresoComun.dcmbBeneficiario_Change
    frmComprobanteEgresoComun.txtValor = Valor
    frmComprobanteEgresoComun.VSFG.Rows = 3
    frmComprobanteEgresoComun.VSFG.TextMatrix(1, 4) = Valor
    frmComprobanteEgresoComun.VSFG.TextMatrix(2, 3) = Valor
    frmComprobanteEgresoComun.VSFG.TextMatrix(2, 4) = 0
    frmComprobanteEgresoComun.VSFG.TextMatrix(2, 1) = CuentaNomina
    frmComprobanteEgresoComun.TxtTotalDebe = Valor
    frmComprobanteEgresoComun.TxtTotalHaber = Valor
    Me.MousePointer = vbDefault
    Me.Tag = "Nuevo"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    SacarCheque
End Sub

Private Sub Form_Activate()
    If PrimeraVez = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    For i = 1 To frmSelEstadoCuenta.VSFG.Rows - 1
        If frmSelEstadoCuenta.VSFG.IsSubtotal(i) = False Then
            PonerCuenta frmSelEstadoCuenta.VSFG.TextMatrix(i, 8), frmSelEstadoCuenta.VSFG.TextMatrix(i, 10), 0, frmSelEstadoCuenta.VSFG.TextMatrix(i, 1)
            PonerCuenta frmSelEstadoCuenta.VSFG.TextMatrix(i, 9), 0, frmSelEstadoCuenta.VSFG.TextMatrix(i, 10), frmSelEstadoCuenta.VSFG.TextMatrix(i, 1)
        End If
    Next i
    
    'Pa que ordene por código cuenta y codigo línea
    VSFG.Col = 5
    VSFG.Sort = 1
    VSFG.Col = 1
    VSFG.Sort = 1
    CalcuTotal
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
        If VSFG.TextMatrix(i, 1) = Me.CuentaNomina Then
            Valor = FormatoD2(VSFG.TextMatrix(i, 4))
        End If
    Next i
    
    PrimeraVez = False
    Me.Frame2.Caption = "Asiento Roles de Pago - " & Fecha
    
    'Poner el valor del cheque
        
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)

    PrimeraVez = True
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster

    strSql = " select concat('L',lin_codigo) as lin_codigo, lin_nombre " & _
             " FROM linea " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY lin_nombre "
    clsSql.Ejecutar (strSql)
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsSql.adorec_Def, "lin_codigo, *lin_nombre", "lin_codigo")
End Sub

Private Sub CalcuTotal()
   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    Dim Diferencia As Double

    'Calcula total debe

    For i = 1 To VSFG.Rows - 1
        Diferencia = FormatoD2(VSFG.TextMatrix(i, 3)) - FormatoD2(VSFG.TextMatrix(i, 4))
        
        If Diferencia >= 0 Then
            VSFG.TextMatrix(i, 3) = Diferencia
            VSFG.TextMatrix(i, 4) = 0
        Else
            VSFG.TextMatrix(i, 3) = 0
            VSFG.TextMatrix(i, 4) = Diferencia * -1
        End If
        SumaDebe = SumaDebe + FormatoD2(VSFG.TextMatrix(i, 3))
        SumaHaber = SumaHaber + FormatoD2(VSFG.TextMatrix(i, 4))
    Next i
    TxtTotal1Debe = Format(SumaDebe, "#,##0.00")
    TxtTotal1Haber = Format(SumaHaber, "#,##0.00")
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 Then Cancel = True
End Sub

Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub
