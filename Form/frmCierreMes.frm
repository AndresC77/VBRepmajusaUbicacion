VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierreMes 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Mes"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmCierreMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8190
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "&Limpiar"
      Height          =   360
      Left            =   4148
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Cerrar mes"
      Height          =   360
      Left            =   1508
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnAbrirMes 
      Caption         =   "&Reabrir Mes"
      Height          =   360
      Left            =   2828
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5468
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cierre de Mes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7905
      Begin VB.TextBox txtDescripcion 
         Height          =   855
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM"
         Format          =   59244547
         CurrentDate     =   37463
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2880
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   7980
      _cx             =   14076
      _cy             =   5080
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCierreMes.frx":030A
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
End
Attribute VB_Name = "frmCierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSQL As String
Private Tipo As String
Private Tipo2 As String
Private Sub IniDato()
    Tipo = "Cierre de Mes"
    Tipo2 = "Cierre de Mes"
    Me.Caption = Tipo
End Sub

Private Sub Carga()
    strSQL = " SELECT cie_mes_ano,cie_mes_mes,cie_mes_mes,cie_mes_descripcion,cie_mes_fechamod,cie_mes_usumod " & _
             " FROM cierre_mes " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = CStr(i)
        VSFG.TextMatrix(i, 3) = MostrarMes(FormatoD0(VSFG.TextMatrix(i, 3)))
    Next i
    
    btnAbrirMes.Enabled = False
    
    strSQL = " SELECT COALESCE(count(*),0) " & _
             " FROM cierre_mes " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cie_mes_ano=" & Year(HoyDia) & " AND cie_mes_mes=" & Month(HoyDia) & " "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        If clsCon_Def.adorec_Def(0) > 0 Then
            cmbAceptar.Enabled = False
            Exit Sub
        End If
    End If
    cmbAceptar.Enabled = True
End Sub

Private Sub btnAbrirMes_Click()
    If MsgBox("Está seguro que desea reabrir el mes de " & MostrarMes(Month(dtpFecha.Value)) & " de " & Year(dtpFecha.Value) & "?", vbQuestion + vbYesNo, "Cerrar Mes") = vbYes Then
        strSQL = " DELETE FROM cierre_mes " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cie_mes_ano='" & Year(dtpFecha.Value) & "' " & _
                 " AND cie_mes_mes='" & Month(dtpFecha.Value) & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        MsgBox "Se ha reabierto el mes con éxito", vbInformation, "Cierre de Mes"
        Limpiar
    End If
End Sub

Private Sub btnLimpiar_Click()
    Limpiar
End Sub

Private Sub cmbAceptar_Click()
    If MsgBox("Está seguro que desea cerrar el mes de " & MostrarMes(Month(dtpFecha.Value)) & " de " & Year(dtpFecha.Value) & "?", vbQuestion + vbYesNo, "Cerrar Mes") = vbYes Then
    
        strSQL = " SELECT cie_mes_ano,cie_mes_mes " & _
                 " FROM cierre_mes " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY cie_mes_ano DESC,cie_mes_mes DESC LIMIT 1 "
        clsCon_Def.Ejecutar strSQL
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            'If DateAdd("m", 1, CDate(Format(clsCon_Def.adorec_Def(0) & "-" & Format(clsCon_Def.adorec_Def(1), "00") & "-01"))) <> CDate(Format(Year(dtpFecha.value) & "-" & Format(Month(dtpFecha.value), "00") & "-01")) Then
            '    MsgBox "No existe el cierre del mes anterior a la fecha seleccionada," & vbNewLine & "no se ha podido cerrar el mes", vbInformation, "Cierre de Mes"
            '    Exit Sub
            'End If
        End If
    
        strSQL = " INSERT INTO cierre_mes(emp_codigo,cie_mes_ano,cie_mes_mes,cie_mes_descripcion,cie_mes_fechamod,cie_mes_usumod) " & _
                " VALUES ('" & strEmpresa & "'," & Year(dtpFecha.Value) & "," & Month(dtpFecha.Value) & ",'" & UCase(txtDescripcion.Text) & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsCon_Def.Ejecutar strSQL, "M"
        MsgBox "Se ha cerrado el mes con éxito", vbInformation, "Cierre de Mes"
        Limpiar
    End If
End Sub


Private Sub VSFG_DblClick()
    If VSFG.Row > 0 Then
        dtpFecha.Value = Format(VSFG.TextMatrix(VSFG.Row, 1) & "-" & VSFG.TextMatrix(VSFG.Row, 2) & "-01", "yyyy-mm-dd")
        txtDescripcion.Text = VSFG.TextMatrix(VSFG.Row, 4)
        cmbAceptar.Enabled = False
        btnAbrirMes.Enabled = True
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
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
    IniDato
    Limpiar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Limpiar()
    Carga
    txtDescripcion.Text = ""
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Function MostrarMes(Index As Integer) As String
    If Index = 1 Then
        MostrarMes = "Enero"
    ElseIf Index = 2 Then
        MostrarMes = "Febrero"
    ElseIf Index = 3 Then
        MostrarMes = "Marzo"
    ElseIf Index = 4 Then
        MostrarMes = "Abril"
    ElseIf Index = 5 Then
        MostrarMes = "Mayo"
    ElseIf Index = 6 Then
        MostrarMes = "Junio"
    ElseIf Index = 7 Then
        MostrarMes = "Julio"
    ElseIf Index = 8 Then
        MostrarMes = "Agosto"
    ElseIf Index = 9 Then
        MostrarMes = "Septiembre"
    ElseIf Index = 10 Then
        MostrarMes = "Octubre"
    ElseIf Index = 11 Then
        MostrarMes = "Noviembre"
    ElseIf Index = 12 Then
        MostrarMes = "Diciembre"
    Else
        MostrarMes = ""
    End If
End Function
