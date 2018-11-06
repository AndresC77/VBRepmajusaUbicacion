VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAsiento 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asiento"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "frmAsiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11025
   Begin VSFlex8Ctl.VSFlexGrid VSFGcarga 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
      _cx             =   4471
      _cy             =   1085
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAsiento.frx":030A
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
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   9600
         TabIndex        =   18
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir"
         Height          =   375
         Left            =   8880
         TabIndex        =   17
         Top             =   9360
         Width           =   1095
      End
      Begin NEED2.dtpFecha Fecha1 
         Height          =   315
         Left            =   2992
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Value           =   42892.7188194444
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtAsiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4672
         TabIndex        =   4
         Top             =   600
         Width           =   3360
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
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1815
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox chkRevisado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Revisado"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CheckBox chkMayorizado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Mayorizado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3020
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Tag             =   "7"
         Top             =   3600
         Width           =   10455
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1080
         Width           =   10560
         _cx             =   18627
         _cy             =   2778
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
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
         FormatString    =   $"frmAsiento.frx":036C
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin MSComDlg.CommonDialog cmdArchivo 
         Left            =   10200
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Asiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4672
         TabIndex        =   13
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2992
         TabIndex        =   12
         Top             =   360
         Width           =   1335
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4770
         TabIndex        =   11
         Top             =   2805
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar Asiento"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5565
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   9480
      Picture         =   "frmAsiento.frx":0462
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   9240
      Picture         =   "frmAsiento.frx":058E
      ToolTipText     =   "Elimina una Fila"
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCta As New clsConsulta
Private clsAsi As New clsConsulta
Private Hacer As Boolean
Private HacerConsulta As Boolean
Private strMaximo As String
Private PrimeraVez As Boolean
Public Manual As Boolean
Public objeto As Object
Public ActivoFijo As Boolean
Public Objeto1 As Object
Public FechaMinima As Variant

Private Sub chkMayorizado_Click()
    If Hacer = True Then Exit Sub
    If chkMayorizado.Value = 1 Then
        Hacer = True
        chkMayorizado.Value = 0
        Hacer = False
    ElseIf chkMayorizado.Value = 0 Then
        Hacer = True
        chkMayorizado.Value = 1
        Hacer = False
    End If
End Sub

Private Sub cmdCargar_Click()
    Dim strPath As String
    Dim i As Long
    VSFGCarga.Clear 1
    VSFGCarga.Rows = 1
    VSFG.Clear 1
    VSFG.Rows = 1
    
    strPath = Trim(App.Path)
    cmdArchivo.DialogTitle = "Abrir"
    'cmdArchivo.DefaultExt = strPath
    cmdArchivo.InitDir = strPath
    'cmdArchivo.FileName = Arch
    cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Documento de Excel 2007|*xlsx|Todos los Archivos|*.*"
    cmdArchivo.ShowOpen
    'num = FreeFile
    Archivo = cmdArchivo.FileName
    If Archivo <> "" Then
        VSFGCarga.LoadGrid Archivo, flexFileExcel
        For i = 0 To VSFGCarga.Rows - 1
            VSFG.AddItem ""
            VSFG.TextMatrix(i + 1, 1) = Trim(VSFGCarga.TextMatrix(i, 0))
            VSFG.TextMatrix(i + 1, 3) = Trim(VSFGCarga.TextMatrix(i, 2))
            VSFG.TextMatrix(i + 1, 4) = VSFGCarga.TextMatrix(i, 3)
            VSFG.TextMatrix(i + 1, 5) = VSFGCarga.TextMatrix(i, 1)
            VSFG.TextMatrix(i + 1, VSFG.Cols - 1) = 1
        Next i
        PonerBotones
        CalcuTotal
    End If

End Sub

Private Sub cmdGuardar_Click()
    a = VSFG.Rows - 1
'    For i = 1 To a
'        For j = i + 1 To a
'            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) And VSFG.TextMatrix(i, 5) = VSFG.TextMatrix(j, 5) Then
'                If VSFG.TextMatrix(i, 3) = "" Then
'                    VSFG.TextMatrix(i, 3) = 0
'                End If
'                If VSFG.TextMatrix(j, 3) = "" Then
'                    VSFG.TextMatrix(j, 3) = 0
'                End If
'                VSFG.TextMatrix(i, 3) = CDbl(VSFG.TextMatrix(i, 3)) + CDbl(VSFG.TextMatrix(j, 3))
'                If VSFG.TextMatrix(i, 4) = "" Then
'                    VSFG.TextMatrix(i, 4) = 0
'                End If
'                If VSFG.TextMatrix(j, 4) = "" Then
'                    VSFG.TextMatrix(j, 4) = 0
'                End If
'                VSFG.TextMatrix(i, 4) = CDbl(VSFG.TextMatrix(i, 4)) + CDbl(VSFG.TextMatrix(j, 4))
'                VSFG.RemoveItem j
'                a = a - 1
'                j = j - 1
'            End If
'            If j >= a Then
'                Exit For
'            End If
'        Next j
'    Next i

    If TxtTotal1Debe = 0 Or TxtTotal1Haber = 0 Then
        If MsgBox("El debe o el haber tienen valor cero." & vbNewLine & "Quiere Guardar el Asiento?", vbYesNo + vbQuestion, "Información") = vbNo Then
            VSFG.SetFocus
            Exit Sub
        End If
    End If
    'verifica que el debe y el haber esten cuadrados
    If TxtTotal1Debe <> TxtTotal1Haber Then
        If MsgBox("No está cuadrado el debe y el haber." & vbNewLine & "Quiere Guardar el Asiento?", vbYesNo + vbQuestion, "Información") = vbNo Then
            VSFG.SetFocus
            Exit Sub
        End If
    End If
    
    If VSFG.TextMatrix(1, 1) = "" Then
        MsgBox "No ingresó el detalle del asiento.", vbInformation, "Información"
        VSFG.SetFocus
        Exit Sub
    End If
    
    
    Dim strMaximo2 As String
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    If Me.Tag = "N" Then
        'Ingreso de datos en el asiento
        clsAsiento.NuevoAsiento "D", Fecha1.Value, chkRevisado.Value, 0, Format(TxtTotal1Debe, "#0.00"), txtDescripcion
        strMaximo = clsAsiento.NumAsiento
        txtAsiento = strMaximo
    ElseIf Me.Tag = "M" Then
        clsAsiento.NumAsiento = txtAsiento
        clsAsiento.ModificarAsiento Format(TxtTotal1Debe, "##0.00"), Format(TxtTotal1Haber, "##0.00"), Fecha1.Value, chkRevisado.Value, , txtDescripcion.Text
        clsAsiento.EliminarAsiento False, True
    End If
    
    With VSFG
        For i = 1 To .Rows - 1
            clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
        Next i
    End With
    
    If Me.Tag = "N" Then
        If Manual = True Then
            objeto.Text = clsAsiento.NumAsiento
        End If
        MsgBox "Asiento " & clsAsiento.NumAsiento & " creado.", vbInformation, "Nuevo"
    ElseIf Me.Tag = "M" Then
        MsgBox "Asiento " & clsAsiento.NumAsiento & " modificado.", vbInformation, "Modificar"
        frmVerAsiento.Modificando = True
    End If
    If Manual = False Then
        frmVerAsiento.HacerActivate = True
    End If
    Dim Asien As New frmReporte
    Asien.strAsiento = clsAsiento.NumAsiento
    Asien.strReporte = "rptAsiento"
    Asien.Show
    Set clsAsiento = Nothing
    Unload Me
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez = False Then Exit Sub
    If Me.Tag = "N" Then
        Me.Caption = "Nuevo Asiento"
        VSFG.Rows = 2
        VSFG.TextMatrix(1, VSFG.Cols - 1) = 1
    ElseIf Me.Tag = "M" Then
        Me.Caption = "Modificar Asiento"
        
    End If
    PonerBotones
    PrimeraVez = False
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Screen.MousePointer = vbHourglass
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Fecha1.Value = HoyDia
    PrimeraVez = True
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    
    strSQL = " SELECT cen_cos_codigo, cen_cos_nombre " & _
             " FROM centro_costo " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY cen_cos_nombre "
    clsCta.Ejecutar strSQL
    
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCta.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
    strSQL = " SELECT cta_codigo, cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY cta_codigo "
    clsCta.Ejecutar strSQL
    
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
    strSQL = " SELECT cta_codigo, cta_nombre" & _
             " FROM ctaconta " & _
             " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY cta_nombre "
    clsCta.Ejecutar strSQL
    VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCta = Nothing
    Set clsAsi = Nothing
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
        MsgBox "CUENTA CONTABLE: Datos incompletos.", vbInformation, "Teso/Bancos "
        VSFG.TextMatrix(Row, 3) = 0
        VSFG.TextMatrix(Row, 4) = 0
        ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe

        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "DEBE: Sólo acepta números.", vbInformation, "Teso/Bancos "
            VSFG.TextMatrix(Row, 3) = 0
        End If

        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "HABER: Sólo acepta números.", vbInformation, "Teso/Bancos "
            VSFG.TextMatrix(Row, 4) = 0
        End If
    CalcuTotal
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 0
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <> OldRow Then
        If Year(Fecha1.Value) >= 2018 And VSFG.Rows > 1 Then
            If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'filtra el nombre y codigo de cuenta para los combos del greed
    If Row >= 1 Then
        With VSFG
            If .TextMatrix(Row, Col) <> "" Then
                If Col = 1 Then
                     clsCta.Filtrar ("cta_codigo = '" & Trim(.TextMatrix(Row, 1)) & "'")
                        .TextMatrix(Row, 2) = clsCta.adorec_Def("cta_nombre")
                     clsCta.QuitarFiltro
                 End If
    
                 If Col = 2 Then
                     clsCta.Filtrar ("cta_codigo = '" & Trim(.TextMatrix(Row, 2)) & "'")
                         .TextMatrix(Row, 1) = clsCta.adorec_Def("cta_codigo")
                     clsCta.QuitarFiltro
                 End If
             End If
             If Col = 3 Or Col = 4 Then
                .TextMatrix(Row, Col) = FormatoD2(.TextMatrix(Row, Col))
             End If
        End With
    End If
End Sub

Public Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
    'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes
    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            If Year(Fecha1.Value) >= 2018 And VSFG.Rows > 1 Then
                If Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "4" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "5" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "6" Then
                    VSFG.AddItem ""
                ElseIf VSFG.TextMatrix(VSFG.Row, 5) <> "" Then
                    VSFG.AddItem ""
                End If
            Else
                VSFG.AddItem ""
            End If
       
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            VSFG.TextMatrix(VSFG.Rows - 1, 3) = FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 3))
            VSFG.TextMatrix(VSFG.Rows - 1, 4) = FormatoD2(VSFG.TextMatrix(VSFG.Rows - 1, 4))
            VSFG.TextMatrix(VSFG.Rows - 1, VSFG.Cols - 1) = 1
            PonerBotones
        End If
    End If
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Or Val(VSFG.TextMatrix(r, VSFG.Cols - 1)) = 0 Then Exit Sub
    If (c <> 0 Or r = 0) Then Exit Sub
    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        If r > 0 Then
            ' click was on a button: do the work
            VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
            Mensaje = "¿Está seguro de eliminar la fila " & r & "?"    ' Define el mensaje.
            Estilo = vbYesNo + vbQuestion + vbDefaultButton2   ' Define los botones.
            Título = "Pregunta"   ' Define el título.
            respuesta = MsgBox(Mensaje, Estilo, Título)
            'Recorro el FlexGrid para poner números a las filas
            If respuesta = vbYes Then
                Dim i As Integer
                If VSFG.Rows > 2 Then
                    VSFG.RemoveItem (r)
                Else
                    VSFG.Clear 1
                End If
                PonerBotones
                CalcuTotal
            Else
                VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
            End If
        End If
    End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub

Private Sub CalcuTotal()
   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double

    'Calcula total debe

    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + (IIf(VSFG.TextMatrix(i, 3) = "", 0, VSFG.TextMatrix(i, 3)))
    Next i
    TxtTotal1Debe = Format(SumaDebe, "###0.00")
    TxtTotal2Debe = TxtTotal1Debe
    'Calcula total haber

    For i = 1 To VSFG.Rows - 1
        SumaHaber = SumaHaber + (IIf(VSFG.TextMatrix(i, 4) = "", 0, VSFG.TextMatrix(i, 4)))
    Next i
    TxtTotal1Haber = Format(SumaHaber, "###0.00")
    TxtTotal2Haber = TxtTotal1Haber
    TxtTotal.Text = FormatoD2(TxtTotal1Debe.Text) - FormatoD2(TxtTotal1Haber.Text)
End Sub

Private Sub VSFG_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            Set frmSelecCtaConta.objEscribir = VSFG
            frmSelecCtaConta.Show
        End If
    End If
End Sub

Private Sub VSFG_Validate(Cancel As Boolean)
    If Year(Fecha1.Value) >= 2018 And VSFG.Rows > 1 Then
        If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
            Cancel = True
        End If
    End If
End Sub
