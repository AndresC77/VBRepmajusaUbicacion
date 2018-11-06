VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJustificar 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Justificación de Atrasos"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmJustificar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11055
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8505
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   4440
         TabIndex        =   10
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Width           =   255
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
            Enabled         =   0   'False
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
            Height          =   255
            Left            =   480
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.ComboBox cmbMesI 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmJustificar.frx":030A
            Left            =   1320
            List            =   "frmJustificar.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Value           =   -1  'True
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
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
            Left            =   480
            TabIndex        =   15
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57671683
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
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
            Left            =   1920
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57671683
            CurrentDate     =   37463
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Empleado"
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dcmbEmpleado 
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empleado"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   495
         Width           =   3255
      End
   End
   Begin VB.Frame fraBotones 
      BackColor       =   &H00DDDDDD&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   10815
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Cerrar"
         Height          =   360
         Left            =   5600
         TabIndex        =   4
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   3755
         TabIndex        =   3
         Top             =   240
         Width           =   1700
      End
   End
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   10815
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   855
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4095
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   10335
         _cx             =   18230
         _cy             =   7223
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmJustificar.frx":039E
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
         ExplorerBar     =   3
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
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   3120
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   9120
      Picture         =   "frmJustificar.frx":0592
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmJustificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsSql As New clsConsulta
Private strSql As String
Private FechaI As Variant
Private FechaF As Variant


Private Sub btnAceptar_Click()
    Dim i As Long, control As Integer, codigo As String
    control = 0
    If VSFG.Rows > 1 Then
        For i = 1 To VSFG.Rows - 1
            If FormatoD0(VSFG.TextMatrix(i, VSFG.Cols - 1)) <> 0 Then
                If CBool(VSFG.TextMatrix(i, 5)) = True Then
                    strSql = " UPDATE asistencia SET " & _
                             " ast_tiempo_atraso='00:00:00', " & _
                             " ast_justificado='1', " & _
                             " ast_multa='0', " & _
                             " ast_observacion='" & VSFG.TextMatrix(i, 6) & "', " & _
                             " ast_fechamod=CURRENT_TIMESTAMP, " & _
                             " ast_usumod='" & strUsuario & "' " & _
                             " WHERE epl_codigo='" & VSFG.TextMatrix(i, 2) & "' " & _
                             " AND ast_hora='" & VSFG.TextMatrix(i, 1) & "' " & _
                             " AND ast_fecha='" & VSFG.TextMatrix(i, 4) & "' " & _
                             " AND emp_codigo='" & strEmpresa & "' "
                    clsSql.Ejecutar strSql, "M"
                End If
            End If
        Next i
        Limpiar
    End If
End Sub


Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub chkFechas_Click()
    If chkFechas.value = 1 Then
        Label22.Caption = "Fecha Inicial"
        Label23.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Label22.Caption = "Fecha"
        Label23.Enabled = False
        Fecha2.Enabled = False
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.value = 1 Then
        dcmbEmpleado.Enabled = True
    Else
        dcmbEmpleado.Enabled = False
    End If
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1", "yyyy-mm-dd")
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal, "yyyy-mm-dd")
        DiaFinal = DiaFinal - 1
    Wend
End Sub

Private Sub chkFiltroFecha_Click()
    If chkFiltroFecha.value = 1 Then
        fraFecha.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        If Option1.value = True Then
            lblMes.Enabled = True
            cmbMesI.Enabled = True
        ElseIf Option2.value = True Then
            Fecha1.Enabled = True
            Label22.Enabled = True
            Fecha1.Enabled = True
            chkFechas.Enabled = True
            If chkFechas.value = 1 Then
                Label23.Enabled = True
                Fecha2.Enabled = True
            End If
        End If
    Else
        fraFecha.Enabled = False
        
        Fecha2.Enabled = False
        Label22.Enabled = False
        Fecha1.Enabled = False
        Label23.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        
        Option1.Enabled = False
        Option2.Enabled = False
        lblMes.Enabled = False
        cmbMesI.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    CargarDatos
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    chkFechas.Enabled = True
    chkFiltroFecha.value = 1
    Option2.value = True
    
    CargarEmpleados
    Dim i As Integer
    Fecha1 = Format(HoyDia, "yyyy-mm-dd")
    Fecha2 = Format(HoyDia, "yyyy-mm-dd")
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    
    CargarDatos
End Sub

Private Sub Limpiar()
    CargarDatos
End Sub

Private Sub CargarDatos()
    strSql = " SELECT TIME_FORMAT(ast_hora,'%H:%i:%s'),asistencia.epl_codigo,CONCAT(epl_apellidos,' ',epl_nombres) as nombre," & _
             " ast_fecha,'' as justif,COALESCE(ast_observacion,''),TIME_FORMAT(ast_entrada_esp,'%H:%i:%s'),TIME_FORMAT(ast_entrada,'%H:%i:%s')," & _
             " TIME_FORMAT(ast_salida_esp,'%H:%i:%s'),TIME_FORMAT(ast_salida,'%H:%i:%s'),TIME_FORMAT(ast_tiempo_atraso,'%H:%i:%s'),COALESCE(ast_multa,0),'0' as modif " & _
             " FROM asistencia " & _
             " INNER JOIN empleado " & _
             " ON empleado.emp_codigo=asistencia.emp_codigo " & _
             " AND empleado.epl_codigo=asistencia.epl_codigo " & _
             " WHERE asistencia.emp_codigo='" & strEmpresa & "' " & _
             " AND ast_atrasado='1' "
    
    If chkFiltroCodigo.value = 1 Then
        strSql = strSql & " AND CONCAT(epl_apellidos,' ',epl_nombres) LIKE '" & dcmbEmpleado.Text & "' "
    End If
    
    If chkFiltroFecha.value = 1 Then
        If Option1.value = True Then
            strSql = strSql & " AND ast_fecha BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
        ElseIf Option2.value = True Then
           If chkFechas.value = 0 Then
                strSql = strSql & " AND ast_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha1 & "' "
            Else
                strSql = strSql & " AND ast_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
            End If
        End If
    End If
    
    strSql = strSql & " ORDER BY epl_apellidos,epl_nombres,ast_fecha,ast_entrada_esp,ast_entrada,ast_salida_esp,ast_salida "
    clsSql.Ejecutar strSql
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    
    Dim i As Long
    
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
        'If CBool(FormatoD0(VSFG.TextMatrix(i, 5))) = True Then
            VSFG.Cell(flexcpData, i, 5) = "Tiempo de Atraso:" & vbCrLf & VSFG.TextMatrix(i, 11)
            VSFG.Cell(flexcpPicture, i, 5) = imgNote
            VSFG.Cell(flexcpAlignment, i, 5) = 4
            VSFG.Cell(flexcpPictureAlignment, i, 5) = flexAlignCenterCenter
'flexPicAlignRightTop
        'End If
        
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub



Private Sub Option1_Click()
    If Option1.value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Label22.Enabled = False
        Fecha1.Enabled = False
        Label23.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        Label22.Enabled = True
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.value = 1 Then
            Label23.Enabled = True
            Fecha2.Enabled = True
        End If
    End If
End Sub


Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 5 Then
        If FormatoD0(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Then
            VSFG.TextMatrix(Row, VSFG.Cols - 1) = "1"
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 And Col <> 6 Then
        Cancel = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub CargarEmpleados()
    strSql = " SELECT epl_codigo as codigo,CONCAT(epl_apellidos,' ',epl_nombres) as nombre " & _
             " FROM empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set dcmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    dcmbEmpleado.ListField = "nombre"
    dcmbEmpleado.BoundColumn = "codigo"
End Sub

Private Sub VSFG_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static lNoteRow&, lNoteCol&, r&, c&

    ' clicking? no work
    If Button <> 0 Then Exit Sub
    
    ' get mouse coordinates
    r = VSFG.MouseRow
    c = VSFG.MouseCol
    
    ' same cell or neighbour? no work
    If r = lNoteRow And c = lNoteCol Then Exit Sub
    If r = lNoteRow And c = lNoteCol + 1 Then Exit Sub
    
    ' other cell, hide current note, if any
    If lNoteRow >= 0 And lNoteCol >= 0 Then
        txtNote.Visible = False
        txtNote.Text = ""
        VSFG.SetFocus
        lNoteRow = -1
        lNoteCol = -1
    End If
    
    ' no note to show? then bail out
    If r <= 0 Or c <= 0 Then Exit Sub
    If TypeName(VSFG.Cell(flexcpData, r, c)) <> "String" Then Exit Sub
    
    ' show new note
    'Dim nt As New frmFlexNote
    'nt.txtNote = VSFG.Cell(flexcpData, r, c)
    'nt.ShowNote VSFG, r, c
    
    lblNote.Caption = VSFG.Cell(flexcpData, r, c)
    
    Dim m_lRow As Long, m_lCol As Long
    m_lRow = VSFG.MouseRow
    m_lCol = VSFG.MouseCol

    
    ' calculate note position
    Dim fLeft!, fTop!, fWid!, fHei!
    fWid = 0
    fHei = 0
    With VSFG
        fLeft = .ColPos(c) + .ColWidth(c) + 200
        fTop = .RowPos(r) + 300
    End With
    
    
    
    txtNote.Width = lblNote.Width
    txtNote.Height = lblNote.Height
    ' calculate note size
    fWid = txtNote.Width + 300
    fHei = txtNote.Height + 150
    
    ' make sure note is not off the screen
    If fLeft + fWid > Screen.Width Then fLeft = fLeft - fWid - VSFG.ColWidth(c) - 200
    If fTop + fHei > Screen.Height - 300 Then fTop = Screen.Height - fHei - 300
    
    txtNote = lblNote
    ' show note (we stay up until deactivated)
    txtNote.Move fLeft, fTop, fWid, fHei
    'txtNote.SelStart = 32000
    txtNote.Visible = True
  
    
    
    ' save coordinates for next time
    lNoteRow = r
    lNoteCol = c
End Sub
