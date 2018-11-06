VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJustificarVisitas 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Justificación de Visitas"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   Icon            =   "frmJustificarVisitas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10770
   Begin VB.Frame fraEmpleado 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Selección de Empleado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
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
         Left            =   6600
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   6600
         TabIndex        =   28
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   210
            Value           =   -1  'True
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmJustificarVisitas.frx":030A
            Left            =   1320
            List            =   "frmJustificarVisitas.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   960
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
            TabIndex        =   6
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
            Format          =   57802755
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
            TabIndex        =   8
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
            Format          =   57802755
            CurrentDate     =   37463
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Inicial"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
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
            TabIndex        =   3
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo cmbEmpleado 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblEmpresas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   3743
      TabIndex        =   18
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   5453
      TabIndex        =   19
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Frame fraDetalle 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   10575
      Begin VB.TextBox txtObservacion 
         Height          =   645
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   6615
      End
      Begin VB.CheckBox chkEntrada 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Entrada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame fraEntrada 
         BackColor       =   &H00DDDDDD&
         Height          =   1095
         Left            =   6000
         TabIndex        =   24
         Top             =   360
         Width           =   3015
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   57802753
            CurrentDate     =   39615
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   57802753
            CurrentDate     =   39615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Entrada:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   660
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   720
            TabIndex        =   25
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.CheckBox chkSalida 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Salida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame fraSalida 
         BackColor       =   &H00DDDDDD&
         Height          =   1095
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   3015
         Begin MSComCtl2.DTPicker dtpFechaS 
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   57802753
            CurrentDate     =   39615
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   57802753
            CurrentDate     =   39615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   615
            TabIndex        =   23
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Salida:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   660
            Width           =   870
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3735
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   10095
         _cx             =   17806
         _cy             =   6588
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmJustificarVisitas.frx":039E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   1650
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmJustificarVisitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private strSql As String
Private FechaI As Variant
Private FechaF As Variant


Private Sub chkEntrada_Click()
    If chkEntrada.value = 1 Then
        dtpFecha.Enabled = True
        dtpEntrada.Enabled = True
        If chkSalida.value = 1 Then
            dtpFecha.value = dtpFechaS.value
        End If
    Else
        dtpFecha.Enabled = False
        dtpEntrada.Enabled = False
    End If
    'CrearHora dtpFecha
    CrearHora dtpEntrada
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

Private Sub cmbEmpleado_Change()
    Cargar
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

Private Sub chkSalida_Click()
    If chkSalida.value = 1 Then
        dtpFechaS.Enabled = True
        dtpSalida.Enabled = True
        If chkEntrada.value = 1 Then
            dtpFechaS.value = dtpFecha.value
        End If
    Else
        dtpFechaS.Enabled = False
        dtpSalida.Enabled = False
    End If
    'CrearHora dtpFechaS
    CrearHora dtpSalida
End Sub

Private Function diferenciaHoraria(ByVal beginTime As Date, endTime As Date) As String
    Dim lngTemp As Long
    Dim hrs, min, sec As Integer
    
    lngTemp = DateDiff("n", beginTime, endTime)
    hrs = Int(lngTemp / 60)
    
    If hrs <= 0 Then hrs = 0
    
    lngTemp = DateDiff("n", beginTime, endTime)
    min = Int(lngTemp Mod 60)
    
    If min <= 0 Then min = 0
    
    lngTemp = DateDiff("s", beginTime, endTime)
    sec = Int(lngTemp Mod 60)
    
    If sec <= 0 Then sec = 0
    
    diferenciaHoraria = hrs & ":" & min & ":" & sec
End Function


Private Sub cmdAceptar_Click()
    Dim Observacion As String
    Dim Multa As Double, Atrasado As Integer, Atraso As String, Registrado As Integer
    Dim horE As String, horS As String, salTol As String, entTol As String
    Dim extra As String
    
    extra = ""
    If cmbEmpleado.Text = "" Then
        MsgBox "Seleccione primero un Empleado", vbInformation, "Aceptar"
        Exit Sub
    End If
    Observacion = UCase(Trim(txtObservacion.Text))
    If Observacion = "" Then Observacion = "JUSTIFICACIÓN MANUAL"
    
    If chkSalida.value = 1 Then
        If chkEntrada.value = 1 Then
            
            If dtpEntrada.value <= dtpSalida.value Then
                MsgBox "Defina correctamente la hora de entrada de la visita", vbInformation, "Justificación de Visitas"
                Exit Sub
            End If
            
            strSql = " SELECT det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s'),TIME_FORMAT(det_hor_salida,'%H:%i:%s'),TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s'),TIME_FORMAT(det_hor_entrada+COALESCE(hor_tol_entrada_max,'00:00:00'),'%H:%i:%s') " & _
                     " FROM horario " & _
                     " INNER JOIN det_horario " & _
                     " ON det_horario.emp_codigo=horario.emp_codigo " & _
                     " AND det_horario.hor_codigo=horario.hor_codigo " & _
                     " INNER JOIN horario_empleado " & _
                     " ON horario_empleado.emp_codigo=horario.emp_codigo " & _
                     " AND horario_empleado.hor_codigo=horario.hor_codigo " & _
                     " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                     " AND horario_empleado.epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND det_horario.det_hor_dia='" & DatePart("w", dtpFechaS.value, vbMonday) - 1 & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN TIME_FORMAT(det_hor_entrada-COALESCE(hor_tol_entrada_min,'00:00:00'),'%H:%i:%s') AND TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s') " & _
                     " AND '" & Format(dtpEntrada.value, "HH:mm:SS") & "' BETWEEN TIME_FORMAT(det_hor_entrada-COALESCE(hor_tol_entrada_min,'00:00:00'),'%H:%i:%s') AND TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s') "
            clsSql.Ejecutar strSql
            
            If clsSql.adorec_Def.RecordCount = 0 Then
                MsgBox "No se encontró un horario para la fecha/hora definida", vbInformation, "Justificación de Visitas"
                Exit Sub
            End If
            
            horE = Format(clsSql.adorec_Def(1), "hh:mm:ss")
            horS = Format(clsSql.adorec_Def(2), "hh:mm:ss")
            salTol = Format(clsSql.adorec_Def(3), "hh:mm:ss")
            entTol = Format(clsSql.adorec_Def(4), "hh:mm:ss")
            
            strSql = " SELECT COUNT(*) " & _
                     " FROM asistencia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND ast_fecha='" & Format(dtpFechaS.value, "yyyy-MM-dd") & "' " & _
                     " AND ast_hora='" & horS & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN COALESCE(ast_entrada,ast_entrada_esp) AND COALESCE(ast_salida,ast_salida_esp) "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsSql.adorec_Def(0)) = 0 Then
                    MsgBox "No existe un registro de asistencia en la fecha/hora definida", vbInformation, "Justificación de Visitas"
                    Exit Sub
                End If
            End If
            
            
            strSql = " SELECT COUNT(*) " & _
                     " FROM visita " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND vis_fecha='" & Format(dtpFechaS.value, "yyyy-MM-dd") & "' " & _
                     " AND vis_hora='" & horS & "' " & _
                     " AND ('" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN COALESCE(vis_salida,'00:00:00') AND COALESCE(vis_entrada,vis_hora) " & _
                     " OR '" & Format(dtpEntrada.value, "HH:mm:SS") & "' BETWEEN COALESCE(vis_salida,'00:00:00') AND COALESCE(vis_entrada,vis_hora)) "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsSql.adorec_Def(0)) > 0 Then
                    MsgBox "Existe un registro de visita en la fecha/hora definida", vbInformation, "Justificación de Visitas"
                    Exit Sub
                End If
            End If
            
            
            Registrado = 0
            If Format(dtpEntrada.value, "hh:mm:ss") <> "00:00:00" Then
                Registrado = 1
            End If
            
            strSql = " INSERT INTO visita(emp_codigo,epl_codigo, vis_fecha, vis_dia, vis_hora, " & _
                     " vis_salida, vis_entrada, vis_fechamod,vis_usumod,vis_salida_tol,vis_observacion," & _
                     " vis_registrado) " & _
                     " VALUES ('" & strEmpresa & "','" & cmbEmpleado.BoundText & "', '" & Format(dtpFechaS.value, "yyyy-MM-dd") & "','" & DatePart("w", dtpFechaS.value, vbMonday) - 1 & "', '" & horS & "','" & _
                     Format(dtpSalida.value, "hh:mm:ss") & "','" & Format(dtpEntrada.value, "hh:mm:ss") & "',CURRENT_TIMESTAMP,'" & strUsuario & "','" & salTol & "','" & Observacion & "','" & _
                     Registrado & "') "
            clsSql.Ejecutar strSql, "M"
             
        Else
            strSql = " SELECT det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s'),TIME_FORMAT(det_hor_salida,'%H:%i:%s'),TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s'),TIME_FORMAT(det_hor_entrada+COALESCE(hor_tol_entrada_max,'00:00:00'),'%H:%i:%s') " & _
                     " FROM horario " & _
                     " INNER JOIN det_horario " & _
                     " ON det_horario.emp_codigo=horario.emp_codigo " & _
                     " AND det_horario.hor_codigo=horario.hor_codigo " & _
                     " INNER JOIN horario_empleado " & _
                     " ON horario_empleado.emp_codigo=horario.emp_codigo " & _
                     " AND horario_empleado.hor_codigo=horario.hor_codigo " & _
                     " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                     " AND horario_empleado.epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND det_horario.det_hor_dia='" & DatePart("w", dtpFechaS.value, vbMonday) - 1 & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN TIME_FORMAT(det_hor_entrada-COALESCE(hor_tol_entrada_min,'00:00:00'),'%H:%i:%s') AND TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s') "
            clsSql.Ejecutar strSql
            
            If clsSql.adorec_Def.RecordCount = 0 Then
                MsgBox "No se encontró un horario para la fecha/hora definida", vbInformation, "Justificación de Visitas"
                Exit Sub
            End If
            
            horE = Format(clsSql.adorec_Def(1), "hh:mm:ss")
            horS = Format(clsSql.adorec_Def(2), "hh:mm:ss")
            salTol = Format(clsSql.adorec_Def(3), "hh:mm:ss")
            entTol = Format(clsSql.adorec_Def(4), "hh:mm:ss")
            
            strSql = " SELECT COUNT(*) " & _
                     " FROM asistencia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND ast_fecha='" & Format(dtpFechaS.value, "yyyy-MM-dd") & "' " & _
                     " AND ast_hora='" & horS & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN COALESCE(ast_entrada,ast_entrada_esp) AND COALESCE(ast_salida,ast_salida_esp) "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsSql.adorec_Def(0)) = 0 Then
                    MsgBox "No existe un registro de asistencia en la fecha/hora definida", vbInformation, "Justificación de Visitas"
                    Exit Sub
                End If
            End If
                  
            strSql = " SELECT COUNT(*) " & _
                     " FROM visita " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND vis_fecha='" & Format(dtpFechaS.value, "yyyy-MM-dd") & "' " & _
                     " AND vis_hora='" & horS & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' BETWEEN COALESCE(vis_salida,'00:00:00') AND COALESCE(vis_entrada,vis_hora) "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsSql.adorec_Def(0)) > 0 Then
                    MsgBox "Existe un registro de visita que se cruza con la fecha/hora definida", vbInformation, "Justificación de Visitas"
                    Exit Sub
                End If
            End If
                               
            strSql = " SELECT TIME_FORMAT(vis_salida,'%H:%i:%s') " & _
                     " FROM visita " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND vis_fecha='" & Format(dtpFechaS.value, "yyyy-MM-dd") & "' " & _
                     " AND vis_hora='" & horS & "' " & _
                     " AND '" & Format(dtpSalida.value, "HH:mm:SS") & "' < COALESCE(vis_salida,vis_hora) " & _
                     " ORDER BY 1 LIMIT 1 "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                extra = Format(clsSql.adorec_Def(0), "hh:mm:ss")
            End If
            
            Registrado = 0
            
            If extra = "" Then
                strSql = " INSERT INTO visita(emp_codigo,epl_codigo, vis_fecha, vis_dia, vis_hora, " & _
                         " vis_salida,vis_fechamod,vis_usumod,vis_salida_tol,vis_observacion,vis_registrado) " & _
                         " VALUES ('" & strEmpresa & "','" & cmbEmpleado.BoundText & "', '" & Format(dtpFechaS.value, "yyyy-MM-dd") & "','" & DatePart("w", dtpFechaS.value, vbMonday) - 1 & "', '" & horS & "','" & _
                         Format(dtpSalida.value, "hh:mm:ss") & "',CURRENT_TIMESTAMP,'" & strUsuario & "','" & salTol & "','" & Observacion & "','" & Registrado & "') "
                clsSql.Ejecutar strSql, "M"
            Else
                strSql = " INSERT INTO visita(emp_codigo,epl_codigo, vis_fecha, vis_dia, vis_hora, " & _
                         " vis_salida,vis_entrada,vis_fechamod,vis_usumod,vis_salida_tol,vis_observacion,vis_registrado) " & _
                         " VALUES ('" & strEmpresa & "','" & cmbEmpleado.BoundText & "', '" & Format(dtpFechaS.value, "yyyy-MM-dd") & "','" & DatePart("w", dtpFechaS.value, vbMonday) - 1 & "', '" & horS & "','" & _
                         Format(dtpSalida.value, "hh:mm:ss") & "','" & extra & "',CURRENT_TIMESTAMP,'" & strUsuario & "','" & salTol & "','" & Observacion & "','" & Registrado & "') "
                clsSql.Ejecutar strSql, "M"
            End If
        End If
    Else
        strSql = " SELECT det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s'),TIME_FORMAT(det_hor_salida,'%H:%i:%s'),TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s'),TIME_FORMAT(det_hor_entrada+COALESCE(hor_tol_entrada_max,'00:00:00'),'%H:%i:%s') " & _
                 " FROM horario " & _
                 " INNER JOIN det_horario " & _
                 " ON det_horario.emp_codigo=horario.emp_codigo " & _
                 " AND det_horario.hor_codigo=horario.hor_codigo " & _
                 " INNER JOIN horario_empleado " & _
                 " ON horario_empleado.emp_codigo=horario.emp_codigo " & _
                 " AND horario_empleado.hor_codigo=horario.hor_codigo " & _
                 " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
                 " AND horario_empleado.epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                 " AND det_horario.det_hor_dia='" & DatePart("w", dtpFecha.value, vbMonday) - 1 & "' " & _
                 " AND '" & Format(dtpEntrada.value, "HH:mm:SS") & "' BETWEEN TIME_FORMAT(det_hor_entrada-COALESCE(hor_tol_entrada_min,'00:00:00'),'%H:%i:%s') AND TIME_FORMAT(det_hor_salida+COALESCE(hor_tol_salida_max,'00:00:00'),'%H:%i:%s') "
        clsSql.Ejecutar strSql
        
        If clsSql.adorec_Def.RecordCount = 0 Then
            MsgBox "No se encontró un horario para la fecha/hora defina", vbInformation, "Justificación de Visitas"
            Exit Sub
        End If
            
        horE = Format(clsSql.adorec_Def(1), "hh:mm:ss")
        horS = Format(clsSql.adorec_Def(2), "hh:mm:ss")
        salTol = Format(clsSql.adorec_Def(3), "hh:mm:ss")
        entTol = Format(clsSql.adorec_Def(4), "hh:mm:ss")
        
        strSql = " SELECT COUNT(*) " & _
                 " FROM asistencia " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                 " AND ast_fecha='" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
                 " AND ast_hora='" & horS & "' " & _
                 " AND '" & Format(dtpEntrada.value, "HH:mm:SS") & "' BETWEEN COALESCE(ast_entrada,ast_entrada_esp) AND COALESCE(ast_salida,ast_salida_esp) "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD0(clsSql.adorec_Def(0)) = 0 Then
                MsgBox "No existe un registro de asistencia en la fecha/hora definida", vbInformation, "Justificación de Visitas"
                Exit Sub
            End If
        End If
        
        strSql = " SELECT COUNT(*) " & _
                 " FROM visita " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                 " AND vis_fecha='" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
                 " AND vis_hora='" & horS & "' " & _
                 " AND vis_entrada IS NULL " & _
                 " AND '" & Format(dtpEntrada.value, "HH:mm:SS") & "'> vis_salida "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD0(clsSql.adorec_Def(0)) = 0 Then
                MsgBox "No existe o se cruza un registro de visita en la fecha/hora definida", vbInformation, "Justificación de Visitas"
                Exit Sub
            End If
        End If
        
        Registrado = 0
        If Format(dtpEntrada.value, "hh:mm:ss") <> "00:00:00" Then
            Registrado = 1
        End If
        
        strSql = "UPDATE visita SET " & _
                " vis_entrada = '" & Format(dtpEntrada.value, "hh:mm:ss") & "'," & _
                " vis_registrado='" & Registrado & "'," & _
                " vis_observacion= TRIM(CONCAT('" & Observacion & "',IF(vis_observacion IS NULL OR vis_observacion='','',CONCAT(' - ',vis_observacion) ))), " & _
                " vis_fechamod=CURRENT_TIMESTAMP, " & _
                " vis_usumod='" & strUsuario & "' " & _
                " WHERE emp_codigo='" & strEmpresa & "' " & _
                " AND epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                " AND vis_fecha='" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
                " AND vis_hora='" & horS & "' " & _
                " AND vis_entrada IS NULL "
        clsSql.Ejecutar strSql, "M"
    End If
    Cargar
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdMostrar_Click()
    Cargar
End Sub

Private Sub Cargar()
    Limpiar
    '" SELECT TIME_FORMAT(ast_hora,'%H:%i:%s'),asistencia.epl_codigo,DATE_FORMAT(ast_fecha,'%w'), "
    strSql = " SELECT DATE_FORMAT(vis_fecha,'%w'), " & _
             " vis_fecha,TIME_FORMAT(vis_salida,'%H:%i:%s'),TIME_FORMAT(vis_entrada,'%H:%i:%s')," & _
             " COALESCE(vis_observacion,'') " & _
             " FROM visita " & _
             " INNER JOIN empleado " & _
             " ON empleado.emp_codigo=visita.emp_codigo " & _
             " AND empleado.epl_codigo=visita.epl_codigo " & _
             " WHERE visita.emp_codigo='" & strEmpresa & "' " & _
             " AND empleado.epl_codigo LIKE '" & cmbEmpleado.BoundText & "' "
       
    If chkFiltroFecha.value = 1 Then
        If Option1.value = True Then
            strSql = strSql & " AND vis_fecha BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
        ElseIf Option2.value = True Then
           If chkFechas.value = 0 Then
                strSql = strSql & " AND vis_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha1 & "' "
            Else
                strSql = strSql & " AND vis_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
            End If
        End If
    End If
    
    strSql = strSql & " ORDER BY vis_fecha,vis_salida,vis_entrada "
    clsSql.Ejecutar strSql
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
    
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(i, 1)
        VSFG.TextMatrix(i, 1) = dia(FormatoD0(VSFG.TextMatrix(i, 1)))
    Next i

    VSFG.MergeCol(0) = True
    VSFG.MergeCol(1) = True
    VSFG.MergeCol(2) = True

    Dim aux As Long
    aux = 1
    For i = 1 To VSFG.Rows - 2
        If VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(i + 1, 0) Then
            VSFG.TextMatrix(i, 0) = aux
        Else
            VSFG.TextMatrix(i, 0) = CStr(aux)
            aux = aux + 1
        End If
        
    Next i
    If i <= VSFG.Rows - 1 Then
        VSFG.TextMatrix(i, 0) = CStr(aux)
    End If
End Sub

Private Function dia(d As Integer) As String
    If d = 0 Then
        dia = "Domingo"
    ElseIf d = 1 Then
        dia = "Lunes"
    ElseIf d = 2 Then
        dia = "Martes"
    ElseIf d = 3 Then
        dia = "Miércoles"
    ElseIf d = 4 Then
        dia = "Jueves"
    ElseIf d = 5 Then
        dia = "Viernes"
    ElseIf d = 6 Then
        dia = "Sábado"
    End If
End Function




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



Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    chkFiltroFecha.value = 1
    Option1.value = True
    
    CargaEmpleados
    
    
    
    Dim i As Integer
    Fecha1 = Format(HoyDia, "yyyy-mm-dd")
    Fecha2 = Format(HoyDia, "yyyy-mm-dd")
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    Cargar
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

Private Sub Limpiar()
    chkEntrada.value = 0
    chkEntrada_Click
    chkSalida.value = 0
    chkSalida_Click
    txtObservacion.Text = ""
    dtpFecha.value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaS.value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub CargaEmpleados()
    strSql = " SELECT CONCAT(epl_apellidos,' ',epl_nombres) as nombre,epl_codigo as codigo " & _
             " FROM empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 1 "
    clsSql.Ejecutar strSql
    Set cmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    cmbEmpleado.ListField = "nombre"
    cmbEmpleado.BoundColumn = "codigo"
End Sub


Private Sub CrearHora(objeto As DTPicker)
    objeto.Format = dtpCustom
    objeto.CustomFormat = "HH:mm:ss"
    objeto.UpDown = True
    objeto.Hour = "00"
    objeto.Minute = "00"
    objeto.Second = "00"
    objeto.value = "00:00:00"
End Sub








