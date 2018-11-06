VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConstruccionCalculoRedes 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contruccion y Calculo de Redes"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmConstruccionCalculoRedes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdEnvioCorreos 
      Caption         =   "Envio Correos"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenImpPDF 
      Caption         =   "Gen/Imp PDF"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdRecalificar 
         Caption         =   "Recalificar"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFechaFin 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   1980
      End
      Begin VB.TextBox txtFechaInicio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1980
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   255
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbCampania 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin de Facturación"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   13
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio de Facturación"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaña:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2390
      TabIndex        =   1
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4190
      TabIndex        =   0
      Top             =   6480
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4080
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   7980
      _cx             =   14076
      _cy             =   7197
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
      Rows            =   2
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConstruccionCalculoRedes.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
End
Attribute VB_Name = "frmConstruccionCalculoRedes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private TipoRed As Integer

Private Sub cmbCampania_Validate(Cancel As Boolean)
    strSql = " SELECT cam_fecha_fac_inicial, cam_fecha_fac_final " & _
             " FROM campaniafecha " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND concat(cam_anio,'-',cam_mes)='" & cmbCampania.BoundText & "' "
    clsCon_Def.Ejecutar strSql
    txtFechaInicio.Text = Left(clsCon_Def.adorec_Def("cam_fecha_fac_inicial"), 10)
    txtFechaFin.Text = Left(clsCon_Def.adorec_Def("cam_fecha_fac_final"), 10)
End Sub

Private Sub cmdCargar_Click()
    Dim clsRed As New clsConsulta
    clsRed.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT LTRIM(RTRIM(CONCAT(N1.per_apellido,' ',N1.per_nombre))) as NN1," & _
             " LTRIM(RTRIM(CONCAT(N2.per_apellido,' ',N2.per_nombre))) as NN2," & _
             " LTRIM(RTRIM(CONCAT(N3.per_apellido,' ',N3.per_nombre))) as NN3,LTRIM(RTRIM(CONCAT(N4.per_apellido,' ',N4.per_nombre))) as NN4," & _
             " LTRIM(RTRIM(CONCAT(N5.per_apellido,' ',N5.per_nombre))) as NN5,LTRIM(RTRIM(CONCAT(N6.per_apellido,' ',N6.per_nombre))) as NN6," & _
             " LTRIM(RTRIM(CONCAT(N7.per_apellido,' ',N7.per_nombre))) as NN7,LTRIM(RTRIM(CONCAT(N8.per_apellido,' ',N8.per_nombre))) as NN8," & _
             " LTRIM(RTRIM(CONCAT(N9.per_apellido,' ',N9.per_nombre))) as NN9,LTRIM(RTRIM(CONCAT(N10.per_apellido,' ',N10.per_nombre))) as NN10," & _
             " li.per_codigo as li_per_codigo,concat(li.per_apellido,' ',li.per_nombre) as lider," & _
             " COALESCE(mul_nombre,'EJECUTIVO') as nivel," & _
             " sum(rc.red_cam_venta_directa+rc.red_cam_venta_directa_no_comi+IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)+IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta_no_comi)) as ventaneta," & _
             " sum(rc.red_cam_activo_directo+IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_activo_indirecto)) as activos," & _
             " sum(rc.red_cam_venta_directa_no_comi+IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta_no_comi)) as dctos," & _
             " ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2) as comision," & _
             " ROUND(ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2)*0.12,2) as IVA," & _
             " ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2) +" & _
             " ROUND(ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2)*0.12,2) as TotalFAC," & _
             " ROUND(ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2)*0.08,2) as RETFUENTE," & _
             " ROUND(ROUND(ROUND(sum(rc.red_cam_venta_directa*det_red_campania.mul_comision/100)+sum(IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta)*det_red_campania.mul_comision/100),2)*0.12,2)*0.7,2) as RETIVA," & _
             " CONCAT(li.per_email,'; ',IIF(red_campania.per_codigo<>N10.per_codigo AND N10.per_email<>'',N10.per_email, IIF(red_campania.per_codigo<>N9.per_codigo AND N9.per_email<>'',N9.per_email, IIF(red_campania.per_codigo<>N8.per_codigo AND N8.per_email<>'',N8.per_email, IIF(red_campania.per_codigo<>N7.per_codigo AND N7.per_email<>'',N7.per_email, IIF(red_campania.per_codigo<>N6.per_codigo AND N6.per_email<>'',N6.per_email, IIF(red_campania.per_codigo<>N5.per_codigo AND N5.per_email<>'',N5.per_email, IIF(red_campania.per_codigo<>N4.per_codigo AND N4.per_email<>'',N4.per_email, IIF(red_campania.per_codigo<>N3.per_codigo AND N3.per_email<>'',N3.per_email, IIF(red_campania.per_codigo<>N2.per_codigo AND N2.per_email<>'',N2.per_email, IIF(red_campania.per_codigo<>N1.per_codigo AND N1.per_email<>'',N1.per_email,''))))))))))) AS email"
    strSql = strSql & " FROM red_campania inner join persona li " & _
             " on red_campania.emp_codigo=li.emp_codigo" & _
             " and red_campania.per_codigo=li.per_codigo" & _
             " inner join det_red_campania" & _
             " on red_campania.emp_codigo=det_red_campania.emp_codigo" & _
             " and red_campania.cam_anio=det_red_campania.cam_anio" & _
             " and red_campania.cam_mes=det_red_campania.cam_mes" & _
             " and red_campania.per_codigo=det_red_campania.per_papa_codigo" & _
             " inner join persona em" & _
             " on det_red_campania.emp_codigo=em.emp_codigo" & _
             " and det_red_campania.per_codigo=em.per_codigo" & _
             " inner join red_campania rc" & _
             " on det_red_campania.emp_codigo=rc.emp_codigo" & _
             " and det_red_campania.per_codigo=rc.per_codigo" & _
             " and red_campania.cam_anio=rc.cam_anio" & _
             " and red_campania.cam_mes=rc.cam_mes" & _
             " inner join persona n1 on li.emp_codigo=n1.emp_codigo" & _
             " and li.per_codigo_ref=n1.per_codigo" & _
             " left join multinivel" & _
             " on red_campania.emp_codigo=multinivel.emp_codigo" & _
             " and red_campania.mul_codigo=multinivel.mul_codigo"
    strSql = strSql & " left join cobrador c on li.emp_codigo=c.emp_codigo" & _
             " and li.cob_codigo=c.cob_codigo" & _
             " left join persona n2 on li.emp_codigo=n2.emp_codigo and li.per_codigo_ref2=n2.per_codigo" & _
             " left join persona n3 on li.emp_codigo=n3.emp_codigo and li.per_codigo_ref3=n3.per_codigo" & _
             " left join persona n4 on li.emp_codigo=n4.emp_codigo and li.per_codigo_ref4=n4.per_codigo" & _
             " left join persona n5 on li.emp_codigo=n5.emp_codigo and li.per_codigo_ref5=n5.per_codigo" & _
             " left join persona n6 on li.emp_codigo=n6.emp_codigo and li.per_codigo_ref6=n6.per_codigo" & _
             " left join persona n7 on li.emp_codigo=n7.emp_codigo and li.per_codigo_ref7=n7.per_codigo" & _
             " left join persona n8 on li.emp_codigo=n8.emp_codigo and li.per_codigo_ref8=n8.per_codigo" & _
             " left join persona n9 on li.emp_codigo=n9.emp_codigo and li.per_codigo_ref9=n9.per_codigo" & _
             " left join persona n10 on li.emp_codigo=n10.emp_codigo and li.per_codigo_ref10=n10.per_codigo" & _
             " where red_campania.emp_codigo='RYB'" & _
             " and red_campania.cam_anio='" & Left(cmbCampania.BoundText, 4) & "'" & _
             " and red_campania.cam_mes='" & Right(cmbCampania.BoundText, 2) & "'" & _
             " and (rc.red_cam_venta_directa!=0 or rc.red_cam_venta_directa_no_comi!=0" & _
             " or rc.red_cam_venta_indirecta!=0 or rc.red_cam_venta_indirecta_no_comi!=0" & _
             " or rc.red_cam_activo_directo!=0 or rc.red_cam_activo_indirecto!=0)"
    strSql = strSql & " group by LTRIM(RTRIM(CONCAT(N1.per_apellido,' ',N1.per_nombre)))," & _
             " LTRIM(RTRIM(CONCAT(N2.per_apellido,' ',N2.per_nombre)))," & _
             " LTRIM(RTRIM(CONCAT(N3.per_apellido,' ',N3.per_nombre))),LTRIM(RTRIM(CONCAT(N4.per_apellido,' ',N4.per_nombre)))," & _
             " LTRIM(RTRIM(CONCAT(N5.per_apellido,' ',N5.per_nombre))),LTRIM(RTRIM(CONCAT(N6.per_apellido,' ',N6.per_nombre)))," & _
             " LTRIM(RTRIM(CONCAT(N7.per_apellido,' ',N7.per_nombre))),LTRIM(RTRIM(CONCAT(N8.per_apellido,' ',N8.per_nombre)))," & _
             " LTRIM(RTRIM(CONCAT(N9.per_apellido,' ',N9.per_nombre))),LTRIM(RTRIM(CONCAT(N10.per_apellido,' ',N10.per_nombre)))," & _
             " li.per_codigo,concat(li.per_apellido,' ',li.per_nombre)," & _
             " COALESCE(mul_nombre,'EJECUTIVO')," & _
             " CONCAT(li.per_email,'; ',IIF(red_campania.per_codigo<>N10.per_codigo AND N10.per_email<>'',N10.per_email, IIF(red_campania.per_codigo<>N9.per_codigo AND N9.per_email<>'',N9.per_email, IIF(red_campania.per_codigo<>N8.per_codigo AND N8.per_email<>'',N8.per_email, IIF(red_campania.per_codigo<>N7.per_codigo AND N7.per_email<>'',N7.per_email, IIF(red_campania.per_codigo<>N6.per_codigo AND N6.per_email<>'',N6.per_email, IIF(red_campania.per_codigo<>N5.per_codigo AND N5.per_email<>'',N5.per_email, IIF(red_campania.per_codigo<>N4.per_codigo AND N4.per_email<>'',N4.per_email, IIF(red_campania.per_codigo<>N3.per_codigo AND N3.per_email<>'',N3.per_email, IIF(red_campania.per_codigo<>N2.per_codigo AND N2.per_email<>'',N2.per_email, IIF(red_campania.per_codigo<>N1.per_codigo AND N1.per_email<>'',N1.per_email,'')))))))))))" & _
             " ORDER BY NN1,NN2,NN3,NN4,NN5,NN6,NN7,NN8,NN9,NN10"
    clsRed.Ejecutar strSql
    Set VSFG.DataSource = clsRed.adorec_Def.DataSource
End Sub

Private Sub cmdEnvioCorreos_Click()
    Dim Destino As String
    Dim Ruta As String
    Dim i As Long
    Destino = Buscar_Carpeta(Me.hwnd, "Carpetas a Subir")
    '404
    For i = 1 To VSFG.Rows - 1
        VSFG.ShowCell i, 12
        VSFG.Select i, 12
        If Len(Trim(VSFG.TextMatrix(i, 22))) > 3 And VSFG.TextMatrix(i, 17) <> 0 Then
            Ruta = Replace(Destino & "\" & Trim(IIf(VSFG.TextMatrix(i, 1) <> "", VSFG.TextMatrix(i, 1) & "\", "")) & _
                                           Trim(IIf(VSFG.TextMatrix(i, 2) <> "", VSFG.TextMatrix(i, 2) & "\", "")) & _
                                           Trim(IIf(VSFG.TextMatrix(i, 3) <> "", VSFG.TextMatrix(i, 3) & "\", "")), " ", "_")
            EnviarMail NombreComercial & " Seguimiento", CorreoServicioAlCliente, VSFG.TextMatrix(i, 12), Trim(VSFG.TextMatrix(i, 22)), "", "Comisiones " & cmbCampania.Text, _
                            "Estimad@" & vbNewLine & _
                            VSFG.TextMatrix(i, 12) & vbNewLine & _
                            "Adjunto encontrarás el Reporte de Comisiones donde se encuentran los valores con los que debes llenar tu factura para el pago de la comisión." & vbNewLine & _
                            "Si tiene alguna novedad comunicate con tu ejecutivo de servicio al cliente." & vbNewLine & _
                            "Saludos Cordiales" & vbNewLine & _
                            "Servicio al Cliente" & vbNewLine & _
                            NombreComercial, Replace(Ruta & Trim(VSFG.TextMatrix(i, 12)) & ".pdf", " ", "_")
                            j = j + 1
        End If
    Next i
    MsgBox "Envios terminados " & Now
    
End Sub

Private Sub cmdProcesar_Click()
    Dim tv As Double
    Dim tv2 As Double
    Dim ta As Long
    Dim ta2 As Long
    'Persona
    TipoRed = 0
    'venta
    'TipoRed = 1
    RevisarRed Left(cmbCampania.BoundText, 4), Right(cmbCampania.BoundText, 2)
    CalificarNivel Left(cmbCampania.BoundText, 4), Right(cmbCampania.BoundText, 2)
    CrearComision Left(cmbCampania.BoundText, 4), Right(cmbCampania.BoundText, 2)
    MsgBox Now
End Sub

Private Sub CrearComision(CamAnio As String, CamMes As String)
    Dim clsRed As New clsConsulta
    Dim clsMod As New clsConsulta
    Dim clsMul As New clsConsulta
    Dim Comi As Double
    clsRed.Inicializar AdoConn, AdoConnMaster
    clsMod.Inicializar AdoConn, AdoConnMaster
    clsMul.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT mul_codigo,mul_activos_min,mul_activos_max,mul_ventas_min,mul_ventas_max,mul_comision " & _
             " FROM multinivel " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY mul_activos_min DESC,mul_ventas_min DESC "
    clsMul.Ejecutar strSql
    strSql = " DELETE " & _
             " FROM det_red_campania " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
    clsRed.Ejecutar strSql, "M"
    strSql = " SELECT red_campania.per_codigo,per_papa_codigo," & _
             " red_campania.mul_codigo," & _
             " COALESCE(mul_comision,0) as mul_comision " & _
             " FROM red_campania INNER JOIN persona ON red_campania.emp_codigo=persona.emp_codigo " & _
             " AND red_campania.per_codigo=persona.per_codigo " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " LEFT JOIN multinivel ON red_campania.emp_codigo=multinivel.emp_codigo" & _
             " AND red_campania.mul_codigo=multinivel.mul_codigo" & _
             " WHERE red_campania.emp_codigo='" & strEmpresa & "'" & _
             " AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'" & _
             " ORDER BY red_cam_usado "
    clsRed.Ejecutar strSql
    While Not clsRed.adorec_Def.EOF
        strSql = " INSERT INTO det_red_campania (emp_codigo, per_papa_codigo, per_codigo, " & _
                 " cam_anio, cam_mes, mul_codigo, mul_comision, " & _
                 " det_red_cam_fechamod , det_red_cam_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & clsRed.adorec_Def("per_codigo") & "','" & clsRed.adorec_Def("per_codigo") & "'," & _
                 " '" & CamAnio & "','" & CamMes & "','" & clsRed.adorec_Def("mul_codigo") & "','" & clsRed.adorec_Def("mul_comision") & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsMod.Ejecutar strSql, "M"
        If clsRed.adorec_Def("per_papa_codigo") <> "%" Then
            strSql = " SELECT COALESCE(mul_comision,0) as mul_comision " & _
                     " FROM red_campania LEFT JOIN multinivel ON red_campania.emp_codigo=multinivel.emp_codigo" & _
                     " AND red_campania.mul_codigo=multinivel.mul_codigo" & _
                     " WHERE red_campania.emp_codigo='" & strEmpresa & "'" & _
                     " AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'" & _
                     " AND per_codigo='" & clsRed.adorec_Def("per_papa_codigo") & "'" & _
                     " ORDER BY red_cam_usado "
            clsMod.Ejecutar strSql
            
            If clsMod.adorec_Def.RecordCount > 0 Then
                Comi = clsMod.adorec_Def("mul_comision")
            Else
                Comi = 0
            End If
            strSql = " INSERT INTO det_red_campania (emp_codigo, per_papa_codigo, per_codigo, " & _
                     " cam_anio, cam_mes, mul_codigo, mul_comision, " & _
                     " det_red_cam_fechamod , det_red_cam_usumod) " & _
                     " VALUES('" & strEmpresa & "','" & clsRed.adorec_Def("per_papa_codigo") & "','" & clsRed.adorec_Def("per_codigo") & "'," & _
                     " '" & CamAnio & "','" & CamMes & "','" & clsRed.adorec_Def("mul_codigo") & "','" & FormatoD2(Comi) - FormatoD2(clsRed.adorec_Def("mul_comision")) & "'," & _
                     " CURRENT_TIMESTAMP,'" & strUsuario & "')"
            clsMod.Ejecutar strSql, "M"
            
        End If
        clsRed.adorec_Def.MoveNext
        
    Wend
    MsgBox "COMISION CREADAS"
End Sub

Private Sub CalificarNivel(CamAnio As String, CamMes As String)
    Dim clsRed As New clsConsulta
    Dim clsMod As New clsConsulta
    Dim clsMul As New clsConsulta
    Dim Nivel As String
    Dim pasoActivo As Boolean
    Dim pasoVenta As Boolean
    clsRed.Inicializar AdoConn, AdoConnMaster
    clsMod.Inicializar AdoConn, AdoConnMaster
    clsMul.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT mul_codigo,mul_activos_min,mul_activos_max,mul_ventas_min,mul_ventas_max,mul_comision " & _
             " FROM multinivel " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY mul_activos_min DESC,mul_ventas_min DESC "
    clsMul.Ejecutar strSql
    strSql = " SELECT red_campania.per_codigo," & _
             " red_cam_venta_directa+red_cam_venta_indirecta+red_cam_venta_directa_no_comi+red_cam_venta_indirecta_no_comi as totVenta," & _
             " red_cam_activo_directo+red_cam_activo_indirecto as totActivo " & _
             " FROM red_campania INNER JOIN persona ON red_campania.emp_codigo=persona.emp_codigo " & _
             " AND red_campania.per_codigo=persona.per_codigo " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " WHERE red_campania.emp_codigo='" & strEmpresa & "'" & _
             " AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
    clsRed.Ejecutar strSql
    
    While Not clsRed.adorec_Def.EOF
        Nivel = ""
        clsMul.adorec_Def.MoveFirst
        pasoActivo = False
        pasoVenta = False
        While Not clsMul.adorec_Def.EOF
            If clsMul.adorec_Def("mul_activos_min") <= clsRed.adorec_Def("totActivo") And clsRed.adorec_Def("totActivo") <= clsMul.adorec_Def("mul_activos_max") Then
                pasoActivo = True
            End If
            If clsMul.adorec_Def("mul_ventas_min") <= clsRed.adorec_Def("totVenta") And clsRed.adorec_Def("totVenta") <= clsMul.adorec_Def("mul_ventas_max") Then
                pasoVenta = True
            End If
            If pasoActivo = True And pasoVenta = True Then
                Nivel = clsMul.adorec_Def("mul_codigo")
                clsMul.adorec_Def.MoveLast
            End If
            clsMul.adorec_Def.MoveNext
        Wend
        strSql = " UPDATE red_campania " & _
                 " SET mul_codigo='" & Nivel & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & clsRed.adorec_Def("per_codigo") & "' " & _
                 " AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
        clsMod.Ejecutar strSql, "M"
        clsRed.adorec_Def.MoveNext
    Wend
End Sub

Private Sub RevisarRed(CamAnio As String, CamMes As String)
    Dim clsRed As New clsConsulta
    Dim clsMod As New clsConsulta
    Dim TVentaDirecta As Double
    Dim TVentaDirectaNC As Double
    Dim TActivoDirecto As Long
    Dim strLider As String
    Dim strPapa As String
    Dim Seguir As Boolean
    clsRed.Inicializar AdoConn, AdoConnMaster
    clsMod.Inicializar AdoConn, AdoConnMaster
    If strLider <> "" Then
        VentaYActivosDirecto strLider, TVentaDirecta, TVentaDirectaNC, TActivoDirecto
    End If
    If TipoRed = 0 Then
        strSql = " SELECT per_codigo,CASE WHEN per_codigo=per_codigo_ref THEN '%' " & _
                 " WHEN per_codigo=per_codigo_ref2 THEN IIF(per_codigo_ref<>'',per_codigo_ref,'%') " & _
                 " WHEN per_codigo=per_codigo_ref3 THEN CASE WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END " & _
                 " WHEN per_codigo=per_codigo_ref4 THEN CASE WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END " & _
                 " WHEN per_codigo=per_codigo_ref5 THEN CASE WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END " & _
                 " WHEN per_codigo=per_codigo_ref6 THEN CASE WHEN per_codigo_ref5<>'' THEN per_codigo_ref5 WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END" & _
                 " WHEN per_codigo=per_codigo_ref7 THEN CASE WHEN per_codigo_ref6<>'' THEN per_codigo_ref6 WHEN per_codigo_ref5<>'' THEN per_codigo_ref5 WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END" & _
                 " WHEN per_codigo=per_codigo_ref8 THEN CASE WHEN per_codigo_ref7<>'' THEN per_codigo_ref7 WHEN per_codigo_ref6<>'' THEN per_codigo_ref6 WHEN per_codigo_ref5<>'' THEN per_codigo_ref5 WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END " & _
                 " WHEN per_codigo=per_codigo_ref9 THEN CASE WHEN per_codigo_ref8<>'' THEN per_codigo_ref8 WHEN per_codigo_ref7<>'' THEN per_codigo_ref7 WHEN per_codigo_ref6<>'' THEN per_codigo_ref6 WHEN per_codigo_ref5<>'' THEN per_codigo_ref5 WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END" & _
                 " WHEN per_codigo=per_codigo_ref10 THEN CASE WHEN per_codigo_ref9<>'' THEN per_codigo_ref9 WHEN per_codigo_ref8<>'' THEN per_codigo_ref8 WHEN per_codigo_ref7<>'' THEN per_codigo_ref7 WHEN per_codigo_ref6<>'' THEN per_codigo_ref6 WHEN per_codigo_ref5<>'' THEN per_codigo_ref5 WHEN per_codigo_ref4<>'' THEN per_codigo_ref4 WHEN per_codigo_ref3<>'' THEN per_codigo_ref3 WHEN per_codigo_ref2<>'' THEN per_codigo_ref2 WHEN per_codigo_ref<>'' THEN per_codigo_ref ELSE '%' END ELSE '%' END as papa " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND (per_es_gz=1 OR per_es_di=1 OR per_es_em=1 OR per_es_ee=1 " & _
                 " OR per_es_n5=1 OR per_es_n6=1 OR per_es_n7=1 OR per_es_n8=1 OR per_es_n9=1 OR per_es_n10=1) " & _
                 " ORDER BY per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4, " & _
                 " per_codigo_ref5, per_codigo_ref6, per_codigo_ref7, per_codigo_ref8, per_codigo_ref9, per_codigo_ref10"
    ElseIf TipoRed = 1 Then
        strSql = " SELECT per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                 " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
                 " per_codigo_ref9,per_codigo_ref10," & _
                 " IIF(SUM(TMovimiento)>1,1,0) as ac, SUM(TMovimiento) as ve, SUM(TMovimientoNC) as veNC"
        strSql = strSql & " FROM (" & _
                 " SELECT persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4," & _
                 " persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8," & _
                 " persona.per_codigo_ref9,pe.rsona.per_codigo_ref10," & _
                 " SUM(ROUND((det_egr_cantidad*det_egr_precio-det_egr_dcto)*(COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimiento," & _
                 " SUM(ROUND((det_egr_cantidad*det_egr_precio-det_egr_dcto)*(1-*COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimientoNC" & _
                 " FROM persona INNER JOIN egreso ON persona.emp_codigo=egreso.emp_codigo" & _
                 " AND persona.per_codigo=egreso.per_codigo" & _
                 " AND egreso.tip_egr_codigo IN ('FAC','NOT')" & _
                 " AND egreso.egr_fecha BETWEEN '" & txtFechaInicio.Text & "' AND '" & txtFechaFin.Text & "'" & _
                 " AND egreso.egr_anulado=0" & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo" & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo" & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo" & _
                 " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo" & _
                 " AND det_egreso.prd_codigo=producto.prd_codigo"
        strSql = strSql & " LEFT JOIN producto_comision pc ON producto.emp_codigo=pc.emp_codigo " & _
                 " AND producto.prd_codigo=pc.prd_codigo " & _
                 " AND pc.cam_anio='" & CamAnio & "' " & _
                 " AND pc.cam_mes='" & CamMes & "' " & _
                 " AND persona.tip_ped_codigo=pc.tip_ped_codigo "
        strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " AND persona.cat_p_tipo='C'" & _
                 " GROUP BY persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4," & _
                 " persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8," & _
                 " Persona.per_codigo_ref9 , Persona.per_codigo_ref10"
        strSql = strSql & " Union " & _
                 " SELECT persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4," & _
                 " persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8," & _
                 " persona.per_codigo_ref9,persona.per_codigo_ref10," & _
                 " -1*SUM(ROUND((det_ing_cantidad*det_ing_precio-det_ing_dcto)*(COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimiento," & _
                 " -1*SUM(ROUND((det_ing_cantidad*det_ing_precio-det_ing_dcto)*(1-*COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimientoNC" & _
                 " FROM persona INNER JOIN ingreso ON persona.emp_codigo=ingreso.emp_codigo" & _
                 " AND persona.per_codigo=ingreso.per_codigo" & _
                 " AND ingreso.tip_ing_codigo='DCL'" & _
                 " AND ingreso.ing_fecha BETWEEN '" & txtFechaInicio.Text & "' AND '" & txtFechaFin.Text & "'" & _
                 " AND ingreso.ing_anulado=0" & _
                 " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo" & _
                 " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo" & _
                 " AND ingreso.ing_codigo=det_ingreso.ing_codigo" & _
                 " INNER JOIN producto ON det_ingreso.emp_codigo=producto.emp_codigo" & _
                 " AND det_ingreso.prd_codigo=producto.prd_codigo "
        strSql = strSql & " LEFT JOIN producto_comision pc ON producto.emp_codigo=pc.emp_codigo " & _
                 " AND producto.prd_codigo=pc.prd_codigo " & _
                 " AND pc.cam_anio='" & CamAnio & "' " & _
                 " AND pc.cam_mes='" & CamMes & "' " & _
                 " AND persona.tip_ped_codigo=pc.tip_ped_codigo "
        strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " AND persona.cat_p_tipo='C'" & _
                 " GROUP BY persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4," & _
                 " persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8," & _
                 " Persona.per_codigo_ref9 , Persona.per_codigo_ref10" & _
                 " ) TVenta"
        strSql = strSql & " GROUP BY per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                 " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
                 " per_codigo_ref9,per_codigo_ref10"
    End If
    clsRed.Ejecutar strSql
    If clsRed.adorec_Def.RecordCount > 0 Then
        While Not clsRed.adorec_Def.EOF
'            MsgBox "AAA"

            If TipoRed = 0 Then
                strLider = clsRed.adorec_Def("per_codigo")
                strPapa = clsRed.adorec_Def("papa")
            ElseIf TipoRed = 1 Then
                If clsRed.adorec_Def("per_codigo_ref10") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref10")
                    If clsRed.adorec_Def("per_codigo_ref9") <> "" And clsRed.adorec_Def("per_codigo_ref9") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref9")
                    ElseIf clsRed.adorec_Def("per_codigo_ref8") <> "" And clsRed.adorec_Def("per_codigo_ref8") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref8")
                    ElseIf clsRed.adorec_Def("per_codigo_ref7") <> "" And clsRed.adorec_Def("per_codigo_ref7") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref7")
                    ElseIf clsRed.adorec_Def("per_codigo_ref6") <> "" And clsRed.adorec_Def("per_codigo_ref6") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref6")
                    ElseIf clsRed.adorec_Def("per_codigo_ref5") <> "" And clsRed.adorec_Def("per_codigo_ref5") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref5")
                    ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref9") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref9")
                    If clsRed.adorec_Def("per_codigo_ref8") <> "" And clsRed.adorec_Def("per_codigo_ref8") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref8")
                    ElseIf clsRed.adorec_Def("per_codigo_ref7") <> "" And clsRed.adorec_Def("per_codigo_ref7") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref7")
                    ElseIf clsRed.adorec_Def("per_codigo_ref6") <> "" And clsRed.adorec_Def("per_codigo_ref6") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref6")
                    ElseIf clsRed.adorec_Def("per_codigo_ref5") <> "" And clsRed.adorec_Def("per_codigo_ref5") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref5")
                    ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref8") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref8")
                    If clsRed.adorec_Def("per_codigo_ref7") <> "" And clsRed.adorec_Def("per_codigo_ref7") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref7")
                    ElseIf clsRed.adorec_Def("per_codigo_ref6") <> "" And clsRed.adorec_Def("per_codigo_ref6") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref6")
                    ElseIf clsRed.adorec_Def("per_codigo_ref5") <> "" And clsRed.adorec_Def("per_codigo_ref5") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref5")
                    ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref7") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref7")
                    If clsRed.adorec_Def("per_codigo_ref6") <> "" And clsRed.adorec_Def("per_codigo_ref6") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref6")
                    ElseIf clsRed.adorec_Def("per_codigo_ref5") <> "" And clsRed.adorec_Def("per_codigo_ref5") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref5")
                    ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref6") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref6")
                    If clsRed.adorec_Def("per_codigo_ref5") <> "" And clsRed.adorec_Def("per_codigo_ref5") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref5")
                    ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref5") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref5")
                    If clsRed.adorec_Def("per_codigo_ref4") <> "" And clsRed.adorec_Def("per_codigo_ref4") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref4")
                    ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref4") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref4")
                    If clsRed.adorec_Def("per_codigo_ref3") <> "" And clsRed.adorec_Def("per_codigo_ref3") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref3")
                    ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref3") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref3")
                    If clsRed.adorec_Def("per_codigo_ref2") <> "" And clsRed.adorec_Def("per_codigo_ref2") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref2")
                    ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref2") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref2")
                    If clsRed.adorec_Def("per_codigo_ref") <> "" And clsRed.adorec_Def("per_codigo_ref") <> strLider Then
                        strPapa = clsRed.adorec_Def("per_codigo_ref")
                    Else
                        strPapa = "%"
                    End If
                ElseIf clsRed.adorec_Def("per_codigo_ref") <> "" Then
                    strLider = clsRed.adorec_Def("per_codigo_ref")
                    strPapa = "%"
                End If
            End If
            
'            If strLider = "C117713" Or strPapa = "C117713" Then
'                MsgBox "AAA"
'            End If
            strSql = " SELECT emp_codigo FROM red_campania WHERE emp_codigo='" & strEmpresa & "' AND per_codigo='" & strLider & "' AND per_papa_codigo='" & strPapa & "' AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
            clsMod.Ejecutar strSql, "M"
            If clsMod.adorec_Def.RecordCount = 0 Then
'                If strLider = "C104942" Then
'                    MsgBox "AAA"
'                End If
                VentaYActivosDirecto strLider, TVentaDirecta, TVentaDirectaNC, TActivoDirecto
                strSql = " INSERT INTO red_campania(emp_codigo, per_codigo, per_papa_codigo, cam_anio, cam_mes," & _
                         " red_cam_venta_directa,red_cam_venta_directa_no_comi, red_cam_activo_directo, " & _
                         " red_cam_venta_indirecta,red_cam_venta_indirecta_no_comi, red_cam_activo_indirecto,red_cam_usado, " & _
                         " red_cam_fechamod,red_cam_usumod) " & _
                         " VALUES('" & strEmpresa & "','" & strLider & "','" & strPapa & "','" & CamAnio & "','" & CamMes & "'," & _
                         " '" & TVentaDirecta & "','" & TVentaDirectaNC & "','" & TActivoDirecto & "'," & _
                         " 0,0,0,0," & _
                         " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsMod.Ejecutar strSql, "M"
            End If
            clsRed.adorec_Def.MoveNext
        Wend
    End If
    Seguir = True
    Dim i As Long
    i = 1
    While Seguir = True
        strSql = " SELECT r1.per_codigo,r1.per_papa_codigo, r1.red_cam_venta_directa+r1.red_cam_venta_indirecta as vd, r1.red_cam_venta_directa_no_comi+r1.red_cam_venta_indirecta_no_comi as vdNC, " & _
                 " r1.red_cam_activo_directo+r1.red_cam_activo_indirecto as ad" & _
                 " FROM red_campania r1 LEFT JOIN red_campania r2 " & _
                 " ON r1.emp_codigo=r2.emp_codigo " & _
                 " AND r1.per_codigo=r2.per_papa_codigo " & _
                 " AND r1.cam_anio=r2.cam_anio " & _
                 " AND r2.red_cam_usado=0 AND r1.cam_mes=r2.cam_mes AND r2.cam_anio='" & CamAnio & "' AND r2.cam_mes='" & CamMes & "'" & _
                 " WHERE r1.emp_codigo='" & strEmpresa & "'" & _
                 " AND r1.red_cam_usado=0 AND r1.cam_anio='" & CamAnio & "' AND r1.cam_mes='" & CamMes & "'" & _
                 " AND r2.emp_codigo IS NULL"
        clsRed.Ejecutar strSql
        If clsRed.adorec_Def.RecordCount > 0 Then
            While Not clsRed.adorec_Def.EOF
                strSql = " UPDATE red_campania " & _
                         " SET red_cam_venta_indirecta=red_cam_venta_indirecta+" & clsRed.adorec_Def("vd") & ", " & _
                         " red_cam_venta_indirecta_no_comi=red_cam_venta_indirecta_no_comi+" & clsRed.adorec_Def("vdNC") & ", " & _
                         " red_cam_activo_indirecto=red_cam_activo_indirecto+" & clsRed.adorec_Def("ad") & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND per_codigo='" & clsRed.adorec_Def("per_papa_codigo") & "' AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
                clsMod.Ejecutar strSql, "M"
                strSql = " UPDATE red_campania " & _
                         " SET red_cam_usado='" & i & "'" & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND per_codigo='" & clsRed.adorec_Def("per_codigo") & "' AND cam_anio='" & CamAnio & "' AND cam_mes='" & CamMes & "'"
                clsMod.Ejecutar strSql, "M"
                clsRed.adorec_Def.MoveNext
            Wend
        Else
            Seguir = False
        End If
        i = i + 1
    Wend
End Sub

Private Sub VentaYActivosDirecto(strLider As String, ByRef TVentaDirecta As Double, ByRef TVentaDirectaNC As Double, ByRef TActivoDirecto As Long)
    Dim clsVentas As New clsConsulta
    clsVentas.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT emp_codigo, COALESCE(SUM(ac),0) as tac, COALESCE(sum(ve),0) as tve, COALESCE(sum(veNC),0) as tveNC " & _
             " FROM ( " & _
                " SELECT emp_codigo,per_codigo,IIF(SUM(TMovimiento)>1,1,0) as ac, SUM(TMovimiento) as ve, SUM(TMovimientoNC) as veNC " & _
                " FROM ( " & _
                       " SELECT persona.emp_codigo,persona.per_codigo," & _
                       " SUM(ROUND((det_egr_cantidad*det_egr_precio-det_egr_dcto)*(COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimiento," & _
                       " SUM(ROUND((det_egr_cantidad*det_egr_precio-det_egr_dcto)*(1-COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimientoNC" & _
                       " FROM persona INNER JOIN egreso ON persona.emp_codigo=egreso.emp_codigo " & _
                       " AND persona.per_codigo=egreso.per_codigo " & _
                       " AND egreso.tip_egr_codigo IN ('FAC','NOT') " & _
                       " AND egreso.egr_fecha BETWEEN '" & txtFechaInicio.Text & "' AND '" & txtFechaFin.Text & "' " & _
                       " AND egreso.egr_anulado=0 " & _
                       " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                       " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                       " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                       " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
                       " AND det_egreso.prd_codigo=producto.prd_codigo "
    strSql = strSql & " LEFT JOIN producto_comision pc ON producto.emp_codigo=pc.emp_codigo " & _
                        " AND producto.prd_codigo=pc.prd_codigo " & _
                        " AND pc.cam_anio='" & Left(cmbCampania.BoundText, 4) & "' " & _
                        " AND pc.cam_mes='" & Right(cmbCampania.BoundText, 2) & "' " & _
                        " AND persona.tip_ped_codigo=pc.tip_ped_codigo "
    strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                        " AND persona.cat_p_tipo='C' AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                        " AND ((per_codigo_ref='" & strLider & "' AND per_codigo_ref2='' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref2='" & strLider & "' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref3='" & strLider & "' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref4='" & strLider & "' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref5='" & strLider & "' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref6='" & strLider & "' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref7='" & strLider & "' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref8='" & strLider & "' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref9='" & strLider & "' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref10='" & strLider & "') " & _
                        " )" & _
                        " GROUP BY persona.emp_codigo,persona.per_codigo "
    strSql = strSql & " UNION " & _
                        " SELECT persona.emp_codigo,persona.per_codigo," & _
                        " -1*SUM(ROUND((det_ing_cantidad*det_ing_precio-det_ing_dcto)*(COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimiento," & _
                        " -1*SUM(ROUND((det_ing_cantidad*det_ing_precio-det_ing_dcto)*(1-COALESCE(pc.pro_com_comision,100.00)/100.00),2)) as TMovimientoNC" & _
                        " FROM persona INNER JOIN ingreso ON persona.emp_codigo=ingreso.emp_codigo " & _
                        " AND persona.per_codigo=ingreso.per_codigo " & _
                        " AND ingreso.tip_ing_codigo='DCL' " & _
                        " AND ingreso.ing_fecha BETWEEN '" & txtFechaInicio.Text & "' AND '" & txtFechaFin.Text & "' " & _
                        " AND ingreso.ing_anulado=0 " & _
                        " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                        " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                        " AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                        " INNER JOIN producto ON det_ingreso.emp_codigo=producto.emp_codigo " & _
                        " AND det_ingreso.prd_codigo=producto.prd_codigo "
    strSql = strSql & " LEFT JOIN producto_comision pc ON producto.emp_codigo=pc.emp_codigo " & _
                        " AND producto.prd_codigo=pc.prd_codigo " & _
                        " AND pc.cam_anio='" & Left(cmbCampania.BoundText, 4) & "' " & _
                        " AND pc.cam_mes='" & Right(cmbCampania.BoundText, 2) & "' " & _
                        " AND persona.tip_ped_codigo=pc.tip_ped_codigo "
    strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                        " AND persona.cat_p_tipo='C' AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                        " AND ((per_codigo_ref='" & strLider & "' AND per_codigo_ref2='' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref2='" & strLider & "' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref3='" & strLider & "' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref4='" & strLider & "' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref5='" & strLider & "' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref6='" & strLider & "' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref7='" & strLider & "' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref8='" & strLider & "' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref9='" & strLider & "' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref10='" & strLider & "') " & _
                        " )" & _
                        " GROUP BY persona.emp_codigo,persona.per_codigo "
    strSql = strSql & " UNION " & _
                        " SELECT persona.emp_codigo,persona.per_codigo,SUM(ingreso.ing_dcto) as TMovimiento,0 as TMovimientoNC " & _
                        " FROM persona INNER JOIN ingreso ON persona.emp_codigo=ingreso.emp_codigo " & _
                        " AND persona.per_codigo=ingreso.per_codigo " & _
                        " AND ingreso.tip_ing_codigo='DCL' " & _
                        " AND ingreso.ing_fecha BETWEEN '" & txtFechaInicio.Text & "' AND '" & txtFechaFin.Text & "' " & _
                        " AND ingreso.ing_anulado=0 " & _
                        " LEFT JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                        " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                        " AND ingreso.ing_codigo=det_ingreso.ing_codigo "
    strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                        " AND det_ingreso.emp_codigo is null AND persona.cat_p_tipo='C' AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                        " AND ((per_codigo_ref='" & strLider & "' AND per_codigo_ref2='' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref2='" & strLider & "' AND per_codigo_ref3='' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref3='" & strLider & "' AND per_codigo_ref4='' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref4='" & strLider & "' AND per_codigo_ref5='' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref5='" & strLider & "' AND per_codigo_ref6='' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref6='" & strLider & "' AND per_codigo_ref7='' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref7='" & strLider & "' AND per_codigo_ref8='' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref8='" & strLider & "' AND per_codigo_ref9='' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref9='" & strLider & "' AND per_codigo_ref10='') " & _
                                " OR (per_codigo_ref10='" & strLider & "') " & _
                        " )" & _
                        " GROUP BY persona.emp_codigo,persona.per_codigo "
    strSql = strSql & ") " & _
                    " TVenta " & _
                    " GROUP BY emp_codigo,per_codigo " & _
                " ) Total " & _
                " GROUP BY emp_codigo "
    clsVentas.Ejecutar strSql
    'End If
    If clsVentas.adorec_Def.RecordCount > 0 Then
        TVentaDirecta = clsVentas.adorec_Def("tve")
        TVentaDirectaNC = clsVentas.adorec_Def("tveNC")
        TActivoDirecto = clsVentas.adorec_Def("tac")
    Else
        TVentaDirecta = 0
        TVentaDirectaNC = 0
        TActivoDirecto = 0
    End If
    
End Sub

Private Sub cmdGenImpPDF_Click()
    Dim frmCom As New frmReporte
    Dim strFechaMax As String
    Dim Destino As String
    Dim Ruta As String
    Dim i As Long
    Dim PdfImpresos As Long
    PdfImpresos = 0
    Destino = Buscar_Carpeta(Me.hwnd, "Carpetas a Subir")
    strFechaMax = InputBox("Fecha limite de entrega", Comsiones, txtFechaFin.Text)
    For i = 1 To VSFG.Rows - 1
        Me.Caption = i & " / " & VSFG.Rows - 1
        If FormatoD2(VSFG.TextMatrix(i, 17)) <> 0 Then
            frmCom.strNumero = cmbCampania.BoundText
            frmCom.strTipo = strFechaMax
            frmCom.strAsiento = Me.txtFechaInicio.Text & " al " & Me.txtFechaFin.Text
            frmCom.Atencion = VSFG.TextMatrix(i, 11)
            frmCom.strReporte = "rptLiqComision"
            frmCom.Show
            frmCom.Form_Activate
            frmCom.VSPrint.PrintDoc
            Ruta = Replace(Destino & "\" & Trim(IIf(VSFG.TextMatrix(i, 1) <> "", VSFG.TextMatrix(i, 1) & "\", "")) & _
                                           Trim(IIf(VSFG.TextMatrix(i, 2) <> "", VSFG.TextMatrix(i, 2) & "\", "")) & _
                                           Trim(IIf(VSFG.TextMatrix(i, 3) <> "", VSFG.TextMatrix(i, 3) & "\", "")), " ", "_")
            VerRuta Ruta
            frmCom.VSRpt.RenderToFile Replace(Ruta & Trim(VSFG.TextMatrix(i, 12)) & ".pdf", " ", "_"), vsrPDF
            If PdfImpresos = 0 Then
                MsgBox "Continuar?"
            End If
            PdfImpresos = PdfImpresos + 1
        End If
    Next i
    Me.Caption = "Construccion y Calculo de redes"
End Sub

Private Function VerRuta(Ruta As String) As Boolean

    Set VLman_arch = New FileSystemObject
    
    If Not VLman_arch.FolderExists(Ruta) Then
        If VerRuta(Left(Ruta, InStrRev(Left(Ruta, Len(Ruta) - 1), "\"))) Then
            MkDir (Ruta)
            VerRuta = True
        End If
    Else
        VerRuta = True
    End If

End Function

Private Sub cmdRecalificar_Click()
    CalificarNivel Left(cmbCampania.BoundText, 4), Right(cmbCampania.BoundText, 2)
    CrearComision Left(cmbCampania.BoundText, 4), Right(cmbCampania.BoundText, 2)
    MsgBox Now
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
    Me.Left = 200
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    
    'Tipo de negocios
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
    
    strSql = " SELECT concat(cam_anio,'-',cam_mes) as cam_codigo, cam_nombre " & _
             " FROM campaniafecha " & _
             " ORDER BY cam_nombre DESC "
    clsCon_Def.Ejecutar strSql
    Set cmbCampania.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbCampania.ListField = "cam_nombre"
    cmbCampania.BoundColumn = "cam_codigo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub VSFG_DblClick()
    Dim frmCom As New frmReporte
    Dim strFechaMax As String
    strFechaMax = InputBox("Fecha limite de entrega", Comsiones, txtFechaFin.Text)
        frmCom.strNumero = cmbCampania.BoundText
        frmCom.strTipo = strFechaMax
        frmCom.strAsiento = Me.txtFechaInicio.Text & " al " & Me.txtFechaFin.Text
        frmCom.Atencion = VSFG.TextMatrix(VSFG.Row, 11)
        frmCom.strReporte = "rptLiqComision"
        frmCom.Show
        frmCom.Form_Activate
End Sub
