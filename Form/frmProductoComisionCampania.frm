VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProductoComisionCampania 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Productos comision para Campaña"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductoComisionCampania.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   9195
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6840
         TabIndex        =   15
         Top             =   5400
         Width           =   1980
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   14
         Top             =   4800
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   1980
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1980
      End
      Begin VB.CommandButton cmdCargarActual 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4695
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   6495
         _cx             =   11456
         _cy             =   8281
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmProductoComisionCampania.frx":030A
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
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   4080
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         TabIndex        =   1
         Top             =   960
         Width           =   375
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   4560
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   240
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
         Left            =   960
         TabIndex        =   6
         Top             =   585
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio de Facturación"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin de Facturación"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6840
         TabIndex        =   12
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   210
         TabIndex        =   8
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaña:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   225
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmProductoComisionCampania"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private strSql As String
Dim i_flag As Integer

Private Sub cmbCampania_Validate(Cancel As Boolean)
    strSql = " SELECT cam_fecha_fac_inicial, cam_fecha_fac_final " & _
             " FROM campaniafecha " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND concat(cam_anio,'-',cam_mes)='" & cmbCampania.BoundText & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
    txtFechaInicio.Text = Left(clsCon_Def.adorec_Def("cam_fecha_fac_inicial"), 10)
    txtFechaFin.Text = Left(clsCon_Def.adorec_Def("cam_fecha_fac_final"), 10)
    End If
End Sub

Private Sub cmdCargarActual_Click()
    strSql = " SELECT producto_comision.prd_codigo,prd_nombre,pro_com_comision " & _
             " FROM producto_comision INNER JOIN producto " & _
             " ON producto_comision.emp_codigo=producto.emp_codigo " & _
             " AND producto_comision.prd_codigo=producto.prd_codigo " & _
             " WHERE producto_comision.emp_codigo='" & strEmpresa & "'" & _
             " AND producto_comision.cam_anio='" & Left(cmbCampania.BoundText, 4) & "'" & _
             " AND producto_comision.cam_mes='" & Right(cmbCampania.BoundText, 2) & "'" & _
             " ORDER BY prd_nombre"
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    cmdGuardar.Enabled = False
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    VSFG.Clear
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivo.Text <> "") Then

        VSFG.LoadGrid txtArchivo.Text, flexFileTabText
        VSFG.Cols = 2
        VSFG.Cols = 4
        VSFG.ColPosition(1) = 2
        VSFG.TextMatrix(0, 0) = "Codigo"
        VSFG.TextMatrix(0, 1) = "Producto"
        VSFG.TextMatrix(0, 2) = "%Comi."
        VSFG.TextMatrix(0, 3) = "Aplica"
        VSFG.ColHidden(3) = True
        VSFG.ColWidth(0) = 1500
        VSFG.ColWidth(1) = 3800
        VSFG.ColWidth(2) = 800
        For i = 1 To VSFG.Rows - 1
            strSql = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                     
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 3) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 2) = 0
                VSFG.TextMatrix(i, 3) = 0
            End If
        Next i
        cmdGuardar.Enabled = True
    End If
    Me.MousePointer = 0
End Sub

Private Sub cmdGuardar_Click()
    Dim i As Long
    If cmbNegocio.MatchedWithList = True And cmbCampania.MatchedWithList = True Then
        For i = 1 To VSFG.Rows - 1
            If FormatoD0(VSFG.TextMatrix(i, 3)) <> 0 Then
                strSql = " INSERT INTO producto_comision (emp_codigo,prd_codigo,cam_anio,cam_mes," & _
                         " tip_ped_codigo,pro_com_comision,pro_com_fechamod,pro_com_usumod)" & _
                         " VALUES('" & strEmpresa & "','" & VSFG.TextMatrix(i, 0) & "','" & Left(cmbCampania.BoundText, 4) & "','" & Right(cmbCampania.BoundText, 2) & "'," & _
                         "'" & cmbNegocio.BoundText & "','" & VSFG.TextMatrix(i, 2) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsCon_Def.Ejecutar strSql, "M"
            End If
        Next i
    End If
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster

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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

