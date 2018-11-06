VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCargaClientesFidelizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Clientes Fidelizacion"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaClientesFidelizacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   8895
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cargar Clientes"
      TabPicture(0)   =   "frmCargaClientesFidelizacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmbFidelizacion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbNegocio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cdArchivo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "VSFG"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAplicar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtArchivo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdExplorar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   4080
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   6840
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   8415
         _cx             =   14843
         _cy             =   8705
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaClientesFidelizacion.frx":0326
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
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   5760
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   480
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
      Begin MSDataListLib.DataCombo cmbFidelizacion 
         Height          =   315
         Left            =   1665
         TabIndex        =   9
         Top             =   840
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fidelización:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   720
         TabIndex        =   10
         Top             =   885
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   975
         TabIndex        =   7
         Top             =   525
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1035
         TabIndex        =   5
         Top             =   1320
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCargaClientesFidelizacion"
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

Private Sub cmdActualizar_Click()
    
    strSql = " SELECT pedido.ped_codigo,ped_fecha,pedido.per_codigo,per_ruc, CONCAT(per_apellido, ' ',per_nombre) as cli,inc_loc_nombre,producto.prd_codigo,prd_nombre,inc_loc_cantidad,inc_loc_precio,inc_loc_dcto " & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
             " AND pedido.per_codigo=persona.per_codigo " & _
             " AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio2.BoundText & "' " & _
             " INNER JOIN incentivo_local ON persona.emp_codigo=incentivo_local.emp_codigo " & _
             " AND persona.per_codigo=incentivo_local.per_codigo " & _
             " AND persona.tip_ped_codigo=incentivo_local.tip_ped_codigo " & _
             " AND LEFT(pedido.ped_fecha,10) BETWEEN LEFT(incentivo_local.inc_loc_fecha_desde,10) AND LEFT(incentivo_local.inc_loc_fecha_hasta,10) " & _
             " AND incentivo_local.inc_loc_estado=0 " & _
             " INNER JOIN producto ON incentivo_local.emp_codigo=producto.emp_codigo " & _
             " AND incentivo_local.prd_codigo=producto.prd_codigo " & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND pedido.ped_estado=0 " & _
             " AND pedido.ped_fecha BETWEEN '" & Format(dtpFechaInicio2.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(dtpFechaFin2.Value, "yyyy-mm-dd hh:mm") & ":59'" & _
             " ORDER BY per_ruc,producto.prd_codigo "
             

    clsCon_Def.Ejecutar strSql
    Set VSFGPeds.DataSource = clsCon_Def.adorec_Def.DataSource
End Sub

Private Sub cmdAplicar_Click()
    Dim i As Long
    Dim j As Long
    
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    Me.MousePointer = 11
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, VSFG.Cols - 1)) = 1 Then
                strSql = " UPDATE persona " & _
                         " SET fid_codigo='" & cmbFidelizacion.BoundText & "', " & _
                         " per_fechamod=CURRENT_TIMESTAMP, " & _
                         " per_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                         " AND per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                         " AND cat_p_tipo='C'"
                clsCon_Def.Ejecutar strSql, "M"
            Else
                Exit For
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Incentivos"
    
    Unload Me

End Sub

Private Sub cmdAplicar2_Click()
    Dim i As Long
    For i = 1 To VSFGPeds.Rows - 1
        strSql = " SELECT COALESCE(count(*),0) as n " & _
                 " FROM incentivo_local " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND per_codigo='" & VSFGPeds.TextMatrix(i, 2) & "' " & _
                 " AND LEFT('" & VSFGPeds.TextMatrix(i, 1) & "',10) BETWEEN LEFT(inc_loc_fecha_desde,10) AND LEFT(inc_loc_fecha_hasta,10)" & _
                 " AND inc_loc_estado=0 " & _
                 " AND ped_codigo is null " & _
                 " AND prd_codigo='" & VSFGPeds.TextMatrix(i, 6) & "' "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If clsCon_Def.adorec_Def("n") > 0 Then
                strSql = " UPDATE incentivo_local " & _
                         " SET inc_loc_estado=1, " & _
                         " ped_codigo='" & VSFGPeds.TextMatrix(i, 0) & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND per_codigo='" & VSFGPeds.TextMatrix(i, 2) & "' " & _
                         " AND LEFT('" & VSFGPeds.TextMatrix(i, 1) & "',10) BETWEEN LEFT(inc_loc_fecha_desde,10) AND LEFT(inc_loc_fecha_hasta,10)" & _
                         " AND inc_loc_estado=0 " & _
                         " AND ped_codigo is null " & _
                         " AND prd_codigo='" & VSFGPeds.TextMatrix(i, 6) & "' "
                clsCon_Def.Ejecutar strSql
                strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                         " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                         " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto) " & _
                         " VALUES ('" & strEmpresa & "','PRI',0," & _
                         " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                         " '" & VSFGPeds.TextMatrix(i, 0) & "','" & VSFGPeds.TextMatrix(i, 6) & "','" & VSFGPeds.TextMatrix(i, 8) & "'," & _
                         " '" & VSFGPeds.TextMatrix(i, 8) & "','" & VSFGPeds.TextMatrix(i, 9) & "','" & VSFGPeds.TextMatrix(i, 10) & "') "
                clsCon_Def.Ejecutar strSql
            End If
        End If
    Next i
    MsgBox "Carga Finalizada"
    Unload Me
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    Dim clsCon_DefP As New clsConsulta
    clsCon_DefP.Inicializar AdoConn, AdoConnMaster
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    VSFG.Clear flexClearScrollable
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivo.Text <> "") Then
        VSFG.Rows = 1
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 4
        VSFG.TextMatrix(0, 0) = "CI RUC"
        VSFG.TextMatrix(0, 1) = "Cod.Cliente"
        VSFG.TextMatrix(0, 2) = "Cliente"
        VSFG.TextMatrix(0, 3) = "Aplica"
        For i = 1 To VSFG.Rows - 1
            VSFG.ShowCell i, 0
            If Trim(VSFG.TextMatrix(i, 0)) <> "" And VSFG.TextMatrix(i, 0) <> VSFG.TextMatrix(i - 1, 0) Then
                strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND per_ruc like '%" & VSFG.TextMatrix(i, 0) & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
                clsCon_DefP.Ejecutar strSql
                If clsCon_DefP.adorec_Def.RecordCount > 0 Then
                    VSFG.TextMatrix(i, 1) = clsCon_DefP.adorec_Def("per_codigo")
                    VSFG.TextMatrix(i, 2) = clsCon_DefP.adorec_Def("cli")
                    VSFG.TextMatrix(i, VSFG.Cols - 1) = 1
                Else
                    VSFG.TextMatrix(i, 1) = ""
                    VSFG.TextMatrix(i, 2) = VSFG.TextMatrix(i, VSFG.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                    VSFG.TextMatrix(i, VSFG.Cols - 1) = 0
                    VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbRed
                End If
            Else
                VSFG.TextMatrix(i, 1) = ""
                VSFG.TextMatrix(i, 2) = VSFG.TextMatrix(i, VSFG.Cols - 2) & "  --  CLIENTE NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 0
                VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbRed
            
            End If
            
        Next i
    
        VSFG.Select 1, VSFG.Cols - 1
        VSFG.Sort = flexSortGenericAscending
    
    End If
    Me.MousePointer = 0
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
    'Consulta las listas de precios que estan disponibles
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
        
        
        'crea combo de fidelizacion
        strSql = " SELECT fid_codigo, fid_nombre" & _
                    " FROM fidelizacion " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " ORDER BY fid_nombre"
        clsCon_Def.Ejecutar strSql
        Set cmbFidelizacion.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbFidelizacion.ListField = "fid_nombre"
        cmbFidelizacion.BoundColumn = "fid_codigo"
     
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

Private Sub op_coleccion_Click()
       i_flag = 0
       dcmbColeccion.Enabled = True
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = False
       dtpFechaFin.Enabled = False
End Sub

Private Sub op_disponibleventa_Click()
        i_flag = 2
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = True
       dtpFechaFin.Enabled = True
End Sub

Private Sub op_lista_Click()
       i_flag = 1
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = True
       dtpFechaInicio.Enabled = False
       dtpFechaFin.Enabled = False
End Sub
