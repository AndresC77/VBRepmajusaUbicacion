VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportarCobros 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Cobros"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportarCobros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3825
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Migracion"
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
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   840
         Width           =   3255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4815
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   10335
         _cx             =   1987135286
         _cy             =   1987125549
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmExportarCobros.frx":030A
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
         ExplorerBar     =   5
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   308
         Width           =   2640
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   10080
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   9600
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   330
         Left            =   3480
         TabIndex        =   7
         Top             =   300
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   930
         TabIndex        =   10
         Top             =   308
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6720
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5505
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
End
Attribute VB_Name = "frmExportarCobros"
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

Private Sub cmdAplicar_Click()
    Dim i As Long
    Me.MousePointer = 11
    
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, VSFG.Cols - 1)) = 1 Then
            
            strSql = " SELECT CONCAT('C',LPAD(ROUND(COALESCE(MAX(RIGHT(per_codigo,LEN(per_codigo)-1)),0)+1,0),5,'0')) as cod " & _
                     " FROM persona " & _
                     " WHERE cat_p_tipo='C'" & _
                     " AND emp_codigo='" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
            clsCon_Def.Ejecutar strSql
            VSFG.TextMatrix(i, 3) = clsCon_Def.adorec_Def("cod")
'            strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
'                     " cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
'                     " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
'                     " per_credito,per_dcto,for_pag_codigo,ven_codigo,tip_ped_codigo,per_codigo_ref,per_codigo_ref2," & _
'                     " per_observacion,per_fac_flete,per_sec_publico,per_siniva, " & _
'                     " per_fechamod,per_usumod,per_fechaing,per_usuing,per_perdesde,per_inactivo,per_aplica_nc) " & _
'                     " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 4) & "','" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "', " & _
'                     " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & VSFG.TextMatrix(i, 17) & "','" & VSFG.TextMatrix(i, 18) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & VSFG.TextMatrix(i, 20) & "','" & FormatoD2(VSFG.TextMatrix(i, 21)) & "','" & FormatoD4(VSFG.TextMatrix(i, 22)) & "','" & VSFG.TextMatrix(i, 23) & "','" & _
'                     VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "','" & VSFG.TextMatrix(i, 26) & "','" & VSFG.TextMatrix(i, 27) & "','" & UCase(VSFG.TextMatrix(i, 28)) & "'," & _
'                     " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 30))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "'," & _
'                     " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_DATE,'" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "')"
'            clsCon_Def.Ejecutar strSql, "M"
        Else
            Exit For
        End If
    Next i
    
'    For i = 1 To VSFG.Rows - 1
'        If Val(VSFG.TextMatrix(i, VSFG.Cols - 1)) = 1 Then
'            strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
'                     " cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
'                     " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
'                     " per_credito,per_dcto,for_pag_codigo,ven_codigo,tip_ped_codigo,per_codigo_ref,per_codigo_ref2," & _
'                     " per_observacion,per_fac_flete,per_sec_publico,per_siniva, " & _
'                     " per_fechamod,per_usumod,per_fechaing,per_usuing,per_perdesde,per_inactivo,per_aplica_nc) " & _
'                     " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 4) & "','" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "', " & _
'                     " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & VSFG.TextMatrix(i, 17) & "','" & VSFG.TextMatrix(i, 18) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & VSFG.TextMatrix(i, 20) & "','" & FormatoD2(VSFG.TextMatrix(i, 21)) & "','" & FormatoD4(VSFG.TextMatrix(i, 22)) & "','" & VSFG.TextMatrix(i, 23) & "','" & _
'                     VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "','" & VSFG.TextMatrix(i, 26) & "','" & VSFG.TextMatrix(i, 27) & "','" & UCase(VSFG.TextMatrix(i, 28)) & "'," & _
'                     " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 30))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "'," & _
'                     " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_DATE,'" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "')"
'            clsCon_Def.Ejecutar strSql, "M"
'        Else
'            Exit For
'        End If
'    Next i
    Me.MousePointer = 0
    MsgBox "Carga de clintes", vbInformation, "Clientes"
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    If cdArchivo.FileName <> "" Then
        VSFG.SaveGrid cdArchivo.FileName, flexFileExcel, flexXLSaveFixedRows

        MsgBox "Termino la migracion de " & VSFG.Rows & " registros.", vbInformation, "Migracion de Cobros"
    Else
        MsgBox "Seleccione primero un archivo", vbInformation, "Migracion de Cobros"
    End If
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    sDir = CurDir
    cdArchivo.Filter = "File Types (*.xls)|*.xls|All Types (*.*)|*.*|"
    cdArchivo.ShowSave
    Me.txtArchivo.Text = cdArchivo.FileName
    ChDir sDir
End Sub

Private Sub cmdMostrar_Click()
    strSql = " SELECT doc_pago.doc_pag_codigo, COALESCE(tip_doc_pag_codigo,'CONT'),per_ruc,IF(LEN(per_ruc)>=13,'R','C'),CONCAT(LPAD(LEFT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),LEN(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7))-3),3,'0'),'-',RIGHT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),3),'-',RIGHT(cue_p_c_egr_codigo,7)+0),doc_pag_fecha_recepcion,RIGHT(doc_pag_fechamod,8),pag_monto " & _
             " FROM doc_pago INNER JOIN pago ON doc_pago.emp_codigo=pago.emp_codigo " & _
             " AND doc_pago.doc_pag_codigo=pago.doc_pag_codigo " & _
             " INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo " & _
             " AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo " & _
             " AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
             " AND doc_pago.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo " & _
             " AND doc_pago.per_codigo=persona.per_codigo " & _
             " WHERE doc_pago.emp_codigo='" & strEmpresa & "' " & _
             " AND doc_pag_estado!='ANULADO' " & _
             " AND doc_pag_valor!=0 " & _
             " AND doc_pag_fecha_recepcion='" & dtpFecha.Value & "' "
    strSql = strSql & " UNION " & _
             " SELECT doc_pago.doc_pag_codigo, COALESCE(tip_doc_pag_codigo,'CONT'),per_ruc,IF(LEN(per_ruc)>=13,'R','C'),CONCAT(LPAD(LEFT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),LEN(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7))-3),3,'0'),'-',RIGHT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),3),'-',RIGHT(cue_p_c_egr_codigo,7)+0),doc_pag_fecha_recepcion,RIGHT(doc_pag_fechamod,8),com_ret_total " & _
             " FROM doc_pago INNER JOIN pago ON doc_pago.emp_codigo=pago.emp_codigo " & _
             " AND doc_pago.doc_pag_codigo=pago.doc_pag_codigo " & _
             " INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo " & _
             " AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo " & _
             " AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
             " AND doc_pago.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
             " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
             " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
             " INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo " & _
             " AND doc_pago.per_codigo=persona.per_codigo " & _
             " WHERE doc_pago.emp_codigo='" & strEmpresa & "' " & _
             " AND tip_doc_pag_codigo='RET' " & _
             " AND doc_pag_valor=0 " & _
             " AND doc_pag_fecha_recepcion='" & dtpFecha.Value & "' "
    strSql = strSql & " UNION " & _
             " SELECT CONCAT('NC-',pago.pag_no_doc), 'NC',per_ruc,IF(LEN(per_ruc)>=13,'R','C'),CONCAT(LPAD(LEFT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),LEN(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7))-3),3,'0'),'-',RIGHT(LEFT(cue_p_c_egr_codigo,LEN(cue_p_c_egr_codigo)-7),3),'-',RIGHT(cue_p_c_egr_codigo,7)+0),pag_fecha,RIGHT(pag_fechamod,8),pag_monto " & _
             " FROM pago INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo " & _
             " AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo " & _
             " AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
             " AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
             " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
             " INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
             " WHERE pago.emp_codigo='" & strEmpresa & "' " & _
             " AND pago.asi_numasiento LIKE '%A%' " & _
             " AND pag_monto!=0 " & _
             " AND pag_fecha='" & dtpFecha.Value & "' "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    'VSFG.Cell(flexcpTextStyle
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
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    dtpFecha.Value = HoyDia
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
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
