VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBloqueoDesbloqueo 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo y Desbloqueo Automático"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBloqueoDesbloqueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   6135
   Begin VB.TextBox txtTiempo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Text            =   "3"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   6600
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Bloqueo"
      TabPicture(0)   =   "frmBloqueoDesbloqueo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ucrtVSFG1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "VSFG"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdConsulta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBloquear"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "DESBloqueo"
      TabPicture(1)   =   "frmBloqueoDesbloqueo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ucrtVSFG2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "VSFG2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdConsulta2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdDESBloquear"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdDESBloquear 
         Caption         =   "&DESBloquear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   6000
         Width           =   2055
      End
      Begin VB.CommandButton cmdConsulta2 
         Caption         =   "Consulta DESBloqueo"
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdBloquear 
         Caption         =   "&Bloquear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   6000
         Width           =   2055
      End
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "Consulta Bloqueo"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   5655
         _cx             =   9975
         _cy             =   7858
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBloqueoDesbloqueo.frx":0342
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
      Begin NEED2.uctrVSFG ucrtVSFG1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   7
         Top             =   1440
         Width           =   5655
         _cx             =   9975
         _cy             =   7858
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBloqueoDesbloqueo.frx":04B2
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
      Begin NEED2.uctrVSFG ucrtVSFG2 
         Height          =   375
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias de Gracia"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -71640
         TabIndex        =   14
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias de Gracia"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3360
         TabIndex        =   13
         Top             =   682
         Width           =   1065
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Negocio:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   172
      Width           =   630
   End
End
Attribute VB_Name = "frmBloqueoDesbloqueo"
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

Private Sub ParaBloquear()
    Dim i As Long
    strSql = " SELECT cuenta_p_c.cue_p_c_codigo as c1,cuenta_p_c.per_codigo as per_codigo,CONCAT(per_apellido, ' ',per_nombre) as cli,IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, " & _
             " RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, " & _
             " cue_p_c_fechaemision, cue_p_c_fechapropuesta," & _
             " cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
             " FROM  cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo" & _
             " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
             " AND persona.per_bloqueado=0 " & _
             " AND persona.per_es_gz=0 AND persona.per_es_di=0 AND persona.per_es_em=0 AND persona.per_es_ee=0" & _
             " AND persona.per_es_n5=0 AND persona.per_es_n6=0 AND persona.per_es_n7=0 AND persona.per_es_n8=0" & _
             " AND persona.per_es_n9=0 " & _
             " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo AND forma_pago.for_pag_tiempo!=0 " & _
             " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
             " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
             " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 AND cue_p_c_fechapropuesta<=DATEADD(d,-" & FormatoD0(txtTiempo.Text) & ",CURRENT_TIMESTAMP)" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.per_codigo,per_apellido, per_nombre,per_ruc,cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,cue_p_c_valor,com_ret_total " & _
             " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)>1.50 " & _
             " ORDER BY cue_p_c_egr_codigo,c1"
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    cmdBloquear.Enabled = True
End Sub
Private Sub ParaDesBloquear()
    Dim i As Long
    Dim clsCliente As New clsConsulta
    clsCliente.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT DISTINCT persona.per_codigo,CONCAT(per_apellido, ' ',per_nombre) as cli " & _
             " FROM persona INNER JOIN doc_pago ON persona.emp_codigo=doc_pago.emp_codigo" & _
             " AND persona.per_codigo=doc_pago.per_codigo" & _
             " AND doc_pago.doc_pag_fecha_recepcion>= DATEADD(d,-" & FormatoD0(txtTiempo.Text) & ",CURRENT_TIMESTAMP)" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " AND per_bloqueado<>0 " & _
             " AND persona.per_es_gz=0 AND persona.per_es_di=0 AND persona.per_es_em=0 AND persona.per_es_ee=0" & _
             " AND persona.per_es_n5=0 AND persona.per_es_n6=0 AND persona.per_es_n7=0 AND persona.per_es_n8=0" & _
             " AND persona.per_es_n9=0 " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'"
    clsCliente.Ejecutar strSql
    VSFG2.Rows = 1
    While Not clsCliente.adorec_Def.EOF
    
        strSql = " SELECT RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, " & _
                 " cue_p_c_fechaemision, cue_p_c_fechapropuesta," & _
                 " cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                 " FROM  cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo" & _
                 " AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " AND cuenta_p_c.per_codigo='" & clsCliente.adorec_Def("per_codigo") & "'" & _
                 " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and cue_p_c_fechapropuesta<=DATEADD(d,-5,CURRENT_TIMESTAMP)" & _
                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cue_p_c_egr_codigo,cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor,com_ret_total " & _
                 " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)>1.50 " & _
                 " ORDER BY cue_p_c_egr_codigo"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount = 0 Then
            VSFG2.AddItem clsCliente.adorec_Def(0) & vbTab & clsCliente.adorec_Def(1)
        End If
        clsCliente.adorec_Def.MoveNext
    Wend
    cmdDESBloquear.Enabled = True
End Sub

Private Sub cmdBloquear_Click()
    Bloquear
End Sub

Private Sub Bloquear()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        strSql = " UPDATE persona " & _
                 " SET per_bloqueado=1," & _
                 " per_observacion_mod = CONCAT(COALESCE(per_observacion_mod,''),CURRENT_TIMESTAMP,' - " & strUsuario & "',' - CLIENTE BLOQUEADO AUTO CARTERA" & vbNewLine & "')," & _
                 " per_usumod='" & strUsuario & "'," & _
                 " per_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND per_codigo_ref NOT IN ('C02443') " & _
                 " AND per_codigo='" & VSFG.TextMatrix(i, 1) & "'"
        clsCon_Def.Ejecutar strSql, "M"
        
    Next i
    MsgBox "Bloqueo Terminado", vbInformation
    cmdBloquear.Enabled = False
End Sub

Private Sub DesBloquear()
    Dim i As Long
    For i = 1 To VSFG2.Rows - 1
        strSql = " UPDATE persona " & _
                 " SET per_bloqueado=0," & _
                 " per_observacion_mod = CONCAT(COALESCE(per_observacion_mod,''),CURRENT_TIMESTAMP,' - " & strUsuario & "',' - CLIENTE DESBLOQUEADO AUTO CARTERA" & vbNewLine & "')," & _
                 " per_usumod='" & strUsuario & "'," & _
                 " per_fechamod=CURRENT_TIMESTAMP " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND per_codigo='" & VSFG2.TextMatrix(i, 0) & "'"
        clsCon_Def.Ejecutar strSql, "M"
        
    Next i
    MsgBox "DesBloqueo Terminado", vbInformation
    cmdDESBloquear.Enabled = False
End Sub

Private Sub cmdConsulta_Click()
    ParaBloquear
End Sub

Private Sub cmdConsulta2_Click()
    ParaDesBloquear
End Sub

Private Sub cmdDESBloquear_Click()
    DesBloquear
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
    Set ucrtVSFG1.VSFGControl = VSFG
    Set ucrtVSFG2.VSFGControl = VSFG2
    ucrtVSFG1.Inicializar False, False, False
    ucrtVSFG2.Inicializar False, False, False
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        
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

Private Sub txtTiempo_Validate(Cancel As Boolean)
    If FormatoD0(txtTiempo.Text) < 0 Then
        MsgBox "El tiempo no puede ser menor a 0 dias"
        txtTiempo.Text = 5
    End If
    If IsNumeric(txtTiempo.Text) = False Then
        MsgBox "El campo debe tener numeros "
        txtTiempo.Text = 5
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub
