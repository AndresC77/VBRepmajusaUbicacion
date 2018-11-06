VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmArchivoCarteraBanco 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmArchivoCarteraBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12285
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7680
      TabIndex        =   12
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optPichincha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pichincha"
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
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optProdubanco 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Produbanco"
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
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmbGuardar 
      Caption         =   "&Guardar"
      Height          =   360
      Left            =   4392
      TabIndex        =   3
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
      TabIndex        =   2
      Top             =   7080
      Width           =   1700
   End
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7185
      Begin VB.CheckBox chkDirector 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Director"
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
         Left            =   3720
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkGerente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Gerente de Zona"
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1200
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Director"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gerente de Zona"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   495
         Width           =   3255
      End
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   4680
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   12060
      _cx             =   1992905496
      _cy             =   1992892479
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
      Cols            =   100
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmArchivoCarteraBanco.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   3840
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   5160
      Picture         =   "frmArchivoCarteraBanco.frx":0AF1
      ToolTipText     =   "Elimina una Fila"
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmArchivoCarteraBanco"
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
    Tipo = ""
    Tipo2 = ""
    Me.Caption = Tipo
End Sub

Private Sub chkGerente_Click()
    If chkGerente.Value = 1 Then
        cmbGerente.Enabled = True
    Else
        cmbGerente.Enabled = False
    End If
End Sub

Private Sub chkDirector_Click()
    If chkDirector.Value = 1 Then
        cmbDirector.Enabled = True
    Else
        cmbDirector.Enabled = False
    End If
End Sub

Private Sub cmbGuardar_Click()
    Dim strPath As String
    Dim Archivo As String
    strPath = Trim(App.Path)
    cmdArchivo.DialogTitle = "Guardar"
    'cmdArchivo.DefaultExt = strPath
    cmdArchivo.InitDir = strPath
    cmdArchivo.FileName = Arch
    cmdArchivo.Filter = "Archivos tipo texto|*.txt"
    cmdArchivo.ShowSave
    Archivo = cmdArchivo.FileName
    If Archivo <> "" Then
        VSFG.SaveGrid Archivo, flexFileTabText
    End If

End Sub

Private Sub cmdMostrar_Click()
    Carga
End Sub
Private Sub Carga()
    Dim strSqlDI As String
    Dim strSqlGZ As String
    Dim strBanco As String
    
    If chkGerente.Value = 1 Then
        strSqlGZ = " AND (persona.per_codigo_ref LIKE '" & cmbGerente.BoundText & "' OR persona.per_codigo LIKE '" & cmbGerente.BoundText & "') "
    End If
    If chkDirector.Value = 1 Then
        strSqlDI = " AND (persona.per_codigo_ref2 LIKE '" & cmbDirector.BoundText & "' OR persona.per_codigo LIKE '" & cmbDirector.BoundText & "') "
    End If
    
    strSQL = " CREATE TEMPORARY TABLE Abo ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " abono decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO Abo " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo," & _
             " COALESCE(sum(pag_monto),0.000) as abono " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             strSqlGZ & strSqlDI & _
             " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = pago.emp_codigo ) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " GROUP BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TEMPORARY TABLE RetFech ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " reten decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO RetFech " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo," & _
             " COALESCE(comprobante_retencion.com_ret_total,0.000) as reten " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             strSqlGZ & strSqlDI & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo " & _
             " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TEMPORARY TABLE Ret ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " reten decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO Ret " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo," & _
             " COALESCE(RetFech.reten,0.000) as reten " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             strSqlGZ & strSqlDI & _
             " LEFT JOIN RetFech ON cuenta_p_c.cue_p_c_codigo = RetFech.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = RetFech.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = RetFech.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TEMPORARY TABLE Cob ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " abono decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO Cob " & _
             " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo," & _
             " COALESCE(Abo.abono,0.000) as abono " & _
             " FROM (cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
             strSqlGZ & strSqlDI & _
             " LEFT JOIN Abo ON cuenta_p_c.cue_p_c_codigo = Abo.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = Abo.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = Abo.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " ORDER BY cue_p_c_codigo"
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE Abo "
    clsCon_Def.Ejecutar strSQL
    
    If optProdubanco.Value = True Then
        strBanco = " 'CO' as CodigoOrientacion, " & _
                   " '" & Format("02005168127", "00000000000") & "' as CuentaEmpresa, " & _
                   " RIGHT(cue_p_c_egr_codigo,7) as SecuencialCobro, " & _
                   " RIGHT(cue_p_c_egr_codigo,20) as ComprobanteDeCobro, " & _
                   " LEFT('RB IMPORTADORES',20) as Contrapartida, " & _
                   " LEFT('USD',3) as Moneda, " & _
                   " RIGHT(LPAD(TRUNCATE((ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Ret.reten,0.000),2)*100.00),0),13,'0'),13) as Valor, " & _
                   " LEFT('REC',3) as FormaDePago, " & _
                   " RIGHT('0036',4) as CodigoDeBanco, " & _
                   " '' as TipoDeCuenta, " & _
                   " '' as NumeroDeCuenta, " & _
                   " if(LEN(replace(per_ruc,'\r\n',''))=10,'C',if(LEN(replace(per_ruc,'\r\n',''))=13,'R','P')) as TipoIdClienteDeudor, " & _
                   " LEFT(replace(per_ruc,'\r\n',''),13) as NumeroIdClienteDeudor, " & _
                   " LEFT(CONCAT(replace(per_apellido,'\r\n',''), ' ', replace(per_nombre,'\r\n','')),40) as NombreDelClienteDeudor, " & _
                   " LEFT(replace(per_direccion,'\r\n',''),40) as DireccionDeudor, " & _
                   " LEFT(ciu_nombre,20) as CiudadDeudor, " & _
                   " LEFT(replace(per_telf,'\r\n',''),20) as TelefonoDeudor, " & _
                   " '' as LocalidadDeCobro, " & _
                   " LEFT(cue_p_c_egr_codigo,200) as Referencia, " & _
                   " LEFT(CONCAT('Fecha: ',LEFT(cue_p_c_fechaemision,10)),100) as ReferenciaAdicional, " & _
                   " '' as BaseImponible"
    ElseIf optPichincha.Value = True Then
        strBanco = " 'CO' as CodigoOrientacion, " & _
                   " LEFT(replace(per_ruc,'\r\n',''),20) as Contrapartida, " & _
                   " LEFT('USD',3) as Moneda, " & _
                   " RIGHT(LPAD(TRUNCATE((ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Ret.reten,0.000),2)*100.00),0),13,'0'),13) as Valor, " & _
                   " LEFT('REC',3) as FormaDePago, " & _
                   " '' as TipoDeCuenta, " & _
                   " '' as NumeroDeCuenta, " & _
                   " LEFT(cue_p_c_egr_codigo,40) as Referencia, " & _
                   " if(LEN(replace(per_ruc,'\r\n',''))=10,'C',if(LEN(replace(per_ruc,'\r\n',''))=13,'R','P')) as TipoIdCliente, " & _
                   " LEFT(replace(per_ruc,'\r\n',''),14) as NumeroIdCliente, " & _
                   " LEFT(CONCAT(replace(per_apellido,'\r\n',''), ' ', replace(per_nombre,'\r\n','')),41) as NombreDelCliente "
    
    End If

    
    strSQL = " SELECT " & strBanco & _
             " FROM (((cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo = 'C' " & _
             strSqlGZ & strSqlDI & ")" & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
             " INNER JOIN Ret ON cuenta_p_c.cue_p_c_codigo = Ret.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = Ret.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = Ret.emp_codigo) " & _
             " INNER JOIN Cob ON cuenta_p_c.cue_p_c_codigo = Cob.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = Cob.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = Cob.emp_codigo) " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
             " AND ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Ret.reten,0.000),2)>0 " & _
             " ORDER BY CONCAT(per_apellido, ' ', per_nombre),cue_p_c_egr_codigo"
    
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
    strSQL = " DROP TABLE Ret "
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE RetFech "
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE Cob "
    clsCon_Def.Ejecutar strSQL
    
    '' ucrtVSFG.PonerNum
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    On Error Resume Next
    Unload frmContacto
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
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    IniDato
    
    strSQL = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN persona as p2 ON p1.per_codigo=p2.per_codigo_ref AND p1.emp_codigo=p2.emp_codigo AND p1.cat_p_tipo=p2.cat_p_tipo " & _
             " INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbGerente.RowSource = clsCon_Def.adorec_Def
    cmbGerente.BoundColumn = "codigo"
    cmbGerente.ListField = "nombre"
    
    strSQL = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN persona as p2 ON p1.per_codigo=p2.per_codigo_ref2 AND p1.emp_codigo=p2.emp_codigo AND p1.cat_p_tipo=p2.cat_p_tipo " & _
             " INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbDirector.RowSource = clsCon_Def.adorec_Def
    cmbDirector.BoundColumn = "codigo"
    cmbDirector.ListField = "nombre"
  
    'Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
        If VSFG.TextMatrix(Row, 16) = "" Then
            VSFG.TextMatrix(Row, 16) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 34) = "" Then
            VSFG.TextMatrix(Row, 34) = HoyDia
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
        If VSFG.TextMatrix(Row, 34) = "" Then
            VSFG.TextMatrix(Row, 34) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 16) = "" Then
            VSFG.TextMatrix(Row, 16) = HoyDia
        End If
    End If
    If Col = 16 Then
        If Not IsDate(VSFG.TextMatrix(Row, 16)) Then
            VSFG.TextMatrix(Row, 16) = HoyDia
        End If
    End If
    If Col = 34 Then
        If Not IsDate(VSFG.TextMatrix(Row, 34)) Then
            VSFG.TextMatrix(Row, 34) = HoyDia
        End If
    End If
    If Col = 7 Then
        If VSFG.TextMatrix(Row, 7) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, 7)) Then
            If MsgBox("Está ingresando en el campo CI/RUC valores no numéricos, desea continuar?", vbQuestion + vbYesNo, "CI/RUC") = vbNo Then
                VSFG.TextMatrix(Row, 7) = ""
            End If
        End If
    End If
End Sub


Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
