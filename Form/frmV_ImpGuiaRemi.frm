VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmV_ImpGuiaRemi 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Guías de Remisón"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmV_ImpGuiaRemi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdVisPreIng 
      Caption         =   "&Vista Previa Ingreso"
      Height          =   375
      Left            =   2093
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Guías Remisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   113
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Tipos de Guías"
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
         Height          =   615
         Left            =   1020
         TabIndex        =   7
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton optSimple 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Simples"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optProyecto 
            BackColor       =   &H00DDDDDD&
            Caption         =   "A Proyectos"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1800
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo cmbEgreso 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   1560
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProyecto 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIngreso 
         Height          =   330
         Left            =   1080
         TabIndex        =   10
         Top             =   1920
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ingreso"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Egreso"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proyecto"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   1260
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4013
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa Egreso"
      Height          =   375
      Left            =   173
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "frmV_ImpGuiaRemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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
Private strSQL As String
Private clsSql As New clsConsulta

Private Sub Command1_Click()
    
End Sub

Private Sub cmdVisPreIng_Click()
'    If dcmbCodigo.MatchedWithList Then
        
    'End If
    If optProyecto = True And cmbProyecto <> "" And cmbIngreso <> "" Then
        drptDevGuia.strTipoGuia = "DPR"
        drptDevGuia.Tag = cmbIngreso.BoundText
        drptDevGuia.Show
        
    ElseIf optSimple = True And cmbIngreso <> "" Then
        drptDevGuia.strTipoGuia = "DRE"
        drptDevGuia.Tag = cmbIngreso.BoundText
        drptDevGuia.Show
    Else
        MsgBox "No ha seleccionado un egreso", vbInformation, "Egreso"
    End If
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

Private Sub cmbProyecto_Change()
    If optProyecto = True And cmbProyecto.MatchedWithList = True Then
        strSQL = " SELECT CONCAT(COALESCE(egreso.egr_factura,''),' - ',egreso.egr_codigo) egre, egreso.egr_codigo " & _
                 " FROM det_pro_tra " & _
                 " INNER JOIN egreso ON (det_pro_tra.det_pro_tra_codigo = egreso.egr_codigo) " & _
                 " AND (det_pro_tra.det_pro_tra_tipo = egreso.tip_egr_codigo) AND (det_pro_tra.emp_codigo = egreso.emp_codigo) " & _
                 " Where egreso.emp_codigo='" & strEmpresa & "' And egreso.tip_egr_codigo='NRP' " & _
                 " And det_pro_tra.pro_tra_codigo=" & cmbProyecto.BoundText & _
                 " ORDER BY egreso.egr_factura+1 "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbEgreso.RowSource = clsSql.adorec_Def.DataSource
        cmbEgreso.ListField = "egre"
        cmbEgreso.BoundColumn = "egr_codigo"
        cmbEgreso = ""
        strSQL = " SELECT CONCAT(COALESCE(ingreso.ing_factura,''),' - ',ingreso.ing_codigo) ingre, ingreso.ing_codigo " & _
                 " FROM det_pro_tra " & _
                 " INNER JOIN ingreso ON (det_pro_tra.det_pro_tra_codigo = ingreso.ing_codigo) " & _
                 " AND (det_pro_tra.det_pro_tra_tipo = ingreso.tip_ing_codigo) AND (det_pro_tra.emp_codigo = ingreso.emp_codigo) " & _
                 " Where ingreso.emp_codigo='" & strEmpresa & "' And ingreso.tip_ing_codigo='DPR' " & _
                 " And det_pro_tra.pro_tra_codigo=" & cmbProyecto.BoundText & _
                 " ORDER BY ingreso.ing_factura+1 "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbIngreso.RowSource = clsSql.adorec_Def.DataSource
        cmbIngreso.ListField = "ingre"
        cmbIngreso.BoundColumn = "ing_codigo"
        cmbIngreso = ""
    ElseIf optSimple = True And cmbProyecto.MatchedWithList = True Then
        strSQL = " SELECT CONCAT(COALESCE(egreso.egr_factura,''),' - ',egreso.egr_codigo) AS egre, egreso.egr_codigo " & _
                " FROM egreso " & _
                " Where egreso.emp_codigo='" & strEmpresa & "' And egreso.tip_egr_codigo='GRE' " & _
                " AND per_codigo = '" & cmbProyecto.BoundText & "' " & _
                " ORDER BY egreso.egr_factura+1 "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbEgreso.RowSource = clsSql.adorec_Def.DataSource
        cmbEgreso.ListField = "egre"
        cmbEgreso.BoundColumn = "egr_codigo"
        cmbEgreso = ""
        strSQL = " SELECT CONCAT(COALESCE(ingreso.ing_factura,''), ' - ',ingreso.ing_codigo) AS ingre, ingreso.ing_codigo " & _
                 " FROM ingreso " & _
                 " Where ingreso.emp_codigo='" & strEmpresa & "' And ingreso.tip_ing_codigo='DRE' " & _
                 " AND per_codigo = '" & cmbProyecto.BoundText & "' " & _
                 " ORDER BY ingreso.ing_factura+1 "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbIngreso.RowSource = clsSql.adorec_Def.DataSource
        cmbIngreso.ListField = "ingre"
        cmbIngreso.BoundColumn = "ing_codigo"
        cmbIngreso = ""
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If optProyecto = True And cmbProyecto <> "" And cmbEgreso <> "" Then
'        drptGuiaRemisionPR.strTipoGuia = "NRP"
'        drptGuiaRemisionPR.PROY = cmbProyecto.BoundText
'        drptGuiaRemisionPR.Tag = cmbEgreso.BoundText
'        drptGuiaRemisionPR.Show
        frmReporte.strNumero = cmbEgreso.BoundText
        frmReporte.strTipo = "NRP"
        frmReporte.strReporte = "rptGuiaRemision2"
        frmReporte.Show
    ElseIf optSimple = True And cmbEgreso <> "" Then
        frmReporte.strNumero = cmbEgreso.BoundText
        frmReporte.strTipo = "GRE"
        frmReporte.strReporte = "rptGuiaRemision2"
        frmReporte.Show
    Else
        MsgBox "No ha seleccionado un egreso", vbInformation, "Egreso"
    End If
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - mdiPrincipal.Height / 40
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    optSimple.Value = True
    lblTipo.Caption = "Cliente"
    cmbProyecto.Enabled = True
    CargarTipo "C"
    strSQL = " SELECT CONCAT(COALESCE(egreso.egr_factura,''),' - ',egreso.egr_codigo) AS egre, egreso.egr_codigo " & _
             " FROM egreso " & _
             " Where egreso.emp_codigo='" & strEmpresa & "' And egreso.tip_egr_codigo='GRE' " & _
             " ORDER BY egreso.egr_factura+1 "
    clsSql.Ejecutar strSQL, "LOCAL"
    Set cmbEgreso.RowSource = clsSql.adorec_Def.DataSource
    cmbEgreso.ListField = "egre"
    cmbEgreso.BoundColumn = "egr_codigo"
    cmbEgreso = ""
    
    strSQL = " SELECT CONCAT(COALESCE(ingreso.ing_factura,''),' - ',ingreso.ing_codigo)  AS ingre, ingreso.ing_codigo " & _
             " FROM ingreso " & _
             " Where ingreso.emp_codigo='" & strEmpresa & "' And ingreso.tip_ing_codigo='DRE' " & _
             " ORDER BY ingreso.ing_factura "
    clsSql.Ejecutar strSQL, "LOCAL"
    Set cmbIngreso.RowSource = clsSql.adorec_Def.DataSource
    cmbIngreso.ListField = "ingre"
    cmbIngreso.BoundColumn = "ing_codigo"
    cmbIngreso = ""
End Sub

Private Sub CargarTipo(tip As String)
    If tip = "C" Then
        strSQL = " SELECT CONCAT(per_apellido,' ', per_nombre) as nombre, per_codigo " & _
             " FROM persona " & _
             " Where emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY 1 "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbProyecto.RowSource = clsSql.adorec_Def.DataSource
        cmbProyecto.ListField = "nombre"
        cmbProyecto.BoundColumn = "per_codigo"
    Else
        strSQL = " SELECT CONCAT(per_apellido,' ', per_nombre,' (',pro_tra_codigo, ') ',SUBSTRING(pro_ven_descricion,1,40),'...') as pro, pro_tra_codigo " & _
             " FROM proyecto_venta INNER JOIN persona ON proyecto_venta.per_codigo=persona.per_codigo AND proyecto_venta.emp_codigo=persona.emp_codigo " & _
             " INNER JOIN pro_trabajo ON (proyecto_venta.pro_ven_codigo = pro_trabajo.pro_ven_codigo) " & _
             " AND (proyecto_venta.emp_codigo = pro_trabajo.emp_codigo) " & _
             " Where proyecto_venta.emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY pro_tra_fecha, pro_tra_codigo "
        clsSql.Ejecutar strSQL, "LOCAL"
        Set cmbProyecto.RowSource = clsSql.adorec_Def.DataSource
        cmbProyecto.ListField = "pro"
        cmbProyecto.BoundColumn = "pro_tra_codigo"
    End If
End Sub



Private Sub optProyecto_Click()
    lblTipo.Caption = "Proyecto"
    cmbProyecto.Enabled = True
    Set cmbEgreso.RowSource = Nothing
    cmbProyecto = ""
    cmbEgreso = ""
    cmbIngreso = ""
    CargarTipo "P"
End Sub

Private Sub optSimple_Click()
    lblTipo.Caption = "Cliente"
    cmbProyecto.Enabled = True
    Set cmbEgreso.RowSource = Nothing
    cmbProyecto = ""
    cmbEgreso = ""
    cmbIngreso = ""
    CargarTipo "C"
'    strSql = " SELECT CONCAT(COALESCE(egreso.egr_factura,''),' - ',egreso.egr_codigo) AS egre, egreso.egr_codigo " & _
'             " FROM egreso " & _
'             " Where egreso.emp_codigo='" & strEmpresa & "' And egreso.tip_egr_codigo='GRE' " & _
'             " ORDER BY egreso.egr_factura+1 "
'    clsSql.Ejecutar strSql, "LOCAL"
'    Set cmbEgreso.RowSource = clsSql.adorec_Def.DataSource
'    cmbEgreso.ListField = "egre"
'    cmbEgreso.BoundColumn = "egr_codigo"
'    cmbProyecto = ""
'    cmbEgreso = ""
'    strSql = " SELECT CONCAT(COALESCE(ingreso.ing_factura,''), ' - ',ingreso.ing_codigo) AS ingre, ingreso.ing_codigo " & _
'             " FROM ingreso " & _
'             " Where ingreso.emp_codigo='" & strEmpresa & "' And ingreso.tip_ing_codigo='DRE' " & _
'             " ORDER BY ingreso.ing_factura+1 "
'    clsSql.Ejecutar strSql, "LOCAL"
'    Set cmbIngreso.RowSource = clsSql.adorec_Def.DataSource
'    cmbIngreso.ListField = "ingre"
'    cmbIngreso.BoundColumn = "ing_codigo"
'    cmbIngreso = ""
    
End Sub
