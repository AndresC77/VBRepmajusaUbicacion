VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerManifiestoCarga 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifiesto de Carga"
   ClientHeight    =   5940
   ClientLeft      =   6285
   ClientTop       =   3150
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerManifiestoCarga.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdImprimirEntregasCliente 
      Caption         =   "Imp.Entregas Cli."
      Height          =   375
      Left            =   6240
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimirEntregas 
      Caption         =   "Imp.Entregas"
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar Guias"
      Height          =   375
      Left            =   2888
      TabIndex        =   26
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1238
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtPlaca 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtOperador 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1560
      Width           =   7815
   End
   Begin VB.TextBox txtManifiesto 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1170
      Width           =   1815
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   5160
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2175
      Left            =   90
      TabIndex        =   15
      Top             =   2880
      Width           =   8775
      _cx             =   2088778870
      _cy             =   2088767228
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerManifiestoCarga.frx":030A
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
   Begin VB.TextBox TxtObserv 
      Height          =   525
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   8
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox txtResponsable 
      Height          =   315
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton cmdImprimirManifiesto 
      Caption         =   "Imp.Manifiesto"
      Height          =   375
      Left            =   4568
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6248
      TabIndex        =   6
      Top             =   5520
      Width           =   1455
   End
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   315
      Left            =   6930
      TabIndex        =   10
      Top             =   1170
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      Value           =   41836.5404166667
      Enabled         =   0   'False
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
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFFactura 
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFContenedor 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFManifiesto 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contenedor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manifiesto"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manifiesto:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   1215
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Guias:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   17
      Top             =   5175
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1965
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4680
      TabIndex        =   12
      Top             =   1935
      Width           =   990
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   6240
      TabIndex        =   11
      Top             =   1215
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1620
      Width           =   735
   End
End
Attribute VB_Name = "frmVerManifiestoCarga"
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
Private strSql As String
Private clsSql As New clsConsulta

Private Sub LimpiarForm()
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    TxtTotal.Text = 0
End Sub

Private Sub cmdAgregar_Click()
    frmManifiestoCarga.Tag = txtManifiesto.Text
    frmManifiestoCarga.cmbCourier.BoundText = txtOperador.Tag
    frmManifiestoCarga.txtPlaca.Text = txtPlaca.Text
    frmManifiestoCarga.txtResponsable.Text = txtResponsable.Text
    frmManifiestoCarga.TxtObserv.Text = TxtObserv.Text
    frmManifiestoCarga.cmbCourier.Locked = True
    frmManifiestoCarga.Show
    Unload Me
End Sub

Private Sub cmdConsultar_Click()
    BuscarManifiesto txtFManifiesto.Text, txtFContenedor.Text, txtFFactura.Text
End Sub

Private Sub cmdImprimirEntregas_Click()
    frmReporte.strNumero = txtManifiesto.Text
    'frmReporte.VSPrint
    frmReporte.strReporte = "rptManifiestoEntregas"
    frmReporte.Show

End Sub

Private Sub cmdImprimirEntregasCliente_Click()
    frmReporte.strNumero = txtManifiesto.Text
    'frmReporte.VSPrint
    frmReporte.strReporte = "rptManifiestoEntregasCliente"
    frmReporte.Show
End Sub

Private Sub cmdImprimirManifiesto_Click()
    frmReporte.strNumero = txtManifiesto.Text
    frmReporte.strReporte = "rptManifiestoCarga"
    frmReporte.Show
End Sub

Private Sub cmdNuevo_Click()
    frmManifiestoCarga.Tag = "N"
    frmManifiestoCarga.cmbCourier.Locked = False
    frmManifiestoCarga.Show
    Unload Me
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    clsSql.Inicializar AdoConn, AdoConnMaster
'    strSql = " SELECT par_texto FROM parametro WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='IED'"
'    clsSql.Ejecutar strSql
'    ImpresoraEtiqueta = clsSql.adorec_Def("par_texto")
        
    dtpFecha.Value = HoyDia
End Sub

Private Sub BuscarManifiesto(strManifiesto As String, strContenedor As String, strFactura As String)
    If Trim(strManifiesto) <> "" Then
        strManifiesto = strManifiesto
    ElseIf Trim(strContenedor) <> "" Then
        strSql = " SELECT det_manifiesto_carga.man_car_codigo " & _
                 " FROM det_manifiesto_carga " & _
                 " WHERE det_manifiesto_carga.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_manifiesto_carga.con_codigo='" & strContenedor & "'"
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strManifiesto = clsSql.adorec_Def("man_car_codigo")
        Else
            strManifiesto = ""
        End If
    ElseIf Trim(strFactura) <> "" Then
        strSql = " SELECT det_manifiesto_carga.man_car_codigo " & _
                 " FROM det_contenedor INNER JOIN det_manifiesto_carga " & _
                 " ON det_contenedor.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                 " AND det_contenedor.con_codigo=det_manifiesto_carga.con_codigo " & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_contenedor.egr_codigo='" & strFactura & "'"
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strManifiesto = clsSql.adorec_Def("man_car_codigo")
        Else
            strManifiesto = ""
        End If
    End If
    strSql = " SELECT man_car_fecha, cou_nombre, " & _
             " manifiesto_carga.cou_codigo,cou_nombre, man_car_placa,man_car_responsable,man_car_observacion " & _
             " FROM manifiesto_carga INNER JOIN courier " & _
             " ON manifiesto_carga.emp_codigo=courier.emp_codigo " & _
             " AND manifiesto_carga.cou_codigo=courier.cou_codigo " & _
             " WHERE manifiesto_carga.emp_codigo='" & strEmpresa & "' " & _
             " AND manifiesto_carga.man_car_codigo='" & strManifiesto & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        txtManifiesto.Text = strManifiesto
        dtpFecha.Value = clsSql.adorec_Def("man_car_fecha")
        txtOperador.Text = clsSql.adorec_Def("cou_nombre")
        txtOperador.Tag = clsSql.adorec_Def("cou_codigo")
        txtPlaca.Text = clsSql.adorec_Def("man_car_placa")
        txtResponsable.Text = clsSql.adorec_Def("man_car_responsable")
        TxtObserv.Text = clsSql.adorec_Def("man_car_observacion")
        
        strSql = " SELECT det_manifiesto_carga.con_codigo,con_guia," & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,det_manifiesto_carga.paq_env_codigo,paq_env_nombre,COUNT(det_man_car_codigo) " & _
                 " FROM det_manifiesto_carga INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                 " INNER JOIN paquete_envio ON det_manifiesto_carga.emp_codigo=paquete_envio.emp_codigo " & _
                 " AND det_manifiesto_carga.paq_env_codigo=paquete_envio.paq_env_codigo " & _
                 " INNER JOIN persona pd ON contenedor.emp_codigo=pd.emp_codigo " & _
                 " AND contenedor.per_codigo=pd.per_codigo AND pd.cat_p_tipo='C' " & _
                 " WHERE det_manifiesto_carga.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_manifiesto_carga.man_car_codigo='" & strManifiesto & "' " & _
                 " GROUP BY det_manifiesto_carga.con_codigo,con_guia,pd.per_apellido,pd.per_nombre,det_manifiesto_carga.paq_env_codigo,paq_env_nombre"
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        TxtTotal.Text = VSFG.Rows - 1
    Else
        txtManifiesto.Text = ""
        dtpFecha.Value = HoyDia
        txtOperador.Text = ""
        txtPlaca.Text = ""
        txtResponsable.Text = ""
        TxtObserv.Text = ""
        VSFG.Clear 1
        VSFG.Rows = 1
        MsgBox "No se encuentra registro", vbInformation, "Manifiesto de Carga"
    End If
End Sub
