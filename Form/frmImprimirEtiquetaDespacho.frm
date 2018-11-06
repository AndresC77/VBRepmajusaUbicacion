VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImprimirEtiquetaDespacho 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Etiqueta para Despacho"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImprimirEtiquetaDespacho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Pedido"
      TabPicture(0)   =   "frmImprimirEtiquetaDespacho.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtLector"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdImpGuia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkFacturaTicket"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cliente"
      TabPicture(1)   =   "frmImprimirEtiquetaDespacho.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "lblCodigo"
      Tab(1).Control(2)=   "cmbCliente"
      Tab(1).Control(3)=   "cmbNegocio"
      Tab(1).Control(4)=   "cmdImpGuiaCli"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Impresora"
      TabPicture(2)   =   "frmImprimirEtiquetaDespacho.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtImpresora"
      Tab(2).Control(1)=   "cmdCambiar"
      Tab(2).Control(2)=   "VSPrinterAUX"
      Tab(2).Control(3)=   "Label9"
      Tab(2).ControlCount=   4
      Begin VB.CheckBox chkFacturaTicket 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Ticket"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpGuiaCli 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtImpresora 
         Height          =   525
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   9
         Top             =   600
         Width           =   3945
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar Impresora"
         Height          =   375
         Left            =   -74640
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdImpGuia 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtLector 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3975
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   -74160
         TabIndex        =   2
         Top             =   360
         Width           =   3465
         _ExtentX        =   6112
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   -74160
         TabIndex        =   3
         Top             =   720
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
         _cx             =   7435
         _cy             =   661
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   -2.58236865538736
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   3
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
         AutoLinkNavigate=   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74880
         TabIndex        =   12
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Etiquetas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74730
         TabIndex        =   10
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmImprimirEtiquetaDespacho"
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

Private Sub cmdCambiar_Click()
    VSPrinterAUX.PrintDialog pdPrint
    ImpresoraEtiqueta = VSPrinterAUX.Device
    txtImpresora.Text = ImpresoraEtiqueta
    GuardarImpresoras
End Sub

Private Sub cmdImpGuiaCli_Click()
    Dim RepStk2 As New frmReporte
    If ImpresoraEtiqueta = "" Then
        RepStk2.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = RepStk2.VSPrint.Device
        txtImpresora.Text = ImpresoraEtiqueta
        GuardarImpresoras
    End If
    RepStk2.VSPrint.Device = ImpresoraEtiqueta
    RepStk2.VSPrint.PaperWidth = 7669.292
    RepStk2.VSPrint.PaperHeight = 3885.039
    RepStk2.strNumero = cmbCliente.BoundText
    RepStk2.strReporte = "rptSTKIdCaja"
    RepStk2.Show
    RepStk2.Form_Activate
    RepStk2.VSPrint.PrintDoc
    Unload RepStk2
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

Private Sub cmdImpGuia_Click()
    Dim RepStk As New frmReporte
    Dim clsauxfac As New clsConsulta
    clsauxfac.Inicializar AdoConn, AdoConnMaster
    If ImpresoraEtiqueta = "" Then
        RepStk.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = RepStk.VSPrint.Device
        txtImpresora.Text = ImpresoraEtiqueta
        GuardarImpresoras
    End If
    
    If chkFacturaTicket.Value = 1 Then
    
        clsauxfac.Ejecutar "SELECT ped_egr_codigo FROM pedido where emp_codigo='" & strEmpresa & "' and ped_codigo='" & Trim(txtLector.Text) & "'"
        
        If clsauxfac.adorec_Def("ped_egr_codigo") <> "" Then
        
        frmImpresionDirecta.strNumero = Trim(txtLector.Text)
        frmImpresionDirecta.strReporte = "rptSTKDespacho"
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.Show
        frmImpresionDirecta.cmdImprimir_Click
        frmImpresionDirecta.CmdCerrar_Click
                
        
        
        frmImpresionDirecta.strNumero = clsauxfac.adorec_Def("ped_egr_codigo")
        frmImpresionDirecta.strReporte = "rptFacturaSola"
        frmImpresionDirecta.Show
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.cmdImprimir_Click
        Else
            MsgBox "Pedido no facturado"
        End If
    Else
    
        RepStk.VSPrint.Device = ImpresoraEtiqueta
        
        RepStk.VSPrint.PaperWidth = 7669.292
        RepStk.VSPrint.PaperHeight = 3885.039
        RepStk.strNumero = Trim(txtLector.Text)
        RepStk.strReporte = "rptSTKDespacho"
        RepStk.strTipo = 2
        RepStk.Show
        RepStk.Form_Activate
        RepStk.VSPrint.PrintDoc
        Unload RepStk
    End If
    txtLector.Text = ""
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
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
        End If
    Else
        Exit Sub
    End If
    
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    txtImpresora.Text = ImpresoraEtiqueta
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If


End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdImpGuia_Click
    End If
End Sub
