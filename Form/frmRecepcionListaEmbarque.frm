VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecepcionListaEmbarque 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion Lista de Embarque"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmRecepcionListaEmbarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8940
   Begin VB.TextBox txtGuia 
      Height          =   315
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   870
      Width           =   1815
   End
   Begin VB.TextBox TxtObserv 
      Height          =   525
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   7
      Top             =   1230
      Width           =   6615
   End
   Begin VB.TextBox txtContenedor 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtCliente 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   510
      Width           =   7815
   End
   Begin VB.TextBox txtOperador 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   870
      Width           =   3375
   End
   Begin VB.TextBox txtPeso 
      Height          =   315
      Left            =   7650
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4725
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6285
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   315
      Left            =   6930
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      Value           =   41836.5404166667
      Enabled         =   0   'False
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   17
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "PEDIDOS"
      TabPicture(0)   =   "frmRecepcionListaEmbarque.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "VSFG"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtmail"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtLector"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotalRecibidos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Impresora"
      TabPicture(1)   =   "frmRecepcionListaEmbarque.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtImpresora"
      Tab(1).Control(1)=   "cmdCambiar"
      Tab(1).Control(2)=   "chkImprimirSTK"
      Tab(1).Control(3)=   "Label9"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtTotalRecibidos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtImpresora 
         Height          =   525
         Left            =   -72870
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   26
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar Impresora"
         Height          =   375
         Left            =   -72840
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkImprimirSTK 
         Caption         =   "NO Imprimir Stiker"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -74280
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtLector 
         Height          =   315
         Left            =   990
         TabIndex        =   20
         Top             =   420
         Width           =   2415
      End
      Begin VB.TextBox txtmail 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   4215
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   -71280
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
         _cx             =   5953
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3855
         Left            =   120
         TabIndex        =   21
         Top             =   780
         Width           =   8655
         _cx             =   15266
         _cy             =   6800
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRecepcionListaEmbarque.frx":0342
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Recibidos:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2280
         TabIndex        =   29
         Top             =   4725
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Etiquetas:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74400
         TabIndex        =   27
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Pedido:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   495
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   4740
         Width           =   1005
      End
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   570
      Width           =   525
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   6240
      TabIndex        =   15
      Top             =   165
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "No.Guia:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4560
      TabIndex        =   14
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   915
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenedor:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7200
      TabIndex        =   10
      Top             =   885
      Width           =   405
   End
End
Attribute VB_Name = "frmRecepcionListaEmbarque"
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
Private emailLista As String
Private emailPapaLista As String


Private Sub chkImprimirSTK_Click()
    If chkImprimirSTK.Value = 1 Then
        ImprimirEtiquetaDespacho = True
    Else
        ImprimirEtiquetaDespacho = False
    End If
    
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.MatchedWithList = True Then
        LimpiarForm
    End If
End Sub

Private Sub LimpiarForm()
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    VSFG2.Clear flexClearScrollable
    VSFG2.Rows = 1
    TxtTotal.Text = 0
    txtmail.Text = ""
    
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND (per_codigo_ref='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref2='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref3='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref4='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref5='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref6='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref7='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref8='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref9='" & cmbCliente.BoundText & "' " & _
             " OR per_codigo_ref10='" & cmbCliente.BoundText & "') " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente2.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente2.ListField = "nombC"
    cmbCliente2.BoundColumn = "per_codigo"
    
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
    LimpiarForm
    cmbCliente.BoundText = ""
     
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nombC " & _
             " FROM persona " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND (per_es_gz=1 " & _
             " OR per_es_di=1 " & _
             " OR per_es_em=1 " & _
             " OR per_es_ee=1 " & _
             " OR per_es_n5=1 " & _
             " OR per_es_n6=1 " & _
             " OR per_es_n7=1 " & _
             " OR per_es_n8=1 " & _
             " OR per_es_n9=1 " & _
             " OR per_es_n10=1) " & _
             " ORDER BY nombC "
    clsSql.Ejecutar (strSql)
    
    Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
        
    cmbCliente.ListField = "nombC"
    cmbCliente.BoundColumn = "per_codigo"
    
End Sub

Private Sub cmdAceptar_Click()
    Dim lngContenedor As Long
    Dim i As Long
    
    For i = 1 To VSFG.Rows - 1
        If VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Rows - 1) = vbCyan Then
            strSql = " UPDATE det_contenedor  " & _
                     " SET det_con_estado=1, " & _
                     " det_con_fecha=CURRENT_TIMESTAMP, " & _
                     " det_con_fechamod=CURRENT_TIMESTAMP, " & _
                     " det_con_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND con_codigo='" & txtContenedor.Text & "' " & _
                     " AND egr_codigo='" & VSFG.TextMatrix(i, 1) & "'"
            clsSql.Ejecutar strSql, "M"
        End If
    Next i
    MsgBox "Pedidos Recibidos de la Lista de Embarque " & txtContenedor.Text, vbInformation, "Contenedores"
    
    frmVerListaEmbarque.Show
    frmVerListaEmbarque.cmdRecibirLista.Visible = True
    frmVerListaEmbarque.cmdNuevo.Visible = False
    frmVerListaEmbarque.cmdCambiarOperador.Visible = False
    frmVerListaEmbarque.cmdImprimirListado.Visible = True
    frmVerListaEmbarque.cmdImprimirEtiqueta.Visible = False
    frmVerListaEmbarque.cmdEnviarCorreo.Visible = False
    Unload Me
End Sub

Private Sub cmdAgregar_Click()
    AgregarDetalle cmbCliente2.BoundText, UCase(txtDescripcion.Text)
    cmbCliente2.BoundText = ""
    txtDescripcion.Text = ""
End Sub

Private Sub cmdCambiar_Click()
    VSPrinterAUX.PrintDialog pdPrint
    ImpresoraEtiqueta = VSPrinterAUX.Device
    txtImpresora.Text = ImpresoraEtiqueta
End Sub

Private Sub Command1_Click()
'    Dim RepStk As New frmReporte
'    RepStk.VSPrint.Device = ImpresoraEtiqueta
'    RepStk.VSPrint.PaperWidth = 7669.292
'    RepStk.VSPrint.PaperHeight = 3885.039
'    RepStk.strNumero = 1
'    RepStk.strReporte = "rptSTKListaEmbarque"
'    RepStk.Show
'    RepStk.Form_Activate
'    RepStk.VSPrint.PrintDoc
'    Unload RepStk
'    Dim RepEmpaque As New frmReporte
'    RepEmpaque.strNumero = 1
'    RepEmpaque.strReporte = "rptListaEmbarque"
'    RepEmpaque.Show
'    RepEmpaque.Form_Activate
'    RepEmpaque.VSPrint.PrintDoc
'    Unload RepEmpaque
Dim RepEmpaque As New frmReporte
    RepEmpaque.strNumero = 1
    RepEmpaque.strReporte = "rptListaEmbarque"
    'RepEmpaque.Show
    RepEmpaque.Form_Activate
    RepEmpaque.VSRpt.RenderToFile "ListaEmbarque1.pdf", vsrPDF
    'RepEmpaque.VSPrint.PrintDoc
    Unload RepEmpaque
    EnviarMail NombreComercial & " Despachos", "acevallos@rbimportadores.com", "Sandra Alvarez", _
    "it@rbimportadores.com", "it@rbimportadores.com", "Lista de Embarque", _
    "Adjunto esta la lista de embarque", "ListaEmbarque1.pdf"
    Kill "ListaEmbarque1.pdf"
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
    If ImprimirEtiquetaDespacho = True Then
        chkImprimirSTK.Value = 1
    Else
        chkImprimirSTK.Value = 0
    End If
    
    strSql = " SELECT est_codigo,est_descripcion FROM est_contenedor ORDER BY est_codigo"
    clsSql.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
    VSFG.AutoSize 7
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    txtImpresora.Text = ImpresoraEtiqueta
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarPedido UCase(txtLector.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub AgregarPedido(strPedido As String)
    Dim j As Long
    Dim booEncontro As Boolean
    booEncontro = False
    For j = 1 To VSFG.Rows - 1
        If strPedido = VSFG.TextMatrix(j, 0) Then
            If VSFG.Cell(flexcpBackColor, j, 0, j, VSFG.Cols - 1) = vbWhite Or VSFG.Cell(flexcpBackColor, j, 0, j, VSFG.Cols - 1) = 0 Then
                VSFG.TextMatrix(j, 7) = 1
                VSFG.Cell(flexcpBackColor, j, 0, j, VSFG.Cols - 1) = vbCyan
            Else
                MsgBox "Ese pedido ya lo escaneó", vbInformation, "Recepción De Pedidos"
            End If
            booEncontro = True
        End If
    Next j
    txtTotalRecibidos.Text = 0
    For j = 0 To VSFG.Rows - 1
        If VSFG.Cell(flexcpBackColor, j, 0, j, VSFG.Cols - 1) <> vbWhite And VSFG.Cell(flexcpBackColor, j, 0, j, VSFG.Cols - 1) <> 0 Then
                txtTotalRecibidos.Text = FormatoD0(txtTotalRecibidos.Text) + 1
        End If
    Next j
    
    If booEncontro = False Then
        MsgBox "El pedido " & strPedido & " no se encuentra en esta lista de embarque", vbInformation, "Recepción De Pedidos"
    End If
    
End Sub

Private Sub AgregarDetalle(strCliente As String, strDescripcion As String)
    Dim clsSqlAux As New clsConsulta
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT persona.per_codigo,CONCAT(per_apellido,' ',per_nombre) as cli,'" & strDescripcion & "' as descr, " & _
             " per_direccion2, ciu_nombre, zon_nombre " & _
             " FROM persona " & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
             " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
             " AND (persona.per_codigo_ref='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref2='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref3='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref4='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref5='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref6='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref7='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref8='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref9='" & cmbCliente.BoundText & "'" & _
             " OR persona.per_codigo_ref10='" & cmbCliente.BoundText & "')" & _
             " AND persona.per_codigo='" & strCliente & "'"
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        VSFG2.AddItem clsSql.adorec_Def("per_codigo") & vbTab & _
                     clsSql.adorec_Def("cli") & vbTab & _
                     clsSql.adorec_Def("descr") & vbTab & _
                     clsSql.adorec_Def("per_direccion2") & vbTab & _
                     clsSql.adorec_Def("ciu_nombre") & vbTab & _
                     clsSql.adorec_Def("zon_nombre")
        TxtTotal.Text = VSFG.Rows - 1 + VSFG2.Rows - 1
    Else
        MsgBox "El cliente no es de la red, error en zona, error en ciudad ", vbInformation, "Despachos"
    End If
    
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 And (Col = 7 Or Col = 8) Then
        If VSFG.TextMatrix(Row, 7) = 1 And Trim(VSFG.TextMatrix(Row, 8)) <> "" Then
            VSFG.Cell(flexcpBackColor, Row, 0, Row, VSFG.Cols - 1) = vbBlue
            VSFG.Cell(flexcpForeColor, Row, 0, Row, VSFG.Cols - 1) = vbWhite
        ElseIf VSFG.TextMatrix(Row, 7) = 1 Then
            VSFG.Cell(flexcpBackColor, Row, 0, Row, VSFG.Cols - 1) = vbCyan
        End If
    End If
End Sub

