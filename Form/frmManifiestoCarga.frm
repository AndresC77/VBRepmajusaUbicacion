VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManifiestoCarga 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifiesto de Carga"
   ClientHeight    =   7005
   ClientLeft      =   7470
   ClientTop       =   1935
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
   Icon            =   "frmManifiestoCarga.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   8940
   Begin VB.TextBox txtResponsable 
      Height          =   315
      Left            =   5610
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtPlaca 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2925
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4485
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   15
      Top             =   1560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "GUIAS"
      TabPicture(0)   =   "frmManifiestoCarga.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbltipo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "VSFG"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLectorGuia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtLectorPaquete"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optGuia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optContenedor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.OptionButton optContenedor 
         Caption         =   "Contenedor"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   473
         Width           =   1335
      End
      Begin VB.OptionButton optGuia 
         Caption         =   "Guia"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   473
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtLectorPaquete 
         Height          =   315
         Left            =   6840
         TabIndex        =   5
         Top             =   443
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   -71280
         TabIndex        =   20
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
      Begin VB.TextBox txtLectorGuia 
         Height          =   315
         Left            =   3465
         TabIndex        =   4
         Top             =   443
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmManifiestoCarga.frx":0326
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paquete:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6195
         TabIndex        =   23
         Top             =   495
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos:"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   4740
         Width           =   1005
      End
      Begin VB.Label lbltipo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Guia:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2700
         TabIndex        =   17
         Top             =   495
         Width           =   645
      End
   End
   Begin VB.TextBox TxtObserv 
      Height          =   525
      Left            =   1290
      MaxLength       =   250
      TabIndex        =   3
      Top             =   960
      Width           =   7575
   End
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   285
      Left            =   7530
      TabIndex        =   10
      Top             =   90
      Width           =   1335
      _ExtentX        =   3201
      _ExtentY        =   503
      Value           =   41836.5404166667
      Enabled         =   0   'False
   End
   Begin MSDataListLib.DataCombo cmbCourier 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3825
      _ExtentX        =   6747
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4560
      TabIndex        =   22
      Top             =   495
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   525
      Width           =   435
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   165
      Width           =   735
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmManifiestoCarga"
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

Private Sub LimpiarForm()
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    TxtTotal.Text = 0
    
'****** Tipo de paquetes
    'Recupera todas las bodegas de una empresa
    strSql = " SELECT paq_env_codigo, paq_env_nombre " & _
             " FROM paquete_envio " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By paq_env_nombre "
    clsSql.Ejecutar (strSql)
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsSql.adorec_Def, "*paq_env_codigo, paq_env_nombre", "paq_env_codigo")
    
End Sub

Private Sub cmbCourier_Validate(Cancel As Boolean)
    LimpiarForm
End Sub

Private Sub cmdAceptar_Click()
    Dim lngManifiesto As Long
    Dim i As Long
    
    Dim RepEmpaque As New frmReporte
    If FormatoD0(VSFG.Rows) < 2 Then
        If MsgBox("No ha ingresado ningun item a despachar" & vbNewLine & _
            "Quiere modificar empaquetado", vbInformation + vbYesNo, "Despacho") = vbNo Then
            Exit Sub
        End If
    ElseIf cmbCourier.MatchedWithList = False Or cmbCourier.BoundText = "" Then
        MsgBox "No ha seleccionado un Operador", vbInformation, "Despacho"
        Exit Sub
    End If
    
    If Revisarguias = False Then
        For i = 1 To VSFG.Rows - 1
            If VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbWhite Then
                MsgBox "Se esta despachando una guia (" & VSFG.TextMatrix(i, 0) & ") incompleta el numero ce cajas", vbInformation, "Despacho"
                Exit Sub
            End If
        Next i
    End If
    
    If frmManifiestoCarga.Tag = "N" Then
        strSql = " SELECT COALESCE(MAX(man_car_codigo),0) as n" & _
                 " FROM manifiesto_carga " & _
                 " WHERE emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            lngManifiesto = FormatoD0(clsSql.adorec_Def("n")) + 1
        Else
            lngManifiesto = 1
        End If
    
        strSql = " INSERT INTO manifiesto_carga (emp_codigo, man_car_codigo, cou_codigo, man_car_placa, " & _
                 " man_car_responsable, man_car_fecha, man_car_observacion, man_car_fechamod, man_car_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & lngManifiesto & "','" & cmbCourier.BoundText & "','" & UCase(txtPlaca.Text) & "', " & _
                 " '" & UCase(txtResponsable.Text) & "','" & dtpFecha.Value & "','" & UCase(TxtObserv.Text) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
    Else
        lngManifiesto = frmManifiestoCarga.Tag
    
        strSql = " UPDATE manifiesto_carga  " & _
                 " SET man_car_placa='" & UCase(txtPlaca.Text) & "'," & _
                 " man_car_responsable='" & UCase(txtResponsable.Text) & "', " & _
                 " man_car_observacion='" & UCase(TxtObserv.Text) & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND man_car_codigo='" & lngManifiesto & "' "
        clsSql.Ejecutar strSql, "M"
    End If
    For i = 1 To VSFG.Rows - 1
        strSql = " INSERT INTO det_manifiesto_carga (emp_codigo, man_car_codigo, con_codigo, " & _
                 " paq_env_codigo, det_man_car_fechamod, det_man_car_usumod) " & _
                 " VALUES('" & strEmpresa & "','" & lngManifiesto & "','" & VSFG.TextMatrix(i, 0) & "', " & _
                 " '" & VSFG.TextMatrix(i, 3) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
    Next i
    MsgBox "Manifiesto de Carga Guardado No. " & lngManifiesto, vbInformation, "Despachos"
    
    RepEmpaque.strNumero = lngManifiesto
    RepEmpaque.strReporte = "rptManifiestoCarga"
    RepEmpaque.Show
    'RepEmpaque.Form_Activate
    frmVerManifiestoCarga.Show
    RepEmpaque.SetFocus
    Unload Me
    
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
    EnviarMail NombreComercial & " Despachos", "acevallos@rbimportadores.com", "Sandra Alvarez", "it@rbimportadores.com", "it@rbimportadores.com", "Lista de Embarque", "Adjunto esta la lista de embarque", "ListaEmbarque1.pdf"
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
    
End Sub

Private Sub cargarCombos()
    
    strSql = " SELECT cou_codigo, cou_nombre " & _
             " FROM courier " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbCourier.RowSource = clsSql.adorec_Def.DataSource
    cmbCourier.ListField = "cou_nombre"
    cmbCourier.BoundColumn = "cou_codigo"
    
'****** Tipo de paquetes
    'Recupera todas las bodegas de una empresa
    strSql = " SELECT paq_env_codigo, paq_env_nombre " & _
             " FROM paquete_envio " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " Order By paq_env_nombre "
    clsSql.Ejecutar (strSql)
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsSql.adorec_Def, "*paq_env_codigo, paq_env_nombre", "paq_env_codigo")
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLectorGuia" And Screen.ActiveControl.Name <> "txtLectorPaquete" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    cargarCombos
    dtpFecha.Value = HoyDia
End Sub

Private Sub optContenedor_Click()
    If optContenedor.Value = True Then
        lblTipo.Caption = "No.Cont:"
    Else
        lblTipo.Caption = "No.Guia:"
    End If
End Sub

Private Sub optGuia_Click()
    If optGuia.Value = True Then
        lblTipo.Caption = "No.Guia:"
    Else
        lblTipo.Caption = "No.Cont:"
    End If
End Sub

Private Sub txtLectorGuia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'txtLectorPaquete.SetFocus
        AgregarItem2 optGuia.Value, UCase(txtLectorGuia.Text)
        txtLectorGuia.Text = ""
        txtLectorGuia.SetFocus
    End If
End Sub

Private Sub txtLectorPaquete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarItem optGuia.Value, UCase(txtLectorGuia.Text), Left(UCase(txtLectorPaquete.Text), 3)
        txtLectorGuia.Text = ""
        txtLectorPaquete.Text = ""
        txtLectorGuia.SetFocus
    End If
End Sub

Private Function AgregarItem2(booTipoGuia As Boolean, strGuia As String) As Integer
    Dim clsSqlAux As New clsConsulta
    Dim i As Long
    Dim Encontro1 As Boolean
    Dim Encontro2 As Boolean
    Dim strContenedor As String
    Dim strCliente As String
    strGuia = Trim(strGuia)
    
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT contenedor.con_codigo,contenedor.con_guia, " & _
             " CONCAT(per_apellido,' ',per_nombre) as cli,paq_env_codigo,det_con_caj_peso " & _
             " FROM contenedor INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo " & _
             " AND contenedor.per_codigo=persona.per_codigo " & _
             " INNER JOIN det_contenedor_caja ON contenedor.emp_codigo=det_contenedor_caja.emp_codigo" & _
             " AND contenedor.con_codigo=det_contenedor_caja.con_codigo" & _
             " WHERE contenedor.emp_codigo='" & strEmpresa & "' "
    If booTipoGuia = True Then
        strSql = strSql & " AND contenedor.con_guia='" & strGuia & "'"
    Else
        strSql = strSql & " AND contenedor.con_codigo='" & strGuia & "'"
    End If
    strSql = strSql & " AND contenedor.cou_codigo='" & cmbCourier.BoundText & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount = 1 Then
    
        If FormatoD2(clsSql.adorec_Def("det_con_caj_peso")) <> 0 Then
            strContenedor = clsSql.adorec_Def("con_codigo")
            strGuia = clsSql.adorec_Def("con_guia")
            strCliente = clsSql.adorec_Def("cli")
            strPaquete = clsSql.adorec_Def("paq_env_codigo")
            VSFG.AddItem strContenedor & vbTab & _
                         strGuia & vbTab & _
                         strCliente & vbTab & _
                         strPaquete
            VSFG.Cell(flexcpBackColor, VSFG.Rows - 1, 0, VSFG.Rows - 1, VSFG.Cols - 1) = vbYellow
            VSFG.ShowCell VSFG.Rows - 1, VSFG.Col
            TxtTotal.Text = FormatoD0(TxtTotal.Text) + 1
        Else
            MsgBox "La guia no puede ser despachada." & vbNewLine & _
                    "No tiene ingresado el peso. ", vbCritical, "Despachos"
        End If
    
    ElseIf clsSql.adorec_Def.RecordCount > 1 Then
        strContenedor = clsSql.adorec_Def("con_codigo")
        strGuia = clsSql.adorec_Def("con_guia")
        strCliente = clsSql.adorec_Def("cli")
        strPaquete = clsSql.adorec_Def("paq_env_codigo")
        Encontro1 = False
        Encontro2 = False
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 0) = strContenedor And VSFG.TextMatrix(i, 1) = strGuia And VSFG.TextMatrix(i, 2) = strCliente Then
                Encontro1 = True
                If VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbWhite Then
                    VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
                    Encontro2 = True
                    TxtTotal.Text = FormatoD0(TxtTotal.Text) + 1
                    Exit For
                End If
            End If
        Next i
        
        If Encontro1 = False And Encontro2 = False Then
            i = 0
            While Not clsSql.adorec_Def.EOF
                strContenedor = clsSql.adorec_Def("con_codigo")
                strGuia = clsSql.adorec_Def("con_guia")
                strCliente = clsSql.adorec_Def("cli")
                strPaquete = clsSql.adorec_Def("paq_env_codigo")
                VSFG.AddItem strContenedor & vbTab & _
                             strGuia & vbTab & _
                             strCliente & vbTab & _
                             strPaquete
                VSFG.Cell(flexcpBackColor, VSFG.Rows - 1, 0, VSFG.Rows - 1, VSFG.Cols - 1) = vbWhite
                VSFG.ShowCell VSFG.Rows - 1, VSFG.Col
                If i = 0 Then
                    VSFG.Cell(flexcpBackColor, VSFG.Rows - 1, 0, VSFG.Rows - 1, VSFG.Cols - 1) = vbYellow
                    TxtTotal.Text = FormatoD0(TxtTotal.Text) + 1
                    i = i + 1
                End If
                clsSql.adorec_Def.MoveNext
            Wend
        ElseIf Encontro1 = True And Encontro2 = False Then
            MsgBox "La caja no puede ser despachada." & vbNewLine & _
                    "Esta sacando mas cajas de las ingresadas. ", vbCritical, "Despachos"
        End If
    Else
            MsgBox "La guia no puede ser despachada." & vbNewLine & _
                    "Error en guia o en operador. ", vbCritical, "Despachos"
    End If
    
End Function

Private Function AgregarItem(booTipoGuia As Boolean, strGuia As String, strPaquete As String) As Integer
    Dim clsSqlAux As New clsConsulta
    Dim j As Long
    Dim strContenedor As String
    Dim strCliente As String
    strGuia = Trim(strGuia)
    strPaquete = Trim(strPaquete)
    
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT paq_env_nombre " & _
             " FROM paquete_envio " & _
             " WHERE paquete_envio.emp_codigo='" & strEmpresa & "' " & _
             " AND paquete_envio.paq_env_codigo='" & strPaquete & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        If booTipoGuia = True Then
            strSql = " SELECT contenedor.con_codigo,contenedor.con_guia, " & _
                     " CONCAT(per_apellido,' ',per_nombre) as cli,con_peso " & _
                     " FROM contenedor INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo " & _
                     " AND contenedor.per_codigo=persona.per_codigo " & _
                     " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                     " AND contenedor.con_guia='" & strGuia & "'" & _
                     " AND contenedor.cou_codigo='" & cmbCourier.BoundText & "'"
        Else
            strSql = " SELECT contenedor.con_codigo,contenedor.con_guia, " & _
                     " CONCAT(per_apellido,' ',per_nombre) as cli,con_peso " & _
                     " FROM contenedor INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo " & _
                     " AND contenedor.per_codigo=persona.per_codigo " & _
                     " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                     " AND contenedor.con_codigo='" & strGuia & "'" & _
                     " AND contenedor.cou_codigo='" & cmbCourier.BoundText & "'"
        End If
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            If FormatoD2(clsSql.adorec_Def("con_peso")) <> 0 Then
                strContenedor = clsSql.adorec_Def("con_codigo")
                strGuia = clsSql.adorec_Def("con_guia")
                strCliente = clsSql.adorec_Def("cli")
                
                VSFG.AddItem strContenedor & vbTab & _
                             strGuia & vbTab & _
                             strCliente & vbTab & _
                             strPaquete
                             
                VSFG.ShowCell VSFG.Rows - 1, VSFG.Col
                TxtTotal.Text = VSFG.Rows - 1
            Else
                MsgBox "La guia no puede ser despachada." & vbNewLine & _
                        "No tiene ingresado el peso. ", vbCritical, "Despachos"
            End If
        Else
            MsgBox "La guia no puede ser despachada." & vbNewLine & _
                    "Ese numero de guia incorrecto " & vbNewLine & _
                    "o error en Courier ", vbCritical, "Despachos"
        End If
    Else
            MsgBox "Tiene error en el tipo de paquete.", vbCritical, "Despachos"
    End If
    
End Function

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
