VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerIngInventario 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Conteos de Inventario"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmVerIngInventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3817
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5497
      TabIndex        =   9
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   98
      TabIndex        =   3
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtFechaIng 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4950
         TabIndex        =   8
         Top             =   435
         Width           =   1410
      End
      Begin VB.TextBox txtObs 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   6705
         Width           =   10110
      End
      Begin MSDataListLib.DataCombo dcmbIngImp 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   427
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfgDetalleImp 
         Height          =   5520
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   10140
         _cx             =   17886
         _cy             =   9737
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerIngInventario.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         FrozenCols      =   1
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblObserv 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   6480
         Width           =   1410
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4080
         TabIndex        =   5
         Top             =   487
         Width           =   495
      End
      Begin VB.Label lblNumIngreso 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Conteo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   487
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   2152
      TabIndex        =   1
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7177
      TabIndex        =   2
      Top             =   7680
      Width           =   1455
   End
End
Attribute VB_Name = "frmVerIngInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para visualizar los ingresos de mercadería realizados por concepto de #
'#  Importaciones,  esta forma es solo de visualización, no permite la edición. #
'#  frmVerIngImp V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los ingresos de mercadería a una determinada  emp-   #
'#  presa por concepto de Importaciones.                                        #
'#  En esta ventana solo se puede visuallizar cualesquiera de los ingresos por  #
'#  este concepto pero no se puede realizar ningún cambio.                      #
'#  Se puede escoger el número de documento o ingresar dicho número en el combo #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se consulta los egresos realizados de tipo     #
'#                 INI.                                                         #
'#    persona    : En esta tabla se consulta los datos del proveedor al que se  #
'#                 le adquirió la mercadería y se importó.                      #
'#    det_ingreso: En esta tabla se consulta los detalles del ingreso.          #
'#    producto   : En esta tabla se consulta el nombre del producto.            #
'#    deposito   : En esta tabla se consulta el nombre del depósito.            #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#    limpiarFxGD() : Permite borrar el flexgrid utilizado para cuando se       #
'#                    realiza un cambio de documento.                           #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsCon_det As New clsConsulta

Private Sub cmdImprimir_Click()
    Dim rpConteo As New frmReporte
    rpConteo.strNumero = dcmbIngImp.Text
    rpConteo.strReporte = "rptConteo"
    rpConteo.Show
End Sub

Private Sub cmdModificar_Click()
    frmIngInventario.NInv = Me.dcmbIngImp
    frmIngInventario.MN = "M"
    frmIngInventario.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_det = Nothing
End Sub


Private Sub cmdNuevo_Click()
    frmIngInventario.MN = "N"
    frmIngInventario.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbIngImp_Change()
    Dim i As Integer
    If dcmbIngImp.MatchedWithList = True Then
        'Despliego los datos segú el dato ingresado o seleccioado en el data combo
        clsConsu.adorec_Def.MoveFirst
        If (dcmbIngImp.Text = "") Then
            Call limpiarFxGD
            txtFechaIng.Text = ""
            txtObs.Text = ""
            Exit Sub
        Else
            clsConsu.Filtrar "inv_codigo='" & dcmbIngImp.Text & "'"
            'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
            txtFechaIng.Text = Format(clsConsu.adorec_Def("inv_fecha"), "yyyy-mmm-dd")
            txtObs.Text = clsConsu.adorec_Def("inv_observacion")
            Call limpiarFxGD
            
            'llenar flexgrid
            strSql = " select dep_nombre, det_inventario.prd_codigo, prd_nombre, det_inv_cantidad" & _
                     " from det_inventario INNER JOIN deposito ON det_inventario.emp_codigo=deposito.emp_codigo AND det_inventario.dep_codigo=deposito.dep_codigo " & _
                     " INNER JOIN producto ON det_inventario.emp_codigo=producto.emp_codigo AND  det_inventario.prd_codigo=producto.prd_codigo " & _
                     " WHERE det_inventario.emp_codigo = '" & strEmpresa & "'" & _
                     " AND inv_codigo = '" & dcmbIngImp.Text & "' "
            clsCon_det.Ejecutar (strSql)
            If (clsCon_det.adorec_Def.RecordCount > 0) Then
                clsCon_det.adorec_Def.MoveFirst
                i = 1
                While Not clsCon_det.adorec_Def.EOF
                    vsfgDetalleImp.AddItem ""
                    vsfgDetalleImp.TextMatrix(i, 0) = i
                    vsfgDetalleImp.TextMatrix(i, 1) = clsCon_det.adorec_Def("dep_nombre")
                    vsfgDetalleImp.TextMatrix(i, 2) = clsCon_det.adorec_Def("prd_codigo")
                    vsfgDetalleImp.TextMatrix(i, 3) = clsCon_det.adorec_Def("prd_nombre")
                    vsfgDetalleImp.TextMatrix(i, 4) = clsCon_det.adorec_Def("det_inv_cantidad")
                    clsCon_det.adorec_Def.MoveNext
                    i = i + 1
                Wend
            End If
        End If
    End If
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub
    
Private Sub dcmbIngImp_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
End If
End Sub

Private Sub Form_Activate()
  clsConsu.Actualizar
  If (clsConsu.adorec_Def.RecordCount <> 0) Then
        Set dcmbIngImp.RowSource = clsConsu.adorec_Def.DataSource
        dcmbIngImp.ListField = "inv_codigo"
    End If

End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 200
    Set ucrtVSFG.VSFGControl = vsfgDetalleImp
    ucrtVSFG.Inicializar False, False, False, True, True, True, False, False, False
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    
    'Ejecuta un SQL contra la base de datosCONCAT(a.ing_codigo) AS ing_codigo
    strSql = " SELECT inv_codigo,inv_fecha,inv_observacion " & _
             " FROM inventario " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY inv_codigo"
    clsConsu.Ejecutar (strSql)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If clsConsu.adorec_Def.EOF Then
        MsgBox "No existen ingresos por Importaciones almacenados en el Sistema", vbInformation, "SisAdmi"
        'Unload Me
    Else
        Set dcmbIngImp.RowSource = clsConsu.adorec_Def.DataSource
        dcmbIngImp.ListField = "inv_codigo"
    End If
    
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Integer
    vsfgDetalleImp.Tag = "N"
    'vsfgDetalleImp.Rows = 2
    
    
'    For X = 1 To vsfgDetalleImp.Rows - 1
'       For Y = 1 To vsfgDetalleImp.Cols - 1
'           vsfgDetalleImp.TextMatrix(X, Y) = ""
'        Next Y
'    Next X
    vsfgDetalleImp.Rows = 1
    vsfgDetalleImp.Clear 1
    vsfgDetalleImp.Tag = "T"
    
End Sub

