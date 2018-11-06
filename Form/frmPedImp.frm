VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPedImp 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frm"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmPedImp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10785
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   98
      TabIndex        =   5
      Top             =   120
      Width           =   10575
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   7200
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         _extentx        =   3201
         _extenty        =   556
      End
      Begin VB.TextBox TxtTotalT 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   18
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtCodPedido 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3900
         TabIndex        =   14
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir"
         Height          =   375
         Left            =   9360
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   10335
         Begin VSFlex8Ctl.VSFlexGrid VSFG 
            Height          =   4935
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   3975
            _cx             =   53287779
            _cy             =   53289473
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vsfgDetalleImp 
            Height          =   4455
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   10095
            _cx             =   94324110
            _cy             =   94314162
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPedImp.frx":030A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
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
            Left            =   7080
            TabIndex        =   1
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Image imgBtnDn 
            Height          =   210
            Left            =   6120
            Picture         =   "frmPedImp.frx":03DC
            Top             =   4680
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Image imgBtnUp 
            Height          =   210
            Left            =   5880
            Picture         =   "frmPedImp.frx":0508
            Top             =   4680
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label lblTotal 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   6360
            TabIndex        =   10
            Top             =   4740
            Width           =   735
         End
      End
      Begin VB.TextBox txtObs 
         Height          =   570
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   6600
         Width           =   10335
      End
      Begin VB.TextBox txtNumIngreso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   1785
      End
      Begin MSComDlg.CommonDialog cmdArchivo 
         Left            =   7440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dcmbCodP 
         Height          =   330
         Left            =   1080
         TabIndex        =   16
         Top             =   720
         Width           =   4725
         _ExtentX        =   8334
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
      Begin VB.Label Label2 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   788
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   300
         Width           =   735
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
         Left            =   120
         TabIndex        =   8
         Top             =   6345
         Width           =   1410
      End
      Begin VB.Label lblFecha 
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
         Left            =   6525
         TabIndex        =   7
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lblNumIngreso 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3686
      TabIndex        =   3
      Top             =   7590
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5644
      TabIndex        =   4
      Top             =   7590
      Width           =   1455
   End
End
Attribute VB_Name = "frmPedImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso de mercadería a los depòsitos por concepto de         #
'#  importaciones se permite crear estos ingresos                               #
'#  frmIngImportacion  V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite ingresar los productos a los diferentes depòsitos       #
'#  de la compañía por concepto de importaciones , solo se permite el ingreso   #
'#  de tales datos para posteriormente actualizar las existencias.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se almacenan los nuevos ingresos de mercadería #
'#    det_ingreso: En estatabla se almacena los detalles de cada ingreso        #
'#    persona    : Se consulta los proveedores de la empresa                    #
'#    deposito   : Se consulta los depositos o bodegas de la empresa            #
'#    producto   : Se consulta los productos de la empresa                      #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#               limpiarFxGD()   Permite borrar los datos que se encuentran     #
'#                               en el flexGrid para realizar un nuevo ingreso  #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsCon_Def As New clsConsulta
Private clsCon_Pro As New clsConsulta
Private strSql As String
Public NInv As Long
Public MN As String

Private Sub cmdAbrir_Click()
    Dim strPath As String
    Dim Archivo As String
    Dim j As Long
    strPath = Trim(App.Path)
    cmdArchivo.DialogTitle = "Abrir"
    cmdArchivo.InitDir = strPath
    cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Todos los Archivos|*.*"
    cmdArchivo.ShowOpen
    Archivo = cmdArchivo.FileName
    'VSFG.Visible = True
    If Archivo <> "" Then
        VSFG.LoadGrid Archivo, flexFileExcel
        j = 1
        For i = 0 To VSFG.Rows - 1
            strSql = " SELECT count(*) as N FROM persona_producto " & _
                     " WHERE persona_producto.emp_codigo='" & strEmpresa & "' " & _
                     " AND per_codigo='" & dcmbCodP.BoundText & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def("N") > 0 Then
                vsfgDetalleImp.ShowCell j - 1, 0
                vsfgDetalleImp.TextMatrix(j, 1) = VSFG.TextMatrix(i, 0)
                'vsfgDetalleImp.TextMatrix(j, 2) = VSFG.TextMatrix(i, 0)
                vsfgDetalleImp.TextMatrix(j, 3) = VSFG.TextMatrix(i, 1)
                j = j + 1
            Else
                MsgBox "El producto " & VSFG.TextMatrix(i, 0) & vbNewLine & _
                       "NO EXISTE y fue contado" & vbNewLine & _
                       VSFG.TextMatrix(i, 1) & " unidades", vbInformation, "Conteos"
            End If
        Next i
        SumaCantidades
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Long
    strSql = " INSERT INTO pedido_importacion " & _
             "(ped_imp_codigo, emp_codigo, ver_codigo, emb_codigo, age_afi_codigo, per_codigo" & _
             ", ped_imp_numero,ped_imp_fecha_ped, ped_imp_fecha_imp, ped_imp_solicitud_insp" & _
             ", ped_imp_fecha_sol_i, ped_imp_estado, ped_imp_total_pedido" & _
             ", ped_imp_total_gastos, ped_imp_observacion, ped_imp_fechamod, ped_imp_usumod) " & _
             "VALUES (" & txtNumIngreso.Text & ", '" & strEmpresa & "', '', '', '', '" & dcmbCodP.BoundText & "','" & UCase(txtCodPedido.Text) & "'" & _
             ", '" & dtpFecha.Value & "', '', ''" & _
             ", '', 'PP', " & TxtTotalT.Text & ", null, '" & UCase(txtObs.Text) & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
    clsCon_Def.Ejecutar strSql
    For i = 1 To vsfgDetalleImp.Rows - 1
        If Trim(vsfgDetalleImp.TextMatrix(i, 1)) <> "" Then
            strSql = " SELECT count(*) as n " & _
                     " FROM det_pedido_imp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND ped_imp_codigo = '" & txtNumIngreso.Text & "'" & _
                     " AND prd_codigo='" & vsfgDetalleImp.TextMatrix(i, 1) & "'"
            clsCon_Def.Ejecutar strSql
            If FormatoD0(clsCon_Def.adorec_Def("n")) = 0 Then
                strSql = " INSERT INTO det_pedido_imp " & _
                         "(ped_imp_codigo, emp_codigo, prd_codigo," & _
                         " det_ped_imp_cantidad, det_ped_imp_precio," & _
                         " det_ped_fechamod, det_ped_usumod)" & _
                         " VALUES (" & txtNumIngreso.Text & ", '" & strEmpresa & "', '" & vsfgDetalleImp.TextMatrix(i, 1) & "'" & _
                         ", '" & vsfgDetalleImp.TextMatrix(i, 3) & "', " & vsfgDetalleImp.TextMatrix(i, 4) & "," & _
                         " CURRENT_TIMESTAMP, '" & strUsuario & "')"
            Else
                strSql = " UPDATE det_pedido_imp " & _
                         " SET det_ped_imp_cantidad=det_ped_imp_cantidad+" & FormatoD0(vsfgDetalleImp.TextMatrix(i, 3)) & " " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND ped_imp_codigo = '" & txtNumIngreso.Text & "'" & _
                         " AND prd_codigo='" & vsfgDetalleImp.TextMatrix(i, 1) & "'"
            End If
            clsCon_Def.Ejecutar strSql
        End If
    Next i
    MsgBox "Pedido Cargado"
    Unload Me
End Sub

Private Sub dcmbCodP_Validate(Cancel As Boolean)

    ' Agrega un combo productos
    strSql = " SELECT persona_producto.prd_codigo,prd_nombre " & _
             " FROM persona_producto INNER JOIN producto ON persona_producto.emp_codigo=producto.emp_codigo " & _
             " AND persona_producto.prd_codigo=producto.prd_codigo " & _
             " WHERE persona_producto.emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo='" & dcmbCodP.BoundText & "' " & _
             " ORDER BY prd_nombre"
    clsCon_Pro.Ejecutar strSql
    vsfgDetalleImp.ColComboList(2) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "prd_codigo,*prd_nombre", "prd_codigo")
    ' Agrega un combo productos
    strSql = " SELECT persona_producto.prd_codigo,prd_nombre,per_pro_precio " & _
             " FROM persona_producto INNER JOIN producto ON persona_producto.emp_codigo=producto.emp_codigo " & _
             " AND persona_producto.prd_codigo=producto.prd_codigo " & _
             " WHERE persona_producto.emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo='" & dcmbCodP.BoundText & "' " & _
             " ORDER BY prd_codigo"
    clsCon_Pro.Ejecutar strSql
    vsfgDetalleImp.ColComboList(1) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*prd_codigo,prd_nombre", "prd_codigo")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_Def = Nothing
    Set clsCon_Pro = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()

Dim d As String
Dim m As Integer
Dim Y As String
Dim ff As Variant
Dim var As Long

    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 350
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_Pro.Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
'Ejecuta un SQL contra la base de datos
    strSql = " select per_codigo, CONCAT(per_apellido, ' ',per_nombre) as per_nombre, " & _
             " per_apellido, per_ruc, per_direccion, " & _
             " per_telf, per_fax from persona " & _
             " where emp_codigo= '" & strEmpresa & "' and cat_p_tipo = 'P' " & _
             " order by per_apellido,per_nombre"
    clsCon_Def.Ejecutar (strSql)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If (clsCon_Def.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Proveedores ingresados en el Sistema", vbInformation, "SisAdmi"
        Exit Sub
    Else
        Set dcmbCodP.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbCodP.ListField = "per_nombre"
        dcmbCodP.BoundColumn = "per_codigo"
    End If
    
    'Descompone la fecha actual  en día, mes y año
    
    dtpFecha.Value = HoyDia
    
    'Consulta del nùmero de ingreso último, se agrega uno para el nuevo ingreso
    strSql = "select COALESCE(max(ped_imp_codigo),0) as num FROM pedido_importacion WHERE emp_codigo = '" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsConsu.Ejecutar (strSql), "M"
    If clsConsu.adorec_Def.EOF Then
        txtNumIngreso.Text = "1"
    Else
        txtNumIngreso.Text = clsConsu.adorec_Def("num") + 1
        
    End If
    txtNumIngreso.Enabled = False
    PonerBotones
    
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
       
        End Select
End Sub

Private Sub vsfgDetalleImp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    Dim NuevaFila As Boolean
    NuevaFila = True
    For i = 1 To vsfgDetalleImp.Cols - 1
        If vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, i) = "" Then
            NuevaFila = False
            Exit For
        End If
    Next i
    If NuevaFila = True Then
        vsfgDetalleImp.AddItem ""
        PonerBotones
    End If
End Sub

Private Sub vsfgDetalleImp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col >= vsfgDetalleImp.Cols - 2 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDetalleImp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleImp.MouseRow
    c = vsfgDetalleImp.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = vsfgDetalleImp.Rows) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleImp.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
   
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleImp.Cell(flexcpLeft, r, c) + vsfgDetalleImp.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleImp.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "AHORA DEBE ELIMINAR ESTA FILA!"
    
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        
        'Recorro el FlexGrid para almacenar los detalles del ingreso
        
        If respuesta = vbYes Then
            Dim i As Long
        
            TxtTotal.Text = Format(CStr(Val(TxtTotal.Text) - Val(vsfgDetalleImp.TextMatrix(r, 4))), "####0.00")
            vsfgDetalleImp.RemoveItem (r)
            For i = 1 To (vsfgDetalleImp.Rows - 1)
                vsfgDetalleImp.TextMatrix(i, 0) = i
                vsfgDetalleImp.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalleImp.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            Next i
        Else
            vsfgDetalleImp.Cell(flexcpPicture, r, c) = imgBtnUp
        
        End If
    
        
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Se da un tab al presionar enter para que al ingresar un dato pase al siguiente campo
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If

End Sub

Private Sub SumaCantidades()
    Dim Suma As Long
    Dim SumaT As Double
    Dim i As Long
    Suma = 0
    SumaT = 0
    For i = 1 To Me.vsfgDetalleImp.Rows - 1
        Suma = Suma + Val(Format(vsfgDetalleImp.TextMatrix(i, 3), "#0"))
        SumaT = SumaT + FormatoD2(vsfgDetalleImp.TextMatrix(i, 5))
    Next i
    TxtTotal.Text = Suma
    TxtTotalT.Text = SumaT
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (vsfgDetalleImp.Rows - 1)
        vsfgDetalleImp.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            vsfgDetalleImp.Cell(flexcpPicture, i, 0) = imgBtnUp
            vsfgDetalleImp.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub

Private Sub VerPrecio(fila As Long)
    clsCon_Pro.Filtrar "prd_codigo='" & vsfgDetalleImp.TextMatrix(fila, 1) & "'"
    vsfgDetalleImp.TextMatrix(fila, 4) = FormatoD4(clsCon_Pro.adorec_Def("per_pro_precio"))
End Sub

Private Sub vsfgDetalleImp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        vsfgDetalleImp.TextMatrix(Row, 2) = vsfgDetalleImp.TextMatrix(Row, 1)
        VerPrecio Row
    ElseIf Col = 2 Then
        vsfgDetalleImp.TextMatrix(Row, 1) = vsfgDetalleImp.TextMatrix(Row, 2)
        VerPrecio Row
    ElseIf Col = 3 Or Col = 4 Then
        vsfgDetalleImp.TextMatrix(Row, 5) = FormatoD4(FormatoD0(vsfgDetalleImp.TextMatrix(Row, 3)) * FormatoD4(vsfgDetalleImp.TextMatrix(Row, 4)))
        SumaCantidades
    End If
    vsfgDetalleImp_AfterRowColChange Row, Col, Row, Col
End Sub
