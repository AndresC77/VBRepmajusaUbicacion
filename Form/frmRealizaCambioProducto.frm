VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRealizaCambioProducto 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar Cambio de Productos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmRealizaCambioProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdDevolver 
      Caption         =   "&Devolver Prenda"
      Height          =   375
      Left            =   10920
      TabIndex        =   20
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   12375
      Begin VB.TextBox txtCantIng 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2640
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   12075
         _cx             =   1974031155
         _cy             =   1974013904
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRealizaCambioProducto.frx":030A
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
      Begin VB.Label lblCantINg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   7320
         TabIndex        =   16
         Top             =   2670
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo de Negocio:"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtRuc 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   375
         Width           =   5055
         _ExtentX        =   8916
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
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbFormulario 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CI/RUC:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1485
         Width           =   795
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   420
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4950
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos del Cambio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      Begin VB.TextBox TxtObserv 
         Height          =   555
         Left            =   1080
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   4695
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   270
         Width           =   1695
         _extentx        =   2990
         _extenty        =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observ:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   675
         Width           =   585
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
         Left            =   240
         TabIndex        =   10
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   600
      Picture         =   "frmRealizaCambioProducto.frx":046C
      Top             =   5160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   360
      Picture         =   "frmRealizaCambioProducto.frx":0598
      Top             =   5160
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmRealizaCambioProducto"
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

Public Neg As String
Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private clsCon_Prd2 As New clsConsulta
Private strSQL As String

Private Sub cmbCliente_Validate(Cancel As Boolean)
    cmbFormulario = ""
    
    strSQL = " SELECT DISTINCT cambio.cam_codigo " & _
             " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
             " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
             " AND det_cambio.prd_codigo_ped!='' AND det_cambio.prd_codigo_ped IS NOT NULL " & _
             " AND det_cambio.tip_ing_codigo='' AND det_cambio.ing_codigo='0' " & _
             " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
             " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
             " INNER JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
             " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
             " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
             " AND cambio.per_codigo='" & cmbCliente.BoundText & "' " & _
             " ORDER BY cambio.cam_codigo "
    clsCon_Def.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbFormulario.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbFormulario.ListField = "cam_codigo"
        cmbFormulario.BoundColumn = "cam_codigo"
    Else
        cmbFormulario = "No hay formularios del cliente "
    End If
End Sub


Private Sub cmbFactura_Validate(Cancel As Boolean)
    CargaProductos
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        strSQL = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSQL
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If strPtoFactura <> clsCon_Def.adorec_Def(0) Then
                LimpiarTodo
            End If
            strPtoFactura = clsCon_Def.adorec_Def(0)
        End If
    Else
        Exit Sub
    End If
    strSQL = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSQL)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
End Sub

Private Sub cmdAceptar_Click()
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim clsCambio As New clsCambio
    Dim booGuardar As Boolean
    Dim i As Long, j As Long, cIng As Long, no As Boolean, cEgr As Long
    no = False
        
    
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsCambio.Inicializar AdoConn, AdoConnMaster
    Dim DocCambio As String
    
    booGuardar = clsIngreso.NuevoIng(True, "ICA", False, strSucursal, strPtoFactura, , , , dtpFecha.Value, , , UCase(TxtObserv.Text))
    If booGuardar = True Then
        clsEgreso.NuevoEgr True, "ECA", False, strSucursal, strPtoFactura, Right(clsIngreso.strDoc, 7), , , dtpFecha.Value, , , UCase(TxtObserv.Text)
        With VSFG
            For i = 1 To .Rows - 1
                If Abs(.TextMatrix(i, 0)) = 1 Then
                    clsIngreso.NuevoDetIng .TextMatrix(i, 3), "PRI", .TextMatrix(i, 5), FormatoD8(.TextMatrix(i, 4)), FormatoD4(.TextMatrix(i, 4)), 0, 1
                    clsEgreso.NuevoDetEgr .TextMatrix(i, 7), "PRI", .TextMatrix(i, 5), FormatoD4(.TextMatrix(i, 9)), FormatoD4(.TextMatrix(i, 9)), 0, 1
                    clsCambio.AsignarIngreso .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 7), clsIngreso.strTipo, clsIngreso.strDoc
                End If
            Next i
            InicializarContenedorRecurrente
        End With
    End If
    MsgBox " Los datos han sido ingresados" & vbNewLine & clsCambio.strDoc, vbInformation, "Cambio de Productos"
    
    Set clsCambio = Nothing
        
    Unload Me
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    Dim i As Long
    Dim j As Long
    strSQL = " SELECT '0' as sel, cambio.cam_codigo,mot_aju_codigo," & _
             " prd_codigo_ing,ping.prd_nombre,det_cam_cantidad,ping.prd_costo," & _
             " prd_codigo_ped,pped.prd_nombre,pped.prd_costo" & _
             " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo" & _
             " AND cambio.cam_codigo=det_cambio.cam_codigo" & _
             " AND det_cambio.prd_codigo_ped!='' AND det_cambio.prd_codigo_ped IS NOT NULL " & _
             " AND det_cambio.tip_ing_codigo='' AND det_cambio.ing_codigo='0' " & _
             " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
             " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
             " INNER JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
             " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
             " WHERE cambio.emp_codigo='" & strEmpresa & "' " & _
             " AND cambio.cam_codigo='" & cmbFormulario.BoundText & "' " & _
             " AND cambio.per_codigo='" & cmbCliente.BoundText & "' " & _
             " ORDER BY cambio.cam_codigo,mot_aju_codigo,ping.prd_nombre"
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    i = 1
    j = 0
    While Not clsCon_Def.adorec_Def.EOF
        For j = 0 To VSFG.Cols - 1
            VSFG.TextMatrix(i, j) = clsCon_Def.adorec_Def(j)
        Next j
        i = i + 1
        clsCon_Def.adorec_Def.MoveNext
    Wend
End Sub

Private Sub cmdDevolver_Click()
    Dim clsCambio As New clsCambio
    Dim i As Long, j As Long, cIng As Long, no As Boolean, cEgr As Long
    no = False
        
    clsCambio.Inicializar AdoConn, AdoConnMaster
    Dim DocCambio As String
    
    With VSFG
        For i = 1 To .Rows - 1
            If Abs(.TextMatrix(i, 0)) = 1 Then
                clsCambio.AsignarIngreso .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 7), "NOA", "0000000"
            End If
        Next i
    End With
    MsgBox " Los datos han sido ingresados" & vbNewLine & clsCambio.strDoc, vbInformation, "Cambio de Productos"
    
    Set clsCambio = Nothing
    
'    Dim rpTra As New frmReporte
'    rpTra.strNumero = DocIng
'    rpTra.strTipo = DocEgr
'    rpTra.strReporte = "rptCambioProducto"
'    rpTra.Show
        
    Unload Me

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSQL = ""
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_Prd2 = Nothing
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd2.Inicializar AdoConn, AdoConnMaster
    dtpFecha.Value = HoyDia
    dtpFecha.Enabled = False
    
    cargarTipoPedido
    
    'PonerBotones
End Sub

Private Sub VsfgDetalleIng_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleIng.MouseRow
    c = vsfgDetalleIng.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = (vsfgDetalleIng.Rows - 1)) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleIng.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleIng.Cell(flexcpLeft, r, c) + vsfgDetalleIng.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         vsfgDetalleIng.RemoveItem (r)
         'PonerBotones
         CalculaTotal
    Else
        vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub cargarTipoPedido()
    strSQL = " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo like '" & Neg & "'" & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSQL
    Set cmbNegocio.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSQL = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo like '" & Neg & "'" & _
             " AND tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
End Sub


Private Sub CargaProductos()

    'Carga los motivos de ajuste
    strSQL = " SELECT mot_aju_codigo,mot_aju_nombre " & _
             " FROM motivo_ajuste " & _
             " Where emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY mot_aju_nombre "
    clsCon_Def.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(1) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "*mot_aju_codigo, mot_aju_nombre", "mot_aju_codigo")
    
    'Consulto los productos de la factura
    
    strSQL = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
             " AND det_egreso.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_egreso.tip_egr_codigo = 'FAC' " & _
             " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
             " AND prd_baja=0 ORDER BY prd_codigo "
    clsCon_Prd.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(2) = vsfgDetalleIng.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    'Consulto los productos de la factura
    strSQL = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
             " AND det_egreso.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_egreso.tip_egr_codigo = 'FAC' " & _
             " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
             " AND prd_baja=0 ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(3) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    
    'Consulto los productos para el cambio
    
    strSQL = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM producto INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
             " AND lis_pre_p_precio!=0 " & _
             " AND prd_baja=0 ORDER BY prd_codigo "
    clsCon_Prd2.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(6) = vsfgDetalleIng.BuildComboList(clsCon_Prd2.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    
    'Consulto los productos de la factura
    strSQL = " SELECT producto.prd_codigo, prd_nombre,lis_pre_p_precio as prd_precio " & _
             " FROM producto INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
             " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
             " AND lis_pre_p_precio!=0 " & _
             " AND prd_baja=0 ORDER BY prd_nombre "
    clsCon_Def.Ejecutar strSQL
    vsfgDetalleIng.ColComboList(7) = vsfgDetalleIng.BuildComboList(clsCon_Def.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")

End Sub

Private Sub vsfgDetalleIng_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If vsfgDetalleIng.TextMatrix(Row, Col) <> "" And Not IsNumeric(vsfgDetalleIng.TextMatrix(Row, Col)) Then
            MsgBox "Ingrese valores numéricos en Cantidad", vbInformation, "Detalle"
            vsfgDetalleIng.TextMatrix(Row, Col) = 0
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Or Col = 3 Then
        If vsfgDetalleIng.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero un motivo", vbInformation, "Motivos"
            vsfgDetalleIng.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
    End If
    If Col = 2 Then
        vsfgDetalleIng.TextMatrix(Row, 3) = vsfgDetalleIng.TextMatrix(Row, 2)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalleIng.TextMatrix(Row, 6) = 0
    ElseIf Col = 3 Then
        vsfgDetalleIng.TextMatrix(Row, 2) = vsfgDetalleIng.TextMatrix(Row, 3)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_costo")
        vsfgDetalleIng.TextMatrix(Row, 6) = 0
    ElseIf Col = 4 Then
        vsfgDetalleIng.TextMatrix(Row, 6) = FormatoD4(FormatoD4(vsfgDetalleIng.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalleIng.TextMatrix(Row, 5)))
    End If
    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 2) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 3) <> "" And Val(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 4)) <> 0 Then
        vsfgDetalleIng.AddItem ""
        vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 0) = vsfgDetalleIng.Rows - 1
        vsfgDetalleIng.Cell(flexcpPicture, vsfgDetalleIng.Rows - 1, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, vsfgDetalleIng.Rows - 1, 0) = flexAlignRightCenter
        If vsfgDetalleIng.Rows > 2 Then
             vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) = vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 2, 1)
        End If
    End If
    CalculaTotal
End Sub

Private Sub CalculaTotal()
    Dim i As Long
    Dim CantIng As Double
    CantIng = 0
    
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            CantIng = CantIng + FormatoD4(VSFG.TextMatrix(i, 5))
        End If
    Next i
    
    txtCantIng.Text = FormatoD4(CantIng)
    
End Sub

Private Sub txtRuc_Validate(Cancel As Boolean)
    
    If Trim(txtRuc.Text) <> "" Then
        strSQL = " SELECT per_codigo " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' " & _
                 " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND per_ruc='" & txtRuc.Text & "'"
        clsCon_Def.Ejecutar strSQL
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If cmbCliente.BoundText <> clsCon_Def.adorec_Def(0) Then
                cmbCliente.BoundText = clsCon_Def.adorec_Def(0)
                cmbCliente_Validate False
            End If
        Else
            MsgBox "No se encontró un cliente con CI/RUC " & txtRuc.Text, vbInformation, "CI/RUC"
            If cmbCliente.Text <> "" Then
                LimpiarTodo
            Else
                txtRuc.Text = ""
            End If
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 And Row > 0 Then
        If Abs(VSFG.TextMatrix(Row, 0)) = 0 Then
            VSFG.Select Row, 1, Row, VSFG.Cols - 1
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
        Else
            VSFG.Select Row, 1, Row, VSFG.Cols - 1
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HC0FFFF
        End If
        CalculaTotal
    End If
    
End Sub
