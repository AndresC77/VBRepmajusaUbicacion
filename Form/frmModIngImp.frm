VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmModIngImp 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Ingresos de Importaciones"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmModIngImp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7530
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingresos por Importaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   98
      TabIndex        =   12
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4200
         Width           =   975
      End
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
         TabIndex        =   4
         Top             =   435
         Width           =   1410
      End
      Begin VB.TextBox txtCodCli 
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
         Left            =   1125
         TabIndex        =   1
         Top             =   795
         Width           =   2220
      End
      Begin VB.TextBox txtFaxProveedor 
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
         Left            =   4935
         TabIndex        =   7
         Top             =   1470
         Width           =   2130
      End
      Begin VB.TextBox txtTelProveedor 
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
         Left            =   4935
         TabIndex        =   6
         Top             =   1125
         Width           =   2130
      End
      Begin VB.TextBox txtRucProveedor 
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
         Left            =   4935
         TabIndex        =   5
         Top             =   795
         Width           =   2130
      End
      Begin VB.TextBox txtDirProveedor 
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
         Left            =   1125
         TabIndex        =   3
         Top             =   1470
         Width           =   2205
      End
      Begin VB.TextBox txtNomP 
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
         Left            =   1125
         TabIndex        =   2
         Top             =   1125
         Width           =   2220
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
         Height          =   930
         Left            =   615
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4665
         Width           =   6510
      End
      Begin MSDataListLib.DataCombo dcmbIngImp 
         Height          =   330
         Left            =   1125
         TabIndex        =   0
         Top             =   420
         Width           =   1410
         _ExtentX        =   2487
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
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalleImp 
         Height          =   1890
         Left            =   120
         TabIndex        =   8
         Top             =   2265
         Width           =   7050
         _cx             =   12435
         _cy             =   3334
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmModIngImp.frx":030A
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
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
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   5520
         TabIndex        =   24
         Top             =   4220
         Width           =   390
      End
      Begin VB.Label lblDetalle 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE"
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
         Left            =   3435
         TabIndex        =   22
         Top             =   1965
         Width           =   780
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
         TabIndex        =   21
         Top             =   4320
         Width           =   1410
      End
      Begin VB.Label lblFaxProveedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
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
         Left            =   4110
         TabIndex        =   20
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label lblTelProveedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
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
         Left            =   4065
         TabIndex        =   19
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label lblRucProveedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "RUC:"
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
         TabIndex        =   18
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lblDirProveedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
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
         TabIndex        =   17
         Top             =   1515
         Width           =   720
      End
      Begin VB.Label lblNomProveedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
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
         TabIndex        =   16
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label lblCodProveedor 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   540
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
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblNumIngreso 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Ingreso:"
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
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Acpetar"
      Height          =   375
      Left            =   2269
      TabIndex        =   10
      Top             =   5925
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3806
      TabIndex        =   11
      Top             =   5925
      Width           =   1455
   End
End
Attribute VB_Name = "frmModIngImp"
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
Private clsExis As New clsConsulta


Private Sub cmdModificar_Click()
    Dim Resp  As Integer
    If cmdModificar.Caption = "Modificar" Then
        cmdModificar.Caption = "Aceptar"
        vsfgDetalleImp.Editable = flexEDKbdMouse
    ElseIf cmdModificar.Caption = "Aceptar" Then
        Resp = MsgBox("¿Está seguro de aceptar los cambios?", vbQuestion + vbYesNo, "SisAdmi - Importaciones")
        If Resp = vbYes Then
            Modificar
            Unload Me
        Else
            Exit Sub
        End If
    Else
        cmdModificar.Caption = "Modificar"
    End If
End Sub
Private Sub Modificar()
    Dim clsIngreso As New clsInventario
    Dim i As Long
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    
    clsIngreso.strTipo = "IIM"
    clsIngreso.strDoc = dcmbIngImp.Text
    With vsfgDetalleImp
        For i = 1 To .Rows - 1
            clsIngreso.EliminarDetIng .TextMatrix(i, 2), .TextMatrix(i, 1)
        Next i
        
        For i = 1 To .Rows - 1
            clsIngreso.NuevoDetIng .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD4(.TextMatrix(i, 4))
        Next i
        InicializarContenedorRecurrente
    End With
    

    
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbIngImp_Change()
    Dim i As Integer
    
    'Despliego los datos segú el dato ingresado o seleccioado en el data combo
    clsConsu.adorec_Def.MoveFirst
    If (dcmbIngImp.Text = "") Then
        Call limpiarFxGD
        txtCodCli.Text = ""
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtFechaIng.Text = ""
        txtObs.Text = ""
        Exit Sub
    End If
    
    If (CSng(dcmbIngImp.Text) > 99999) Then
        Call limpiarFxGD
        txtCodCli.Text = ""
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtFechaIng.Text = ""
        txtObs.Text = ""
        Exit Sub
    End If
    clsConsu.adorec_Def.Find "ing_codigo = '" & dcmbIngImp.Text & "'", , adSearchForward
   
    If clsConsu.adorec_Def.EOF = False Then
        'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
        txtCodCli.Text = clsConsu.adorec_Def("per_codigo")
        txtNomP.Text = clsConsu.adorec_Def("per_apellido") & " " & clsConsu.adorec_Def("per_nombre")
        txtFechaIng.Text = Format(clsConsu.adorec_Def("ing_fecha"), "yyyy-mmm-dd")
        txtRucProveedor.Text = clsConsu.adorec_Def("per_ruc")
        txtDirProveedor.Text = clsConsu.adorec_Def("per_direccion")
        txtTelProveedor.Text = clsConsu.adorec_Def("per_telf")
        txtFaxProveedor.Text = clsConsu.adorec_Def("per_fax")
        Call limpiarFxGD
        
        'llenar flexgrid
        strSQL = " select det_ingreso.dep_codigo, det_ingreso.prd_codigo, producto.prd_nombre," & _
                 " det_ingreso.det_ing_cantidad,det_ing_precio" & _
                 " from det_ingreso" & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo = producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " WHERE det_ingreso.emp_codigo = '" & strEmpresa & "'" & _
                 " and det_ingreso.tip_ing_codigo = 'IIM' " & _
                 " and det_ingreso.ing_codigo = " & clsConsu.adorec_Def("ing_codigo")
        clsCon_det.Ejecutar (strSQL)
        If (clsCon_det.adorec_Def.RecordCount > 0) Then
            clsCon_det.adorec_Def.MoveFirst
            For i = 1 To clsCon_det.adorec_Def.RecordCount
                vsfgDetalleImp.AddItem ""
                vsfgDetalleImp.TextMatrix(i, 0) = i
                vsfgDetalleImp.TextMatrix(i, 1) = clsCon_det.adorec_Def("dep_codigo")
                vsfgDetalleImp.TextMatrix(i, 2) = clsCon_det.adorec_Def("prd_codigo")
                vsfgDetalleImp.TextMatrix(i, 3) = clsCon_det.adorec_Def("prd_nombre")
                vsfgDetalleImp.TextMatrix(i, 4) = clsCon_det.adorec_Def("det_ing_cantidad")
                vsfgDetalleImp.TextMatrix(i, 5) = clsCon_det.adorec_Def("det_ing_precio")
                clsCon_det.adorec_Def.MoveNext
            Next i
        End If
        txtObs.Text = clsConsu.adorec_Def("ing_observacion")
    Else
        Call limpiarFxGD
        txtCodCli.Text = ""
        txtNomP.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtFechaIng.Text = ""
        txtObs.Text = ""
        
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

Private Sub Form_Load()
    Dim strSQL As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0 '((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 200
    
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    clsExis.Inicializar AdoConn, AdoConnMaster
    'Ejecuta un SQL contra la base de datosCONCAT(a.ing_codigo) AS ing_codigo
    strSQL = " CREATE TABLE #ing_imp ( " & _
             " emp_codigo char(3) NOT NULL default '', " & _
             " ing_codigo decimal(12,0) NOT NULL default '0', " & _
             " tip_ing_codigo char(3) NOT NULL default '', " & _
             " PRIMARY KEY  (emp_codigo,ing_codigo,tip_ing_codigo))"
    clsConsu.Ejecutar strSQL
    strSQL = " INSERT INTO ing_imp " & _
             " SELECT DISTINCT emp_codigo,ing_codigo,tip_ing_codigo " & _
             " FROM det_ingreso_imp " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo  = 'IIM'"
    clsConsu.Ejecutar strSQL
    strSQL = " SELECT CONCAT(a.ing_codigo) AS ing_codigo, a.per_codigo, a.ing_fecha," & _
             " a.ing_observacion, b.per_codigo, b.per_nombre," & _
             " b.per_apellido, b.per_ruc, b.per_direccion," & _
             " b.per_telf, b.per_fax FROM ingreso a INNER JOIN persona b ON a.per_codigo = b.per_codigo AND  a.emp_codigo = b.emp_codigo" & _
             " LEFT JOIN ing_imp c ON a.emp_codigo=c.emp_codigo AND a.ing_codigo=c.ing_codigo AND a.tip_ing_codigo=c.tip_ing_codigo " & _
             " WHERE a.tip_ing_codigo  = 'IIM' AND a.emp_codigo = '" & strEmpresa & "'" & _
             " AND c.ing_codigo IS NULL " & _
             " ORDER BY a.ing_codigo"
    clsConsu.Ejecutar (strSQL)
    strSQL = " DROP TABLE #ing_imp "
    clsCon_det.Ejecutar strSQL
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If clsConsu.adorec_Def.EOF Then
        MsgBox "No existen ingresos por Importaciones almacenados en el Sistema", vbInformation, "SisAdmi"
    Else
        Set dcmbIngImp.RowSource = clsConsu.adorec_Def.DataSource
        dcmbIngImp.ListField = "ing_codigo"
    End If
    
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Integer
    vsfgDetalleImp.Tag = "N"
    vsfgDetalleImp.Rows = 1
    vsfgDetalleImp.Clear 1
    vsfgDetalleImp.Tag = "T"
    
End Sub

Private Sub vsfgDetalleImp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If Not IsNumeric(vsfgDetalleImp.TextMatrix(Row, 4)) Then vsfgDetalleImp.TextMatrix(Row, 4) = "0"
        TxtTotal = ""
        Tot = 0
        For i = 1 To vsfgDetalleImp.Rows - 1
            If vsfgDetalleImp.TextMatrix(i, 4) <> "" Then
                Tot = Tot + FormatoD4(vsfgDetalleImp.TextMatrix(i, 4))
            End If
        Next i
        TxtTotal.Text = Tot
    End If
End Sub

Private Sub vsfgDetalleImp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 4 Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDetalleImp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If Not IsNumeric(vsfgDetalleImp.TextMatrix(Row, 4)) Then vsfgDetalleImp.TextMatrix(Row, 4) = "0"
        TxtTotal = ""
        Tot = 0
        For i = 1 To vsfgDetalleImp.Rows - 1
            If vsfgDetalleImp.TextMatrix(i, 4) <> "" Then
                Tot = Tot + FormatoD4(vsfgDetalleImp.TextMatrix(i, 4))
            End If
        Next i
        TxtTotal.Text = Tot
    End If
End Sub
