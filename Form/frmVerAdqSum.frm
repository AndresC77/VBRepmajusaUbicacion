VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerAdqSum 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Adquisicion Suministro"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmVerAdqSum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7650
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3870
      TabIndex        =   15
      Top             =   6900
      Width           =   1700
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   2081
      TabIndex        =   14
      Top             =   6900
      Width           =   1700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Height          =   6855
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtNomProveedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   690
            Width           =   2220
         End
         Begin VB.TextBox txtDirProveedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1035
            Width           =   2205
         End
         Begin VB.TextBox txtRucProveedor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   2130
         End
         Begin VB.TextBox txtTelProveedor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   690
            Width           =   2130
         End
         Begin VB.TextBox txtFaxProveedor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1035
            Width           =   2130
         End
         Begin VB.TextBox txtCodProveedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   2220
         End
         Begin VB.Label Label2 
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
            Left            =   240
            TabIndex        =   28
            Top             =   412
            Width           =   540
         End
         Begin VB.Label Label3 
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
            Left            =   240
            TabIndex        =   27
            Top             =   742
            Width           =   600
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   26
            Top             =   1087
            Width           =   720
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   3600
            TabIndex        =   25
            Top             =   390
            Width           =   375
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label7 
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
            Height          =   255
            Left            =   3600
            TabIndex        =   23
            Top             =   1065
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Adquisión Suministro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtAdq_fecha 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtNum_orden 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dcmbAdq_codigo 
            Height          =   330
            Left            =   960
            TabIndex        =   0
            Top             =   352
            Width           =   1290
            _ExtentX        =   2275
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
         Begin VB.Label Label13 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "# Orden:"
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
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   390
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
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label8 
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
            Height          =   255
            Left            =   5040
            TabIndex        =   20
            Top             =   390
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Detalle Adquisión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   7335
         Begin VB.TextBox txtObs 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   3360
            Width           =   7050
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2400
            Width           =   1095
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFDetalle 
            Height          =   1695
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   7095
            _cx             =   12515
            _cy             =   2990
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmVerAdqSum.frx":030A
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
            Editable        =   1
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   3120
            Width           =   1110
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5160
            TabIndex        =   30
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblIva 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5160
            TabIndex        =   29
            Top             =   2430
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   5160
            TabIndex        =   18
            Top             =   2790
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmVerAdqSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para visualizar los Suministros realizados por concepto de            #
'#  Adquisiciones,  esta forma es solo de visualización, no permite la edición. #
'#  frmVerAdqSum V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los ingresos de Suministros a una determinada  emp-  #
'#  presa por concepto de Compra.                                               #
'#  En esta ventana solo se puede visuallizar cualesquiera de los ingresos por  #
'#  este concepto pero no se puede realizar ningún cambio.                      #
'#  Se puede escoger el número de documento o ingresar dicho número en el combo #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    suministro : En esta tabla se consulta el codigo,nombre y descripcion     #
'#                       del suministro.                                        #
'#    det_adquisicion_su : En esta tabla se consulta el detalle de adquisicion  #
'#                        de Suministros.                                        #
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
Private clsCon_Iva As New clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_det = Nothing
    Set clsCon_Iva = Nothing
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFDetalle.Rows - 1)
        VSFDetalle.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub CalcuTotal()
   'Calcula totales
    Dim SubTotal As Double
    'Calcula Subtotal
    SubTotal = 0
    For i = 1 To VSFDetalle.Rows - 1
        VSFDetalle.TextMatrix(i, 5) = Val(VSFDetalle.TextMatrix(i, 3)) * Val(VSFDetalle.TextMatrix(i, 4))
        SubTotal = SubTotal + Val(VSFDetalle.TextMatrix(i, 5))
    Next i
    TxtSubTotal = FormatoD2(SubTotal)
    TxtIva = FormatoD2(Val(TxtSubTotal) * (TxtIva.Tag / 100))
    TxtTotal = FormatoD2(Val(TxtSubTotal.Text) + Val(TxtIva.Text))
End Sub
Private Sub cmdNuevo_Click()
    Me.Hide
    frmAdqSum.Show
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub dcmbAdq_codigo_Change()
    Dim i As Integer
    
    'Despliego los datos según el dato ingresado o seleccioado en el data combo
    If (dcmbAdq_codigo.Text = "") Then
        borrar_datos
        Call limpiarFxGD
        Exit Sub
    End If
    
    If (CSng(dcmbAdq_codigo.Text) > 99999) Then
        borrar_datos
        Call limpiarFxGD
        Exit Sub
    End If
    
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "adq_codigo = '" & dcmbAdq_codigo.Text & "'", , adSearchForward
    If clsConsu.adorec_Def.EOF = False Then
        'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
        txtCodProveedor.Text = clsConsu.adorec_Def("per_codigo")
        txtNomProveedor.Text = clsConsu.adorec_Def("per_nombre") & " " & clsConsu.adorec_Def("per_apellido")
        txtAdq_fecha.Text = Format(clsConsu.adorec_Def("adq_fecha"), "yyyy-mmm-dd")
        txtRucProveedor.Text = clsConsu.adorec_Def("per_ruc")
        txtDirProveedor.Text = clsConsu.adorec_Def("per_direccion")
        txtTelProveedor.Text = clsConsu.adorec_Def("per_telf")
        txtFaxProveedor.Text = clsConsu.adorec_Def("per_fax")
        txtNum_orden.Text = clsConsu.adorec_Def("adq_numdoc")
        TxtIva.Text = Format(clsConsu.adorec_Def("adq_impuesto"), "###0.00")
        
        Call limpiarFxGD
        'llenar flexgrid
        strSql = " SELECT det_adquisicion_su.sum_codigo, sum_nombre,det_adq_su_cantidad, sum_ultimo_precio,det_adq_su_precio  " & _
                 " FROM det_adquisicion_su INNER JOIN suministro ON  det_adquisicion_su.emp_codigo = suministro.emp_codigo" & _
                 "                                              AND det_adquisicion_su.sum_codigo = suministro.sum_codigo " & _
                 " WHERE det_adquisicion_su.emp_codigo = '" & strEmpresa & "' and det_adquisicion_su.adq_codigo =  " & dcmbAdq_codigo.Text & " "
        clsCon_det.Ejecutar (strSql)
        
        If (clsCon_det.adorec_Def.RecordCount > 0) Then
            TxtSubTotal.Text = 0
            TxtTotal.Text = 0
            Set VSFDetalle.DataSource = clsCon_det.adorec_Def.DataSource
            PonerBotones
            If clsCon_Iva.adorec_Def.EOF Then
                     LblIva.Caption = "IVA 0 %"
                     TxtIva.Tag = " 0 "
                Else
                    LblIva.Caption = "IVA " & Format(clsCon_Iva.adorec_Def.Fields("par_numero").Value, "###0.00") & "%"
                    TxtIva.Tag = clsCon_Iva.adorec_Def.Fields("par_numero")
                End If
            CalcuTotal
            txtObs.Text = clsConsu.adorec_Def("adq_observacion")
        Else
            TxtSubTotal.Text = " "
            LblIva.Caption = " IVA 0 %"
            TxtIva.Text = " "
            TxtTotal.Text = " "
            txtObs.Text = " "
            End If
    Else
        Call limpiarFxGD
        txtCodProveedor.Text = ""
        txtNomProveedor.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtAdq_fecha.Text = ""
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
    
Private Sub dcmbAdq_codigo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
End If
End Sub

Private Sub Form_Activate()
clsConsu.Actualizar
  If (clsConsu.adorec_Def.RecordCount <> 0) Then
        Set dcmbAdq_codigo.RowSource = clsConsu.adorec_Def.DataSource
        dcmbAdq_codigo.ListField = "adq_codigo"
    End If

End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    clsCon_Iva.Inicializar AdoConn, AdoConnMaster
    'Busco el valor del IVA para la compra
    strSql = " SELECT par_numero,par_texto " & _
             " FROM  parametro " & _
             " WHERE par_codigo ='IVAC' AND emp_codigo='" & strEmpresa & "' "
    clsCon_Iva.Ejecutar (strSql)
    
    'Ejecuta un SQL contra la base de datos
    strSql = " SELECT concat(a.adq_codigo)as adq_codigo, a.per_codigo, a.adq_fecha,a.adq_subtotal," & _
             " a.adq_observacion,a.adq_numdoc,a.adq_impuesto, b.per_codigo, b.per_nombre," & _
             " b.per_apellido, b.per_ruc, b.per_direccion," & _
             " b.per_telf, b.per_fax " & _
             " FROM adquisicion a, persona b" & _
             " WHERE  a.per_codigo = b.per_codigo and a.emp_codigo = '" & strEmpresa & "' AND b.cat_p_tipo =  'p' " & _
             " ORDER BY a.adq_codigo"
'            " where  a.per_codigo = b.per_codigo and tip_adq_codigo  = 'ICA' and a.emp_codigo = '" & strEmpresa & "'" &
    clsConsu.Ejecutar (strSql)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existe adquisicion de Suministros almacenados en el Sistema", vbInformation, "Sis-Admin"
        Exit Sub
    Else
        Set dcmbAdq_codigo.RowSource = clsConsu.adorec_Def.DataSource
        dcmbAdq_codigo.ListField = "adq_codigo"
    End If
    
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim X, Y  As Integer
    VSFDetalle.Tag = "N"
    VSFDetalle.Rows = 1
    VSFDetalle.Clear 1
    VSFDetalle.Tag = "T"
    
End Sub

Public Sub borrar_datos()
        txtCodProveedor.Text = ""
        txtNomProveedor.Text = ""
        txtRucProveedor.Text = ""
        txtDirProveedor.Text = ""
        txtTelProveedor.Text = ""
        txtFaxProveedor.Text = ""
        txtAdq_fecha.Text = ""
        txtObs.Text = ""
        LblIva.Caption = "IVA = 0 %"
        TxtIva.Text = ""
        TxtSubTotal.Text = ""
        TxtTotal.Text = ""
        txtNum_orden.Text = ""

End Sub

Private Sub VSFDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Or Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Then
Cancel = True
End If
End Sub

