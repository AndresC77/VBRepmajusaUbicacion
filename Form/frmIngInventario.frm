VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIngInventario 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Conteo"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmIngInventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10785
   Begin VB.CommandButton btn_cargar 
      Caption         =   "Cargar Toma Fisica del ""NS"""
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   840
      Width           =   1695
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
      Height          =   7815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10575
      Begin MSDataListLib.DataCombo dcbo_conteo 
         Height          =   315
         Left            =   1500
         TabIndex        =   19
         Top             =   840
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Value           =   41947.6801273148
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
         Height          =   5655
         Left            =   120
         TabIndex        =   9
         Top             =   1185
         Width           =   10335
         Begin VB.TextBox txtLector 
            Height          =   285
            Left            =   7440
            TabIndex        =   15
            Top             =   360
            Width           =   2415
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFG 
            Height          =   1695
            Left            =   3000
            TabIndex        =   13
            Top             =   3600
            Visible         =   0   'False
            Width           =   7215
            _cx             =   12726
            _cy             =   2990
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
            Top             =   840
            Width           =   10095
            _cx             =   1981826446
            _cy             =   1981816498
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmIngInventario.frx":030A
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
            Left            =   9000
            TabIndex        =   1
            Top             =   5280
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Left            =   6720
            TabIndex        =   16
            Top             =   435
            Width           =   555
         End
         Begin VB.Image imgBtnDn 
            Height          =   210
            Left            =   8040
            Picture         =   "frmIngInventario.frx":03AF
            Top             =   5280
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Image imgBtnUp 
            Height          =   210
            Left            =   7800
            Picture         =   "frmIngInventario.frx":04DB
            Top             =   5280
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
            Left            =   8280
            TabIndex        =   10
            Top             =   5340
            Width           =   735
         End
      End
      Begin VB.TextBox txtObs 
         Height          =   570
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   7080
         Width           =   10335
      End
      Begin VB.TextBox txtNumIngreso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
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
      Begin VB.Label Label2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "No. Conteo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1095
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
         Top             =   6825
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
         Left            =   3885
         TabIndex        =   7
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lblNumIngreso 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Conteo:"
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
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3686
      TabIndex        =   3
      Top             =   8070
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5644
      TabIndex        =   4
      Top             =   8070
      Width           =   1455
   End
End
Attribute VB_Name = "frmIngInventario"
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

'Variables Globales
Public i_conteo As Integer


Private Sub btn_cargar_Click()
 
 vsfgDetalleImp.Clear 1

 Dim tField As ADODB.Field
    '
    ' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
    
        'Se conecta a la base de SQL Server 2005
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "SELECT dcon_bodega,dcon_codproduc, dcon_cantidad FROM detconteo where  detconteo.cont_id = " & i_conteo, cnn, adOpenDynamic, adLockOptimistic
        
        Dim variable As Integer
        Dim contador As Integer
        
        For i = 1 To rst.RecordCount
         With rst
            If .EOF And .BOF Then
                lblData.Caption = "No hay ningún registro activo"
            Else
                vsfgDetalleImp.Rows = vsfgDetalleImp.Rows + 1
                
                'Insertamos el botón de eliminar en cada una de las filas
                
                vsfgDetalleImp.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalleImp.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        
                vsfgDetalleImp.TextMatrix(i, 1) = rst.Fields("dcon_bodega")
                vsfgDetalleImp.TextMatrix(i, 2) = rst.Fields("dcon_codproduc")
                vsfgDetalleImp.TextMatrix(i, 4) = rst.Fields("dcon_cantidad")
                rst.MoveNext
            End If
        End With
       Next i
        vsfgDetalleImp.Rows = vsfgDetalleImp.Rows - 1
    
       rst.Close
       cnn.Close
    End If
End Sub

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
    If Archivo <> "" Then
        VSFG.LoadGrid Archivo, flexFileExcel
        j = 1
        For i = 0 To VSFG.Rows - 1
            strSql = " SELECT count(*) as N FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
            clsCon_Def.Ejecutar strSql
            VSFG.ShowCell i, 0
            vsfgDetalleImp.ShowCell i, 0
            If clsCon_Def.adorec_Def("N") > 0 Then
                vsfgDetalleImp.TextMatrix(j, 1) = VSFG.TextMatrix(i, 0)
                vsfgDetalleImp_BeforeRowColChange j, 1, j, 2, False
                vsfgDetalleImp.TextMatrix(j, 2) = VSFG.TextMatrix(i, 1)
                vsfgDetalleImp_BeforeRowColChange j, 2, j, 4, False
                vsfgDetalleImp.TextMatrix(j, 4) = VSFG.TextMatrix(i, 2)
                j = j + 1
            Else
                MsgBox "El producto " & VSFG.TextMatrix(i, 1) & vbNewLine & _
                       "NO EXISTE y fue contado" & vbNewLine & _
                       VSFG.TextMatrix(i, 2) & " unidades", vbInformation, "Conteos"
            End If
        Next i
    End If
End Sub


Private Sub dcbo_conteo_Validate(Cancel As Boolean)
 If (dcbo_conteo.BoundText <> "") Then
      i_conteo = CInt(dcbo_conteo.BoundText)
  End If
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

Private Sub cmdAceptar_Click()
    Dim i As Long
    Dim j As Long
    Dim ff As String
    Dim aux As Long
    
    Dim numR As Long

           
    ff = Format(dtpFecha.Value, "yyyy-mm-dd")
    'Verifica si la fecha ingresada es correcta
    If (IsDate(ff)) = False Then
        MsgBox "La fecha de Conteo no es correcta", vbExclamation, "Inventario"
        Exit Sub
    End If
    'Verifica que existan datos en el FlexGrid
    If vsfgDetalleImp.Rows = 1 Then
        MsgBox "El conteo no tiene detalle", vbInformation, "Inventario"
        vsfgDetalleImp.AddItem ""
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) = vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1)
        vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
        vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
    Else
     
     If (TxtTotal.Text = "") Then
         Exit Sub
     End If
     If (vsfgDetalleImp.Rows - 1 <> 0) Then ' Si existen detalles, almaceno.
     
         Mensaje = "Existen " & vsfgDetalleImp.Rows - 1 & " detalle(s) en el conteo, desea guardar?" ' Define el mensaje.
         Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
         Título = "Inventario"   ' Define el título.
         respuesta = MsgBox(Mensaje, Estilo, Título)
         
         'Recorro el FlexGrid para almacenar los detalles del ingreso
         If respuesta = vbYes Then
             'Compacta la Matriz
             numR = 1
             Dim booBandera As Boolean
             ReDim prdDet(2, numR) As Variant
             prdDet(0, numR) = vsfgDetalleImp.TextMatrix(1, 1)
             prdDet(1, numR) = vsfgDetalleImp.TextMatrix(1, 2)
             prdDet(2, numR) = vsfgDetalleImp.TextMatrix(1, 4)
             
             For i = 2 To vsfgDetalleImp.Rows - 1
                booBandera = False
                For j = 1 To numR
                   ' si encontro repetido
                   If prdDet(0, j) = vsfgDetalleImp.TextMatrix(i, 1) And prdDet(1, j) = vsfgDetalleImp.TextMatrix(i, 2) Then
                       prdDet(2, j) = Val(Format(prdDet(2, j), "###0")) + Val(Format(vsfgDetalleImp.TextMatrix(i, 4), "###0"))
                       booBandera = True
                       Exit For
                   End If
                Next j
                'no encontro igual
                If booBandera = False Then
                    'inserta en matriz item para facturar
                    numR = numR + 1
                    ReDim Preserve prdDet(2, numR) As Variant
                    prdDet(0, numR) = vsfgDetalleImp.TextMatrix(i, 1)
                    prdDet(1, numR) = vsfgDetalleImp.TextMatrix(i, 2)
                    prdDet(2, numR) = vsfgDetalleImp.TextMatrix(i, 4)
                End If
             Next i
             
             clsCon_Def.Inicializar AdoConn, AdoConnMaster
             If MN = "N" Then
                strSql = " SELECT COALESCE(max(inv_codigo),0)" & _
                         " FROM inventario where emp_codigo = '" & strEmpresa & "' " & _
                         " GROUP BY emp_codigo"
                clsConsu.Ejecutar (strSql), "M"
        
                If (IsNull(clsConsu.adorec_Def.Fields(0).Value)) Then
                    aux = 1
                Else
                    aux = clsConsu.adorec_Def.Fields(0).Value + 1
                End If
             Else
                aux = txtNumIngreso.Text
                strSql = " DELETE FROM inventario " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND inv_codigo='" & aux & "'"
                clsCon_Def.Ejecutar strSql, "M"
                strSql = " DELETE FROM det_inventario " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND inv_codigo='" & aux & "'"
                clsCon_Def.Ejecutar strSql, "M"
             End If
             strSql = " INSERT INTO inventario (emp_codigo,inv_codigo, inv_fecha,inv_estado, inv_observacion, " & _
                      " inv_fechamod, inv_usumod) values " & _
                      "( '" & strEmpresa & "','" & aux & "','" & ff & "',0, '" & txtObs.Text & "' ," & _
                      " CURRENT_TIMESTAMP,substring_index(USER(),'@',1))"
                      
             clsCon_Def.Ejecutar strSql, "M"
             For i = 1 To numR
                If Val(prdDet(2, i)) <> 0 Or Trim(prdDet(1, i)) <> "" Then
                 strSql = " INSERT INTO det_inventario (emp_codigo, inv_codigo, dep_codigo, prd_codigo, det_inv_cantidad, det_inv_fechamod," & _
                          " det_inv_usumod) values ('" & strEmpresa & "' ," & _
                          " " & aux & " , '" & prdDet(0, i) & "'," & _
                          " '" & prdDet(1, i) & "', " & prdDet(2, i) & " ," & _
                          " CURRENT_TIMESTAMP, substring_index(USER(),'@',1))"
                 clsCon_Def.Ejecutar strSql, "M"
                 End If
             Next i
             MsgBox "Conteo almacenado", vbInformation, "Inventario"
                
            Call actualizar_conteo(i_conteo)
                
            Dim rpConteo As New frmReporte
            rpConteo.strNumero = aux
            rpConteo.strReporte = "rptConteo"
            rpConteo.Show
             
         End If
     End If
    End If
    Unload Me
End Sub


Private Sub actualizar_conteo(p_actualizar As Integer)

' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        
        'Se conecta a la base de SQL Server 2005
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "UPDATE conteo set cont_estado=" & 2 & " where cont_id = " & p_actualizar, cnn, adOpenDynamic, adLockOptimistic
    End If

End Sub




Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
    ActualizarCombosGrid
End Sub

Private Sub Form_Load()

Dim d As String
Dim m As Integer
Dim Y As String
Dim ff As Variant
Dim var As Long

' Objetos de conexion para SQL Server
    
    ' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        'Se conecta a la base de SQL Server 2005
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "SELECT cont_id,cont_codigo FROM conteo where cont_estado = 1", cnn, adOpenDynamic, adLockOptimistic
        
        
        Set dcbo_conteo.RowSource = rst
        dcbo_conteo.ListField = "cont_codigo"
        dcbo_conteo.BoundColumn = "cont_id"
    End If

    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 350
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_Pro.Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    'txtFechaIng.Text = Format(Date, "dd/mm/yyyy")
    
    'Descompone la fecha actual  en día, mes y año
    
    dtpFecha.Value = HoyDia
    
    'Consulta del nùmero de ingreso último, se agrega uno para el nuevo ingreso
    strSql = "select COALESCE(max(inv_codigo),0) as num FROM inventario WHERE emp_codigo = '" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsConsu.Ejecutar (strSql), "M"
    If clsConsu.adorec_Def.EOF Then
        txtNumIngreso.Text = "1"
    Else
        txtNumIngreso.Text = clsConsu.adorec_Def("num") + 1
        
    End If
    txtNumIngreso.Enabled = False
    
    'Insertamos el botón de eliminar en cada una de las filas
    
    ' initializa el flexgrid
    vsfgDetalleImp.Editable = flexEDKbdMouse
    vsfgDetalleImp.AllowUserResizing = flexResizeBoth
    
    ' Agrega un botón en el grid
    
    vsfgDetalleImp.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalleImp.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    i = 1
    If MN = "M" Then
        vsfgDetalleImp.Rows = 1
        strSql = " SELECT inv_observacion " & _
                 " FROM inventario " & _
                 " WHERE inv_codigo='" & NInv & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsConsu.Ejecutar (strSql)
        Me.txtObs.Text = clsConsu.adorec_Def("inv_observacion")
        strSql = " SELECT dep_codigo,det_inventario.prd_codigo,prd_nombre,det_inventario.det_inv_cantidad " & _
                 " FROM det_inventario INNER JOIN producto ON det_inventario.emp_codigo=producto.emp_codigo " & _
                 " AND det_inventario.prd_codigo=producto.prd_codigo " & _
                 " WHERE inv_codigo='" & NInv & "' " & _
                 " AND det_inventario.emp_codigo='" & strEmpresa & "'"
        clsConsu.Ejecutar (strSql)
        Set vsfgDetalleImp.DataSource = clsConsu.adorec_Def.DataSource
        clsConsu.adorec_Def.MoveFirst
        vsfgDetalleImp.TextMatrix(0, 2) = "Código"
        vsfgDetalleImp.TextMatrix(1, 2) = clsConsu.adorec_Def("prd_codigo")
        vsfgDetalleImp.TextMatrix(1, 4) = clsConsu.adorec_Def("det_inv_cantidad")
        txtNumIngreso.Text = NInv
        PonerBotones
    End If
    
    
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
       
        End Select
End Sub

Private Sub txtFechaIng_KeyPress(KeyAscii As Integer)
    'Validación de caracteres ingresados para que solo ingrese números y el caracter "/"
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProd UCase(txtLector.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub AgregarProd(codigo As String, Optional EsAux As Boolean = True)
  Dim i As Long
  Dim pas As Boolean
  pas = False
  
  clsCon_Pro.Filtrar "prd_codigo='" & codigo & "'"
  
  If clsCon_Pro.adorec_Def.RecordCount = 1 Then
    vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Row, 2) = codigo
    vsfgDetalleImp.ShowCell vsfgDetalleImp.Row, 2
  Else
    MsgBox "No se ha encontrado el producto con el código especificado." & vbCr & "Asegúrese el tipo de código del producto y que el mismo se encuentre en lista.", vbCritical, "Error de codigo"
  End If

End Sub


Private Sub vsfgDetalleImp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
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
    d = vsfgDetalleImp.Cell(flexcpLeft, r, c) + vsfgDetalleImp.Cell(flexcpWidth, r, c) - X
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

Private Sub vsfgDetalleImp_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 2) <> "" Then
        vsfgDetalleImp.AddItem "" & vbTab & vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) & vbTab & "" & vbTab & "" & vbTab & 0
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
        vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
    End If
End Sub

Private Sub vsfgDetalleImp_CellChanged(ByVal Row As Long, ByVal Col As Long)
If Row > 0 Then
   If (Col = 2 And vsfgDetalleImp.TextMatrix(Row, Col) <> "") Then
'        clsCon_Pro.adorec_Def.MoveFirst
'        clsCon_Pro.adorec_Def.Find "prd_codigo = '" & vsfgDetalleImp.TextMatrix(Row, Col) & "' ", , adSearchForward
'        If (clsCon_Pro.adorec_Def.EOF = False) Then
'            vsfgDetalleImp.TextMatrix(Row, Col + 1) = clsCon_Pro.adorec_Def("prd_nombre")
'        End If
        vsfgDetalleImp.TextMatrix(Row, 3) = vsfgDetalleImp.TextMatrix(Row, 2)
        vsfgDetalleImp.TextMatrix(Row, 4) = 1
    ElseIf (Col = 3 And vsfgDetalleImp.TextMatrix(Row, Col) <> "") Then
        vsfgDetalleImp.TextMatrix(Row, 2) = vsfgDetalleImp.TextMatrix(Row, 3)
        vsfgDetalleImp.TextMatrix(Row, 4) = 1
    End If
'    SumaCantidades
    If vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 2) <> "" Then
        vsfgDetalleImp.AddItem "" & vbTab & vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 1) & vbTab & "" & vbTab & "" & vbTab & 0
        vsfgDetalleImp.TextMatrix(vsfgDetalleImp.Rows - 1, 0) = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.Cell(flexcpPicture, (vsfgDetalleImp.Rows - 1), 0) = imgBtnUp
        vsfgDetalleImp.Cell(flexcpPictureAlignment, (vsfgDetalleImp.Rows - 1), 0) = flexAlignRightCenter
        vsfgDetalleImp.Row = vsfgDetalleImp.Rows - 1
        vsfgDetalleImp.Col = 2
    End If

End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub SumaCantidades()
    Dim Suma As Long
    Dim i As Long
    Suma = 0
    For i = 1 To Me.vsfgDetalleImp.Rows - 1
        Suma = Suma + Val(Format(vsfgDetalleImp.TextMatrix(i, 4), "#0"))
    Next i
    TxtTotal.Text = Suma
End Sub

Private Sub ActualizarCombosGrid()
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_Pro.Ejecutar (strSql)
    
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgDetalleImp
    'vsfgGrupo.BuildComboList(clsCon_Def.adorec_Def, "*gru_nombre, gru_codigo", "gru_nombre")
    vsfgDetalleImp.ColComboList(1) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    'Consulto los productos de la empresa
    strSql = " SELECT prd_codigo, prd_nombre " & _
             " FROM producto " & _
             " Where emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
             " ORDER BY prd_nombre "
    clsCon_Pro.Ejecutar (strSql)
    
    vsfgDetalleImp.ColComboList(3) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
        
    strSql = " SELECT prd_codigo, prd_nombre " & _
             " FROM producto " & _
             " Where emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
             " ORDER BY prd_codigo "
    clsCon_Pro.Ejecutar (strSql)
    
    'Cargo el código del producto en el combo del FlexGrid en la columna 2
    vsfgDetalleImp.ColComboList(2) = vsfgDetalleImp.BuildComboList(clsCon_Pro.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
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

