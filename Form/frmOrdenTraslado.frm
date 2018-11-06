VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOrdenTraslado 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden Traslado"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmOrdenTraslado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10185
   Begin VSFlex8Ctl.VSFlexGrid VSFGAbrir 
      Height          =   1260
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   3105
      _cx             =   58463589
      _cy             =   58460334
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOrdenTraslado.frx":030A
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
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar detalle"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   7830
      Width           =   1455
   End
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Transferencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   9975
      Begin VB.OptionButton optContenedorMayor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Buscar Contenedor Grande"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6240
         TabIndex        =   25
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   8760
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optUnidadesCargadas 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Unidades Cargadas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtN 
         Height          =   285
         Left            =   2680
         TabIndex        =   22
         Top             =   1440
         Width           =   480
      End
      Begin VB.OptionButton optUnidadesFijas 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Todos   nnnn  Unidades"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optCajaCompleta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Caja Completa"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox TxtObser 
         Height          =   285
         Left            =   1290
         TabIndex        =   18
         Top             =   1080
         Width           =   6000
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   285
         Left            =   7800
         TabIndex        =   14
         Top             =   383
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Value           =   41961.4577777778
      End
      Begin MSDataListLib.DataCombo cmbBodOrigen 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbBodDestino 
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label LblObser 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDestino 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         Caption         =   "Bodega de Destino"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblOrigen 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         Caption         =   "Bodega de Origen"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label15 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   7170
         TabIndex        =   6
         Top             =   420
         Width           =   585
      End
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
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   9975
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox TxtObserv 
         Height          =   285
         Left            =   240
         MaxLength       =   250
         TabIndex        =   11
         Top             =   5160
         Width           =   7335
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalle 
         Height          =   4410
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9735
         _cx             =   94716691
         _cy             =   94707299
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOrdenTraslado.frx":0358
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cantidad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7320
         TabIndex        =   16
         Top             =   4875
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   300
         TabIndex        =   12
         Top             =   4800
         Width           =   1185
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   8400
         Picture         =   "frmOrdenTraslado.frx":044E
         Top             =   5400
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   8640
         Picture         =   "frmOrdenTraslado.frx":0584
         Top             =   5400
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Transferir"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   7830
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   7830
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   4
      Top             =   6120
      Width           =   1410
   End
End
Attribute VB_Name = "frmOrdenTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private strSQL As String
Private strSql_1 As String


'Variables Globales
Public i_transferencia As Integer
Public i_flag As Integer
Public s_bodegaorigen As String
Public s_bodegadestino As String


Private Sub btn_cargar_Click()

    'Obtener_bodegas
    'Limpiar
    'CargaBodegas
    
    ' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        'Se conecta a la base de SQL Server 2005
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        
        rst2.Open "SELECT tra_id,tra_codigo,tra_fechatra FROM transferencia WHERE tra_id=" & i_transferencia, cnn, adOpenDynamic, adLockOptimistic
        dtpFecha.Value = Format(rst2.Fields("tra_fechatra"), "yyyy-mm-dd")
        TxtObserv.Text = UCase("TRANSFERENCIA: " & rst2.Fields("tra_codigo") & vbNewLine & "ID: " & rst2.Fields("tra_id"))
        rst.Open "SELECT dtra_codproducto, dtra_cantidad FROM dettransferencia where tra_id=" & i_transferencia, cnn, adOpenDynamic, adLockOptimistic
        
        
        For i = 1 To rst.RecordCount
         With rst
            If .EOF And .BOF Then
                lblData.Caption = "No hay ningún registro activo"
            Else
                vsfgDetalle.Rows = vsfgDetalle.Rows + 1
    
                'Insertamos el botón de eliminar en cada una de las filas
    
                vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                vsfgDetalle.ShowCell i, 1
                vsfgDetalle.TextMatrix(i, 1) = rst.Fields("dtra_codproducto")
                'vsfgDetalle.TextMatrix(i, 2) = rst.Fields("dtra_codproducto")
                vsfgDetalle.TextMatrix(i, 3) = rst.Fields("dtra_cantidad")
                rst.MoveNext
            End If
        End With
       Next i
    End If
 

End Sub

Private Sub chk_transferencias_Click()

If chk_transferencias.Value = 1 Then
    dcbo_transferencia.Enabled = True
    btn_cargar.Enabled = True
    i_flag = 1
Else
    dcbo_transferencia.Enabled = False
    btn_cargar.Enabled = False
    i_flag = 0
End If
Limpiar
CargaBodegas

End Sub

Private Sub cmbBodOrigen_Validate(Cancel As Boolean)
    If Me.vsfgDetalle.Rows > 1 Then
        Cancel = False
    End If

End Sub

Private Sub cmdAbrir_Click()
    Dim num As Integer
    Dim i As Long, j As Long
    Dim strPath As String
    Dim strLinea As String
    Dim Arch As String
    'Arch = cmbTDoc.Text & ".xls"
    VSFGAbrir.Clear 1
    VSFGAbrir.Rows = 1
    
    If vsfgDetalle.Rows > 1 Then
        strPath = Trim(App.Path)
        cmdArchivo.DialogTitle = "Abrir"
        'cmdArchivo.DefaultExt = strPath
        cmdArchivo.InitDir = strPath
        'cmdArchivo.FileName = Arch
        cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Documento de Excel 2007|*xlsx|Todos los Archivos|*.*"
        cmdArchivo.ShowOpen
        num = FreeFile
        Archivo = cmdArchivo.FileName
        j = 1
        If Archivo <> "" Then
            VSFGAbrir.LoadGrid Archivo, flexFileExcel
            vsfgDetalle.Rows = 1
            With VSFGAbrir
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" Then
                        vsfgDetalle.AddItem "", i
                        vsfgDetalle.TextMatrix(i, 1) = .TextMatrix(i, 0)
                        vsfgDetalle.ShowCell i, 1
                        If optCajaCompleta.Value = True Then
                            vsfgDetalle.TextMatrix(i, 6) = "MAX"
                        ElseIf optUnidadesFijas.Value = True Then
                            vsfgDetalle.TextMatrix(i, 6) = txtN.Text
                        ElseIf optUnidadesCargadas.Value = True Then
                            vsfgDetalle.TextMatrix(i, 6) = .TextMatrix(i, 1)
                        ElseIf optContenedorMayor.Value = True Then
                            vsfgDetalle.TextMatrix(i, 6) = "GRANDE"
                        End If
                        vsfgDetalle.TextMatrix(i, 0) = i
                        vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                        vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                        BuscarContenedores i
                    End If
                Next i
                If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = "" Then
                    vsfgDetalle.RemoveItem vsfgDetalle.Rows - 1
                End If
            End With
        End If
    Else
        MsgBox "No se tiene información para guardar", vbInformation, "Guardar"
    End If
    
    vsfgDetalle.SaveGrid "c:\Transferencia.xls", flexFileExcel
End Sub

Private Sub BuscarContenedores(Linea As Long)
    Dim clsBuscar As New clsConsulta
    Dim clsBuscarAux As New clsConsulta
    Dim clsBuscarAux2 As New clsConsulta
    Dim producto As String
    Dim saldo As Long
    producto = vsfgDetalle.TextMatrix(Linea, 1)
    clsBuscar.Inicializar AdoConn, AdoConnMaster
    clsBuscarAux.Inicializar AdoConn, AdoConnMaster
    clsBuscarAux2.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT contenedor_mercaderia.con_mer_codigo,contenedor_mercaderia.ubi_bod_codigo,prd_ubica_linea,prd_nombre,SUM(if(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
             " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
             " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
             " INNER JOIN producto ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
             " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
             " WHERE contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
             " AND contenedor_mercaderia.est_con_mer_codigo != '-1' " & _
             " AND (contenedor_mercaderia.ubi_bod_codigo LIKE '1%' OR contenedor_mercaderia.ubi_bod_codigo LIKE '2%') " & _
             " AND det_contenedor_mercaderia.prd_codigo = '" & producto & "'" & _
             " GROUP BY contenedor_mercaderia.con_mer_codigo " & _
             " HAVING tot>0 "
             
    If optCajaCompleta.Value = True Or optUnidadesFijas.Value = True Or optContenedorMayor.Value = True Then
        strSQL = strSQL & " ORDER BY tot DESC "
    ElseIf optUnidadesCargadas.Value = True Then
        strSQL = strSQL & " ORDER BY tot ASC "
    End If
    clsBuscar.Ejecutar strSQL
    While Not clsBuscar.adorec_Def.EOF
        If optCajaCompleta.Value = True Then
            strSQL = " SELECT producto.prd_codigo,SUM(if(con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                     " FROM det_contenedor_mercaderia INNER JOIN producto ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo" & _
                     " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo LIKE  '" & clsBuscar.adorec_Def("con_mer_codigo") & "'" & _
                     " GROUP BY producto.prd_codigo " & _
                     " HAVING tot>0" & _
                     " ORDER BY tot DESC "
            clsBuscarAux.Ejecutar strSQL
            If clsBuscarAux.adorec_Def.RecordCount > 0 Then
                vsfgDetalle.TextMatrix(Linea, 6) = clsBuscar.adorec_Def("tot")
                vsfgDetalle.TextMatrix(Linea, 3) = clsBuscar.adorec_Def("con_mer_codigo")
                vsfgDetalle.TextMatrix(Linea, 5) = clsBuscar.adorec_Def("prd_ubica_linea")
                strSQL = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(if(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                         " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                         " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                         " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                         " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                         " AND contenedor_mercaderia.ubi_bod_codigo='" & clsBuscar.adorec_Def("prd_ubica_linea") & "' " & _
                         " AND det_contenedor_mercaderia.prd_codigo = '" & producto & "'" & _
                         " GROUP BY contenedor_mercaderia.con_mer_codigo " & _
                         " HAVING tot>0" & _
                         " ORDER BY tot DESC"
                clsBuscarAux2.Ejecutar strSQL
                If clsBuscarAux2.adorec_Def.RecordCount > 0 Then
                    vsfgDetalle.TextMatrix(Linea, 4) = clsBuscarAux2.adorec_Def("con_mer_codigo")
                Else
                    vsfgDetalle.TextMatrix(Linea, 4) = 0
                End If
                Exit Sub
            End If
        ElseIf optUnidadesFijas.Value = True Or optUnidadesCargadas.Value Then
            If optUnidadesFijas.Value = True Then
                saldo = txtN.Text
            Else
                saldo = vsfgDetalle.TextMatrix(Linea, 6)
            End If
            strSQL = " SELECT producto.prd_codigo,ABS(SUM(if(con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)-" & txtN.Text & ") as orden,SUM(if(con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                     " FROM det_contenedor_mercaderia INNER JOIN producto ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo" & _
                     " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo LIKE  '" & clsBuscar.adorec_Def("con_mer_codigo") & "'" & _
                     " GROUP BY producto.prd_codigo " & _
                     " HAVING tot>0"
            If optUnidadesFijas.Value = True Then
                strSQL = strSQL & " ORDER BY orden ASC "
            Else
                strSQL = strSQL & " ORDER BY orden DESC "
            End If
            clsBuscarAux.Ejecutar strSQL
            While Not clsBuscarAux.adorec_Def.EOF
                If saldo > clsBuscar.adorec_Def("tot") Then
                    vsfgDetalle.TextMatrix(Linea, 6) = clsBuscar.adorec_Def("tot")
                    vsfgDetalle.TextMatrix(Linea, 3) = clsBuscar.adorec_Def("con_mer_codigo")
                    vsfgDetalle.TextMatrix(Linea, 5) = clsBuscar.adorec_Def("prd_ubica_linea")
                    strSQL = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(if(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                             " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                             " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                             " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                             " AND contenedor_mercaderia.ubi_bod_codigo='" & clsBuscar.adorec_Def("prd_ubica_linea") & "' " & _
                             " AND det_contenedor_mercaderia.prd_codigo = '" & producto & "'" & _
                             " GROUP BY contenedor_mercaderia.con_mer_codigo " & _
                             " HAVING tot>0" & _
                             " ORDER BY tot DESC "
                    clsBuscarAux2.Ejecutar strSQL
                    If clsBuscarAux2.adorec_Def.RecordCount > 0 Then
                        vsfgDetalle.TextMatrix(Linea, 4) = clsBuscarAux2.adorec_Def("con_mer_codigo")
                    Else
                        vsfgDetalle.TextMatrix(Linea, 4) = 0
                    End If
                    saldo = saldo - clsBuscar.adorec_Def("tot")
                Else
                    vsfgDetalle.TextMatrix(Linea, 6) = saldo
                    vsfgDetalle.TextMatrix(Linea, 3) = clsBuscar.adorec_Def("con_mer_codigo")
                    vsfgDetalle.TextMatrix(Linea, 5) = clsBuscar.adorec_Def("prd_ubica_linea")
                    strSQL = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(if(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                             " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                             " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                             " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                             " AND contenedor_mercaderia.ubi_bod_codigo='" & clsBuscar.adorec_Def("prd_ubica_linea") & "' " & _
                             " AND det_contenedor_mercaderia.prd_codigo = '" & producto & "'" & _
                             " GROUP BY contenedor_mercaderia.con_mer_codigo " & _
                             " HAVING tot>0" & _
                             " ORDER BY tot DESC "
                    clsBuscarAux2.Ejecutar strSQL
                    If clsBuscarAux2.adorec_Def.RecordCount > 0 Then
                        vsfgDetalle.TextMatrix(Linea, 4) = clsBuscarAux2.adorec_Def("con_mer_codigo")
                    Else
                        vsfgDetalle.TextMatrix(Linea, 4) = 0
                    End If
                    Exit Sub
                End If
                clsBuscarAux.adorec_Def.MoveNext
            Wend
        ElseIf optContenedorMayor.Value = True Then
            vsfgDetalle.TextMatrix(Linea, 2) = clsBuscar.adorec_Def("prd_nombre")
            vsfgDetalle.TextMatrix(Linea, 3) = clsBuscar.adorec_Def("con_mer_codigo")
            vsfgDetalle.TextMatrix(Linea, 4) = clsBuscar.adorec_Def("ubi_bod_codigo")
            vsfgDetalle.TextMatrix(Linea, 5) = clsBuscar.adorec_Def("prd_ubica_linea")
            vsfgDetalle.TextMatrix(Linea, 6) = clsBuscar.adorec_Def("tot")
            Exit Sub
        End If
        clsBuscar.adorec_Def.MoveNext
    Wend
        
    
End Sub

Private Sub cmdAceptar_Click()
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim i As Long
    Dim strObserv As String
    Dim booGuardar As Boolean
    Dim codEgr As String
    Dim codIng As String
    
    If cmbBodOrigen.Text = "" Or cmbBodDestino.Text = "" Then
        MsgBox "Llene los campos de Bodega", vbInformation, "Bodega"
        Exit Sub
    ElseIf cmbBodOrigen.Text = cmbBodDestino.Text Then
        MsgBox "La transferencia no se puede realizar a la misma Bodega", vbInformation, "Bodega"
        Exit Sub
    End If
    For i = 1 To vsfgDetalle.Rows - 1
        If vsfgDetalle.TextMatrix(i, 1) <> "" Or vsfgDetalle.TextMatrix(i, 2) <> "" Then
            If Val(Trim(vsfgDetalle.TextMatrix(i, 3))) <= 0 Then
                MsgBox "La cantidad del producto " & vsfgDetalle.Cell(flexcpTextDisplay, i, 2) & " es inválida", vbInformation, "Cantidad"
                Exit Sub
            End If
        End If
    Next i
    
    strObserv = TxtObserv.Text & vbNewLine & "TRANSFERENCIA DESDE " & cmbBodOrigen.Text & " HACIA " & cmbBodDestino.Text
    'strObserv = strObserv & vbCrLf & UCase(Trim(TxtObserv.Text))
    
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    booGuardar = clsEgreso.NuevoEgr(False, "ETR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , strObserv)
    
    If booGuardar Then
        codEgr = clsEgreso.strDoc
        With vsfgDetalle
            For i = 1 To .Rows - 1
                clsEgreso.NuevoDetEgr .TextMatrix(i, 1), cmbBodOrigen.BoundText, FormatoD2(.TextMatrix(i, 3)), 0, 0, 0, 0
            Next i
        End With
    
        booGuardar = clsIngreso.NuevoIng(False, "ITR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , strObserv)
        codIng = clsIngreso.strDoc
        With vsfgDetalle
            For i = 1 To .Rows - 1
                clsIngreso.NuevoDetIng .TextMatrix(i, 1), cmbBodDestino.BoundText, FormatoD2(.TextMatrix(i, 3)), 0, 0, 0, 0
            Next i
            InicializarContenedorRecurrente
        End With
    
        Set clsEgreso = Nothing
        Set clsIngreso = Nothing
        
        
        'Dim frmGuia As New frmReporte
        'frmGuia.strNumero = codIng
        'frmGuia.strTipo = "ITR"
        'frmGuia.strReporte = "rptGuiaRemisionTransferencia"
        'frmGuia.Show
        
        Dim frmEntrega As New frmReporte
        frmEntrega.strNumero = codIng
        frmEntrega.strTipo = "ETR"
        frmEntrega.strReporte = "rptEgresoMercaderia"
        frmEntrega.Show
        Setear_informacion (i_transferencia)
        Limpiar
        cargar_transferencia
    End If
    
End Sub

Private Sub Setear_informacion(p_actualizar As Integer)


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
        rst.Open "UPDATE transferencia set tra_estado=" & 2 & " where tra_id = " & p_actualizar, cnn, adOpenDynamic, adLockOptimistic
        cargar_transferencia
    End If
End Sub

Private Sub Limpiar()
    cmbBodOrigen.Text = ""
    cmbBodDestino.Text = ""
    
    cmbBodOrigen.BoundText = ""
    cmbBodDestino.BoundText = ""
    
    CargaBodegas
    vsfgDetalle.Clear 1
    vsfgDetalle.Rows = 2
    vsfgDetalle.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalle.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    TxtObserv.Text = ""
    dtpFecha.Value = HoyDia
End Sub



Private Sub cmdLimpiar_Click()
    Limpiar
End Sub

Private Sub dcbo_transferencia_Change()
 Limpiar
 If (dcbo_transferencia.BoundText <> "") Then
     i_transferencia = CInt(dcbo_transferencia.BoundText)
 End If
 Obtener_bodegas
 CargaBodegas
End Sub

Private Sub Obtener_bodegas()
    i_flag = 1
   
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
            rst.Open "SELECT tra_boddid, tra_bodoid FROM transferencia where tra_estado = 1 and tra_id=" & i_transferencia, cnn, adOpenDynamic, adLockOptimistic
            
            s_bodegaorigen = rst.Fields("tra_bodoid")
            s_bodegadestino = rst.Fields("tra_boddid")
            
        rst.Close
        cnn.Close
    End If
  
End Sub

Private Sub Form_Activate()
    CargaProductos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    
    cargar_transferencia

    dtpFecha.Value = HoyDia

        'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
   
    CargaBodegas

'    'Carga los depositos
'    strSql = " SELECT deposito.dep_codigo, deposito.dep_nombre " & _
'             " FROM deposito LEFT JOIN sucursal ON sucursal.dep_codigo=deposito.dep_codigo AND sucursal.emp_codigo=deposito.emp_codigo " & _
'             " AND sucursal.suc_codigo='" & strSucursal & "'" & _
'             " WHERE deposito.emp_codigo = '" & strEmpresa & "'" & _
'             " ORDER BY sucursal.suc_codigo DESC"
'    clsCon_Def.Ejecutar strSql
'    vsfgDetalle.ColComboList(1) = vsfgDetalle.BuildComboList(clsCon_Def.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
'
    'Consulta los recargos que puede manejar una empresa
'    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
'             " FROM ocargos " & _
'             " WHERE emp_codigo='" & strEmpresa & "' " & _
'             " ORDER BY oca_nombre "
'    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
'    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Insertamos el botón de eliminar en cada una de las filas
    vsfgDetalle.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalle.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    
'    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
'    strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
'             " FROM forma_pago " & _
'             " WHERE emp_codigo='" & strEmpresa & "' " & _
'             " ORDER BY for_pag_nombre "
'    clsFPago.Ejecutar (strSql)
'    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
'    CmbFpago.ListField = "for_pag_nombre"
'    CmbFpago.BoundColumn = "for_pag_codigo"
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        End Select
End Sub

Private Sub cargar_transferencia()
' Objetos de conexion para SQL Server
    
    ' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "SELECT DISTINCT transferencia.tra_id,transferencia.tra_codigo,transferencia.tra_fechatra FROM transferencia INNER JOIN dettransferencia ON transferencia.tra_id=dettransferencia.tra_id WHERE tra_estado = 1 AND tra_boddid!=tra_bodoid ORDER BY tra_codigo", cnn, adOpenDynamic, adLockOptimistic
        
        
        Set dcbo_transferencia.RowSource = rst
        dcbo_transferencia.ListField = "tra_codigo"
        dcbo_transferencia.BoundColumn = "tra_id"
    End If
End Sub

Private Sub vsfgDetalle_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalle.MouseRow
    c = vsfgDetalle.MouseCol
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = 1) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalle.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
   
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalle.Cell(flexcpLeft, r, c) + vsfgDetalle.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "AHORA DEBE ELIMINAR ESTA FILA!"
    
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        
        'Recorro el FlexGrid para almacenar los detalles del ingreso
        If respuesta = vbYes Then
            Dim i As Long
            If vsfgDetalle.Rows - 1 = r Then
                vsfgDetalle.RemoveItem (r)
                vsfgDetalle.AddItem ""
            Else
                vsfgDetalle.RemoveItem (r)
            End If
        
            For i = 1 To (vsfgDetalle.Rows - 1)
                vsfgDetalle.TextMatrix(i, 0) = i
                vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            Next i
            
        Else
            vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    Cancel = True
End Sub

Private Sub SumarExist()
    Dim i As Long
    Dim cant As Long
    cant = 0
    For i = 1 To vsfgDetalle.Rows - 1
        cant = cant + FormatoD0(vsfgDetalle.TextMatrix(i, 3))
    Next i
    txtCantidad.Text = FormatoD0(cant)
End Sub

Private Sub vsfgDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


'Private Sub CargarPersonas(Tipo As String)
''Carga los personas
'    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')') per " & _
'             " FROM persona " & _
'             " WHERE emp_codigo = '" & strEmpresa & "' " & _
'             " AND cat_p_tipo LIKE '" & Tipo & "'" & _
'             " ORDER BY CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,')')"
'    clsCon_Def.Ejecutar strSql
'    dcmbCodP.ListField = "per"
'    dcmbCodP.BoundColumn = "per_codigo"
'    Set dcmbCodP.RowSource = clsCon_Def.adorec_Def.DataSource
'End Sub


'Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
'    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 And VSFGReca.TextMatrix(OldRow, 1) <> "" Then
'        VSFGReca.AddItem ""
'        PonerBotones
'    End If
'End Sub

'Private Sub VSFGReca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    'Permite modificar solo la columna 0 del recargo
'    If Col = 2 Then
'        Cancel = True
'    End If
'End Sub

'Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'    With VSFGReca
'        ' only interesetd in left button
'        If Button <> 1 Then Exit Sub
'
'        ' get cell that was clicked
'        Dim r&, c&
'        r = .MouseRow
'        c = .MouseCol
'
'        ' make sure the click was on the sheet
'        If r < 0 Or c < 0 Then Exit Sub
'
'        If (c <> 0 Or r = (.Rows - 1)) Then Exit Sub
'
'        ' make sure the click was on a cell with a button
'        If .Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
'
'        ' make sure the click was on the button (not just on the cell)
'        ' note: this works for right-aligned buttons
'        Dim d!
'        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - X
'        If d > imgBtnDn.Width Then Exit Sub
'
'        ' click was on a button: do the work
'         .Cell(flexcpPicture, r, c) = imgBtnDn
'        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
'        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
'        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
'        Respuesta = MsgBox(Mensaje, Estilo, Título)
'
'        'Recorro el FlexGrid para poner números a las filas
'
'        If Respuesta = vbYes Then
'             Dim i As Integer
'              .RemoveItem (r)
'             PonerBotones
'             CalculaTotal
'        Else
'             .Cell(flexcpPicture, r, c) = imgBtnUp
'        End If
'
'        ' cancel default processing
'        ' note: this is not strictly necessary in this case, because
'        '       the dialog box already stole the focus etc, but let's be safe.
'        Cancel = True
'    End With
'End Sub

'Private Sub VSFGReca_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    'Busca y coloca el valor del recargo seleccionado
'    If Row > 0 And VSFGReca.TextMatrix(Row, 1) <> "" And Col <> 3 Then
'        clsRecargos.Filtrar "oca_codigo='" & VSFGReca.TextMatrix(Row, 1) & "'"
'        VSFGReca.TextMatrix(Row, 2) = clsRecargos.adorec_Def("oca_nombre")
'        VSFGReca.TextMatrix(Row, 3) = clsRecargos.adorec_Def("oca_precio")
'        clsRecargos.QuitarFiltro
'        'Verifica que no se haya escogido antes el mismo recargo, en ese caso suma sus valores
'        For i = 1 To VSFGReca.Rows - 1
'            If VSFGReca.TextMatrix(Row, 1) = VSFGReca.TextMatrix(i, 1) And Row <> i Then
'                VSFGReca.TextMatrix(i, 3) = Val(VSFGReca.TextMatrix(i, 3)) + (VSFGReca.TextMatrix(Row, 3))
'                VSFGReca.RemoveItem Row
'                PonerBotones
'                Exit For
'            End If
'        Next i
'    End If
'    CalculaTotal
'End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With vsfgDetalle
        For i = 1 To (.Rows - 1)
            .TextMatrix(i, 0) = i
            If conBot = True Then
                'Coloca los botones de elimniar fila en el grid
                .Cell(flexcpPicture, i, 0) = imgBtnUp
                .Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub


Private Sub CargaBodegas()
        Dim s_bod_origen As String
        Dim s_bod_destino As String
        
        If i_flag = 0 Then
            strSQL = " SELECT dep_codigo " & _
                  " FROM tipo_pedido " & _
                  " WHERE emp_codigo = '" & strEmpresa & "' AND tip_ped_ptofac='" & strPtoFactura & "' "
            clsCon_Def.Ejecutar strSQL
            s_bod_origen = clsCon_Def.adorec_Def("dep_codigo")
            strSQL = " SELECT dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "' ORDER BY dep_nombre "
        Else
            strSQL = " SELECT dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "'  AND dep_codigo= '" & s_bodegaorigen & "' ORDER BY dep_nombre "
            
            strSql_1 = " SELECT dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "'  AND dep_codigo= '" & s_bodegadestino & "' ORDER BY dep_nombre "
        End If
              
        clsCon_Def.Ejecutar strSQL
        cmbBodOrigen.ListField = "dep_nombre"
        cmbBodOrigen.BoundColumn = "dep_codigo"
        Set cmbBodOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
        
      If i_flag = 0 Then
        cmbBodDestino.ListField = "dep_nombre"
        cmbBodDestino.BoundColumn = "dep_codigo"
        Set cmbBodDestino.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbBodOrigen.BoundText = s_bod_origen
      Else
        clsCon_Def.Ejecutar strSql_1
        cmbBodDestino.ListField = "dep_nombre"
        cmbBodDestino.BoundColumn = "dep_codigo"
        Set cmbBodDestino.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbBodDestino.BoundText = s_bodegadestino
        cmbBodOrigen.BoundText = s_bodegaorigen
      End If
End Sub


Private Sub CargaProductos()
    'Carga los productos
    strSQL = " SELECT producto.prd_codigo, prd_nombre,sum(exi_cantidad) as c " & _
             " FROM producto INNER JOIN existencia ON producto.emp_codigo=existencia.emp_codigo" & _
             " AND producto.prd_codigo=existencia.prd_codigo" & _
             " AND existencia.dep_codigo='" & cmbBodOrigen.BoundText & "'" & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' AND prd_baja=0 " & _
             " group by producto.prd_codigo " & _
             " having sum(exi_cantidad)!=0 " & _
             " ORDER BY producto.prd_codigo "
    clsCon_Prd.Ejecutar strSQL
  '  vsfgDetalle.ColComboList(1) = vsfgDetalle.BuildComboList(clsCon_Prd.adorec_Def, "*prd_codigo, prd_nombre", "prd_codigo")
    vsfgDetalle.ColComboList(2) = vsfgDetalle.BuildComboList(clsCon_Prd.adorec_Def, "prd_codigo, *prd_nombre", "prd_codigo")
    'Consulto los productos de la empresa
'    strSql = " SELECT producto.prd_codigo, prd_nombre " & _
'             " FROM producto " & _
'             " WHERE producto.emp_codigo = '" & strEmpresa & "' AND prd_baja=0 ORDER BY prd_nombre "
'    clsCon_Def.Ejecutar strSql
    
End Sub
