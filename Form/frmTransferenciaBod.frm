VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTransferenciaBod 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia entre Bodegas"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmTransferenciaBod.frx":0000
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
      TabIndex        =   18
      Top             =   7080
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
      FormatString    =   $"frmTransferenciaBod.frx":030A
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
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   9000
      TabIndex        =   17
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar detalle"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
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
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton btn_cargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkContenedorNuevo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Contenedor Nuevo"
         Height          =   255
         Left            =   8040
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtPedido 
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   1320
         Width           =   2655
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   285
         Left            =   7800
         TabIndex        =   13
         Top             =   383
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Value           =   42383.3915509259
      End
      Begin MSDataListLib.DataCombo cmbBodDestino 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGBodOrigen 
         Height          =   1020
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   2625
         _cx             =   4630
         _cy             =   1799
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransferenciaBod.frx":0358
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
      Begin VB.Label lblBodOrigen 
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
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   5625
      End
      Begin VB.Label lbl_transferencia 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         Caption         =   "Pedido"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblDestino 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         Caption         =   "Bodega de Destino"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3240
         TabIndex        =   9
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
         TabIndex        =   7
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
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   2040
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
         TabIndex        =   14
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox TxtObserv 
         Height          =   285
         Left            =   240
         MaxLength       =   250
         TabIndex        =   10
         Top             =   5040
         Width           =   7335
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalle 
         Height          =   4410
         Left            =   120
         TabIndex        =   0
         Top             =   240
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransferenciaBod.frx":03BD
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
         TabIndex        =   15
         Top             =   4755
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
         TabIndex        =   11
         Top             =   4680
         Width           =   1185
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   8400
         Picture         =   "frmTransferenciaBod.frx":048C
         Top             =   5280
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   8640
         Picture         =   "frmTransferenciaBod.frx":05C2
         Top             =   5280
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
Attribute VB_Name = "frmTransferenciaBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private strSQL As String
Private strSql_1 As String
Private cargado As Boolean

'Variables Globales
Public i_transferencia As Integer
Public i_flag As Integer
Public s_bodegaorigen As String
Public s_bodegadestino As String
Private strBodOrigen As String
Private strBodOrigenNombre As String

Private strCampo As String

Private Sub ModBodOrigen()
Dim i As Long
    
    strBodOrigen = ""
    strBodOrigenNombre = ""
    For i = 1 To VSFGBodOrigen.Rows - 1
        If Abs(VSFGBodOrigen.TextMatrix(i, 0)) = 1 Then
            strBodOrigen = strBodOrigen & "'" & VSFGBodOrigen.TextMatrix(i, 1) & "',"
            strBodOrigenNombre = strBodOrigenNombre & VSFGBodOrigen.TextMatrix(i, 2) & ","
        End If
    Next i
    lblBodOrigen = "Origen: " & strBodOrigenNombre
End Sub

Private Sub btn_cargar_Click()
    Dim clsConAUX As New clsConsulta
    'strBodOrigen = Left(strBodOrigen, Len(strBodOrigen) - 1)
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    If MsgBox("Desea buscar solo el PCK?" & vbNewLine & "Si responde que NO se buscara la CANTIDAD", vbQuestion + vbYesNo, "Transferencia") = vbYes Then
        strCampo = "det_ped_cant_entregada"
    Else
        strCampo = "det_ped_cant_pedida"
    End If
    ModBodOrigen
    strSQL = " SELECT dep_codigo,CONCAT(per_apellido,' ',per_nombre) as cli,prd_codigo," & strCampo & " as cant " & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
             " AND pedido.per_codigo=persona.per_codigo " & _
             " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
             " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND ped_estado in(0,1)" & _
             " AND pedido.ped_codigo='" & txtPedido.Text & "'" & _
             " AND " & strCampo & "!=0"
    clsConAUX.Ejecutar strSQL
    If clsConAUX.adorec_Def.RecordCount > 0 Then
        'If Not (cmbBodOrigen.BoundText = clsConAUX.adorec_Def("dep_codigo") Or cmbBodDestino.BoundText = clsConAUX.adorec_Def("dep_codigo")) Then
        'cmbBodOrigen.BoundText = clsConAUX.adorec_Def("dep_codigo")
        'End If
        TxtObserv.Text = "Reserva para " & clsConAUX.adorec_Def("cli") & " PEDIDO No. " & txtPedido.Text
        clsConAUX.adorec_Def.MoveFirst
        vsfgDetalle.Clear 1
        vsfgDetalle.Rows = 2
        cargado = True
        i = 1
        While Not clsConAUX.adorec_Def.EOF
            i = vsfgDetalle.Rows - 1
            vsfgDetalle.TextMatrix(i, 0) = i
            vsfgDetalle.TextMatrix(i, 1) = clsConAUX.adorec_Def("prd_codigo")
            vsfgDetalle.TextMatrix(i, 3) = clsConAUX.adorec_Def("cant")
            clsConAUX.adorec_Def.MoveNext
            If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) <> "" Then
                vsfgDetalle.AddItem ""
            End If
        Wend
    Else
        MsgBox "El pedido no existe, esta vacio, o ya esta confirmado facturado o de baja"
    End If
    SumarExist
    vsfgDetalle.SaveGrid "c:\TR.txt", flexFileTabText
End Sub

Private Sub cmbBodOrigen_Validate(Cancel As Boolean)
    If Me.vsfgDetalle.Rows > 1 Then
        Cancel = False
    End If

End Sub

Private Sub cmdAbrir_Click()
    Dim num As Integer
    
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
        If Archivo <> "" Then
            VSFGAbrir.LoadGrid Archivo, flexFileExcel
            vsfgDetalle.Rows = 1
            With VSFGAbrir
                j = 1
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" Then
                        If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) <> "" Then
                            vsfgDetalle.AddItem ""
                        End If
                        j = vsfgDetalle.Rows - 1
                        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = .TextMatrix(i, 0)
                        vsfgDetalle.ShowCell vsfgDetalle.Rows - 1, 1
                        vsfgDetalle.TextMatrix(j, 3) = .TextMatrix(i, 1)
                        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 0) = i
                        vsfgDetalle.Cell(flexcpPicture, vsfgDetalle.Rows - 1, 0) = imgBtnUp
                        vsfgDetalle.Cell(flexcpPictureAlignment, vsfgDetalle.Rows - 1, 0) = flexAlignRightCenter
                    End If
'                    j = vsfgDetalle.Rows - 1
                Next i
                If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = "" Then
                    vsfgDetalle.RemoveItem vsfgDetalle.Rows - 1
                End If
            End With
        End If
    Else
        MsgBox "No se tiene información para guardar", vbInformation, "Guardar"
    End If

End Sub

Private Sub cmdAceptar_Click()
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim clsConAUX As New clsConsulta
    Dim i As Long
    Dim strObserv As String
    Dim booGuardar As Boolean
    Dim codEgr As String
    Dim codIng As String
    Dim strBodOrigenNombre As String
    
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    If strBodOrigen = "" Or cmbBodDestino.Text = "" Then
        MsgBox "Llene los campos de Bodega", vbInformation, "Bodega"
        Exit Sub
    ElseIf strBodOrigen = cmbBodDestino.Text Then
        MsgBox "La transferencia no se puede realizar a la misma Bodega", vbInformation, "Bodega"
        Exit Sub
    End If
    For i = 1 To vsfgDetalle.Rows - 1
        If vsfgDetalle.TextMatrix(i, 1) <> "" Or vsfgDetalle.TextMatrix(i, 2) <> "" Then
            If Val(Trim(vsfgDetalle.TextMatrix(i, 3))) < 0 Then
                MsgBox "La cantidad del producto " & vsfgDetalle.Cell(flexcpTextDisplay, i, 2) & " es inválida", vbInformation, "Cantidad"
                Exit Sub
            End If
        End If
    Next i
    
    
    strObserv = TxtObserv.Text & vbNewLine & "TRANSFERENCIA DESDE " & strBodOrigenNombre & " HACIA " & cmbBodDestino.Text
    'strObserv = strObserv & vbCrLf & UCase(Trim(TxtObserv.Text))
    
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    booGuardar = clsEgreso.NuevoEgr(True, "ETR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , strObserv)
    booGuardar = clsIngreso.NuevoIng(True, "ITR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , strObserv)
    If booGuardar Then
        codEgr = clsEgreso.strDoc
        With vsfgDetalle
            For i = 1 To .Rows - 1
                If FormatoD4(.TextMatrix(i, 3)) <> 0 Then
                    clsEgreso.NuevoDetEgr .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD4(.TextMatrix(i, 3)), 0, 0, 0, 0
                    .ShowCell i, 0
                End If
            Next i
        End With
    
        
        If chkContenedorNuevo.Value = 1 Then
            strContenedorRecurrente = "111"
        Else
            strContenedorRecurrente = ""
        End If
        'booGuardar = clsIngreso.NuevoIng(True, "ITR", False, strSucursal, strPtoFactura, , , , Format(dtpFecha.Value, "yyyy-MM-dd"), , , strObserv)
        codIng = clsIngreso.strDoc
        If txtPedido.Text <> "" Then
        strSQL = " UPDATE det_pedido " & _
                 " SET det_ped_cant_entregada=0,det_ped_cant_confirmada=0" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ped_codigo='" & txtPedido.Text & "'"
        clsConAUX.Ejecutar strSQL, "M"
        End If
        With vsfgDetalle
            For i = 1 To .Rows - 1
                If FormatoD4(.TextMatrix(i, 3)) <> 0 Then
                    clsIngreso.NuevoDetIng .TextMatrix(i, 1), cmbBodDestino.BoundText, FormatoD4(.TextMatrix(i, 3)), 0, 0, 0, 0
                    If Trim(txtPedido.Text) <> "" Then
                        strSQL = " UPDATE det_pedido " & _
                                 " SET det_ped_cant_entregada=det_ped_cant_entregada+'" & FormatoD4(.TextMatrix(i, 3)) & "',det_ped_cant_confirmada=det_ped_cant_confirmada+'" & FormatoD4(.TextMatrix(i, 3)) & "'" & _
                                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                                 " AND ped_codigo='" & txtPedido.Text & "' AND prd_codigo='" & .TextMatrix(i, 1) & "'"
                        clsConAUX.Ejecutar strSQL, "M"
                    End If
                End If
            Next i
        End With
        InicializarContenedorRecurrente
        Set clsEgreso = Nothing
        Set clsIngreso = Nothing
        If txtPedido.Text <> "" Then
        strSQL = " UPDATE det_pedido " & _
                 " SET dep_codigo='" & cmbBodDestino.BoundText & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ped_codigo='" & txtPedido.Text & "'"
        clsConAUX.Ejecutar strSQL, "M"
        End If
        'Dim frmGuia As New frmReporte
        'frmGuia.strNumero = codIng
        'frmGuia.strTipo = "ITR"
        'frmGuia.strReporte = "rptGuiaRemisionTransferencia"
        'frmGuia.Show
        
        Dim frmEntrega As New frmReporte
        frmEntrega.strNumero = codIng
        frmEntrega.strTipo = "ETR"
        frmEntrega.strReporte = "rptTransferencia"
        frmEntrega.Show
        Limpiar
    End If
    InicializarContenedorRecurrente
    chkContenedorNuevo.Value = 0
End Sub


Private Sub Limpiar()
    strBodOrigen = ""
    strBodOrigenNombre = ""
    cmbBodDestino.Text = ""
    
    strBodOrigenNombre = ""
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
    
    cargado = False
    dtpFecha.Value = HoyDia

        'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
   
    CargaBodegas
    ModBodOrigen
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

Private Sub VSFGBodOrigen_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub VSFGBodOrigen_LostFocus()
    ModBodOrigen
    CargaProductos
End Sub

Private Sub vsfgDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 4 Or cargado = True Then Cancel = True
End Sub

'Private Sub vsfgDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col >= 8 Then
'        Cancel = True
'    End If
'    If ModPreCos = False Then
'        If Col = 5 Or Col = 6 Or Col = 7 Then
'            Cancel = True
'        End If
'    End If
'End Sub

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

Private Sub vsfgDetalle_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim band As Boolean
    Dim booNuevo As Boolean
    Dim booSalir As Boolean
    Dim dblSaldo As Double
    Dim i As Long
'    If cmbBodOrigen.MatchedWithList = False Then
'        MsgBox "Seleccione la bodega de origen", vbInformation, "Bodega"
'        vsfgDetalle.TextMatrix(Row, Col) = ""
'
'        Exit Sub
'    End If
    
    
    If Col = 1 Then
        vsfgDetalle.TextMatrix(Row, 2) = vsfgDetalle.TextMatrix(Row, 1)
    ElseIf Col = 2 Then
        vsfgDetalle.TextMatrix(Row, 1) = vsfgDetalle.TextMatrix(Row, 2)
    End If
    
    strSqlPrd = " SELECT existencia.dep_codigo, producto.prd_codigo, COALESCE(SUM(existencia.exi_cantidad),0) as exi_cantidad, " & _
                 " producto.prd_nombre " & _
                 " FROM producto  INNER JOIN existencia " & _
                 " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo " & _
                 " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND producto.prd_codigo='" & vsfgDetalle.TextMatrix(Row, 1) & "'" & _
                 " AND dep_codigo in (" & strBodOrigen & "'') " & _
                 " GROUP BY existencia.dep_codigo, producto.prd_codigo,producto.prd_nombre HAVING COALESCE(SUM(existencia.exi_cantidad),0)>0" & _
                 " ORDER BY producto.prd_nombre,exi_cantidad ASC "
    If Col = 1 Then
        
        clsCon_Prd.Ejecutar (strSqlPrd)
        If clsCon_Prd.adorec_Def.RecordCount <= 0 Then
            vsfgDetalle.TextMatrix(Row, 1) = vsfgDetalle.TextMatrix(Row, 2)
            Exit Sub
        End If
        'clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalle.TextMatrix(Row, 1) & "'"
        booNuevo = True
        While Not clsCon_Prd.adorec_Def.EOF
            If booNuevo = True Then
                booNuevo = False
                vsfgDetalle.TextMatrix(Row, 3) = 0
                vsfgDetalle.TextMatrix(Row, 4) = clsCon_Prd.adorec_Def("exi_cantidad")
                vsfgDetalle.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("dep_codigo")
            Else
                vsfgDetalle.AddItem Row + 1 & vbTab & clsCon_Prd.adorec_Def("prd_codigo") & _
                vbTab & clsCon_Prd.adorec_Def("prd_codigo") & vbTab & "0" & vbTab & clsCon_Prd.adorec_Def("exi_cantidad") & _
                vbTab & clsCon_Prd.adorec_Def("dep_codigo")
            End If
            clsCon_Prd.adorec_Def.MoveNext
        Wend
    ElseIf Col = 2 Then
        
        clsCon_Prd.Ejecutar (strSqlPrd)
        If clsCon_Prd.adorec_Def.RecordCount <= 0 Then
            vsfgDetalle.TextMatrix(Row, 1) = vsfgDetalle.TextMatrix(Row, 2)
            Exit Sub
        End If
        'clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalle.TextMatrix(Row, 1) & "'"
        booNuevo = True
        While Not clsCon_Prd.adorec_Def.EOF
            If booNuevo = True Then
                booNuevo = False
                vsfgDetalle.TextMatrix(Row, 3) = 0
                vsfgDetalle.TextMatrix(Row, 4) = clsCon_Prd.adorec_Def("exi_cantidad")
                vsfgDetalle.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("dep_codigo")
            Else
                vsfgDetalle.AddItem Row + 1 & vbTab & clsCon_Prd.adorec_Def("prd_codigo") & _
                vbTab & clsCon_Prd.adorec_Def("prd_codigo") & vbTab & "0" & vbTab & clsCon_Prd.adorec_Def("exi_cantidad") & _
                vbTab & clsCon_Prd.adorec_Def("dep_codigo")
            End If
            clsCon_Prd.adorec_Def.MoveNext
        Wend
    ElseIf Col = 3 Then
        If FormatoD4(vsfgDetalle.TextMatrix(Row, 3)) <= 0 Or Not IsNumeric(vsfgDetalle.TextMatrix(Row, 3)) Then
            MsgBox "Ingrese un número válido", vbInformation, "Cantidad"
        Else
            If Row < vsfgDetalle.Rows - 1 Then
                If FormatoD4(vsfgDetalle.TextMatrix(Row, 3)) > FormatoD4(vsfgDetalle.TextMatrix(Row, 4)) And vsfgDetalle.TextMatrix(Row, 1) <> vsfgDetalle.TextMatrix(Row + 1, 1) Then
                    MsgBox "Puede Transferir máximo " & vsfgDetalle.TextMatrix(Row, 4) & " unidades.", vbInformation, "Cantidad"
                    vsfgDetalle.TextMatrix(Row, 3) = vsfgDetalle.TextMatrix(Row, 4)
                ElseIf FormatoD4(vsfgDetalle.TextMatrix(Row, 3)) > FormatoD4(vsfgDetalle.TextMatrix(Row, 4)) And vsfgDetalle.TextMatrix(Row, 1) = vsfgDetalle.TextMatrix(Row + 1, 1) Then
                    i = Row
                    dblSaldo = vsfgDetalle.TextMatrix(Row, 3)
                    While booSalir = False
                        If dblSaldo > vsfgDetalle.TextMatrix(i, 4) Then
                            vsfgDetalle.TextMatrix(i, 3) = vsfgDetalle.TextMatrix(i, 4)
                            dblSaldo = dblSaldo - vsfgDetalle.TextMatrix(i, 4)
                        Else
                            vsfgDetalle.TextMatrix(i, 3) = dblSaldo
                            dblSaldo = 0
                        End If
                        i = i + 1
                        If i > vsfgDetalle.Rows - 1 Then
                            booSalir = True
                        Else
                            If vsfgDetalle.TextMatrix(Row, 1) <> vsfgDetalle.TextMatrix(i, 1) Or dblSaldo = 0 Then
                                booSalir = True
                            End If
                        End If
                            
                    Wend
                End If
            Else
                dblSaldo = vsfgDetalle.TextMatrix(Row, 3)
                vsfgDetalle.ShowCell Row, 3
                i = Row
                If dblSaldo > FormatoD4(vsfgDetalle.TextMatrix(i, 4)) Then
                    vsfgDetalle.TextMatrix(i, 3) = FormatoD4(vsfgDetalle.TextMatrix(i, 4))
                    dblSaldo = dblSaldo - FormatoD4(vsfgDetalle.TextMatrix(i, 4))
                Else
                    vsfgDetalle.TextMatrix(i, 3) = dblSaldo
                    dblSaldo = 0
                End If
            End If
        End If
    End If
    band = False
    If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 2) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 3) <> "0" Then
        For i = 1 To vsfgDetalle.Rows - 1
            If i <> Row Then
                If vsfgDetalle.TextMatrix(i, 1) = vsfgDetalle.TextMatrix(Row, 1) And vsfgDetalle.TextMatrix(i, 5) = vsfgDetalle.TextMatrix(Row, 5) Then
                    vsfgDetalle.TextMatrix(i, 3) = Val(vsfgDetalle.TextMatrix(i, 3)) + Val(vsfgDetalle.TextMatrix(Row, 3))
                    If FormatoD4(vsfgDetalle.TextMatrix(i, 3)) > FormatoD4(vsfgDetalle.TextMatrix(i, 4)) Then
                        MsgBox "Puede Transferir máximo " & FormatoD4(vsfgDetalle.TextMatrix(i, 4)) & " unidades.", vbInformation, "Cantidad"
                        vsfgDetalle.TextMatrix(i, 3) = vsfgDetalle.TextMatrix(i, 4)
                    End If
                    band = True
                End If
            End If
        Next i
        If band Then vsfgDetalle.RemoveItem (Row)
                
        vsfgDetalle.AddItem ""
        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 0) = vsfgDetalle.Rows - 1
        vsfgDetalle.Cell(flexcpPicture, vsfgDetalle.Rows - 1, 0) = imgBtnUp
        vsfgDetalle.Cell(flexcpPictureAlignment, vsfgDetalle.Rows - 1, 0) = flexAlignRightCenter
        vsfgDetalle.Col = 1
        vsfgDetalle.Row = vsfgDetalle.Rows - 1
    End If
    SumarExist
End Sub

Private Sub SumarExist()
    Dim i As Long
    Dim cant As Double
    cant = 0
    For i = 1 To vsfgDetalle.Rows - 1
        cant = cant + FormatoD4(vsfgDetalle.TextMatrix(i, 3))
    Next i
    txtCantidad.Text = FormatoD0(cant)
    
        If FormatoD0(txtCantidad.Text) = 0 Then
               cmdAceptar.Enabled = False
            Else
               cmdAceptar.Enabled = True
        End If
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
        Dim i As Long
        
        If i_flag = 0 Then
            strSQL = " SELECT dep_codigo " & _
                  " FROM tipo_pedido " & _
                  " WHERE emp_codigo = '" & strEmpresa & "' AND tip_ped_ptofac='" & strPtoFactura & "' "
            clsCon_Def.Ejecutar strSQL
            s_bod_origen = clsCon_Def.adorec_Def("dep_codigo")
            strSQL = " SELECT 0 as sel,dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "' ORDER BY dep_nombre "
        Else
            strSQL = " SELECT 0 as sel,dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "'  AND dep_codigo= '" & s_bodegaorigen & "' ORDER BY dep_nombre "
            
            strSql_1 = " SELECT 0 as sel,dep_codigo, dep_nombre " & _
                  " FROM deposito " & _
                  " WHERE emp_codigo = '" & strEmpresa & "'  AND dep_codigo= '" & s_bodegadestino & "' ORDER BY dep_nombre "
        End If
              
        clsCon_Def.Ejecutar strSQL
        Set VSFGBodOrigen.DataSource = clsCon_Def.adorec_Def.DataSource
        
      If i_flag = 0 Then
        cmbBodDestino.ListField = "dep_nombre"
        cmbBodDestino.BoundColumn = "dep_codigo"
        Set cmbBodDestino.RowSource = clsCon_Def.adorec_Def.DataSource
        For i = 1 To VSFGBodOrigen.Rows - 1
            If VSFGBodOrigen.TextMatrix(i, 1) = s_bod_origen Then
                VSFGBodOrigen.TextMatrix(i, 0) = 1
            End If
        Next i
      Else
        clsCon_Def.Ejecutar strSql_1
        cmbBodDestino.ListField = "dep_nombre"
        cmbBodDestino.BoundColumn = "dep_codigo"
        Set cmbBodDestino.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbBodDestino.BoundText = s_bodegadestino
        For i = 1 To VSFGBodOrigen.Rows - 1
            If VSFGBodOrigen.TextMatrix(i, 1) = s_bod_origen Then
                VSFGBodOrigen.TextMatrix(i, 0) = 1
            End If
        Next i
      End If
End Sub


Private Sub CargaProductos()
    'Carga los productos
    strSQL = " SELECT producto.prd_codigo, prd_nombre,sum(exi_cantidad) as c " & _
             " FROM producto INNER JOIN existencia ON producto.emp_codigo=existencia.emp_codigo" & _
             " AND producto.prd_codigo=existencia.prd_codigo" & _
             " AND existencia.dep_codigo in (" & strBodOrigen & "'')" & _
             " WHERE producto.emp_codigo = '" & strEmpresa & "' AND prd_baja=0 " & _
             " group by producto.prd_codigo, prd_nombre " & _
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
