VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCambioProductos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cambio de Codigos"
   ClientHeight    =   5505
   ClientLeft      =   2520
   ClientTop       =   3840
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambioProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11385
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   10935
      _cx             =   19288
      _cy             =   6376
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCambioProductos.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
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
      Height          =   1215
      Left            =   825
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   2640
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Archivo"
         Height          =   315
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar de Consulta"
         Height          =   315
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmbAceptar 
         Caption         =   "&Procesar Cambios"
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7920
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "&Importar de Archivo"
         Height          =   315
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   5640
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargar de: "
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   525
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA: El Orden debe ser CodigoFinal, Codigo Anterior, Descipción, Unidad de Medida, Linea, Marca, Grupo, SubGrupo"
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
      Height          =   210
      Left            =   885
      TabIndex        =   3
      Top             =   1440
      Width           =   9615
   End
   Begin VB.Menu Eliminar 
      Caption         =   "Eliminar"
      Visible         =   0   'False
      Begin VB.Menu menElimina 
         Caption         =   "Eliminar lo q sea"
      End
   End
End
Attribute VB_Name = "frmCambioProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Private clsCon_Ins As New clsConsulta
Public strUniMed As String
Public strMarca As String
Public strLinea As String
Public strGrupo As String

Private Sub cmbAceptar_Click()
    VSFG.Cols = 9
    RevisarColumnas
    CambioCodigosProductos
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdCargar_Click()
    Dim strSql As String
    frmConsultaPrd.Show vbModal
    'VSFG.AddItem "", 0
    VSFG.FixedRows = 1
    strSql = " SELECT prd_codigo as 'CODIGO',prd_codigo as 'CODIGO ACTUAL',prd_nombre as 'NOMBRE',uni_codigo as 'UNIDAD DE MEDIDA',lin_codigo as 'LINEA',mar_codigo as 'MARCA',left(gru_codigo,2) as 'GRUPO',gru_codigo as 'SUBGRUPO' " & _
             " FROM producto " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND uni_codigo like '" & strUniMed & "' " & _
             " AND lin_codigo like '" & strLinea & "' " & _
             " AND mar_codigo like '" & strMarca & "' " & _
             " AND gru_codigo like '" & strGrupo & "' "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    Enumerar
    codigoAnombre
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowSave
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    If cdArchivo.FileName <> "" Then
        VSFG.Cols = 1
        VSFG.Rows = 1
        VSFG.FixedRows = 0
        VSFG.LoadGrid cdArchivo.FileName, flexFileTabText
        VSFG.FixedRows = 1
        VSFG.RemoveItem VSFG.Rows - 1
        Enumerar
    End If
End Sub

Private Sub cmdExportar_Click()
    Dim sDir As String
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowSave
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    VSFG.SaveGrid txtArchivo.Text, flexFileTabText, True
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Ins = Nothing
End Sub

Private Sub menElimina_Click()
    VSFG.RemoveItem VSFG.Row
    Enumerar
End Sub

Private Sub VSFG_AfterSort(ByVal Col As Long, Order As Integer)
    Enumerar
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = 2 And VSFG.MouseCol = 0 And VSFG.MouseCol <> VSFG.MouseRow Then
        VSFG.Row = VSFG.MouseRow
        PopupMenu Eliminar
    End If
End Sub

Private Sub Enumerar()
    Dim i As Long
        VSFG.TextMatrix(0, 0) = "Item"
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub RevisarColumnas()
    Dim i As Long
    Dim strSql As String
    Dim strValAux As String
    Dim strValCambio As String
    VSFG.Select 1, 1
    VSFG.Sort = flexSortGenericAscending
    Enumerar
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) = "" Or VSFG.TextMatrix(i, 3) = "" Then
            MsgBox "Tiene vacio un campo en el item " & VSFG.TextMatrix(i, 0)
            Exit Sub
        End If
    Next i
    strSql = " SELECT prd_codigo,prd_nombre " & _
             " FROM producto " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY prd_codigo"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 2) <> "" Then
            clsCon_Def.Filtrar "prd_codigo='" & VSFG.TextMatrix(i, 2) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                VSFG.TextMatrix(i, 2) = InputBox("El código " & VSFG.TextMatrix(i, 2) & _
                "no existe." & vbNewLine & "Ingrese un código correcto", "Codigo de Producto", VSFG.TextMatrix(i, 2))
            End If
        End If
    Next i
    VSFG.Select 1, 4
    VSFG.Sort = flexSortGenericAscending
    strSql = " SELECT uni_codigo,uni_nombre " & _
             " FROM unidad " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY uni_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "uni_codigo='" & VSFG.TextMatrix(i, 4) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            clsCon_Def.Filtrar "uni_nombre='" & Left(VSFG.TextMatrix(i, 4), 20) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                If strValAux <> Left(VSFG.TextMatrix(i, 4), 20) Then
                    MsgBox "tiene mal la unidad del item " & VSFG.TextMatrix(i, 0)
                    strValAux = Left(VSFG.TextMatrix(i, 4), 20)
                    Set frmSelCaracteristica.clsCaracteristica.adorec_Def = clsCon_Def.adorec_Def
                    frmSelCaracteristica.dcmbNombre.ListField = "uni_nombre"
                    frmSelCaracteristica.dcmbNombre.BoundColumn = "uni_codigo"
                    frmSelCaracteristica.Caption = "Unidad de Medida"
                    frmSelCaracteristica.lblCodigo = VSFG.TextMatrix(i, 1)
                    frmSelCaracteristica.lblNombre = VSFG.TextMatrix(i, 3)
                    frmSelCaracteristica.lblDato = Left(VSFG.TextMatrix(i, 4), 20)
                    frmSelCaracteristica.lblCaracteristica = "Unidad de Medida:"
                    frmSelCaracteristica.Show vbModal
                    strValCambio = Me.Tag
                    VSFG.TextMatrix(i, 4) = strValCambio
                Else
                    VSFG.TextMatrix(i, 4) = strValCambio
                End If
            Else
                VSFG.TextMatrix(i, 4) = clsCon_Def.adorec_Def("uni_codigo")
            End If
        End If
    Next i
    strValAux = ""
    strValCambio = ""
    VSFG.Select 1, 5
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN UNIDAD DE MEDIDA"
    strSql = " SELECT lin_codigo,lin_nombre " & _
             " FROM linea " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY lin_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "lin_codigo='" & VSFG.TextMatrix(i, 5) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            clsCon_Def.Filtrar "lin_nombre='" & Left(VSFG.TextMatrix(i, 5), 20) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                If strValAux <> Left(VSFG.TextMatrix(i, 5), 20) Then
                    MsgBox "tiene mal la linea del item " & VSFG.TextMatrix(i, 0)
                    strValAux = Left(VSFG.TextMatrix(i, 5), 20)
                    Set frmSelCaracteristica.clsCaracteristica.adorec_Def = clsCon_Def.adorec_Def
                    frmSelCaracteristica.dcmbNombre.ListField = "lin_nombre"
                    frmSelCaracteristica.dcmbNombre.BoundColumn = "lin_codigo"
                    frmSelCaracteristica.Caption = "Línea"
                    frmSelCaracteristica.lblCodigo = VSFG.TextMatrix(i, 1)
                    frmSelCaracteristica.lblNombre = VSFG.TextMatrix(i, 3)
                    frmSelCaracteristica.lblDato = Left(VSFG.TextMatrix(i, 5), 20)
                    frmSelCaracteristica.lblCaracteristica = "Línea:"
                    frmSelCaracteristica.Show vbModal
                    strValCambio = Me.Tag
                    VSFG.TextMatrix(i, 5) = strValCambio
                Else
                    VSFG.TextMatrix(i, 5) = strValCambio
                End If
            Else
                VSFG.TextMatrix(i, 5) = clsCon_Def.adorec_Def("lin_codigo")
            End If
        End If
    Next i
    strValAux = ""
    strValCambio = ""
    VSFG.Select 1, 6
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN LINEA"
    strSql = " SELECT mar_codigo,mar_nombre " & _
             " FROM marca " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY mar_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "mar_codigo='" & VSFG.TextMatrix(i, 6) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            clsCon_Def.Filtrar "mar_nombre='" & Left(VSFG.TextMatrix(i, 6), 20) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                If strValAux <> Left(VSFG.TextMatrix(i, 6), 20) Then
                    MsgBox "tiene mal la linea del item " & VSFG.TextMatrix(i, 0)
                    strValAux = Left(VSFG.TextMatrix(i, 6), 20)
                    Set frmSelCaracteristica.clsCaracteristica.adorec_Def = clsCon_Def.adorec_Def
                    frmSelCaracteristica.dcmbNombre.ListField = "mar_nombre"
                    frmSelCaracteristica.dcmbNombre.BoundColumn = "mar_codigo"
                    frmSelCaracteristica.Caption = "Marca"
                    frmSelCaracteristica.lblCodigo = VSFG.TextMatrix(i, 1)
                    frmSelCaracteristica.lblNombre = VSFG.TextMatrix(i, 3)
                    frmSelCaracteristica.lblDato = Left(VSFG.TextMatrix(i, 6), 20)
                    frmSelCaracteristica.lblCaracteristica = "Marca:"
                    frmSelCaracteristica.Show vbModal
                    strValCambio = Me.Tag
                    VSFG.TextMatrix(i, 6) = strValCambio
                Else
                    VSFG.TextMatrix(i, 6) = strValCambio
                End If
            Else
                VSFG.TextMatrix(i, 6) = clsCon_Def.adorec_Def("mar_codigo")
            End If
        End If
    Next i
    strValAux = ""
    strValCambio = ""
    VSFG.Select 1, 7
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN MARCA"
    strSql = " SELECT gru_codigo,gru_nombre " & _
             " FROM grupo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND gru_nivel=1" & _
             " ORDER BY gru_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "gru_codigo='" & VSFG.TextMatrix(i, 7) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            clsCon_Def.Filtrar "gru_nombre='" & Left(VSFG.TextMatrix(i, 7), 20) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                If strValAux <> Left(VSFG.TextMatrix(i, 7), 20) Then
                    MsgBox "tiene mal el grupo del item " & VSFG.TextMatrix(i, 0)
                    strValAux = Left(VSFG.TextMatrix(i, 7), 20)
                    Set frmSelCaracteristica.clsCaracteristica.adorec_Def = clsCon_Def.adorec_Def
                    frmSelCaracteristica.dcmbNombre.ListField = "gru_nombre"
                    frmSelCaracteristica.dcmbNombre.BoundColumn = "gru_codigo"
                    frmSelCaracteristica.Caption = "Grupo"
                    frmSelCaracteristica.lblCodigo = VSFG.TextMatrix(i, 1)
                    frmSelCaracteristica.lblNombre = VSFG.TextMatrix(i, 3)
                    frmSelCaracteristica.lblDato = Left(VSFG.TextMatrix(i, 7), 20)
                    frmSelCaracteristica.lblCaracteristica = "Grupo:"
                    frmSelCaracteristica.Show vbModal
                    strValCambio = Me.Tag
                    VSFG.TextMatrix(i, 7) = strValCambio
                Else
                    VSFG.TextMatrix(i, 7) = strValCambio
                End If
            Else
                VSFG.TextMatrix(i, 7) = clsCon_Def.adorec_Def("gru_codigo")
            End If
        End If
    Next i
    strValAux = ""
    strValCambio = ""
    VSFG.Select 1, 7, 1, 8
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN GRUPO"
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 7) <> VSFG.TextMatrix(i - 1, 7) Then
            strSql = " SELECT gru_codigo,gru_nombre " & _
                     " FROM grupo " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND gru_nivel=2" & _
                     " AND gru_codigo LIKE '" & VSFG.TextMatrix(i, 7) & "%'" & _
                     " ORDER BY LEFT(gru_codigo,2),gru_nombre"
            clsCon_Def.Ejecutar strSql
        End If
        clsCon_Def.Filtrar "gru_codigo='" & VSFG.TextMatrix(i, 8) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            clsCon_Def.Filtrar "gru_nombre='" & Left(VSFG.TextMatrix(i, 8), 20) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                If strValAux <> Left(VSFG.TextMatrix(i, 8), 20) Then
                    MsgBox "tiene mal el subgrupo del item " & VSFG.TextMatrix(i, 0)
                    strValAux = Left(VSFG.TextMatrix(i, 8), 20)
                    Set frmSelCaracteristica.clsCaracteristica.adorec_Def = clsCon_Def.adorec_Def
                    frmSelCaracteristica.dcmbNombre.ListField = "gru_nombre"
                    frmSelCaracteristica.dcmbNombre.BoundColumn = "gru_codigo"
                    frmSelCaracteristica.Caption = "SubGrupo"
                    frmSelCaracteristica.Tag = VSFG.TextMatrix(i, 7)
                    frmSelCaracteristica.lblCodigo = VSFG.TextMatrix(i, 1)
                    frmSelCaracteristica.lblNombre = VSFG.TextMatrix(i, 3)
                    frmSelCaracteristica.lblDato = Left(VSFG.TextMatrix(i, 8), 20)
                    frmSelCaracteristica.lblCaracteristica = "SubGrupo:"
                    frmSelCaracteristica.Show vbModal
                    strValCambio = Me.Tag
                    VSFG.TextMatrix(i, 8) = strValCambio
                    
                Else
                    VSFG.TextMatrix(i, 8) = strValCambio
                End If
            Else
                VSFG.TextMatrix(i, 8) = clsCon_Def.adorec_Def("gru_codigo")
            End If
        End If
    Next i
    VSFG.Select 1, 1
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN SUBGRUPO"
End Sub

Private Sub codigoAnombre()
    Dim i As Long
    Dim strSql As String
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    VSFG.Select 1, 1
    VSFG.Sort = flexSortGenericAscending
    Enumerar
    VSFG.Select 1, 4
    VSFG.Sort = flexSortGenericAscending
    strSql = " SELECT uni_codigo,uni_nombre " & _
             " FROM unidad " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY uni_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "uni_codigo='" & VSFG.TextMatrix(i, 4) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(i, 4) = "NO DEFINIDO"
        Else
            VSFG.TextMatrix(i, 4) = clsCon_Def.adorec_Def("uni_nombre")
        End If
        
    Next i
    VSFG.Select 1, 5
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN UNIDAD DE MEDIDA"
    strSql = " SELECT lin_codigo,lin_nombre " & _
             " FROM linea " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY lin_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "lin_codigo='" & VSFG.TextMatrix(i, 5) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(i, 5) = "NO DEFINIDO"
        Else
            VSFG.TextMatrix(i, 5) = clsCon_Def.adorec_Def("lin_nombre")
        End If
    Next i
    VSFG.Select 1, 6
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN LINEA"
    strSql = " SELECT mar_codigo,mar_nombre " & _
             " FROM marca " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY mar_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "mar_codigo='" & VSFG.TextMatrix(i, 6) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(i, 6) = "NO DEFINIDO"
        Else
            VSFG.TextMatrix(i, 6) = clsCon_Def.adorec_Def("mar_nombre")
        End If
    Next i
    strValAux = ""
    strValCambio = ""
    VSFG.Select 1, 7, 1, 8
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN MARCA"
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 7) <> VSFG.TextMatrix(i - 1, 7) Then
            strSql = " SELECT gru_codigo,gru_nombre " & _
                     " FROM grupo " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND gru_nivel=2" & _
                     " AND gru_codigo LIKE '" & VSFG.TextMatrix(i, 7) & "%'" & _
                     " ORDER BY LEFT(gru_codigo,2),gru_nombre"
            clsCon_Def.Ejecutar strSql
        End If
        clsCon_Def.Filtrar "gru_codigo='" & VSFG.TextMatrix(i, 8) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(i, 8) = "NO DEFINIDO"
        Else
            VSFG.TextMatrix(i, 8) = clsCon_Def.adorec_Def("gru_nombre")
        End If
    Next i
    VSFG.Select 1, 7
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN SUBGRUPO"
    strSql = " SELECT gru_codigo,gru_nombre " & _
             " FROM grupo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND gru_nivel=1" & _
             " ORDER BY gru_nombre"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        clsCon_Def.Filtrar "gru_codigo='" & VSFG.TextMatrix(i, 7) & "'"
        If clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(i, 7) = "NO DEFINIDO"
        Else
            VSFG.TextMatrix(i, 7) = clsCon_Def.adorec_Def("gru_nombre")
        End If
    Next i
    VSFG.Select 1, 1
    VSFG.Sort = flexSortGenericAscending
    MsgBox "FIN GRUPO"
End Sub


Private Sub CambioCodigosProductos()
    Dim i As Long
    Dim strSql As String
'Inserta productos nuevos
    clsCon_Ins.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT prd_codigo,prd_nombre " & _
             " FROM producto " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY prd_codigo"
    clsCon_Def.Ejecutar strSql
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) <> "" Then
            clsCon_Def.Filtrar "prd_codigo='" & VSFG.TextMatrix(i, 1) & "'"
            If clsCon_Def.adorec_Def.EOF Then
                ' Almacenamiento de los datos del nuevo producto
                strSql = " INSERT INTO producto(prd_codigo,uni_codigo,emp_codigo,gru_codigo " & _
                         ",mar_codigo, lin_codigo,prd_nombre,prd_costo,prd_baja,prd_fechamod,prd_usumod) " & _
                         " VALUES ('" & UCase(VSFG.TextMatrix(i, 1)) & "','" & VSFG.TextMatrix(i, 4) & "','" & strEmpresa & _
                         "','" & VSFG.TextMatrix(i, 8) & _
                         "','" & VSFG.TextMatrix(i, 6) & "','" & VSFG.TextMatrix(i, 5) & _
                         "','" & VSFG.TextMatrix(i, 3) & "',0,0" & _
                         ",CURRENT_TIMESTAMP, '" & strUsuario & "')"
                clsCon_Ins.Ejecutar strSql, "M"
            ' Almacenamiento de los datos del nuevo producto en las listas de precios
                strSql = " INSERT INTO lista_precio_p " & _
                         " SELECT lis_pre_codigo, '" & UCase(VSFG.TextMatrix(i, 1)) & "', emp_codigo, " & _
                         " 0/(1-lis_pre_politica/100), " & _
                         " lis_pre_politica,CURRENT_TIMESTAMP, '" & strUsuario & "' " & _
                         " FROM lista_precio WHERE emp_codigo='" & strEmpresa & "' "
                clsCon_Ins.Ejecutar strSql, "M"
            ' Almacenamiento de los datos del nuevo producto en los depositos
                strSql = " INSERT INTO existencia " & _
                         " SELECT '" & UCase(VSFG.TextMatrix(i, 1)) & "',dep_codigo, emp_codigo,'" & Trim(Year(HoyDia)) & "-" & Trim(Month(HoyDia)) & "-" & Trim(Day(HoyDia)) & "', " & _
                         " 0, CURRENT_TIMESTAMP, '" & strUsuario & "' " & _
                         " FROM deposito WHERE emp_codigo='" & strEmpresa & "' "
                clsCon_Ins.Ejecutar strSql, "M"
                'actualiza productos
                strSql = " SELECT prd_codigo,prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " ORDER BY prd_codigo"
                clsCon_Def.Ejecutar strSql
            Else
                strSql = " UPDATE producto " & _
                         " SET uni_codigo='" & VSFG.TextMatrix(i, 4) & "', " & _
                         " gru_codigo='" & VSFG.TextMatrix(i, 8) & "', " & _
                         " mar_codigo='" & VSFG.TextMatrix(i, 6) & "', " & _
                         " lin_codigo='" & VSFG.TextMatrix(i, 5) & "', " & _
                         " prd_nombre='" & VSFG.TextMatrix(i, 3) & "', " & _
                         " prd_fechamod=CURRENT_TIMESTAMP, " & _
                         " prd_usumod='" & strUsuario & "' " & _
                         " WHERE prd_codigo='" & UCase(VSFG.TextMatrix(i, 1)) & "' " & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Ins.Ejecutar strSql, "M"
            End If
        End If
    Next i
    MsgBox "FIN INSERT UPDATE"
'cambio de historial
    For i = 1 To VSFG.Rows - 1
        Me.Caption = "Cambio de Codigos " & Format(i / VSFG.Rows * 100, "##0") & "%"
        'Me.Refresh
        If VSFG.TextMatrix(i, 1) <> "" And VSFG.TextMatrix(i, 2) <> "" Then
            'BACKORDER
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_backorder " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_backorder " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_backorder " & _
                     " SELECT emp_codigo,'" & VSFG.TextMatrix(i, 1) & "',bac_codigo,COALESCE(SUM(det_bac_cantidad),0),COALESCE(SUM(det_bac_cantidad*det_bac_precio),0)/SUM(det_bac_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,bac_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql
            'COTIZACION
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_cotizacion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_cotizacion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_cotizacion " & _
                     " SELECT cot_codigo,'" & VSFG.TextMatrix(i, 1) & "',emp_codigo,COALESCE(SUM(det_cot_cantidad),0),COALESCE(SUM(det_cot_cantidad*det_cot_precio),0)/SUM(det_cot_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,cot_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'EGRESO
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_egreso " & _
                     " SELECT emp_codigo,egr_codigo,tip_egr_codigo,'" & VSFG.TextMatrix(i, 1) & "',dep_codigo,COALESCE(SUM(det_egr_cantidad),0),COALESCE(SUM(det_egr_cantidad*det_egr_precio),0)/SUM(det_egr_cantidad),COALESCE(SUM(det_egr_cantidad*det_egr_costo),0)/SUM(det_egr_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,egr_codigo,tip_egr_codigo,dep_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'INGRESO
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_ingreso " & _
                     " SELECT emp_codigo,ing_codigo,tip_ing_codigo,'" & VSFG.TextMatrix(i, 1) & "',dep_codigo,COALESCE(SUM(det_ing_cantidad),0),COALESCE(SUM(det_ing_cantidad*det_ing_precio),0)/SUM(det_ing_cantidad),COALESCE(SUM(det_ing_cantidad*det_ing_costo),0)/SUM(det_ing_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,ing_codigo,tip_ing_codigo,dep_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'INGRESO DE IMPORTACION
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_ingreso_imp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_ingreso_imp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_ingreso_imp " & _
                     " SELECT emp_codigo,ped_imp_codigo,'" & VSFG.TextMatrix(i, 1) & "',ing_codigo,tip_ing_codigo,COALESCE(SUM(det_ing_imp_cantidad),0),COALESCE(SUM(det_ing_imp_cantidad*det_ing_imp_fob),0)/SUM(det_ing_imp_cantidad),COALESCE(SUM(det_ing_imp_cantidad*det_ing_imp_cif),0)/SUM(det_ing_imp_cantidad),COALESCE(SUM(det_ing_imp_cantidad*det_ing_imp_costofinal),0)/SUM(det_ing_imp_cantidad),COALESCE(SUM(det_ing_imp_cantidad*det_ing_imp_cif*det_ing_imp_arancel),0)/SUM(det_ing_imp_cantidad*det_ing_imp_cif),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,ped_imp_codigo,ing_codigo,tip_ing_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'PEDIDO
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_pedido " & _
                     " SELECT emp_codigo,ped_codigo,'" & VSFG.TextMatrix(i, 1) & "',dep_codigo,COALESCE(SUM(det_ped_cant_pedida),0),COALESCE(SUM(det_ped_cant_entregada),0),COALESCE(SUM(det_ped_cant_entregada*det_ped_precio),0)/SUM(det_ped_cant_entregada),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,ped_codigo,dep_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'PEDIDO IMPORTACION
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_pedido_imp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_pedido_imp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_pedido_imp " & _
                     " SELECT ped_imp_codigo,emp_codigo,'" & VSFG.TextMatrix(i, 1) & "',COALESCE(SUM(det_ped_imp_cantidad),0),COALESCE(SUM(det_ped_imp_cantidad*det_ped_imp_precio),0)/SUM(det_ped_imp_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,ped_imp_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'PRODUCTO COMPUESTO
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM det_prd_com " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM det_prd_com " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO det_prd_com " & _
                     " SELECT emp_codigo,prd_com_codigo,'" & VSFG.TextMatrix(i, 1) & "',COALESCE(SUM(det_prd_com_cantidad),0),COALESCE(SUM(det_prd_com_cantidad*det_prd_com_costo),0)/SUM(det_prd_com_cantidad),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY emp_codigo,prd_com_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'EXISTENCIA
            strSql = " CREATE TEMPORARY TABLE Temp " & _
                     " SELECT * FROM existencia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DELETE FROM existencia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (prd_codigo='" & VSFG.TextMatrix(i, 2) & "' OR prd_codigo='" & VSFG.TextMatrix(i, 1) & "') "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " INSERT INTO existencia " & _
                     " SELECT '" & VSFG.TextMatrix(i, 1) & "',dep_codigo,emp_codigo,COALESCE(SUM(exi_cantidad),0),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
                     " FROM Temp " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " GROUP BY dep_codigo,emp_codigo "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " DROP TABLE Temp "
            clsCon_Ins.Ejecutar strSql, "M"
            'LISTA DE PRECIO
            strSql = " DELETE FROM lista_precio_p " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' "
            clsCon_Ins.Ejecutar strSql, "M"
            strSql = " UPDATE lista_precio_p " & _
                     " SET prd_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 2) & "' "
            clsCon_Ins.Ejecutar strSql, "M"
            'LISTA DE PRECIO PROVEEDOR
            strSql = " UPDATE persona_producto " & _
                     " SET prd_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 2) & "' "
            clsCon_Ins.Ejecutar strSql, "M"
            'ELIMINA PRODUCTO
            strSql = " DELETE FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 2) & "' "
            clsCon_Ins.Ejecutar strSql, "M"
        End If
    Next i
    MsgBox "FIN TODO"
    Me.Caption = "Cambio de Codigos "
End Sub
