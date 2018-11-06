VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaPrd 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsultaPrd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   9630
   Begin MSDataListLib.DataCombo dcmbUnidad 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VSFlex8LCtl.VSFlexGrid vsfgGrupo 
      Height          =   1335
      Left            =   6120
      TabIndex        =   3
      Tag             =   "F"
      Top             =   120
      Width           =   3255
      _cx             =   5741
      _cy             =   2355
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
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsultaPrd.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin MSDataListLib.DataCombo dcmbLinea 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1575
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3225
      TabIndex        =   4
      Top             =   1575
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dcmbMarca 
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Línea:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   660
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   285
      TabIndex        =   7
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos de Producto:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmConsultaPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Productos con los que trabajará la       #
'#  empresa.                                                                    #
'#  frmProducto V1.0                                                            #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de los productos.                   #
'#  Permitirá almacenar en la base de datos nuevos productos y modificar        #
'#  sus datos, esto dependiendo de la propiedad Tag, la cual se cambiará en la  #
'#  ventana frmSelProducto y desde esta se llamará a esta ventana.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     producto: En esta tabla se almacenan los nuevos productos y se           #
'#               modifican los datos de estos.                                  #
'#        grupo: En esta tabla se sacan los grupos a los que puede pertenecer   #
'#               los diferentes productos en sus diferentes niveles.            #
'#       marca : En esta tabla se sacan las marcas a las que se puede asignar a #
'#               los diferentes productos.                                      #
'#       linea : En esta tabla se sacan las lineas a las que se puede asignar a #
'#               los diferentes productos.                                      #
'#      unidad : En esta tabla se sacan las unidades de medida que se puede     #
'#               asignar los productos.                                         #
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

Private clsCon_Def As clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim strSql As String
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSql = " DROP TABLE IF EXISTS Temp"
    clsCon_Def.Ejecutar strSql, "M"
    Set clsCon_Def = Nothing
End Sub

Private Sub cmbAceptar_Click()
    frmCambioProductos.strGrupo = IIf(vsfgGrupo.TextMatrix(1, 1) = "", IIf(vsfgGrupo.TextMatrix(0, 1) = "", "%", vsfgGrupo.TextMatrix(0, 1) & ".%"), vsfgGrupo.TextMatrix(1, 1))
    frmCambioProductos.strLinea = dcmbLinea.BoundText
    frmCambioProductos.strMarca = dcmbMarca.BoundText
    frmCambioProductos.strUniMed = dcmbUnidad.BoundText
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las unidades de medida que estan disponibles
    strSql = " CREATE TEMPORARY TABLE Temp " & _
             " SELECT uni_codigo,uni_nombre " & _
             " FROM unidad " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY uni_nombre "
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " INSERT INTO Temp VALUES('%','TODAS')"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT * FROM Temp ORDER BY uni_codigo"
    clsCon_Def.Ejecutar strSql, "M"
    dcmbUnidad.ListField = "uni_nombre"
    dcmbUnidad.BoundColumn = "uni_codigo"
    Set dcmbUnidad.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbUnidad.BoundText = "%"
    strSql = " DROP TABLE Temp "
    clsCon_Def.Ejecutar strSql, "M"
    'Consulta las lineas que estan disponibles
    strSql = " CREATE TEMPORARY TABLE Temp " & _
             " SELECT lin_codigo,lin_nombre " & _
             " FROM linea " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY lin_nombre "
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " INSERT INTO Temp VALUES('%','TODOS')"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT * FROM Temp ORDER BY lin_codigo "
    clsCon_Def.Ejecutar strSql, "M"
    dcmbLinea.ListField = "lin_nombre"
    dcmbLinea.BoundColumn = "lin_codigo"
    Set dcmbLinea.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbLinea.BoundText = "%"
    strSql = " DROP TABLE Temp "
    clsCon_Def.Ejecutar strSql, "M"
    'Consulta las marcas que estan disponibles
    strSql = " CREATE TEMPORARY TABLE Temp " & _
             " SELECT mar_codigo,mar_nombre " & _
             " FROM marca " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY mar_nombre "
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " INSERT INTO Temp VALUES('%','TODOS')"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT * FROM Temp ORDER BY mar_codigo"
    clsCon_Def.Ejecutar strSql, "M"
    dcmbMarca.ListField = "mar_nombre"
    dcmbMarca.BoundColumn = "mar_codigo"
    Set dcmbMarca.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbMarca.BoundText = "%"
    strSql = " DROP TABLE Temp "
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT max(gru_nivel) as m FROM grupo " & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    vsfgGrupo.Rows = clsCon_Def.adorec_Def("m")
End Sub

Private Sub vsfgGrupo_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    ' Si se cambia el contenido de la celda con los combos
    If vsfgGrupo.Tag = "T" Then
        ' Mueve el recordset la linea seleccionada y escribe el codio en el vsflexgrid
        If vsfgGrupo.ComboIndex >= 0 Then
            clsCon_Def.adorec_Def.Move vsfgGrupo.ComboIndex, adBookmarkFirst
        End If
        If Not clsCon_Def.adorec_Def.EOF Then
            vsfgGrupo.TextMatrix(Row, 1) = clsCon_Def.adorec_Def("gru_codigo")
            For i = Row + 1 To vsfgGrupo.Rows - 1
                vsfgGrupo.TextMatrix(i, 0) = ""
                vsfgGrupo.TextMatrix(i, 1) = ""
            Next i
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub vsfgGrupo_GotFocus()
    vsfgGrupo_RowColChange
End Sub

Private Sub vsfgGrupo_RowColChange()
    Dim strCod As String
    Dim booBand As Boolean
    Dim Row As Long
    Dim strSql As String
    Row = vsfgGrupo.Row
    If Row < 0 Then
        Row = 0
    End If
    ' Si edita el grid por primera ves pone tag=T antes era tag=F
    vsfgGrupo.Tag = "T"
    booBand = False
    ' Si selecciono el combo de la primera fila
    If Row = 0 Then
        strCod = "%"
        booBand = True
    Else ' Si selecciono el combo de las filas 2 en adelante
        If vsfgGrupo.TextMatrix(Row - 1, 1) <> "" Then
            strCod = vsfgGrupo.TextMatrix(Row - 1, 1) & ".%"
            booBand = True
        Else
            booBand = False
        End If
    End If
    If booBand = True Then
        strSql = " DROP TABLE IF EXISTS Temp"
        clsCon_Def.Ejecutar strSql, "M"
        ' Consulta para conocer los grupos que podría seleccionar
        strSql = " CREATE TEMPORARY TABLE Temp " & _
                 " SELECT gru_codigo,gru_nombre " & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND gru_nivel=" & (Row + 1) & _
                 " AND gru_codigo LIKE '" & strCod & "' " & _
                 " ORDER BY gru_nombre "
        clsCon_Def.Ejecutar strSql, "M"
        strSql = " INSERT INTO Temp VALUES('" & strCod & "','TODOS')"
        clsCon_Def.Ejecutar strSql, "M"
        strSql = " SELECT * FROM Temp ORDER BY gru_codigo"
        clsCon_Def.Ejecutar strSql, "M"
        If Col = 0 Then
            ' Genera el combo de los grupos
            vsfgGrupo.ColComboList(0) = vsfgGrupo.BuildComboList(clsCon_Def.adorec_Def, "*gru_nombre,gru_codigo", "gru_nombre")
        End If
    Else
        Cancel = True
    End If
    SendKeys vbKeySpace
End Sub
