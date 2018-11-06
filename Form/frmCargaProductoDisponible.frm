VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCargaProductoDisponible 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Listas de Precio"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaProductoDisponible.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listas de Precio"
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8640
      Begin VB.OptionButton op_coleccion 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Colección"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6720
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton op_lista 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Lista"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   8295
         _cx             =   1963407655
         _cy             =   1963402575
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaProductoDisponible.frx":030A
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
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         Top             =   840
         Width           =   2640
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   8040
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dcmbDescripcion 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   7560
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo dcmbColeccion 
         Height          =   330
         Left            =   1080
         TabIndex        =   9
         Top             =   840
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   615
         Left            =   5400
         TabIndex        =   13
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colección:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4680
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4455
      TabIndex        =   1
      Top             =   7440
      Width           =   1455
   End
End
Attribute VB_Name = "frmCargaProductoDisponible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
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

Private clsCon_Def As New clsConsulta
Private strSql As String
Dim i_flag As Integer

Private Sub cmdAplicar_Click()

If (i_flag = 1) Then
    Dim i As Long
    Me.MousePointer = 11
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, 5)) = 1 Then
            strSql = " REPLACE INTO lista_precio_p " & _
                     " (lis_pre_p_precio,lis_pre_p_politica,lis_pre_p_fechamod,lis_pre_p_usumod,emp_codigo,lis_pre_codigo,prd_codigo) VALUES('" & VSFG.TextMatrix(i, 1) & "', " & _
                     " '" & VSFG.TextMatrix(i, 4) & "' ," & _
                     " CURRENT_TIMESTAMP ," & _
                     " '" & strUsuario & "' " & _
                     " ,'" & strEmpresa & "'" & _
                     " ,'" & dcmbDescripcion.BoundText & "' " & _
                     " ,'" & VSFG.TextMatrix(i, 0) & "')"
            clsCon_Def.Ejecutar strSql, "M"
        End If
    Next i
    Me.MousePointer = 0
    MsgBox "Carga Finalizada", vbInformation, "Lista de Precio"
    Unload Me
End If

If (i_flag = 0) Then
    Dim i_o As Long
    Dim j As Long
    Me.MousePointer = 11
    j = 1
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, 2)) = 1 Then
            strSql = " Update producto set " & _
                     " clc_codigo = " & _
                     " '" & dcmbColeccion.BoundText & "', prd_usumod='" & strUsuario & "',prd_fechamod=CURRENT_TIMESTAMP" & _
                     " WHERE prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql, "M"
            j = j + 1
        End If
    Next i
    Me.MousePointer = 0
    MsgBox "Cambio de Colección realizado a " & j & " items", 64, "Colección"
    Unload Me
End If
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    VSFG.Clear
    'Lee archivo para cargar lista de precio
    If (txtArchivo.Text <> "" And i_flag = 1) Then
        Me.MousePointer = 11
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 2
        VSFG.Cols = 6
        VSFG.TextMatrix(0, 2) = "Producto"
        VSFG.TextMatrix(0, 3) = "Costo"
        VSFG.TextMatrix(0, 4) = "Plotica"
        VSFG.TextMatrix(0, 5) = "Aplicar"
        VSFG.ColWidth(1) = 1200
        VSFG.ColWidth(2) = 3000
        For i = 1 To VSFG.Rows - 1
            strSql = " SELECT prd_nombre,prd_costo " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                     
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 2) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 3) = clsCon_Def.adorec_Def("prd_costo")
                If Val(Format(VSFG.TextMatrix(i, 3), "###0.00")) <> 0 Then
                    VSFG.TextMatrix(i, 4) = 100 - (Val(Format(VSFG.TextMatrix(i, 1), "###0.00")) / Val(Format(VSFG.TextMatrix(i, 3), "###0.00"))) * 100
                Else
                    VSFG.TextMatrix(i, 4) = 100
                End If
                VSFG.TextMatrix(i, 5) = 1
            Else
                VSFG.TextMatrix(i, 2) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 3) = 0
                VSFG.TextMatrix(i, 4) = 100
                VSFG.TextMatrix(i, 5) = 0
            End If
        Next i
        Me.MousePointer = 0
    
    Else
    
       'Lee archivo para cambiar la colección
        Me.MousePointer = 11
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 3
        VSFG.TextMatrix(0, 0) = "Código Producto"
        VSFG.TextMatrix(0, 1) = "Descripción"
        VSFG.TextMatrix(0, 2) = "Estado"
        VSFG.ColWidth(1) = 3500
        For i = 1 To VSFG.Rows - 1
            strSql = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 2) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 2) = 0
                VSFG.Cell(flexcpBackColor, i, 0, i, 2) = &HC0C0FF
            End If
        Next i
        Me.MousePointer = 0
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    op_lista.value = True
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        strSql = " SELECT CONCAT(lis_pre_codigo) AS lis_pre_codigo,lis_pre_descripcion " & _
                 " FROM lista_precio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lis_pre_codigo "
        clsCon_Def.Ejecutar strSql
        Set dcmbDescripcion.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbDescripcion.ListField = "lis_pre_descripcion"
        dcmbDescripcion.BoundColumn = "lis_pre_codigo"
     
     'Consulta las colecciones existentes
        strSql = " SELECT CONCAT(clc_codigo) AS clc_codigo,clc_nombre " & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_codigo "
        clsCon_Def.Ejecutar strSql
        Set dcmbColeccion.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbColeccion.ListField = "clc_nombre"
        dcmbColeccion.BoundColumn = "clc_codigo"
        Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub op_coleccion_Click()
       i_flag = 0
       dcmbColeccion.Enabled = True
       dcmbDescripcion.Enabled = False
End Sub

Private Sub op_lista_Click()
       i_flag = 1
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = True
End Sub
