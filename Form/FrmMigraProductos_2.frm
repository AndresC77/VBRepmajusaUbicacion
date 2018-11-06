VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmMigraProductos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migración de Productos"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "FrmMigraProductos_2.frx":0000
   LinkTopic       =   "Formulario migracion de Productos"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13860
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CmbVerificar 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton bb1 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar Pgb 
      Height          =   255
      Left            =   10800
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdMigraexcel 
      Caption         =   "Migra Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Seleccionar Archivo a Migrar"
      Top             =   120
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13620
      _cx             =   59596248
      _cy             =   59583442
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
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMigraProductos_2.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Sku Duplicado"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Marca error"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "OK"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9960
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Codigo Duplicado"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMigraProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Private strSql As String

Private Sub CmbVerificar_Click()
Call Identificar
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo salirgrabar
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos
    
    Call Identificar
    
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            strSql = " INSERT INTO producto(emp_codigo,prd_codigo,prd_nombre,mar_codigo,lin_codigo,gru_codigo,uni_codigo,tal_codigo,col_codigo,clc_codigo,prd_baja," & _
                     " prd_costo,prd_cambia_precio,prd_iva,prd_sku,prd_fechamod,prd_usumod,PRD_NO_COMISION) " & _
                     " VALUES ('" & strEmpresa & "','" & UCase(VSFG.TextMatrix(i, 1)) & "','" & VSFG.TextMatrix(i, 2) & "','" & UCase(VSFG.TextMatrix(i, 3)) & "','" & UCase(VSFG.TextMatrix(i, 4)) & "', " & _
                     " '" & VSFG.TextMatrix(i, 5) & "','" & VSFG.TextMatrix(i, 6) & "','" & VSFG.TextMatrix(i, 7) & "','" & VSFG.TextMatrix(i, 8) & "','" & VSFG.TextMatrix(i, 9) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 10))) & "','" & FormatoD4(VSFG.TextMatrix(i, 13)) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 14))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 15))) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "'," & _
                     " CURRENT_TIMESTAMP, '" & strUsuario & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 17))) & "')"
            clsCon_Def.Ejecutar strSql, "M"
            
            strSql = " INSERT INTO lista_precio_p " & _
                     " SELECT lis_pre_codigo, '" & UCase(VSFG.TextMatrix(i, 1)) & "', emp_codigo, 0," & _
                     " lis_pre_politica,0,0,CURRENT_TIMESTAMP, substring_index(USER(),'@',1) " & _
                     " FROM lista_precio WHERE emp_codigo='" & strEmpresa & "' "
            clsCon_Def.Ejecutar strSql, "M"
            
            strSql = " INSERT INTO existencia " & _
                     " SELECT '" & UCase(VSFG.TextMatrix(i, 1)) & "',dep_codigo, emp_codigo, " & _
                     " 0, CURRENT_TIMESTAMP, substring_index(USER(),'@',1) " & _
                     " FROM deposito WHERE emp_codigo='" & strEmpresa & "' "
            clsCon_Def.Ejecutar strSql, "M"
        End If
    Next i
    Call Identificar
    Exit Sub
salirgrabar:
    MsgBox Err.Description
End Sub


Private Sub CmdMigraexcel_Click()
On Error GoTo SalirMigrar
    Dim sDir As String
    cdArchivo.ShowOpen
    sDir = cdArchivo.FileName
    If Not LeerExcel(sDir) Then Exit Sub
    If Not Identificar Then Exit Sub
    MsgBox ("Importación desde Excel ha Culminado con exito.")
    Exit Sub
SalirMigrar:
    MsgBox "La Importacion desde Excel ha Fallado, " & Err.Description
End Sub
Function LeerExcel(Archivo As String) As Boolean
On Error GoTo SalirExcel
'dimensiones
LeerExcel = False
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim lngUltimaFila As Long, Fil As Long, Col As Long

Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

'abrir programa Excel
Set xlApp = New Excel.Application
xlApp.Visible = False

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(Archivo, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)

'2. Si no conoces el rango
lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

'lngUltimaFila = 17
Pgb.min = 1
Pgb.Max = lngUltimaFila

If MsgBox("Serán Migrados #" & CStr(lngUltimaFila) & " Desea Continuar?", vbYesNo) = vbNo Then Exit Function
Pgb.Visible = True
VSFG.Rows = 1
For Fil = 2 To lngUltimaFila
    VSFG.Rows = Fil
    Pgb.value = Fil
    For Col = 1 To 17
        VSFG.TextMatrix(Fil - 1, Col) = xlHoja.Range(xlHoja.Cells(Fil, Col), xlHoja.Cells(Fil, Col))
    Next Col
Next Fil
Pgb.Visible = False
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing
LeerExcel = True
Exit Function
SalirExcel:
    LeerExcel = False
    MsgBox Err.Description
End Function



Function Identificar() As Boolean
On Error GoTo salirIdentificar
    Dim Fil As Long, Col As Long, i As Long
    Identificar = False
    i = VSFG.Rows
    VSFG.Cols = VSFG.Cols + 1
    VSFG.ColHidden(VSFG.Cols - 1) = True
    
    For Fil = 1 To i - 1
            strSql = " SELECT prd_nombre,prd_costo " & _
                 " FROM producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND prd_codigo='" & VSFG.TextMatrix(Fil, 1) & "' "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                VSFG.Cell(flexcpBackColor, Fil, 1, Fil, 17) = vbYellow 'RGB(500, 0, 0)
            Else
            
                strSql = " SELECT prd_nombre,prd_costo " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_SKU='" & VSFG.TextMatrix(Fil, 16) & "' "
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                    VSFG.Cell(flexcpBackColor, Fil, 1, Fil, 17) = vbGreen 'RGB(500, 0, 0)
                Else
                
                    VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 3
                    VSFG.Cell(flexcpBackColor, Fil, 1, Fil, 17) = vbWhite
                    Col = 0
                    If VSFG.TextMatrix(Fil, 1) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 1, Fil, 1) = vbCyan
                    End If
                    If VSFG.TextMatrix(Fil, 2) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 2, Fil, 2) = vbCyan
                    End If
                    
                    If VSFG.TextMatrix(Fil, 3) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 3, Fil, 3) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM marca " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND mar_codigo='" & VSFG.TextMatrix(Fil, 3) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 3, Fil, 3) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 4) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 4, Fil, 4) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM linea " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND lin_codigo='" & VSFG.TextMatrix(Fil, 4) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 4, Fil, 4) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 5) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 5, Fil, 5) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM grupo " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND gru_codigo='" & VSFG.TextMatrix(Fil, 5) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 5, Fil, 5) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 6) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 6, Fil, 6) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM unidad " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND uni_codigo='" & VSFG.TextMatrix(Fil, 6) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 6, Fil, 6) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 7) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 7, Fil, 7) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM talla " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND tal_codigo='" & VSFG.TextMatrix(Fil, 7) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 7, Fil, 7) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 8) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 8, Fil, 8) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM COLOR " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND COL_codigo='" & VSFG.TextMatrix(Fil, 8) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 8, Fil, 8) = vbCyan
                        End If
                    End If
                    
                    If VSFG.TextMatrix(Fil, 9) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 9, Fil, 9) = vbCyan
                    Else
                        strSql = " SELECT 1 " & _
                             " FROM COLECCION " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND CLC_codigo='" & VSFG.TextMatrix(Fil, 9) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount = 0 Then
                            VSFG.TextMatrix(Fil, VSFG.Cols - 1) = 2 'No insertar
                            VSFG.Cell(flexcpBackColor, Fil, 9, Fil, 9) = vbCyan
                        End If
                    End If
                    
                    
                    If VSFG.TextMatrix(Fil, 16) = "" Then
                        VSFG.Cell(flexcpBackColor, Fil, 16, Fil, 16) = vbCyan
                    End If
                
                End If
        End If

    Next Fil
    Identificar = True
    Exit Function
salirIdentificar:
    Identificar = False
    MsgBox Err.Description
End Function

Private Sub CmdMigraexcel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'StatusBar.Panels(1).Text = "<Click> Seleccionar el Archivo Excel a Importar."
End Sub
Private Sub CmdGrabar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'StatusBar.Panels(1).Text = "<Click> Grabar Productos."
End Sub
Private Sub CmbVerificar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'StatusBar.Panels(1).Text = "<Click> Varificar Errores antes de Grabar."
End Sub
Private Sub Form_Load()
    'Coneccion a datos
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster

    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    'Referenncia botones a flex
    'Set ucrtVSFG.VSFGControl = VSFG
    'ucrtVSFG.Inicializar
    Cargar
    
End Sub
Private Sub Cargar()
    
    'crea combo de marca
    strSql = " SELECT mar_codigo, mar_nombre" & _
                 " FROM marca " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY mar_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "mar_codigo,*mar_nombre", "mar_codigo")
    
    'crea combo de linea
    strSql = " SELECT lin_codigo, lin_nombre" & _
                 " FROM linea " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lin_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(4) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "lin_codigo,*lin_nombre", "lin_codigo")
    
    'crea combo de grupo
    strSql = " SELECT gru_codigo, CONCAT(REPEAT(' ',(gru_nivel)*2),gru_nombre) as gru_nombre" & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY gru_codigo"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(5) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "gru_codigo,*gru_nombre", "gru_codigo")
    
    'crea combo de unidad de medida
    strSql = " SELECT uni_codigo, uni_nombre" & _
                 " FROM unidad " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY uni_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(6) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "uni_codigo,*uni_nombre", "uni_codigo")
    
    'crea combo de talla
    strSql = " SELECT tal_codigo, tal_nombre" & _
                 " FROM talla " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY tal_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(7) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*tal_nombre", "tal_codigo")
    
    'crea combo de color
    strSql = " SELECT col_codigo, col_nombre" & _
                 " FROM color " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY col_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(8) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "cal_codigo,*col_nombre", "col_codigo")
    
    'crea combo de coleccion
    strSql = " SELECT clc_codigo, clc_nombre" & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_nombre"
     clsCon_Def.Ejecutar strSql
    VSFG.ColComboList(9) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "clc_codigo,*clc_nombre", "clc_codigo")
    

End Sub




