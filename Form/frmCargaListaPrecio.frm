VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCargaListaPrecio 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Listas de Precio"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaListaPrecio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   8880
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   8415
         _cx             =   14843
         _cy             =   7435
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
         FormatString    =   $"frmCargaListaPrecio.frx":030A
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
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   3240
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   7800
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dcmbDescripcion 
         Height          =   330
         Left            =   4560
         TabIndex        =   0
         Top             =   840
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   7320
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo dcmbColeccion 
         Height          =   330
         Left            =   4560
         TabIndex        =   9
         Top             =   1200
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton op_nodisponiblepreventa 
            BackColor       =   &H00DDDDDD&
            Caption         =   "NO Disponible para la PREventa"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1990
            Width           =   2775
         End
         Begin VB.OptionButton op_disponiblePreventa 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Disponible para la PREventa"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   2655
         End
         Begin VB.OptionButton op_nodisponibleventa 
            BackColor       =   &H00DDDDDD&
            Caption         =   "NO Disponible para la venta"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   2295
         End
         Begin VB.OptionButton op_Balanceo 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Balanceo"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   2175
         End
         Begin VB.OptionButton op_disponibleventa 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Disponible para la venta"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton op_lista 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Lista"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton op_coleccion 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Colección"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
      End
      Begin NEED2.dtpFecha dtpFechaInicio 
         Height          =   285
         Left            =   4560
         TabIndex        =   15
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         Value           =   41850.4704976852
      End
      Begin NEED2.dtpFecha dtpFechaFin 
         Height          =   285
         Left            =   4560
         TabIndex        =   16
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         Value           =   41850.4704976852
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3600
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3750
         TabIndex        =   17
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colección"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3705
         TabIndex        =   10
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3885
         TabIndex        =   6
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4110
         TabIndex        =   3
         Top             =   960
         Width           =   345
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
Attribute VB_Name = "frmCargaListaPrecio"
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
Private strSQL As String
Dim i_flag As Integer

Private Sub cmdAplicar_Click()
    Dim i As Long
    Dim j As Long
    Me.MousePointer = 11
    If (i_flag = 1) Then
        
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 5)) = 1 Then
                strSQL = " EXEC SP_lista_precio_p_Mantenimiento " & _
                         " '" & VSFG.TextMatrix(i, 1) & "', " & _
                         " '" & IIf(0 > VSFG.TextMatrix(i, 4), 0, IIf(100 < VSFG.TextMatrix(i, 4), 100, VSFG.TextMatrix(i, 4))) & "' ," & _
                         " " & _
                         " '" & strUsuario & "' " & _
                         " ,'" & strEmpresa & "'" & _
                         " ,'" & dcmbDescripcion.BoundText & "' " & _
                         " ,'" & VSFG.TextMatrix(i, 0) & "',0,0 "
                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Lista de Precio"
    ElseIf (i_flag = 0) Then
        j = 1
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 2)) = 1 Then
                strSQL = " Update producto set " & _
                         " clc_codigo = " & _
                         " '" & dcmbColeccion.BoundText & "', prd_usumod='" & strUsuario & "',prd_fechamod=CURRENT_TIMESTAMP" & _
                         " WHERE prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                j = j + 1
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Cambio de Colección realizado a " & j & " items", 64, "Colección"
    ElseIf (i_flag = 2) Then
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 2)) = 1 Then
                strSQL = " DELETE FROM producto_disponible " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                         " AND pro_dis_fechainicio='" & Format(dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss") & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                strSQL = " INSERT INTO producto_disponible " & _
                         " (emp_codigo,prd_codigo,pro_dis_fechainicio,pro_dis_fechafin,pro_dis_fechamod,pro_dis_usumod) " & _
                         " VALUES('" & strEmpresa & "', " & _
                         " '" & VSFG.TextMatrix(i, 0) & "'," & _
                         " '" & Format(dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss") & "' ," & _
                         " '" & Format(dtpFechaFin.Value, "yyyy-mm-dd HH:mm:ss") & "' ," & _
                         " CURRENT_TIMESTAMP ," & _
                         " '" & strUsuario & "') "
                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Lista de Precio"
    ElseIf (i_flag = 5) Then
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 2)) = 1 Then
                strSQL = " DELETE FROM producto_disponible_pv " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                         " AND pro_dis_fechainicio='" & Format(dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss") & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                strSQL = " INSERT INTO producto_disponible_pv " & _
                         " (emp_codigo,prd_codigo,pro_dis_fechainicio,pro_dis_fechafin,pro_dis_fechamod,pro_dis_usumod) " & _
                         " VALUES('" & strEmpresa & "', " & _
                         " '" & VSFG.TextMatrix(i, 0) & "'," & _
                         " '" & Format(dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss") & "' ," & _
                         " '" & Format(dtpFechaFin.Value, "yyyy-mm-dd HH:mm:ss") & "' ," & _
                         " CURRENT_TIMESTAMP ," & _
                         " '" & strUsuario & "') "
                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Lista de Precio"
    ElseIf (i_flag = 4) Then
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 5)) = 1 Then
                strSQL = " DELETE FROM producto_disponible " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                         " AND pro_dis_fechainicio='" & VSFG.TextMatrix(i, 2) & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                strSQL = " INSERT INTO producto_disponible " & _
                         " (emp_codigo,prd_codigo,pro_dis_fechainicio,pro_dis_fechafin,pro_dis_fechamod,pro_dis_usumod) " & _
                         " VALUES('" & strEmpresa & "', " & _
                         " '" & VSFG.TextMatrix(i, 0) & "'," & _
                         " '" & VSFG.TextMatrix(i, 2) & "' ," & _
                         " '" & VSFG.TextMatrix(i, 4) & "' ," & _
                         " CURRENT_TIMESTAMP ," & _
                         " '" & strUsuario & "') "
                clsCon_Def.Ejecutar strSQL, "M"
                
                strSQL = " REPLACE INTO producto_disponible " & _
                         " (emp_codigo,prd_codigo,pro_dis_fechainicio,pro_dis_fechafin,pro_dis_fechamod,pro_dis_usumod) " & _
                         " VALUES('" & strEmpresa & "', " & _
                         " '" & VSFG.TextMatrix(i, 0) & "'," & _
                         " '" & VSFG.TextMatrix(i, 2) & "' ," & _
                         " '" & VSFG.TextMatrix(i, 4) & "' ," & _
                         " CURRENT_TIMESTAMP ," & _
                         " '" & strUsuario & "') "
'                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Disponible para la venta"
    ElseIf (i_flag = 6) Then
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 5)) = 1 Then
                strSQL = " REPLACE INTO producto_disponible_ " & _
                         " (emp_codigo,prd_codigo,pro_dis_fechainicio,pro_dis_fechafin,pro_dis_fechamod,pro_dis_usumod) " & _
                         " VALUES('" & strEmpresa & "', " & _
                         " '" & VSFG.TextMatrix(i, 0) & "'," & _
                         " '" & VSFG.TextMatrix(i, 2) & "' ," & _
                         " '" & VSFG.TextMatrix(i, 4) & "' ," & _
                         " CURRENT_TIMESTAMP ," & _
                         " '" & strUsuario & "') "
                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Disponible para la venta"
    ElseIf (i_flag = 3) Then
        For i = 1 To VSFG.Rows - 1
            If Val(VSFG.TextMatrix(i, 3)) = 1 Then
                strSQL = " UPDATE producto " & _
                         " SET prd_ubica_linea='" & VSFG.TextMatrix(i, 2) & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "'"
                clsCon_Def.Ejecutar strSQL, "M"
            End If
        Next i
        Me.MousePointer = 0
        MsgBox "Carga Finalizada", vbInformation, "Balanceo"
    End If
    Unload Me

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
    Me.MousePointer = 11
    'Lee archivo para cargar lista de precio
    If (txtArchivo.Text <> "" And i_flag = 1) Then

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
            strSQL = " SELECT prd_nombre,prd_costo " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                     
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 2) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 3) = clsCon_Def.adorec_Def("prd_costo")
                If Val(Format(VSFG.TextMatrix(i, 3), "###0.00")) <> 0 Then
                    VSFG.TextMatrix(i, 4) = FormatoD2(100 - (Val(Format(VSFG.TextMatrix(i, 1), "###0.00")) / Val(Format(VSFG.TextMatrix(i, 3), "###0.00"))) * 100#)
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
    
    ElseIf (txtArchivo.Text <> "" And i_flag = 0) Then
    
       'Lee archivo para cambiar la colección
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 3
        VSFG.TextMatrix(0, 0) = "Código Producto"
        VSFG.TextMatrix(0, 1) = "Descripción"
        VSFG.TextMatrix(0, 2) = "Estado"
        VSFG.ColWidth(1) = 3500
        For i = 1 To VSFG.Rows - 1
            strSQL = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 2) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 2) = 0
                VSFG.Cell(flexcpBackColor, i, 0, i, 2) = &HC0C0FF
            End If
        Next i
    
    ElseIf (txtArchivo.Text <> "" And (i_flag = 2 Or i_flag = 5)) Then
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 3
        VSFG.TextMatrix(0, 0) = "Producto"
        VSFG.TextMatrix(0, 1) = "Descripcion"
        VSFG.TextMatrix(0, 2) = "Para Ingresar"
        VSFG.ColWidth(0) = 1500
        VSFG.ColWidth(1) = 3000
        VSFG.ColWidth(2) = 500
        For i = 1 To VSFG.Rows - 1
            strSQL = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                     
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 2) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 2) = 0
            End If
        Next i
    
    ElseIf (txtArchivo.Text <> "" And i_flag = 3) Then
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 4
        
        VSFG.ColPosition(1) = 2
        
        VSFG.TextMatrix(0, 0) = "Producto"
        VSFG.TextMatrix(0, 1) = "Descripcion"
        VSFG.TextMatrix(0, 2) = "Ubicacion"
        VSFG.TextMatrix(0, 3) = "Para Ingresar"
        VSFG.ColWidth(0) = 1500
        VSFG.ColWidth(1) = 3000
        VSFG.ColWidth(2) = 500
        VSFG.ColWidth(3) = 500
        For i = 1 To VSFG.Rows - 1
            strSQL = " SELECT prd_nombre " & _
                     " FROM producto " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                     
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 3) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 3) = 0
            End If
        Next i
    ElseIf (txtArchivo.Text <> "" And i_flag = 4) Then
        VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.Cols = 6
        VSFG.TextMatrix(0, 0) = "Producto"
        VSFG.TextMatrix(0, 1) = "Descripcion"
        VSFG.TextMatrix(0, 2) = "Fecha Inicio"
        VSFG.TextMatrix(0, 3) = "Fecha Fin"
        VSFG.TextMatrix(0, 4) = "Fecha Fin Nuevo"
        VSFG.TextMatrix(0, 5) = "Para Ingresar"
        VSFG.ColWidth(0) = 1500
        VSFG.ColWidth(1) = 3000
        VSFG.ColWidth(2) = 1000
        VSFG.ColWidth(3) = 1000
        VSFG.ColWidth(4) = 1000
        VSFG.ColWidth(5) = 500
        For i = 1 To VSFG.Rows - 1
            strSQL = " SELECT prd_nombre,pro_dis_fechainicio,pro_dis_fechafin " & _
                     " FROM producto INNER JOIN producto_disponible " & _
                     " ON producto.emp_codigo=producto_disponible.emp_codigo " & _
                     " AND producto.prd_codigo=producto_disponible.prd_codigo " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "'" & _
                     " AND producto.prd_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                     " AND pro_dis_fechafin>'" & Format(Me.dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss") & "' "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("prd_nombre")
                VSFG.TextMatrix(i, 2) = clsCon_Def.adorec_Def("pro_dis_fechainicio")
                VSFG.TextMatrix(i, 3) = clsCon_Def.adorec_Def("pro_dis_fechafin")
                VSFG.TextMatrix(i, 4) = Format(Me.dtpFechaInicio.Value, "yyyy-mm-dd HH:mm:ss")
                VSFG.TextMatrix(i, 5) = 1
            Else
                VSFG.TextMatrix(i, 1) = "NO ENCONTRADO - No se aplicara"
                VSFG.TextMatrix(i, 2) = ""
                VSFG.TextMatrix(i, 3) = ""
                VSFG.TextMatrix(i, 4) = ""
                VSFG.TextMatrix(i, 5) = 0
            End If
        Next i
    
    End If
    Me.MousePointer = 0
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
    op_lista.Value = True
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        strSQL = " SELECT (lis_pre_codigo) AS lis_pre_codigo,lis_pre_descripcion " & _
                 " FROM lista_precio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY lis_pre_codigo "
        clsCon_Def.Ejecutar strSQL
        Set dcmbDescripcion.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbDescripcion.ListField = "lis_pre_descripcion"
        dcmbDescripcion.BoundColumn = "lis_pre_codigo"
     
     'Consulta las colecciones existentes
        strSQL = " SELECT (clc_codigo) AS clc_codigo,clc_nombre " & _
                 " FROM coleccion " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY clc_codigo "
        clsCon_Def.Ejecutar strSQL
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

Private Sub op_Balanceo_Click()
        i_flag = 3
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = False
       dtpFechaFin.Enabled = False

End Sub

Private Sub op_coleccion_Click()
       i_flag = 0
       dcmbColeccion.Enabled = True
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = False
       dtpFechaFin.Enabled = False
End Sub

Private Sub op_disponiblePreventa_Click()
        i_flag = 5
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = True
       dtpFechaFin.Enabled = True
End Sub

Private Sub op_disponibleventa_Click()
        i_flag = 2
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = True
       dtpFechaFin.Enabled = True
End Sub

Private Sub op_lista_Click()
       i_flag = 1
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = True
       dtpFechaInicio.Enabled = False
       dtpFechaFin.Enabled = False
End Sub

Private Sub op_nodisponiblepreventa_Click()
        i_flag = 6
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = True
       dtpFechaFin.Enabled = False
End Sub

Private Sub op_nodisponibleventa_Click()
        i_flag = 4
       dcmbColeccion.Enabled = False
       dcmbDescripcion.Enabled = False
       dtpFechaInicio.Enabled = True
       dtpFechaFin.Enabled = False
End Sub
