VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRangoMultas 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas de Multas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmRangoMultas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4350
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Rangos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2535
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3675
         _cx             =   6482
         _cy             =   4471
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
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRangoMultas.frx":030A
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
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   360
         Picture         =   "frmRangoMultas.frx":03A0
         Top             =   2880
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   120
         Picture         =   "frmRangoMultas.frx":04CC
         ToolTipText     =   "Elimina una Fila"
         Top             =   2880
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tablas de Multas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   248
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton optMinutos 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Minutos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optHoras 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Horas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTiempo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin MSComCtl2.UpDown udcTiempo 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   285
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   99
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2228
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   668
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "frmRangoMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsSqlAux As New clsConsulta
Private strSql As String
Public lonCod As Long

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    Set clsSqlAux = Nothing
End Sub

Private Sub cmdAceptar_Click()
    'Inicializa los objetos de conexión con la base de datos
    Dim i As Long
    Dim intAct As Integer
    Dim parametro As String
    Dim salida As String
    
    txtTiempo.Tag = "1"
    'If txtTiempo.Tag <> "" Then
        strSql = " DELETE FROM multas WHERE emp_codigo='" & strEmpresa & "' AND mul_codigo='" & txtTiempo.Tag & "' "
        clsSql.Ejecutar strSql, "M"
        strSql = " DELETE FROM det_multas WHERE emp_codigo='" & strEmpresa & "' AND mul_codigo='" & txtTiempo.Tag & "' "
        clsSql.Ejecutar strSql, "M"
        lonCod = FormatoD0(txtTiempo.Tag)
'    Else
'        strSql = " SELECT COALESCE(max(mul_codigo)+1,1) as num " & _
'                 " FROM multas " & _
'                 " WHERE emp_codigo='" & strEmpresa & "' "
'        clsSql.Ejecutar strSql
'        If clsSql.adorec_Def.RecordCount > 0 Then
'            lonCod = FormatoD0(clsSql.adorec_Def("num"))
'        Else
'            lonCod = 1
'        End If
'    End If
        
    If optMinutos.value = True Then
        parametro = "M"
    Else
        parametro = "H"
    End If
    
    strSql = " INSERT INTO multas (emp_codigo,mul_codigo,mul_parametro,mul_tiempo,mul_fechamod,mul_usumod)" & _
             " VALUES('" & strEmpresa & "','" & lonCod & "','" & parametro & "','" & _
               txtTiempo.Text & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsSql.Ejecutar strSql, "M"
    For i = 1 To VSFG.Rows - 1
        If i = VSFG.Rows - 1 Then
            If Not IsNumeric(VSFG.TextMatrix(i, 3)) Then
                Exit For
            End If
        End If
        strSql = " INSERT INTO det_multas (emp_codigo,mul_codigo,det_mul_inferior,det_mul_superior,det_mul_valor,det_mul_fechamod,det_mul_usumod)" & _
                 " VALUES('" & strEmpresa & "','" & lonCod & "','" & Format(VSFG.TextMatrix(i, 1), "HH:mm:SS") & "','" & _
                 Format(VSFG.TextMatrix(i, 2), "HH:mm:SS") & "','" & FormatoD4(VSFG.TextMatrix(i, 3)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
    Next i
    
    MsgBox "Se ha generado la Tabla de Multas", vbInformation, "Tabla de Multas"
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSqlAux.Inicializar AdoConn, AdoConnMaster
    
    
    cargar
End Sub




Private Sub txtTiempo_Change()
    If Not IsNumeric(txtTiempo.Text) Then
        txtTiempo.Text = "0"
    End If
    CambiarRango
End Sub

Private Sub udcTiempo_DownClick()
    If FormatoD0(txtTiempo.Text) > 0 Then
        txtTiempo.Text = FormatoD0(txtTiempo.Text) - 1
    Else
        txtTiempo.Text = "0"
    End If
    CambiarRango
End Sub

Private Sub udcTiempo_UpClick()
    If FormatoD0(txtTiempo.Text) < 9998 Then
        txtTiempo.Text = FormatoD0(txtTiempo.Text) + 1
    Else
        txtTiempo.Text = "9999"
    End If
    CambiarRango
End Sub

Private Sub CambiarRango()
    Dim rango As Double, i As Long, param As String
    If optHoras.value = True Then
        param = "h"
    Else
        param = "n"
    End If
    
    rango = FormatoD4(txtTiempo.Text)
    If VSFG.Rows > 1 Then
        VSFG.TextMatrix(1, 2) = Format(DateAdd(param, txtTiempo.Text, "00:00:00"), "HH:mm:SS")
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim param As String
    If optHoras.value = True Then
        param = "h"
    Else
        param = "n"
    End If
    If (Col = 1 Or Col = 2) And Not IsDate(VSFG.TextMatrix(Row, Col)) Then
        MsgBox "Ingrese son formato HH:mm:SS", vbInformation, "Error"
        VSFG.TextMatrix(Row, Col) = "00:00:00"
    End If
    If Col = 3 And Not IsNumeric(VSFG.TextMatrix(Row, Col)) Then
        MsgBox "Ingrese solo números", vbInformation, "Error"
        VSFG.TextMatrix(Row, Col) = 0
    End If
    If Row = VSFG.Rows - 1 And (Col = 2 Or Col = 3) And Val(VSFG.TextMatrix(Row, 3)) >= 0 And Format(VSFG.TextMatrix(Row, 2), "HH:mm:SS") <> "00:00:00" Then
        If VSFG.TextMatrix(Row, 2) <> "23:59:59" Then
            'VSFG.AddItem "" & vbTab & VSFG.TextMatrix(VSFG.Rows - 1, 2) + 0.01 & vbTab & VSFG.TextMatrix(VSFG.Rows - 1, 2) + 0.02 & vbTab & "0.00"
            VSFG.AddItem "" & vbTab & Format(DateAdd(param, txtTiempo.Text, VSFG.TextMatrix(VSFG.Rows - 1, 2)), "HH:mm:SS") & vbTab & Format(DateAdd(param, txtTiempo.Text, VSFG.TextMatrix(VSFG.Rows - 1, 2)), "HH:mm:SS") & vbTab & "0.00"
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
        End If
    End If
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar toda columna menos la 1
    If Col = 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    Dim i As Long
    r = VSFG.MouseRow
    c = VSFG.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If c <> 0 Or r = (VSFG.Rows - 1) Then
        If c = 0 And r = (VSFG.Rows - 1) Then
        Else
            Exit Sub
        End If
    End If
    
    
    
    ' make sure the click was on a cell with a button
    If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "Tabla de Multas"
    ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         VSFG.RemoveItem (r)
         If VSFG.Rows > 1 Then
            For i = 1 To VSFG.Rows - 2
                'VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i, 2) + 0.01
                VSFG.TextMatrix(i + 1, 1) = DateAdd("s", 1, VSFG.TextMatrix(i, 2))
                If DateDiff("s", VSFG.TextMatrix(i + 1, 1), VSFG.TextMatrix(i + 1, 2)) <= 0 And Format(VSFG.TextMatrix(i + 1, 2), "HH:mm:SS") <> "00:00:00" Then
                    'VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i + 1, 2) + 0.01
                    VSFG.TextMatrix(i + 1, 1) = Format(DateAdd("s", 1, VSFG.TextMatrix(i + 1, 2)), "HH:mm:SS")
                End If
            Next i
        End If
         'PonerBotones
         'CalcuTotal
    Else
        VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True

End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Verifica cuando haya datos en una fila del grid tanto en bodega como en producto
    'para obtener la existencia de un producto en bodega
    Dim i As Long
    If Row > 0 Then
        If Col = 2 Then
            If VSFG.TextMatrix(Row, 2) = "" Or Not IsDate(VSFG.TextMatrix(Row, 2)) Then VSFG.TextMatrix(Row, 2) = "00:00:00"
            If (DateDiff("s", VSFG.TextMatrix(Row, 1), VSFG.TextMatrix(Row, 2)) < 0) And Format(VSFG.TextMatrix(Row, 2), "HH:mm:SS") <> "00:00:00" Then
                MsgBox "El rango superior debe ser mayor al inferior", vbInformation, "Error"
                'VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(Row, 1) + 0.01
                VSFG.TextMatrix(Row, 2) = Format(DateAdd("s", 1, VSFG.TextMatrix(Row, 1)), "HH:mm:SS")
            End If
            
            If VSFG.Rows > 1 Then
                For i = 1 To VSFG.Rows - 2
                    'VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i, 2) + 0.01
                    VSFG.TextMatrix(i + 1, 1) = Format(DateAdd("s", 1, VSFG.TextMatrix(i, 2)), "HH:mm:SS")
                    If DateDiff("s", VSFG.TextMatrix(i + 1, 1), VSFG.TextMatrix(i + 1, 2)) <= 0 And Format(VSFG.TextMatrix(i + 1, 2), "HH:mm:SS") <> "00:00:00" Then
                        'VSFG.TextMatrix(i + 1, 2) = VSFG.TextMatrix(i + 1, 1) + 0.01
                        VSFG.TextMatrix(i + 1, 2) = Format(DateAdd("s", 1, VSFG.TextMatrix(i + 1, 1)), "HH:mm:SS")
                    End If
                Next i
            End If
        End If
    End If
End Sub

Private Sub cargar()
    
    strSql = " SELECT mul_codigo,mul_parametro,mul_tiempo " & _
             " FROM multas " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        Dim i As Long
        txtTiempo.Tag = "1"
    
        strSql = " SELECT TIME_FORMAT(det_mul_inferior,'%H:%i:%s') as inferior,TIME_FORMAT(det_mul_superior,'%H:%i:%s') as superior,det_mul_valor " & _
                 " FROM det_multas " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND mul_codigo='" & txtTiempo.Tag & "'" & _
                 " ORDER BY 1,2 "
        clsSqlAux.Ejecutar strSql
        'Set VSFG.DataSource = clsSqlAux.adorec_Def.DataSource
        
        
        i = 1
        VSFG.Rows = 1
        While Not clsSqlAux.adorec_Def.EOF
            VSFG.AddItem ""
            VSFG.TextMatrix(i, 1) = clsSqlAux.adorec_Def("inferior")
            VSFG.TextMatrix(i, 2) = clsSqlAux.adorec_Def("superior")
            VSFG.TextMatrix(i, 3) = clsSqlAux.adorec_Def("det_mul_valor")
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            i = i + 1
            clsSqlAux.adorec_Def.MoveNext
        Wend
        
        
        txtTiempo.Text = clsSql.adorec_Def("mul_tiempo")
        
        If clsSql.adorec_Def("mul_parametro") = "M" Then
            optMinutos.value = True
            optHoras.value = False
        Else
            optHoras.value = True
            optMinutos.value = False
        End If
        
    Else
        VSFG.Rows = 2

        VSFG.TextMatrix(1, 1) = "00:00:01"
        txtTiempo.Text = 0
        VSFG.TextMatrix(1, 2) = Format(DateAdd("n", txtTiempo.Text, "00:00:00"), "HH:mm:SS")
        VSFG.TextMatrix(1, 3) = "0.00"
        
        
        optMinutos.value = True
        optHoras.value = False
        
        
        VSFG.Cell(flexcpPicture, 1, 0) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
        
        txtTiempo.Tag = "1"
    
    End If
End Sub




