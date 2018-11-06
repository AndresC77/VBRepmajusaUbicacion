VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmV_RangoComi 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas de Comisiones"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmV_RangoComi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
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
      Height          =   2775
      Left            =   128
      TabIndex        =   5
      Top             =   1680
      Width           =   4095
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3675
         _cx             =   6482
         _cy             =   3836
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
         FormatString    =   $"frmV_RangoComi.frx":030A
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
         Picture         =   "frmV_RangoComi.frx":03A9
         Top             =   2520
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   120
         Picture         =   "frmV_RangoComi.frx":04D5
         ToolTipText     =   "Elimina una Fila"
         Top             =   2520
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tablas de Comisiones"
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
      Height          =   1455
      Left            =   248
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Format          =   106168321
         CurrentDate     =   39153
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Format          =   106168321
         CurrentDate     =   39153
      End
      Begin VB.CheckBox chkActivo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Activar"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha fin:"
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
         Height          =   285
         Left            =   210
         TabIndex        =   7
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha inicio:"
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
         Height          =   285
         Left            =   210
         TabIndex        =   6
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2228
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   668
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "frmV_RangoComi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para confirmar el stock en bodega de un pedido ya realizado con ante_ #
'#  rioridad.                                                                   #
'#  frmV_VerPedBod V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Opciones que permite:                                                       #
'#  *   En una lista se despliegan los pedidos y sus detalles para que la       #
'#      persona encargada pueda ver en mayor detalle el mismo y así poder       #
'#      confirmar la cantidad de los productos que se está pidiendo.            #
'#                                                                              #
'#  Procesos internos que maneja:                                               #
'#  *   La lista que muestra los distintos pedidos se refresca automáticamente  #
'#      cada 20 segundos para buscar un nuevo pedido generado.                  #
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los   #
'#      detalles del mismo en un segundo grid.                                  #
'#  *   Se controla que el usuario pida como máximo la cantidad de productos    #
'#      solicitada a la bodega.                                                 #
'#  *   Una vez que el pedido ha sido confirmado su estado pasa a revisado.     #
'#  *   Se pueden ver solo los pedidos que aún no están revisados o los que ya  #
'#      se han revisado el día de hoy.                                          #
'#                                                                              #
'#  Tablas que maneja:                                                          #
'#                                                                              #
'#  persona:                                                                    #
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el  #
'#      pedido que se está confirmando.                                         #
'#  *   También se extrae el nombre del vendedor asignado al pedido.            #
'#  pedido:                                                                     #
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.               #
'#  det_pedido:                                                                 #
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.      #
'#                                                                              #
'################################################################################

Private clsSql As New clsConsulta
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
End Sub

Private Sub cmdAceptar_Click()
    'Inicializa los objetos de conexión con la base de datos
    Dim i As Long
    Dim intAct As Integer
    If chkActivo.value = 1 Then
        If MsgBox("Desea activar la nueva tabla desde hoy?", vbQuestion + vbYesNo, "Comisiones") = vbYes Then
            intAct = 1
        Else
            chkActivo.value = 0
            intAct = 0
        End If
    Else
        intAct = 0
    End If
    If Me.Tag = "N" Then
        strSql = " SELECT COALESCE(max(ran_com_codigo)+1,1) as num " & _
                 " FROM rango_comi " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " GROUP BY emp_codigo"
        clsSql.Ejecutar strSql
        lonCod = clsSql.adorec_Def("num")
    Else
        strSql = " DELETE " & _
                 " FROM det_rango_comi " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ran_com_codigo='" & lonCod & "' "
        clsSql.Ejecutar strSql, "M"
        strSql = " DELETE " & _
                 " FROM rango_comi " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ran_com_codigo='" & lonCod & "' "
        clsSql.Ejecutar strSql, "M"
    End If
    strSql = " INSERT INTO rango_comi (emp_codigo,ran_com_codigo,ran_com_fecha,ran_com_fechafin,ran_com_activo,ran_com_fechamod,ran_com_usumod)" & _
             " VALUES('" & strEmpresa & "','" & lonCod & "','" & DTPicker1.value & "','" & _
               DTPicker2.value & "','" & intAct & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsSql.Ejecutar strSql, "M"
    For i = 1 To VSFG.Rows - 1
        strSql = " INSERT INTO det_rango_comi (emp_codigo,ran_com_codigo,det_com_ran_inferior,det_com_ran_superior,det_com_ran_porcentaje,det_com_ran_fechamod,det_com_ran_usumodmod)" & _
                 " VALUES('" & strEmpresa & "','" & lonCod & "','" & VSFG.TextMatrix(i, 1) & "','" & _
                 VSFG.TextMatrix(i, 2) & "','" & VSFG.TextMatrix(i, 3) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsSql.Ejecutar strSql, "M"
    Next i
    If intAct = 1 Then
        strSql = " UPDATE rango_comi SET " & _
                 " ran_com_fechafin='" & DTPicker1.value & _
                 "', ran_com_activo='0',ran_com_fechamod=CURRENT_TIMESTAMP,ran_com_usumod='" & strUsuario & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ran_com_activo=1 AND ran_com_codigo != '" & lonCod & "'"
        clsSql.Ejecutar strSql, "M"
    End If
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = "M" And VSFG.Rows < 3 Then
        strSql = " SELECT det_com_ran_inferior,det_com_ran_superior,det_com_ran_porcentaje " & _
                " FROM det_rango_comi " & _
                " Where emp_codigo='" & strEmpresa & "'" & _
                " AND ran_com_codigo='" & lonCod & "'" & _
                " ORDER BY det_com_ran_inferior "
        clsSql.Ejecutar strSql
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        VSFG.TextMatrix(1, 3) = frmV_VerRangoComi.VSFG.TextMatrix(1, 2)
    End If
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
    If Me.Tag = "N" Then
        cmbAño1.Text = Year(HoyDia)
        cmbAño2.Text = Year(HoyDia)
        cmbMes1.Text = "Ene"
        cmbMes2.Text = "Ene"
        cmbDia1.Text = 1
        cmbDia2.Text = 1
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Not IsNumeric(VSFG.TextMatrix(Row, Col)) Then
        MsgBox "Ingrese solo números.", vbInformation, "Error"
        VSFG.TextMatrix(Row, Col) = 0
    End If
    If Row = VSFG.Rows - 1 And (Col = 2 Or Col = 3) And Val(VSFG.TextMatrix(Row, 3)) >= 0 And Val(VSFG.TextMatrix(Row, 2)) <> 0 Then
        VSFG.AddItem "" & vbTab & VSFG.TextMatrix(VSFG.Rows - 1, 2) + 0.01 & vbTab & VSFG.TextMatrix(VSFG.Rows - 1, 2) + 0.02 & vbTab & "0.00"
        VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
        VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
        VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
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
    
    If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub
     
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
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         VSFG.RemoveItem (r)
         If VSFG.Rows > 1 Then
            For i = 1 To VSFG.Rows - 2
                VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i, 2) + 0.01
                If Val(VSFG.TextMatrix(i + 1, 1)) >= Val(VSFG.TextMatrix(i + 1, 2)) And Val(VSFG.TextMatrix(i + 1, 2)) <> 0 Then
                    VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i + 1, 2) + 0.01
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
            If (Val(VSFG.TextMatrix(Row, 2)) < Val(VSFG.TextMatrix(Row, 1))) And Val(VSFG.TextMatrix(Row, 2)) <> 0 Then
                MsgBox "El rango superior debe ser mayor al inferior", vbInformation, "Error"
                VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(Row, 1) + 0.01
            End If
            If VSFG.Rows > 1 Then
                For i = 1 To VSFG.Rows - 2
                    VSFG.TextMatrix(i + 1, 1) = VSFG.TextMatrix(i, 2) + 0.01
                    If Val(VSFG.TextMatrix(i + 1, 1)) >= Val(VSFG.TextMatrix(i + 1, 2)) And Val(VSFG.TextMatrix(i + 1, 2)) <> 0 Then
                        VSFG.TextMatrix(i + 1, 2) = VSFG.TextMatrix(i + 1, 1) + 0.01
                    End If
                Next i
            End If
        End If
    End If
End Sub
