VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmContenedorMercaderiaVacio 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contenedores Vacios"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmContenedorMercaderiaVacio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7590
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdAnularContenedores 
      Caption         =   "&Anular Contenedores"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   6360
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5880
      TabIndex        =   0
      Top             =   6360
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   5760
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7380
      _cx             =   13017
      _cy             =   10160
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmContenedorMercaderiaVacio.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin VB.Label Label4 
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
      Left            =   4320
      TabIndex        =   3
      Top             =   195
      Width           =   555
   End
End
Attribute VB_Name = "frmContenedorMercaderiaVacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSQL As String

Private Sub cmdAnularContenedores_Click()
    Dim clsCont As New clsContenedor
    Dim i As Long
    Dim n As Long
    clsCont.Inicializar AdoConn, AdoConnMaster
    n = 0
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            clsCont.SetContenedor VSFG.TextMatrix(i, 1)
            clsCont.AnularContenedor "CONTENEDOR VACIO CONFIRMADO"
            n = n + 1
        End If
    Next i
    MsgBox n & " Contenedores Anulados"
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

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    CargaContenedores
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarContenedor UCase(txtLector.Text)
        txtLector.Text = ""
    End If
End Sub

Private Sub AgregarContenedor(strContenedor As String)
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, 1) = strContenedor Then
            VSFG.TextMatrix(i, 0) = 1
        End If
    Next i
End Sub

Private Sub CargaContenedores()
    
    strSQL = " SELECT DISTINCT '0' as sel, con_mer_codigo,con_mer_fecha,est_con_mer_descripcion,dep_nombre,ubi_bod_codigo,con_mer_observacion,con_mer_fechamod,con_mer_usumod " & _
             " FROM contenedor_mercaderia INNER JOIN est_contenedor_mercaderia ON contenedor_mercaderia.est_con_mer_codigo=est_contenedor_mercaderia.est_con_mer_codigo" & _
             " LEFT JOIN deposito ON contenedor_mercaderia.emp_codigo=deposito.emp_codigo AND contenedor_mercaderia.dep_codigo=deposito.dep_codigo " & _
             " WHERE contenedor_mercaderia.emp_codigo = 'RYB' AND contenedor_mercaderia.est_con_mer_codigo!=-1 AND NOT EXISTS ( " & _
                " SELECT DISTINCT con_mer_codigo" & _
                " FROM (" & _
                    " SELECT CM.con_mer_codigo,det_contenedor_mercaderia.prd_codigo, SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) AS tot " & _
                    " FROM contenedor_mercaderia CM " & _
                    " INNER JOIN det_contenedor_mercaderia ON CM.emp_codigo=det_contenedor_mercaderia.emp_codigo AND CM.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo" & _
                    " WHERE CM.emp_codigo = '" & strEmpresa & "' AND CM.est_con_mer_codigo!=-1 " & _
                    " GROUP BY CM.con_mer_codigo,det_contenedor_mercaderia.prd_codigo " & _
                    " HAVING SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)!=0" & _
                " ) lleno" & _
                " WHERE lleno.con_mer_codigo=contenedor_mercaderia.con_mer_codigo" & _
            " ) " & _
            " ORDER BY contenedor_mercaderia.con_mer_codigo "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    
End Sub
