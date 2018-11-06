VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDatosGuia 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecciones el Vehículo"
   ClientHeight    =   1710
   ClientLeft      =   6435
   ClientTop       =   5295
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDatosGuia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2265
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   705
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4280
      _cx             =   7549
      _cy             =   1032
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDatosGuia.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ShowComboButton =   2
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el operador y el vehiculo:"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDatosGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCliente As String
Public strTipoDocumento As String
Public strNumeroDocumento As String
Public strCourier As String
Public strPlaca As String
Public booGuiaCreada As Boolean

Private Sub cmdAceptar_Click()
    If VSFG.TextMatrix(1, 0) <> "" And VSFG.TextMatrix(1, 1) <> "" Then
        strCourier = VSFG.TextMatrix(1, 0)
        strPlaca = VSFG.TextMatrix(1, 1)
        booGuiaCreada = True
        Unload Me
    Else
        MsgBox "No tiene ingresado todos los datos", vbCritical, "Operador"
    End If
End Sub

Private Sub cmdcancelar_Click()
    strCourier = ""
    strPlaca = ""
    booGuiaCreada = False
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim clsCons As New clsConsulta
    clsCons.Inicializar AdoConn, AdoConnMaster
    strCourier = ""
    strPlaca = ""
    booGuiaCreada = False
    strSQL = " SELECT courier_placa.cou_codigo,courier_placa.cou_pla_placa " & _
             " FROM persona INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo " & _
             " AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
             " INNER JOIN courier_placa ON forma_entrega.emp_codigo=courier_placa.emp_codigo " & _
             " AND forma_entrega.cou_codigo=courier_placa.cou_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND persona.per_codigo='" & strCliente & "'"
    clsCons.Ejecutar strSQL
    If clsCons.adorec_Def.RecordCount = 1 Then
        strCourier = clsCons.adorec_Def("cou_codigo")
        strPlaca = clsCons.adorec_Def("cou_pla_placa")
        booGuiaCreada = True
        Unload Me
    Else
        If clsCons.adorec_Def.RecordCount > 1 Then
            strCourier = clsCons.adorec_Def("cou_codigo")
        End If
        VSFG.TextMatrix(1, 0) = ""
        VSFG.TextMatrix(1, 1) = ""
        strSQL = " SELECT cou_codigo,cou_nombre " & _
                 " FROM courier " & _
                 " WHERE courier.emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY cou_nombre "
        clsCons.Ejecutar strSQL
        VSFG.ColComboList(0) = VSFG.BuildComboList(clsCons.adorec_Def, " *cou_nombre", "cou_codigo")
        If strCourier <> "" Then
            VSFG.TextMatrix(1, 0) = strCourier
            VSFG.TextMatrix(1, 1) = ""
            strCourier = ""
        End If
        
    End If
End Sub

Private Sub Form_Load()
'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    strCourier = ""
    strPlaca = ""
    booGuiaCreada = False

End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim clsCons As New clsConsulta
    clsCons.Inicializar AdoConn, AdoConnMaster
    If Col = 0 Then
    
        strSQL = " SELECT cou_pla_placa " & _
                 " FROM courier_placa " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cou_codigo='" & VSFG.TextMatrix(1, 0) & "'" & _
                 " ORDER BY cou_pla_placa"
        clsCons.Ejecutar strSQL
        VSFG.ColComboList(1) = VSFG.BuildComboList(clsCons.adorec_Def, " *cou_pla_placa", "cou_pla_placa")
    
    End If
End Sub
