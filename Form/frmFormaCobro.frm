VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFormaCobro 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de Pago"
   ClientHeight    =   1530
   ClientLeft      =   6435
   ClientTop       =   5295
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFormaCobro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2385
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   825
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4425
      _cx             =   7805
      _cy             =   582
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
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFormaCobro.frx":030A
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
      Caption         =   "Seleccione la Forma de Pago:"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmFormaCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strNegocio As String
Public strCliente As String
Public strFormaCobro As String
Private Sub cmdAceptar_Click()
    strFormaCobro = VSFG.TextMatrix(0, 0)
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    strFormaCobro = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim clsCons As New clsConsulta
    clsCons.Inicializar AdoConn, AdoConnMaster
    strFormaCobro = ""
    strSQL = " SELECT COALESCE(persona_forma_cobro.for_cob_codigo,tipo_pedido_forma_cobro.for_cob_codigo,'0') as fc " & _
             " FROM persona LEFT JOIN persona_forma_cobro ON persona.emp_codigo=persona_forma_cobro.emp_codigo " & _
             " AND persona.per_codigo=persona_forma_cobro.per_codigo " & _
             " LEFT JOIN tipo_pedido_forma_cobro ON persona.emp_codigo=tipo_pedido_forma_cobro.emp_codigo " & _
             " AND persona.tip_ped_codigo=tipo_pedido_forma_cobro.tip_ped_codigo " & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND persona.per_codigo='" & strCliente & "'"
    clsCons.Ejecutar strSQL
    If clsCons.adorec_Def("fc") <> "0" Then
        strFormaCobro = clsCons.adorec_Def("fc")
        Unload Me
    Else
        strSQL = " SELECT for_cob_codigo,for_cob_nombre " & _
                 " FROM forma_cobro " & _
                 " ORDER BY for_cob_codigo"
        clsCons.Ejecutar strSQL
        VSFG.TextMatrix(0, 0) = ""
        VSFG.ColComboList(0) = VSFG.BuildComboList(clsCons.adorec_Def, "*for_cob_nombre,for_cob_codigo", "for_cob_codigo")
        SendKeys vbKeySpace
     End If
End Sub

Private Sub Form_Load()
'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    strFormaCobro = ""
End Sub

