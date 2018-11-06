VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDatos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "frmDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8370
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2435
      TabIndex        =   2
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4235
      TabIndex        =   1
      Top             =   6480
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8145
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   5400
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   7935
         _cx             =   13996
         _cy             =   9525
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
         AllowUserResizing=   3
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
         FormatString    =   $"frmDatos.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
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
         WordWrap        =   -1  'True
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
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Public VSFGOrigen As VSFlexGrid

Private Sub cmbAceptar_Click()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        If Trim(VSFG.TextMatrix(i, 2)) = "" Then
            VSFGOrigen.Cell(flexcpText, VSFGOrigen.Tag, i) = VSFG.Cell(flexcpText, i, 1)
        Else
            VSFGOrigen.Cell(flexcpText, VSFGOrigen.Tag, i) = VSFG.Cell(flexcpText, i, 3)
        End If
    Next i
    Unload Me
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = 0 Or Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = 2 Or Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = -2 Then
        If Row >= VSFG.Rows - 3 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = 3 Or Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = -3 Then
        If Row = 1 Or Row >= VSFG.Rows - 3 Then
            Cancel = True
        End If
    End If
    VSFG.ComboList = Trim(VSFG.TextMatrix(Row, 2))
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
    
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = -2 Then
        VSFG.TextMatrix(VSFG.Rows - 1, 1) = 2
    ElseIf Val(VSFG.TextMatrix(VSFG.Rows - 1, 1)) = -3 Then
        VSFG.TextMatrix(VSFG.Rows - 1, 1) = 3
    End If
    If VSFG.TextMatrix(Row, 2) <> "" Then
        VSFG.TextMatrix(Row, 3) = VSFG.ComboData
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    VSFG.AutoSize Col, , False
End Sub
