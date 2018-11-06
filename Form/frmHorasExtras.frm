VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmHorasExtras 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horas Extras"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmHorasExtras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6285
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      TabIndex        =   0
      Tag             =   "3"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3195
      TabIndex        =   1
      Tag             =   "6"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Horas Extras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Valor Hora Regular = Sueldo Básico / 160"
      Top             =   120
      Width           =   6015
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5535
         _cx             =   9763
         _cy             =   2143
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHorasExtras.frx":030A
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
         TabBehavior     =   0
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
   End
End
Attribute VB_Name = "frmHorasExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private strSql As String
Public SueldoBasico As Double
Public Normal As Boolean
Public Columna As Long


Private Sub cmdNuevo_Click()
    If Normal = True Then
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.VSFG.Col) = Format(VSFG.TextMatrix(3, 3), "###0.00")
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.INICIO1 - 5) = Format(VSFG.TextMatrix(1, 1), "###0.00")
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.INICIO1 - 4) = Format(VSFG.TextMatrix(2, 1), "###0.00")
        frmIngEgrRol.VSFG_AfterEdit frmIngEgrRol.VSFG.Row, frmIngEgrRol.VSFG.Col
    Else
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.VSFG.Col) = Format(VSFG.TextMatrix(3, 3), "###0.00")
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.EMP1 + 5) = Format(VSFG.TextMatrix(1, 1), "###0.00")
        frmIngEgrRol.VSFG.TextMatrix(frmIngEgrRol.VSFG.Row, frmIngEgrRol.EMP1 + 6) = Format(VSFG.TextMatrix(2, 1), "###0.00")
        frmIngEgrRol.VSFG_AfterEdit frmIngEgrRol.VSFG.Row, frmIngEgrRol.VSFG.Col
    End If
    Unload Me
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    VSFG.SetFocus
    Frame1.Caption = "Horas Extras (" & Format(SueldoBasico / 240, "###0.00") & " hora regular)"
    VSFG.TextMatrix(1, 1) = Format(VSFG.TextMatrix(1, 1), "###0.00")
    VSFG.TextMatrix(2, 1) = Format(VSFG.TextMatrix(2, 1), "###0.00")
    
    VSFG.TextMatrix(1, 2) = (SueldoBasico / 240) * 1.5
    VSFG.TextMatrix(2, 2) = (SueldoBasico / 240) * 2
    
    VSFG.TextMatrix(1, 3) = Format(Val(VSFG.TextMatrix(1, 2)) * Val(VSFG.TextMatrix(1, 1)), "###0.00")
    VSFG.TextMatrix(2, 3) = Format(Val(VSFG.TextMatrix(2, 2)) * Val(VSFG.TextMatrix(2, 1)), "###0.00")
    PonerTotales
End Sub

Private Sub PonerTotales()
    VSFG.Subtotal flexSTSum, -1, 1, "#,###.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
    VSFG.Subtotal flexSTSum, -1, 3, "#,###.00", RGB(230, 230, 230), RGB(120, 0, 0), , "Total"
End Sub

Private Sub Form_Load()
    Normal = True
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 40)
    clsSql.Inicializar AdoConn, AdoConnMaster
    VSFG.SubtotalPosition = flexSTBelow
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    PonerTotales
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    Else
        If VSFG.IsSubtotal(Row) = True Then Cancel = True
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 And Row > 0 Then
        VSFG.Cell(flexcpForeColor, Row, Col) = RGB(120, 0, 0)
    End If
End Sub

Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    VSFG.EditText = FormatoD2(VSFG.EditText)
    VSFG.TextMatrix(Row, 3) = Format(VSFG.TextMatrix(Row, 2) * VSFG.EditText, "###0.00")
End Sub
