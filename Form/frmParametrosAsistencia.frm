VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmParametrosAsistencia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   Icon            =   "frmParametrosAsistencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9030
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2765
      TabIndex        =   4
      Top             =   4320
      Width           =   1700
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4565
      TabIndex        =   3
      Top             =   4320
      Width           =   1700
   End
   Begin VB.Frame fraParametros 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Definición de Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   8295
         _cx             =   14631
         _cy             =   5318
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParametrosAsistencia.frx":030A
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
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmParametrosAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsSql As New clsConsulta
Private strSql As String

Private Sub btnAceptar_Click()
    Dim i As Long, control As Integer
    Dim descripcion As String, valor As String
    
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    For i = 1 To VSFG.Rows - 1
        descripcion = Trim(VSFG.TextMatrix(i, 3))
        valor = Trim(VSFG.TextMatrix(i, 4))
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            If valor = "" Then
                MsgBox "No puede modificar Parámetro, debe ingresar algún valor", vbCritical, "Modificación de " & VSFG.TextMatrix(i, 2)
                control = 1
            Else
                strSql = " UPDATE parametro_asistencia SET " & _
                         " par_descripcion='" & descripcion & "'," & _
                         " par_valor='" & valor & "' " & _
                         " WHERE par_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                clsSql.Ejecutar strSql
            End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    
    If control = 0 Then
        Limpiar
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    ucrtVSFG.Inicializar False, True, False
    Set ucrtVSFG.VSFGControl = VSFG
    CargarParametros
End Sub

Private Sub Limpiar()
    CargarParametros
    ucrtVSFG.SetFocus
End Sub

Private Sub CargarParametros()
    strSql = " SELECT par_codigo,par_nombre,par_descripcion,par_valor,'0' as modo " & _
             " FROM parametro_asistencia " & _
             " ORDER BY par_nombre "
    clsSql.Ejecutar strSql
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.Count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 1 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col = 2 Or Col >= VSFG.Cols - 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
