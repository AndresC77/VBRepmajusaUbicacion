VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerImpuestoRenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas de impuesto a la renta"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmVerImpuestoRenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   3845
      TabIndex        =   0
      Top             =   3480
      Width           =   1445
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impuesto a la renta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1935
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   8415
         _cx             =   14843
         _cy             =   3413
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerImpuestoRenta.frx":030A
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
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmVerImpuestoRenta.frx":03FB
         Left            =   240
         List            =   "frmVerImpuestoRenta.frx":0405
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tabla"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   380
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmVerImpuestoRenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Dim strSql As String

Private Sub cmbTipo_Click()
    BuscarTabla
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    'Call Centrar_Forma
End Sub

Private Sub BuscarTabla()
    Dim strWhere As String
    If cmbTipo.ListIndex = 0 Then
        strWhere = " AND ren_codigo LIKE 'A%'"
    Else
        strWhere = " AND ren_codigo LIKE 'M%'"
    End If
    strSql = " SELECT COALESCE(ren_frac_basica,'') as ren_frac_basica, COALESCE(ren_frac_exceso,'') as ren_frac_exceso, COALESCE(ren_imp_frac_basica,'') as ren_imp_frac_basica, COALESCE(ren_imp_frac_excedente,'') as ren_imp_frac_excedente, ren_codigo" & _
             " FROM parametro_renta" & _
             " WHERE  (emp_codigo = '" & strEmpresa & "')" & strWhere & _
             " ORDER BY ren_codigo"
    clsSql.Ejecutar (strSql)
    Set VSFG.DataSource = clsSql.adorec_Def.DataSource
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    
    cmbTipo.ListIndex = 0
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Me.cmbTipo.ListIndex = 1 Then Cancel = True
    If Col = 0 Then Cancel = True
End Sub

Private Sub VSFG_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim ValorMensual As Double
    Dim strSql1 As String
    VSFG.EditText = FormatoD0(VSFG.EditText)
    ValorMensual = FormatoD0(VSFG.EditText / 12)
    
    If Col = 0 Then
        strSql = " ren_frac_basica='" & VSFG.EditText & "'"
        strSql1 = " ren_frac_basica='" & ValorMensual & "'"
    End If
    If Col = 1 Then
        strSql = " ren_frac_exceso='" & VSFG.EditText & "'"
        strSql1 = " ren_frac_exceso='" & ValorMensual & "'"
    End If
    If Col = 2 Then
        strSql = " ren_imp_frac_basica='" & VSFG.EditText & "'"
        strSql1 = " ren_imp_frac_basica='" & ValorMensual & "'"
    End If
    If Col = 3 Then
        strSql = " ren_imp_frac_excedente='" & VSFG.EditText & "'"
        strSql1 = " ren_imp_frac_excedente='" & VSFG.EditText & "'"
    End If
    strSql = " UPDATE parametro_renta SET " & strSql & " WHERE emp_codigo='" & strEmpresa & "' AND ren_codigo='" & VSFG.TextMatrix(Row, 4) & "'"
    clsSql.Ejecutar strSql, "M"
    strSql1 = " UPDATE parametro_renta SET " & strSql1 & " WHERE emp_codigo='" & strEmpresa & "' AND ren_codigo='M" & Mid(VSFG.TextMatrix(Row, 4), 2) & "'"
    clsSql.Ejecutar strSql1, "M"
    If Col = 1 And Row < 6 Then
        Dim Valor As Double
        Dim Valor2 As Double
        Valor = FormatoD0(VSFG.EditText) + 0.01
        Valor2 = FormatoD0(ValorMensual) + 0.01
        strSql = " UPDATE parametro_renta SET ren_frac_basica = " & Valor & " WHERE emp_codigo='" & strEmpresa & "' AND ren_codigo='" & VSFG.TextMatrix(Row + 1, 4) & "'"
        clsSql.Ejecutar strSql, "M"
        strSql1 = " UPDATE parametro_renta SET ren_frac_basica = " & Valor2 & " WHERE emp_codigo='" & strEmpresa & "' AND ren_codigo='M" & Mid(VSFG.TextMatrix(Row + 1, 4), 2) & "'"
        clsSql.Ejecutar strSql1, "M"
        VSFG.TextMatrix(Row + 1, 0) = Valor
    End If
End Sub
