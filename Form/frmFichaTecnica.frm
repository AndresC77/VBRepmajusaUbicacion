VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFichaTecnica 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Tecnica"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmFichaTecnica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6630
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox txtCostoServicio 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtNombre 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtReferencia 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1425
      TabIndex        =   1
      Top             =   4440
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3225
      TabIndex        =   0
      Top             =   4440
      Width           =   1700
   End
   Begin MSDataListLib.DataCombo cmbProducto 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8LCtl.VSFlexGrid VSFG 
      Height          =   2250
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   6375
      _cx             =   11245
      _cy             =   3969
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFichaTecnica.frx":030A
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo Serv:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4440
      TabIndex        =   7
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "frmFichaTecnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String

Private Sub cmbAceptar_Click()
    Dim i As Long
      
'if para ver q los campos esten llenos
    
    'insert en preproducto_ficha
    
    For i = 1 To VSFG.Rows - 1
      
                'If tres columans llenas Then
                'insert en det_preproducto_ficha
                    strSql = " INSERT INTO zona(zon_codigo,zon_nombre,zon_fechamod,zon_usumod) " & _
                            " VALUES ('" & UCase(VSFG.TextMatrix(i, 1)) & "','" & UCase(VSFG.TextMatrix(i, 2)) & "', " & _
                            " CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsCon_Def.Ejecutar strSql, "M"
                
             'End If
    Next i
    
End Sub

Private Sub cmbProducto_Validate(Cancel As Boolean)
    VSFG.TextMatrix(VSFG.Row, 1) = cmbProducto.BoundText
    cmbProducto.Visible = False
    VSFG.SetFocus
    VSFG.Col = 3
    VSFG.EditCell

End Sub

Private Sub Form_Activate()
    
    strSql = " SELECT COALESCE(MAX(pre_fic_version),1) as n " & _
             " FROM preproducto_ficha " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND pre_codigo='" & txtReferencia.Text & "' "
    
    clsCon_Def.Ejecutar strSql
    
    txtVersion.Text = clsCon_Def.adorec_Def("n")
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

    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo = '" & Trim(VSFG.TextMatrix(Row, Col)) & "' " & _
                 " ORDER BY producto.prd_nombre "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            VSFG.TextMatrix(Row, 2) = clsCon_Def.adorec_Def("prd_nombre")
            VSFG.Col = 3
            VSFG.EditCell
        Else
            MsgBox "No existe el producto", vbInformation
            VSFG.TextMatrix(Row, 1) = ""
            VSFG.Col = 1
            VSFG.EditCell
        End If
    End If
    If VSFG.TextMatrix(Row, 1) <> "" And VSFG.TextMatrix(Row, 2) <> "" And VSFG.TextMatrix(Row, 3) <> "" Then
        VSFG.AddItem ""
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If VSFG.Col = 2 And KeyCode = vbKeyF4 And Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col)) <> "" Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " AND prd_nombre LIKE '" & Trim(VSFG.TextMatrix(VSFG.Row, VSFG.Col)) & "%' " & _
                 " ORDER BY producto.prd_nombre "
        clsAux.Ejecutar strSql
        cmbProducto = ""
        Set cmbProducto.RowSource = clsAux.adorec_Def.DataSource
        cmbProducto.ListField = "prd_nombre"
        cmbProducto.BoundColumn = "prd_codigo"
        cmbProducto.Visible = True
        cmbProducto.SetFocus
    End If
End Sub
