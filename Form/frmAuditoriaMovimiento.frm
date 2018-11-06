VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAuditoriaMovimiento 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoria de Movimientos"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   Icon            =   "frmAuditoriaMovimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   13980
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   12255
      Begin VB.OptionButton optIng 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Ingresos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optEgr 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Egresos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkFiltroPersona 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Tipo de Persona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   90
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkNum 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por No. de Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5160
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2685
      End
      Begin VB.TextBox txtNum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   735
         Width           =   3450
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   8760
         TabIndex        =   10
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmAuditoriaMovimiento.frx":030A
            Left            =   1320
            List            =   "frmAuditoriaMovimiento.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   480
            TabIndex        =   15
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69468163
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69468163
            CurrentDate     =   37463
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8760
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   735
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPersona 
         Height          =   330
         Left            =   90
         TabIndex        =   26
         Top             =   1560
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPersona 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Personas"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   1335
         Width           =   4905
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Ingreso"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   495
         Width           =   4905
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Doc"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   495
         Width           =   3450
      End
   End
   Begin VB.CheckBox chkResta 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Resta"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9240
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtLector 
      Height          =   285
      Left            =   11520
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdAuditar 
      Caption         =   "&Auditar"
      Height          =   360
      Left            =   6840
      TabIndex        =   1
      Top             =   6600
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   12240
      TabIndex        =   0
      Top             =   6600
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGDestino 
      Height          =   3840
      Left            =   6840
      TabIndex        =   2
      Top             =   2640
      Width           =   7065
      _cx             =   12462
      _cy             =   6773
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAuditoriaMovimiento.frx":039E
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
      FrozenRows      =   1
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGOrigen 
      Height          =   3840
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   6585
      _cx             =   11615
      _cy             =   6773
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAuditoriaMovimiento.frx":0440
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
      FrozenRows      =   1
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
   Begin MSDataListLib.DataCombo cmbCotizacion 
      Height          =   330
      Left            =   4125
      TabIndex        =   30
      Top             =   2280
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000050&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número de Doc"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4125
      TabIndex        =   31
      Top             =   2040
      Width           =   4500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   10920
      TabIndex        =   4
      Top             =   2355
      Width           =   555
   End
End
Attribute VB_Name = "frmAuditoriaMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Public clsContenedorOrigen As New clsContenedor
Private clsContenedorDestino As New clsContenedor

Private Sub cmbBodegaOrigen_Validate(Cancel As Boolean)
    CargaUbicaOrigen
End Sub

Private Sub CargaUbicaOrigen()
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & cmbBodegaOrigen.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsCon_Def.Ejecutar strSql
    Set cmbUbicacionOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbUbicacionOrigen.ListField = "ubi_bod_codigo"
    cmbUbicacionOrigen.BoundColumn = "ubi_bod_codigo"
End Sub


Private Sub cmdAuditar_Click()
    Dim i As Long
    Dim j As Long
    Dim Pasa As Long
    
    For i = 2 To VSFGOrigen.Rows - 1
        If FormatoD0(VSFGOrigen.TextMatrix(i, 4)) = 0 And FormatoD0(VSFGOrigen.TextMatrix(i, 3)) <> 0 Then
            For j = 2 To VSFGDestino.Rows - 1
                If FormatoD0(VSFGDestino.TextMatrix(j, 4)) = 0 And FormatoD0(VSFGDestino.TextMatrix(j, 3)) <> 0 Then
                    If VSFGOrigen.TextMatrix(i, 0) = VSFGDestino.TextMatrix(j, 0) Then
                        If FormatoD2(FormatoD2(VSFGOrigen.TextMatrix(i, 2)) - FormatoD2(VSFGOrigen.TextMatrix(i, 3))) = 0 Then
                            If FormatoD2(VSFGOrigen.TextMatrix(i, 3) - VSFGDestino.TextMatrix(j, 3)) = 0 Then
                                VSFGOrigen.TextMatrix(i, 4) = 1
                                VSFGDestino.TextMatrix(j, 4) = 1
                                Exit For
                            End If
                        End If
                    End If
                ElseIf VSFGDestino.TextMatrix(j, 0) = "" And FormatoD0(VSFGDestino.TextMatrix(j, 4)) = 0 And FormatoD0(VSFGDestino.TextMatrix(j, 3)) = 0 Then
                    VSFGDestino.TextMatrix(j, 4) = 1
                End If
            Next j
        ElseIf FormatoD0(VSFGOrigen.TextMatrix(i, 4)) = 0 And FormatoD0(VSFGOrigen.TextMatrix(i, 3)) = 0 Then
            VSFGOrigen.TextMatrix(i, 4) = 1
        End If
    Next i
    Pasa = 0
    For i = 2 To VSFGOrigen.Rows - 1
        If VSFGOrigen.TextMatrix(i, 0) <> "" Then
        If FormatoD0(VSFGOrigen.TextMatrix(i, 4)) = 0 Then
            Pasa = Pasa + 1
        End If
        End If
    Next i
    For i = 2 To VSFGDestino.Rows - 1
        If VSFGDestino.TextMatrix(i, 0) <> "" Then
        If FormatoD0(VSFGDestino.TextMatrix(i, 4)) = 0 Then
            Pasa = Pasa + 1
        End If
        End If
    Next i
    If Pasa = 0 Then
        MsgBox "Contenedor pasa Auditoria", vbInformation, "AUDITORIA"
        clsContenedorOrigen.AgregaObservacion "PASA AUDITORIA " & Ahora & " - " & strUsuario
    Else
        MsgBox "Contenedor NO pasa Auditoria" & vbNewLine & "Tiene " & Pasa & " Error(es)", vbCritical, "AUDITORIA"
        clsContenedorOrigen.AgregaObservacion "NO PASA AUDITORIA " & Ahora & " - " & strUsuario
    End If
    Unload Me
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
    clsContenedorDestino.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set cmbBodegaOrigen.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbBodegaOrigen.ListField = "dep_nombre"
    cmbBodegaOrigen.BoundColumn = "dep_codigo"
    VSFGDestino.SubtotalPosition = flexSTAbove
    VSFGOrigen.SubtotalPosition = flexSTAbove
    VSFGDestino.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFGOrigen.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    VSFGDestino.Cell(flexcpFontSize, 1, 0, 1, VSFGDestino.Cols - 1) = VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) + 2
    VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) = VSFGOrigen.Cell(flexcpFontSize, 1, 0, 1, VSFGOrigen.Cols - 1) + 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtCodigoOrigen_Change()
    clsContenedorOrigen.SetContenedor txtCodigoOrigen.Text
    dtpFechaOrigen.Value = clsContenedorOrigen.strFecha
    cmbBodegaOrigen.BoundText = clsContenedorOrigen.strBodega
    cmbBodegaOrigen_Validate False
    cmbUbicacionOrigen.BoundText = clsContenedorOrigen.strUbicacion
    TxtObserOrigen.Text = clsContenedorOrigen.strObservacion
    Set VSFGOrigen.DataSource = clsContenedorOrigen.adorec_DetalleContenedor
    VSFGOrigen.Cols = VSFGOrigen.Cols + 2
    VSFGOrigen.TextMatrix(0, VSFGOrigen.Cols - 2) = "Descargar"
    VSFGOrigen.TextMatrix(0, VSFGOrigen.Cols - 1) = "Modi"
    
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProd UCase(txtLector.Text), chkResta.Value
        txtLector.Text = ""
        chkResta.Value = False
    End If
End Sub

Private Sub AgregarProd(codigo As String, Optional Resta As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim pas As Boolean
    pas = False
InicioFor:
    For i = 1 To VSFGOrigen.Rows - 1
        If codigo = VSFGOrigen.TextMatrix(i, 0) Then
'            If FormatoD2(VSFGOrigen.TextMatrix(i, 2)) >= FormatoD2(VSFGOrigen.TextMatrix(i, 3)) + 1 Then
                VSFGOrigen.ShowCell i, 0
                VSFGOrigen.Select i, 0
                If Resta = False Then
                    VSFGOrigen.TextMatrix(i, 3) = Val(Format(VSFGOrigen.TextMatrix(i, 3), "###0")) + 1
                Else
                    VSFGOrigen.TextMatrix(i, 3) = Val(Format(VSFGOrigen.TextMatrix(i, 3), "###0")) - 1
                End If
                pas = True
                Exit For
 '           Else
                'MsgBox "No tiene suficiente mercaderia", vbCritical, "Contenedor"
 '               Exit For
 '           End If
        End If
    Next i
    
    If pas = True Then
        pas = False
        For j = 1 To VSFGDestino.Rows - 1
            If codigo = VSFGDestino.TextMatrix(j, 0) Then
                VSFGDestino.ShowCell j, 0
                VSFGDestino.Select j, 0
                If Resta = False Then
                    VSFGDestino.TextMatrix(j, 3) = Val(Format(VSFGDestino.TextMatrix(j, 3), "###0")) + 1
                Else
                    VSFGDestino.TextMatrix(j, 3) = Val(Format(VSFGDestino.TextMatrix(j, 3), "###0")) - 1
                End If
                pas = True
                Exit For
            End If
        Next j
        If pas = False Then
            If Resta = False Then
                VSFGDestino.AddItem VSFGOrigen.TextMatrix(i, 0) & vbTab & VSFGOrigen.TextMatrix(i, 1) & vbTab & "0" & vbTab & "1"
            Else
                VSFGDestino.AddItem VSFGOrigen.TextMatrix(i, 0) & vbTab & VSFGOrigen.TextMatrix(i, 1) & vbTab & "0" & vbTab & "-1"
            End If
        End If
        VSFGDestino.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
        VSFGOrigen.Subtotal flexSTSum, -1, 3, , vbBlue, vbWhite, True, "TOTAL"
    Else
        strSql = " SELECT prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND prd_codigo='" & codigo & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If Resta = False Then
                VSFGDestino.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & "0" & vbTab & "1"
            Else
                VSFGDestino.AddItem clsCon_Def.adorec_Def("prd_codigo") & vbTab & clsCon_Def.adorec_Def("prd_nombre") & vbTab & "0" & vbTab & "-1"
            End If
        Else
            MsgBox "El producto no existe en la base de datos." & vbNewLine & _
                   "No se ingresara.", vbInformation, "Productos"
        End If

    End If
End Sub
