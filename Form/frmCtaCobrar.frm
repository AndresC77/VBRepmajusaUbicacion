VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCtaCobrar 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observaciones Cuentas por Cobrar"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmCtaCobrar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11175
   Begin VB.Frame Frame5 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   255
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker dtpFecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   70451203
         CurrentDate     =   37463
      End
      Begin VB.Label lblFech 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Director:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1065
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gerente Zona:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   705
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   10935
      Begin VB.CommandButton btnMostrar 
         Caption         =   "Mostrar/Recargar"
         Height          =   375
         Left            =   9000
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame FraDatos 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos del Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   10455
         Begin VB.TextBox txtGer 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtDirector 
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtFPago 
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtRuc 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtDireccion 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   600
            Width           =   5895
         End
         Begin VB.TextBox txtTF 
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtVendedor 
            Height          =   285
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblGer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gerente Zona:"
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
            Left            =   180
            TabIndex        =   27
            Top             =   1035
            Width           =   1050
         End
         Begin VB.Label lblDir 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Director:"
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
            Left            =   5520
            TabIndex        =   26
            Top             =   1035
            Width           =   615
         End
         Begin VB.Label lblFormaPago 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma Pago:"
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
            Left            =   3240
            TabIndex        =   24
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CI/RUC:"
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
            Left            =   180
            TabIndex        =   23
            Top             =   270
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
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
            Left            =   180
            TabIndex        =   22
            Top             =   630
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telf/Fax:"
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
            Left            =   7080
            TabIndex        =   21
            Top             =   630
            Width           =   630
         End
         Begin VB.Label lblVend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
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
            Left            =   6495
            TabIndex        =   20
            Top             =   270
            Width           =   765
         End
      End
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   4215
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   10695
         _cx             =   18865
         _cy             =   7435
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCtaCobrar.frx":030A
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
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5595
      TabIndex        =   1
      Top             =   8640
      Width           =   1575
   End
End
Attribute VB_Name = "frmCtaCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsSql As New clsConsulta
Private clsPer As New clsConsulta
Private clsAsi As New clsConsulta
Private strSql As String
Private Descripcion As String

Private Sub btnMostrar_Click()
    If dcmbBeneficiario.BoundText <> "" Then
        CargarCuentas
    Else
        MsgBox "Seleccione un cliente", vbInformation, "Mostrar/Recargar"
    End If
End Sub

Private Sub cmbDirector_Change()
    OptCliente_Click
End Sub

Private Sub cmbGerente_Change()
    OptCliente_Click
End Sub

Private Sub cmbNegocio_Change()
    dcmbBeneficiario.Tag = "SI"
    dcmbBeneficiario.BoundText = ""
    cargarGZDir
    OptCliente_Click
    dcmbBeneficiario.Tag = "NO"
End Sub

Private Sub cmdAceptar_Click()
    
    If dcmbBeneficiario.BoundText = "" Then
        MsgBox "Seleccione un cliente", vbInformation, "Aceptar"
        Exit Sub
    End If
    
    Dim Cuantos As Long, i As Long
    Cuantos = 0
    For i = 1 To VSFG1.Rows - 2
        'If VSFG1.IsSubtotal(i) = False Then
            If CBool(FormatoD0(VSFG1.TextMatrix(i, 0))) = True Then
                Cuantos = Cuantos + 1
            End If
        'End If
    Next i
    If Cuantos = 0 Then
        MsgBox "No ha seleccionado ninguna cuenta por cobrar", vbInformation, "Aceptar"
        Exit Sub
    End If
    
    For i = 1 To VSFG1.Rows - 2
        'If VSFG1.IsSubtotal(i) = False Then
            If CBool(FormatoD0(VSFG1.TextMatrix(i, 0))) = True Then
                strSql = " UPDATE cuenta_p_c SET " & _
                         " cue_p_c_observacion='" & UCase(Trim(VSFG1.TextMatrix(i, VSFG1.Cols - 1))) & "', " & _
                         " cue_p_c_fechamod=CURRENT_TIMESTAMP, " & _
                         " cue_p_c_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cue_p_c_codigo='" & VSFG1.TextMatrix(i, 1) & "' " & _
                         " AND cue_p_c_tipo='C' "
                clsCta.Ejecutar strSql, "M"
            End If
        'End If
    Next i
    
    MsgBox "Se ha ingresado las observaciones correctamente", vbInformation, "Aceptar"
    CargarCuentas
End Sub

Private Sub dtpFecha_Change()
    If dcmbBeneficiario.BoundText <> "" Then
        CargarCuentas
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCta = Nothing
    Set clsSql = Nothing
    Set clsAsi = Nothing
    Set clsPer = Nothing
End Sub


Private Sub CalcuTotal()

   'Calcula totales
'    Dim SumaDebe As Double
'    Dim SumaHaber As Double
'
'    'Calcula total debe
'    For i = 1 To VSFG.Rows - 1
'        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
'        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
'    Next i
'    txtTotalDebe = Format(SumaDebe, "##0.00")
'    txtTotalHaber = Format(SumaHaber, "##0.00")
'   txtTotal = Format(txtTotalDebe - txtTotalHaber, "##0.00")
'
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBeneficiario_Change()
    On Error Resume Next
    Limpiar
    If dcmbBeneficiario.BoundText = "" Then
        cmdAceptar.Enabled = False
        VSFG1.Enabled = False
    End If
    If dcmbBeneficiario.MatchedWithList = True Or dcmbBeneficiario.Tag = "SI" Then
        cmdAceptar.Enabled = True
        VSFG1.Enabled = True
        CargarDatos
        CargarCuentas
    End If
End Sub

Private Sub CargarDatos()
    Limpiar
    '''If optcliente.value = True Then
        strSql = " SELECT persona.per_ruc,for_pag_nombre,CONCAT(ven_apellido,' ',ven_nombre) as vend, " & _
                 " persona.per_direccion, CONCAT(persona.per_telf,'/',persona.per_fax) as telf,COALESCE(CONCAT(p1.per_apellido,' ',p1.per_nombre),'') as gz,COALESCE(CONCAT(p2.per_apellido,' ',p2.per_nombre),'') as directo " & _
                 " FROM persona " & _
                 " INNER JOIN forma_pago ON forma_pago.emp_codigo=persona.emp_codigo AND forma_pago.for_pag_codigo=persona.for_pag_codigo " & _
                 " INNER JOIN vendedor ON vendedor.emp_codigo=persona.emp_codigo AND vendedor.ven_codigo=persona.ven_codigo " & _
                 " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                 " AND persona.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND persona.cat_p_tipo='C' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            txtRuc.Text = clsSql.adorec_Def("per_ruc")
            txtFPago.Text = clsSql.adorec_Def("for_pag_nombre")
            txtVendedor.Text = clsSql.adorec_Def("vend")
            txtDireccion.Text = clsSql.adorec_Def("per_direccion")
            txtTF.Text = clsSql.adorec_Def("telf")
            txtGer.Text = clsSql.adorec_Def("gz")
            txtDirector.Text = clsSql.adorec_Def("directo")
        End If
'''    Else
'''        strSql = " SELECT persona.per_ruc, " & _
'''                 " persona.per_direccion, CONCAT(persona.per_telf,'/',persona.per_fax) as telf " & _
'''                 " FROM persona " & _
'''                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
'''                 " AND persona.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
'''                 " AND persona.cat_p_tipo='P' "
'''        clsSql.Ejecutar strSql
'''        If clsSql.adorec_Def.RecordCount > 0 Then
'''            txtRuc.Text = clsSql.adorec_Def("per_ruc")
'''            txtFPago.Text = "N/A"
'''            txtVendedor.Text = "N/A"
'''            txtDireccion.Text = clsSql.adorec_Def("per_direccion")
'''            txtTF.Text = clsSql.adorec_Def("telf")
'''            txtGer.Text = "N/A"
'''            txtDirector.Text = "N/A"
'''        End If
  '''  End If

End Sub

Private Sub CargarCuentas()
    'Se carga todo el detalle de la cta x cobrar Vendedor
    strSql = " CREATE TEMPORARY TABLE Abo ( " & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
             " cue_p_c_tipo char(1) NOT NULL default ''," & _
             " abono decimal(14,2) default NULL," & _
             " abonoNC decimal(14,2) default NULL," & _
             " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCta.Ejecutar strSql
    
    '''Modificado
    strSql = " INSERT INTO Abo " & _
           " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(sum(if(pag_observacion!='NOTA DE CREDITO',pag_monto,0)),0.000) as abono,COALESCE(sum(if(pag_observacion='NOTA DE CREDITO',pag_monto,0)),0.000) as abonoNC " & _
           " FROM cuenta_p_c " & _
           " INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = pago.emp_codigo AND pago.pag_fecha <='" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
           " AND cuenta_p_c.cue_p_c_tipo='C' AND cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " GROUP BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
           " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCta.Ejecutar strSql
    
    strSql = " CREATE TEMPORARY TABLE RetFech ( " & _
           " emp_codigo char(3) NOT NULL default ''," & _
           " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
           " cue_p_c_tipo char(1) NOT NULL default ''," & _
           " reten decimal(14,2) default NULL," & _
           " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCta.Ejecutar strSql
    
    '''modificado
    strSql = " INSERT INTO RetFech " & _
           " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(comprobante_retencion.com_ret_total,0.000) as reten " & _
           " FROM cuenta_p_c " & _
           " INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo " & _
           " AND comprobante_retencion.com_ret_fecha <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
           " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCta.Ejecutar strSql
    
    strSql = " CREATE TEMPORARY TABLE Ret ( " & _
           " emp_codigo char(3) NOT NULL default ''," & _
           " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
           " cue_p_c_tipo char(1) NOT NULL default ''," & _
           " reten decimal(14,2) default NULL," & _
           " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCta.Ejecutar strSql
    
    '''modificado
    strSql = " INSERT INTO Ret " & _
           " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(RetFech.reten,0.000) as reten " & _
           " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " LEFT JOIN RetFech ON cuenta_p_c.cue_p_c_codigo = RetFech.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = RetFech.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = RetFech.emp_codigo " & _
           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
           " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " ORDER BY cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
    clsCta.Ejecutar strSql
    
    strSql = " CREATE TEMPORARY TABLE Cob ( " & _
           " emp_codigo char(3) NOT NULL default ''," & _
           " cue_p_c_codigo decimal(6,0) NOT NULL default '0'," & _
           " cue_p_c_tipo char(1) NOT NULL default ''," & _
           " abono decimal(14,2) default NULL," & _
           " abonoNC decimal(14,2) default NULL," & _
           " PRIMARY KEY (emp_codigo,cue_p_c_codigo,cue_p_c_tipo))"
    clsCta.Ejecutar strSql
    
    ''''modificado
    strSql = " INSERT INTO Cob " & _
           " SELECT cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,COALESCE(Abo.abono,0.000) as abono,COALESCE(Abo.abonoNC,0.000) as abonoNC " & _
           " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " LEFT JOIN Abo ON cuenta_p_c.cue_p_c_codigo = Abo.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = Abo.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = Abo.emp_codigo " & _
           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
           " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " ORDER BY cue_p_c_codigo"
    clsCta.Ejecutar strSql
    
    strSql = " DROP TABLE Abo "
    clsCta.Ejecutar strSql
    
'''''''    '''modificado
'''''''    strSql = " CREATE TEMPORARY TABLE Cuentas " & _
'''''''           " SELECT '' as sel,CONCAT(if(persona.cat_p_tipo = 'C','Cliente:','Proveedor:'),' ',persona.per_apellido,' ', persona.per_nombre,' (',persona.tip_ped_codigo,') ' ) as persona, CONCAT(cue_p_c_tot_cuenta, '/',cue_p_c_fra_cuenta) as pagos, " & _
'''''''           " cue_p_c_fechaemision as emision, cue_p_c_fechapropuesta as vencimiento, cue_p_c_fechapago as ultimo, " & _
'''''''           " COALESCE(cue_p_c_valor,0.000) as valor, COALESCE(Cob.abono,0.000) as abono, COALESCE(Cob.abonoNC,0.000) as abonoNC, cue_p_c_descripcion as descripcion," & _
'''''''           " cuenta_p_c.cue_p_c_codigo as numero, cue_p_c_egr_codigo, COALESCE(Ret.reten,0.000) as reten, cuenta_p_c.cue_p_c_tipo,persona.per_ruc,persona.per_direccion,persona.per_telf,ciu_nombre,for_pag_nombre,CONCAT(ven_apellido,' ',ven_nombre) as vend, COALESCE(egr_subtotal_o,0) as flete, CONCAT(COALESCE(gz.per_apellido,''),' ',COALESCE(gz.per_nombre,'')) as gerzo, CONCAT(COALESCE(di.per_apellido,''),' ',COALESCE(di.per_nombre,'')) as dir " & _
'''''''           " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
'''''''           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
'''''''           " AND cuenta_p_c.emp_codigo = persona.emp_codigo " & _
'''''''           " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
'''''''           " INNER JOIN forma_pago ON persona.for_pag_codigo=forma_pago.for_pag_codigo AND persona.emp_codigo=forma_pago.emp_codigo " & _
'''''''           " INNER JOIN Ret ON cuenta_p_c.cue_p_c_codigo = Ret.cue_p_c_codigo  " & _
'''''''           " AND cuenta_p_c.cue_p_c_tipo = Ret.cue_p_c_tipo " & _
'''''''           " AND cuenta_p_c.emp_codigo = Ret.emp_codigo " & _
'''''''           " INNER JOIN Cob ON cuenta_p_c.cue_p_c_codigo = Cob.cue_p_c_codigo  " & _
'''''''           " AND cuenta_p_c.cue_p_c_tipo = Cob.cue_p_c_tipo " & _
'''''''           " AND cuenta_p_c.emp_codigo = Cob.emp_codigo " & _
'''''''           " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
'''''''           " LEFT JOIN persona gz ON persona.emp_codigo=gz.emp_codigo AND persona.per_codigo_ref=gz.per_codigo " & _
'''''''           " LEFT JOIN persona di ON persona.emp_codigo=di.emp_codigo AND persona.per_codigo_ref2=di.per_codigo " & _
'''''''           " LEFT JOIN egreso ON cuenta_p_c.emp_codigo=egreso.emp_codigo AND cuenta_p_c.per_codigo=egreso.per_codigo AND cuenta_p_c.cue_p_c_egr_codigo=egreso.egr_codigo AND egreso.tip_egr_codigo='FAC' AND egr_fecha <= '" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
'''''''           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
'''''''           " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
'''''''           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
'''''''           " AND ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Cob.abonoNC,0.000) - COALESCE(Ret.reten,0.000),2)>0 " & _
'''''''           " ORDER BY persona,cue_p_c_egr_codigo, emision"
'''''''    clsCta.Ejecutar strSql
'''''''
'''''''    '''modificado
'''''''    strSql = " INSERT INTO Cuentas " & _
'''''''           " SELECT '' as sel, CONCAT(if(persona.cat_p_tipo = 'C','Cliente:','Proveedor:'),' ',persona.per_apellido,' ', persona.per_nombre,' (',persona.tip_ped_codigo,') ' ) as persona, '1/1' as pagos, " & _
'''''''           " ing_fecha as emision, ing_fecha as vencimiento, ing_fecha as ultimo, " & _
'''''''           " COALESCE(-1 * ing_total,0.000) as valor, '0.000' as abono, COALESCE(-1 * ing_saldo,0.000) as abonoNC, 'NOTA DE CREDITO' as descripcion," & _
'''''''           " 'NC' as numero, ing_codigo, '0.000' as reten, 'P',persona.per_ruc,persona.per_direccion,persona.per_telf,ciu_nombre,for_pag_nombre,CONCAT(ven_apellido,' ',ven_nombre) as vend, COALESCE(ing_subtotal_o,0) as flete, CONCAT(COALESCE(gz.per_apellido,''),' ',COALESCE(gz.per_nombre,'')) as gerzo, CONCAT(COALESCE(di.per_apellido,''),' ',COALESCE(di.per_nombre,'')) as dir " & _
'''''''           " FROM ingreso INNER JOIN persona ON ingreso.per_codigo = persona.per_codigo " & _
'''''''           " AND ingreso.emp_codigo = persona.emp_codigo " & _
'''''''           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
'''''''           " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
'''''''           " INNER JOIN forma_pago ON persona.for_pag_codigo=forma_pago.for_pag_codigo AND persona.emp_codigo=forma_pago.emp_codigo " & _
'''''''           " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
'''''''           " LEFT JOIN persona gz ON persona.emp_codigo=gz.emp_codigo AND persona.per_codigo_ref=gz.per_codigo " & _
'''''''           " LEFT JOIN persona di ON persona.emp_codigo=di.emp_codigo AND persona.per_codigo_ref2=di.per_codigo " & _
'''''''           " WHERE ingreso.emp_codigo = '" & strEmpresa & "' " & _
'''''''           " AND ingreso.tip_ing_codigo='DCL' AND  ingreso.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
'''''''           " AND ingreso.ing_fecha <= '" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
'''''''           " AND ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2)>0 " & _
'''''''           " ORDER BY persona,ing_codigo, emision"
'''''''    clsCta.Ejecutar strSql
    
    
    '''modificado
    strSql = " CREATE TEMPORARY TABLE Cuentas " & _
           " SELECT '' as sel,cuenta_p_c.cue_p_c_codigo as numero, " & _
           " cue_p_c_fechaemision as emision, cue_p_c_fechapropuesta as vencimiento, " & _
           " cue_p_c_egr_codigo, COALESCE(egr_subtotal_o,0) as flete, " & _
           " COALESCE(cue_p_c_valor,0.000) as valor, COALESCE(Cob.abono,0.000) as abono, COALESCE(Cob.abonoNC,0.000) as abonoNC,COALESCE(Ret.reten,0.000) as reten, " & _
           " COALESCE(cue_p_c_valor,0.000)-COALESCE(Cob.abono,0.000)-COALESCE(Cob.abonoNC,0.000)-COALESCE(Ret.reten,0.000) as saldo,cue_p_c_descripcion as descripcion,cue_p_c_observacion " & _
           " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " AND cuenta_p_c.emp_codigo = persona.emp_codigo " & _
           " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
           " INNER JOIN forma_pago ON persona.for_pag_codigo=forma_pago.for_pag_codigo AND persona.emp_codigo=forma_pago.emp_codigo " & _
           " INNER JOIN Ret ON cuenta_p_c.cue_p_c_codigo = Ret.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = Ret.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = Ret.emp_codigo " & _
           " INNER JOIN Cob ON cuenta_p_c.cue_p_c_codigo = Cob.cue_p_c_codigo  " & _
           " AND cuenta_p_c.cue_p_c_tipo = Cob.cue_p_c_tipo " & _
           " AND cuenta_p_c.emp_codigo = Cob.emp_codigo "
     strSql = strSql & " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
           " LEFT JOIN persona gz ON persona.emp_codigo=gz.emp_codigo AND persona.per_codigo_ref=gz.per_codigo " & _
           " LEFT JOIN persona di ON persona.emp_codigo=di.emp_codigo AND persona.per_codigo_ref2=di.per_codigo " & _
           " LEFT JOIN egreso ON cuenta_p_c.emp_codigo=egreso.emp_codigo AND cuenta_p_c.per_codigo=egreso.per_codigo AND cuenta_p_c.cue_p_c_egr_codigo=egreso.egr_codigo AND egreso.tip_egr_codigo='FAC' AND egr_fecha <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
           " AND cuenta_p_c.cue_p_c_tipo='C' AND  cuenta_p_c.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND cue_p_c_fechaemision <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " AND ROUND(COALESCE(cue_p_c_valor,0.000) - COALESCE(Cob.abono,0.000) - COALESCE(Cob.abonoNC,0.000) - COALESCE(Ret.reten,0.000),2)>0 " & _
           " ORDER BY cue_p_c_egr_codigo, emision"
    clsCta.Ejecutar strSql
    
    '''modificado
    strSql = " INSERT INTO Cuentas " & _
           " SELECT '' as sel,'NC' as numero, " & _
           " ing_fecha as emision, ing_fecha as vencimiento, " & _
           " ing_codigo,COALESCE(ing_subtotal_o,0) as flete, " & _
           " COALESCE(-1 * ing_total,0.000) as valor, '0.000' as abono, COALESCE(-1 * ing_saldo,0.000) as abonoNC,'0.000' as reten, " & _
           " COALESCE(-1 * ing_total,0.000)-'0.000'-COALESCE(-1 * ing_saldo,0.000)-'0.000' as saldo,'NOTA DE CREDITO' as descripcion,'' as cue_p_c_observacion " & _
           " FROM ingreso INNER JOIN persona ON ingreso.per_codigo = persona.per_codigo " & _
           " AND ingreso.emp_codigo = persona.emp_codigo " & _
           " AND persona.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' " & _
           " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
           " INNER JOIN forma_pago ON persona.for_pag_codigo=forma_pago.for_pag_codigo AND persona.emp_codigo=forma_pago.emp_codigo " & _
           " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
           " LEFT JOIN persona gz ON persona.emp_codigo=gz.emp_codigo AND persona.per_codigo_ref=gz.per_codigo " & _
           " LEFT JOIN persona di ON persona.emp_codigo=di.emp_codigo AND persona.per_codigo_ref2=di.per_codigo " & _
           " WHERE ingreso.emp_codigo = '" & strEmpresa & "' " & _
           " AND ingreso.tip_ing_codigo='DCL' AND  ingreso.per_codigo='" & dcmbBeneficiario.BoundText & "' " & _
           " AND ingreso.ing_fecha <= '" & Format(dtpFecha.Value, "yyyy-MM-dd") & "' " & _
           " AND ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2)>0 " & _
           " ORDER BY ing_codigo, emision"
    clsCta.Ejecutar strSql
    
    
    strSql = " SELECT * FROM Cuentas ORDER BY cue_p_c_egr_codigo, emision "
    clsCta.Ejecutar strSql
    
    Set VSFG1.DataSource = clsCta.adorec_Def.DataSource
    
    
    Dim Valor As Double, abono As Double, ncred As Double, reten As Double
    Valor = 0: abono = 0: ncred = 0: reten = 0
    strSql = " DROP TABLE Ret "
    clsCta.Ejecutar strSql
    strSql = " DROP TABLE RetFech "
    clsCta.Ejecutar strSql
    strSql = " DROP TABLE Cob "
    clsCta.Ejecutar strSql
    strSql = " DROP TABLE Cuentas "
    clsCta.Ejecutar strSql
    'VSFG1.Sort = flexSortGenericAscending
    'PONER SUBTOTALES
    Dim datos As Boolean
    datos = False
    If VSFG1.Rows = 1 Then
        VSFG1.Rows = 2
        datos = True
    End If
    
    VSFG1.Subtotal flexSTClear
    'VSFG3.subtotal flexSTSum, 3, 9, , &H8000000F, , True, "Total"
    VSFG1.Subtotal flexSTSum, -1, 10, , &H8000000F, &H80&
    ''''VSFG1.Cell(flexcpBackColor, VSFG1.Rows - 1, 0) = VSFG1.Cell(flexcpBackColor, VSFG1.Rows - 1, 1)
    
    For i = 1 To VSFG1.Rows - 2
        Valor = Valor + FormatoD4(VSFG1.TextMatrix(i, 6))
        abono = abono + FormatoD4(VSFG1.TextMatrix(i, 7))
        ncred = ncred + FormatoD4(VSFG1.TextMatrix(i, 8))
        reten = reten + FormatoD4(VSFG1.TextMatrix(i, 9))
    Next i
  
    VSFG1.Cell(flexcpAlignment, VSFG1.Rows - 1, 4) = 7
    VSFG1.TextMatrix(VSFG1.Rows - 1, 4) = "TOTAL"
    VSFG1.TextMatrix(VSFG1.Rows - 1, 6) = Valor
    VSFG1.TextMatrix(VSFG1.Rows - 1, 7) = abono
    VSFG1.TextMatrix(VSFG1.Rows - 1, 8) = ncred
    VSFG1.TextMatrix(VSFG1.Rows - 1, 9) = reten

    If datos = True Then
        cmdAceptar.Enabled = False
        VSFG1.RemoveItem 1
    Else
        cmdAceptar.Enabled = True
    End If
    
    
End Sub




Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub

Private Sub cargarGZDir()
    strSql = " SELECT '-1' as codigo,' Todos los Gerentes de Zona' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona " & _
             " INNER JOIN persona p1 " & _
             " ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND persona.cat_p_tipo = 'C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbGerente.RowSource = clsSql.adorec_Def.DataSource
    cmbGerente.ListField = "nombre"
    cmbGerente.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los Directores' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona " & _
             " INNER JOIN persona p1 " & _
             " ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref2 " & _
             " WHERE persona.emp_codigo= '" & strEmpresa & "' AND persona.cat_p_tipo = 'C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbDirector.RowSource = clsSql.adorec_Def.DataSource
    cmbDirector.ListField = "nombre"
    cmbDirector.BoundColumn = "codigo"
    
    cmbGerente.BoundText = "-1"
    cmbDirector.BoundText = "-1"
    
End Sub


'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub Limpiar()
    txtRuc.Text = ""
    txtFPago.Text = ""
    txtVendedor.Text = ""
    txtDireccion.Text = ""
    txtTF.Text = ""
    txtGer.Text = ""
    txtDirector.Text = ""
    VSFG1.Clear 1
    VSFG1.Rows = 2
    cmdAceptar.Enabled = False
    
    dtpFecha.Value = Format(HoyDia, "yyyy-MM-dd")
End Sub


Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa las clases para hacer distintas consultas
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    
    If dcmbBeneficiario.Text = "" Then
        cmdAceptar.Enabled = False
        VSFG1.Enabled = False
    End If
    
    dtpFecha.Value = Format(HoyDia, "yyyy-MM-dd")
    
    cargarTipoPedido
    cargarGZDir
           
End Sub

Private Sub OptCliente_Click()
    fraDatos.Caption = "Datos del Cliente"
    
'
'    txtFPago.Enabled = True
'
'    txtVendedor.Enabled = True
'
'    txtGer.Enabled = True
'    txtDirector.Enabled = True
'
    Limpiar
    
    p = 0
    
    dcmbBeneficiario.Text = ""
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    If cmbGerente.BoundText <> "-1" And cmbGerente.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref='" & cmbGerente.BoundText & "'"
    End If
    If cmbDirector.BoundText <> "-1" And cmbDirector.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref2='" & cmbDirector.BoundText & "'"
    End If
    strSql = strSql & " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSql
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
    End If
End Sub

''''''Private Sub optproveedor_Click()
''''''    FraDatos.Caption = "Datos del Proveedor"
''''''    p = 1
''''''
''''''
''''''    txtFPago.Enabled = False
''''''    txtVendedor.Enabled = False
''''''    txtGer.Enabled = False
''''''    txtDirector.Enabled = False
''''''
''''''    Limpiar
''''''    dcmbBeneficiario.Text = ""
''''''    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
''''''             " FROM persona " & _
''''''             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
''''''             " ORDER BY per_apellido,per_nombre"
''''''    clsPer.Ejecutar strSql
''''''    If clsPer.adorec_Def.EOF = False Then
''''''        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
''''''        dcmbBeneficiario.ListField = "nombre"
''''''        dcmbBeneficiario.BoundColumn = "per_codigo"
''''''    End If
''''''End Sub


Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If VSFG1.IsSubtotal(Row) = False Then
        If Col = 0 Then
            If CBool(FormatoD0(VSFG1.TextMatrix(Row, Col))) = True Then
                VSFG1.Cell(flexcpBackColor, Row, VSFG1.Cols - 1) = &H80FFFF
            Else
                VSFG1.Cell(flexcpBackColor, Row, VSFG1.Cols - 1) = &HFFFFFF
            End If
        End If
    End If
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFG1.IsSubtotal(Row) = False Then
        If Col <> 12 And Col <> 0 Then
            Cancel = True
        ElseIf Col = 12 Then
            If CBool(FormatoD0(VSFG1.TextMatrix(Row, 0))) = False Then
                Cancel = True
            End If
        End If
    Else
        Cancel = True
    End If
End Sub
