VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEnvioCorreo 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Peso en Lista de Embarque"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEnvioCorreo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   13350
   Begin VB.CommandButton cmdEjemplo 
      Caption         =   "Ejemplo"
      Height          =   375
      Left            =   11880
      TabIndex        =   50
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "frmEnvioCorreo.frx":030A
      Top             =   5880
      Width           =   7095
   End
   Begin VB.TextBox txtHipervinculo 
      Height          =   315
      Left            =   7320
      TabIndex        =   47
      Top             =   6000
      Width           =   5895
   End
   Begin VB.Frame Frame1 
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
      Height          =   5295
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   7080
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar"
         Height          =   285
         Left            =   2520
         TabIndex        =   42
         Top             =   4920
         Width           =   1335
      End
      Begin VB.OptionButton optN9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   3645
         Width           =   255
      End
      Begin VB.OptionButton optN8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   3285
         Width           =   255
      End
      Begin VB.OptionButton optN7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   2925
         Width           =   255
      End
      Begin VB.OptionButton optN6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   2565
         Width           =   255
      End
      Begin VB.OptionButton optN5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   2205
         Width           =   255
      End
      Begin VB.OptionButton optN4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1845
         Width           =   255
      End
      Begin VB.OptionButton optN3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1485
         Width           =   255
      End
      Begin VB.OptionButton optN2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1125
         Width           =   255
      End
      Begin VB.OptionButton optN1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   765
         Width           =   255
      End
      Begin VB.OptionButton optN10 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   705
         TabIndex        =   12
         Top             =   4005
         Width           =   255
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   330
         Left            =   1080
         TabIndex        =   22
         Top             =   720
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   330
         Left            =   1080
         TabIndex        =   23
         Top             =   1080
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   330
         Left            =   1080
         TabIndex        =   24
         Top             =   1440
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   330
         Left            =   1080
         TabIndex        =   25
         Top             =   1800
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   330
         Left            =   1080
         TabIndex        =   26
         Top             =   2160
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   330
         Left            =   1080
         TabIndex        =   27
         Top             =   2520
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   330
         Left            =   1080
         TabIndex        =   28
         Top             =   2880
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   330
         Left            =   1080
         TabIndex        =   29
         Top             =   3240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   330
         Left            =   1080
         TabIndex        =   30
         Top             =   3600
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN10 
         Height          =   330
         Left            =   1065
         TabIndex        =   31
         Top             =   3960
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
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
      Begin MSDataListLib.DataCombo cmbCiudad 
         Height          =   330
         Left            =   1080
         TabIndex        =   45
         Top             =   4440
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   46
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   41
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   38
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   37
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N1:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N2:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N3:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N4:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N10:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   4020
         Width           =   330
      End
   End
   Begin VB.TextBox txtListaCorreos 
      Height          =   1755
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   5895
   End
   Begin VB.TextBox txtCuerpo 
      Height          =   2355
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   5895
   End
   Begin VB.TextBox txtAsunto 
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Top             =   2760
      Width           =   5895
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   330
      ItemData        =   "frmEnvioCorreo.frx":0358
      Left            =   840
      List            =   "frmEnvioCorreo.frx":0362
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9780
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      Height          =   375
      Left            =   8340
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2055
      Left            =   7320
      TabIndex        =   8
      Top             =   360
      Width           =   5895
      _cx             =   10398
      _cy             =   3625
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
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEnvioCorreo.frx":03AB
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hipervinculo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7335
      TabIndex        =   48
      Top             =   5760
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerpo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7320
      TabIndex        =   4
      Top             =   3120
      Width           =   570
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7335
      TabIndex        =   3
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   135
      Width           =   345
   End
End
Attribute VB_Name = "frmEnvioCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta

Private Sub cmbNegocio_Click(Area As Integer)
    CargaRed
End Sub

Private Sub cmbTipo_Validate(Cancel As Boolean)
    If cmbTipo.ListIndex = 0 Then
        txtListaCorreos.Visible = True
        cmdConsultar.Enabled = True
        cmdConsultar.Caption = "Agregar Emails"
        txtAsunto.Locked = False
        txtCuerpo.Locked = False
        txtAsunto.Text = ""
        txtCuerpo.Text = ""
        VSFG.Rows = 1
        VSFG.Cols = 1
        VSFG.TextMatrix(0, 0) = "EMail"
    ElseIf cmbTipo.ListIndex = 1 Then
        txtListaCorreos.Visible = False
        cmdConsultar.Enabled = True
        cmdConsultar.Caption = "Cargar"
        txtAsunto.Locked = False
        txtCuerpo.Locked = False
        txtAsunto.Text = ""
        txtCuerpo.Text = ""
        VSFG.Rows = 1
        VSFG.Cols = 4
        VSFG.TextMatrix(0, 0) = "EMail"
        VSFG.TextMatrix(0, 1) = "Cliente"
        VSFG.TextMatrix(0, 2) = "Codigo"
        VSFG.TextMatrix(0, 3) = "Referencias"
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdConsultar_Click()
    Dim strSql As String
    Dim i As Long
    If cmbTipo.ListIndex = 0 Then
        txtListaCorreos.Visible = True
        txtListaCorreos.Text = ""
    ElseIf cmbTipo.ListIndex = 1 Then
        strSql = " SELECT persona.per_email,"
        strSql = strSql & " CONCAT(persona.per_apellido,' ', persona.per_nombre) as cli, persona.per_codigo " & _
                 " FROM persona "
        strSql = strSql & " LEFT JOIN persona as N1 ON N1.emp_codigo=persona.emp_codigo " & _
                 " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
                 " LEFT JOIN persona as N2 ON N2.emp_codigo=persona.emp_codigo " & _
                 " AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1 " & _
                 " LEFT JOIN persona as N3 ON persona.emp_codigo = N3.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1 " & _
                 " LEFT JOIN persona as N4 ON persona.emp_codigo = N4.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 "
        strSql = strSql & " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
                 " AND persona.ciu_codigo LIKE '" & cmbCiudad.BoundText & "' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND persona.cat_p_tipo='C' AND persona.per_email!='' "
        If cmbGerente.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref='" & cmbGerente.BoundText & "' "
        End If
        If cmbDirector.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref2='" & cmbDirector.BoundText & "' "
        End If
        If cmbEmprendedor.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref3='" & cmbEmprendedor.BoundText & "' "
        End If
        If cmbEjecutivo.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref4='" & cmbEjecutivo.BoundText & "' "
        End If
        If cmbN5.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref5='" & cmbN5.BoundText & "' "
        End If
        If cmbN6.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref6='" & cmbN6.BoundText & "' "
        End If
        If cmbN7.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref7='" & cmbN7.BoundText & "' "
        End If
        If cmbN8.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref8='" & cmbN8.BoundText & "' "
        End If
        If cmbN9.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref9='" & cmbN9.BoundText & "' "
        End If
        If cmbN10.BoundText <> "" Then
            strSql = strSql & " AND persona.per_codigo_ref10='" & cmbN10.BoundText & "' "
        End If
        strSql = strSql & " ORDER BY CONCAT(persona.per_apellido,' ', persona.per_nombre) "
        clsCon_Def.Ejecutar strSql
        Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
        If VSFG.Rows > 2 Then
            i = 1
        Else
            i = 0
        End If
'        txtAsunto.Text = ""
'        txtCuerpo.Text = ""
    End If
End Sub

Private Sub cmdEjemplo_Click()
    Dim Asunto As String
    Dim Cuerpo As String
    
    Asunto = CambioTexto(txtAsunto.Text, 1, True)
    Cuerpo = CambioTexto(txtCuerpo.Text, 1, True)
    
    MsgBox Asunto & vbNewLine & vbNewLine & Cuerpo, vbOKOnly, "Correo"
    
End Sub

Private Function CambioTexto(Texto As String, Linea As Long, Demo As Boolean) As String
    Dim Cabecera As String
    Dim IniCuerpo As String
    Dim FinCuerpo As String
    
    Cabecera = "<html>" & vbNewLine & _
               "<head>" & vbNewLine & _
               "<link rel=""STYLESHEET"" type=""text/css"" href=""http://www.rbimportadores.com/estiloemail.css"">" & vbNewLine & _
               "<base target=""_blank"">" & vbNewLine & _
               "</head>" & vbNewLine
    IniCuerpo = "<body>" & vbNewLine
    FinCuerpo = "</body>" & vbNewLine & _
                "</html>"
    
    Texto = Replace(Texto, "<cliente>", VSFG.TextMatrix(Linea, 1))
    Texto = Replace(Texto, "<email>", VSFG.TextMatrix(Linea, 0))
    If Demo = True Then
        Texto = Replace(Texto, "<hipervinculo>", "Click Aqui")
    Else
        Texto = Replace(Texto, "<hipervinculo>", "<a href=""" & txtHipervinculo.Text & """>Click Aqui</a>")
        Texto = Replace(Texto, vbNewLine, "<br>")
        Texto = Cabecera & IniCuerpo & Texto & FinCuerpo
    End If
    CambioTexto = Texto
End Function

Private Sub cmdEnviar_Click()
    Dim i As Long
    Dim j As Long
    Dim strDireccion As String
    Dim TotalRegistros As Long
    Dim ActRegistros As Long
    If VSFG.Rows > 1 Then
        If Trim(txtAsunto.Text) <> "" Then
            If Trim(txtCuerpo.Text) <> "" Then
                TotalRegistros = VSFG.Rows - 1
                For i = 1 To VSFG.Rows - 1
                    If VSFG.TextMatrix(i, 0) <> "" Then
                        VSFG.Select i, 0
                        VSFG.ShowCell i, 0
                        EnviarMail NombreComercial, CorreoNoticias, "Clientes " & NombreComercial, _
                                    VSFG.TextMatrix(i, 0), "", CambioTexto(txtAsunto.Text, i, True), _
                                    CambioTexto(txtCuerpo.Text, i, False), , True
                                    
                        mdiPrincipal.StatusBar.Panels(3).Text = i & "/" & TotalRegistros
                        mdiPrincipal.StatusBar.Refresh
                    End If
                Next i
                MsgBox "Se han enviado " & VSFG.Rows - 1 & " mensajes", vbInformation, "Envio de Correos"
            Else
                MsgBox "No tiene un cuerpo", vbInformation, "Envio de Correos"
            End If
        Else
            MsgBox "No tiene un asunto", vbInformation, "Envio de Correos"
        End If
    Else
        MsgBox "No tiene una lista de destinatarios", vbInformation, "Envio de Correos"
    End If
    mdiPrincipal.StatusBar.Panels(3).Text = ""
    mdiPrincipal.StatusBar.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    INICIO = False
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    CargaCombos
    CargaRed
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtListaCorreos" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtListaCorreos_LostFocus()
    Dim correos() As String
    Dim i As Long
    correos = Split(txtListaCorreos.Text, vbNewLine)
    VSFG.Rows = 1
    For i = 0 To UBound(correos)
        VSFG.AddItem correos(i)
    Next i
    txtListaCorreos.Visible = False
End Sub


Private Sub cmbN10_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN10.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref9")
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref8")
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN9_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN9.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref8")
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN8_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN8.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN7_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN7.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN6_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN6.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN5_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN5.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEjecutivo_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEjecutivo.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEmprendedor_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEmprendedor.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbDirector_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbDirector.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbGerente_Validate(Cancel As Boolean)
        cmbN10.BoundText = ""
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbDirector.BoundText = ""
End Sub

Private Sub CargaRed()
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_gz=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbGerente.RowSource = clsCon_Def.adorec_Def
    cmbGerente.BoundColumn = "codigo"
    cmbGerente.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_di=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbDirector.RowSource = clsCon_Def.adorec_Def
    cmbDirector.BoundColumn = "codigo"
    cmbDirector.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_em=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbEmprendedor.RowSource = clsCon_Def.adorec_Def
    cmbEmprendedor.BoundColumn = "codigo"
    cmbEmprendedor.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_ee=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbEjecutivo.RowSource = clsCon_Def.adorec_Def
    cmbEjecutivo.BoundColumn = "codigo"
    cmbEjecutivo.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n5=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN5.RowSource = clsCon_Def.adorec_Def
    cmbN5.BoundColumn = "codigo"
    cmbN5.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n6=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN6.RowSource = clsCon_Def.adorec_Def
    cmbN6.BoundColumn = "codigo"
    cmbN6.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n7=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN7.RowSource = clsCon_Def.adorec_Def
    cmbN7.BoundColumn = "codigo"
    cmbN7.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n8=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN8.RowSource = clsCon_Def.adorec_Def
    cmbN8.BoundColumn = "codigo"
    cmbN8.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n9=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN9.RowSource = clsCon_Def.adorec_Def
    cmbN9.BoundColumn = "codigo"
    cmbN9.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n10=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbN10.RowSource = clsCon_Def.adorec_Def
    cmbN10.BoundColumn = "codigo"
    cmbN10.ListField = "nombre"
End Sub

Private Sub CargaCombos()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    If Trim(strPtoFactura) = "" Then
        frmSelNegocio.Show vbModal
    End If
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
    
    
    strSql = " SELECT '%' as codigo,'- Todas las Ciudades -' as nombre UNION SELECT DISTINCT ciu_codigo as codigo, ciu_nombre AS nombre " & _
             " FROM ciudad " & _
             " ORDER BY nombre "
    clsCon_Def.Ejecutar strSql
    Set cmbCiudad.RowSource = clsCon_Def.adorec_Def
    cmbCiudad.BoundColumn = "codigo"
    cmbCiudad.ListField = "nombre"
    
End Sub

Private Sub optN1_Click()
    ActivarCombo
End Sub

Private Sub optN2_Click()
    ActivarCombo
End Sub

Private Sub optN3_Click()
    ActivarCombo
End Sub

Private Sub optN4_Click()
    ActivarCombo
End Sub

Private Sub optN5_Click()
    ActivarCombo
End Sub

Private Sub optN6_Click()
    ActivarCombo
End Sub

Private Sub optN7_Click()
    ActivarCombo
End Sub

Private Sub optN8_Click()
    ActivarCombo
End Sub

Private Sub optN9_Click()
    ActivarCombo
End Sub

Private Sub optN10_Click()
    ActivarCombo
End Sub

Private Sub ActivarCombo()
    cmbGerente.Locked = True
    cmbDirector.Locked = True
    cmbEmprendedor.Locked = True
    cmbEjecutivo.Locked = True
    cmbN5.Locked = True
    cmbN6.Locked = True
    cmbN7.Locked = True
    cmbN8.Locked = True
    cmbN9.Locked = True
    cmbN10.Locked = True
    If optN1.Value = True Then
        cmbGerente.Locked = False
    ElseIf optN2.Value = True Then
        cmbDirector.Locked = False
    ElseIf optN3.Value = True Then
        cmbEmprendedor.Locked = False
    ElseIf optN4.Value = True Then
        cmbEjecutivo.Locked = False
    ElseIf optN5.Value = True Then
        cmbN5.Locked = False
    ElseIf optN6.Value = True Then
        cmbN6.Locked = False
    ElseIf optN7.Value = True Then
        cmbN7.Locked = False
    ElseIf optN8.Value = True Then
        cmbN8.Locked = False
    ElseIf optN9.Value = True Then
        cmbN9.Locked = False
    ElseIf optN10.Value = True Then
        cmbN10.Locked = False
    End If
End Sub

