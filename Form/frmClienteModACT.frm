VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClienteModACT 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmClienteModACT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12285
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4392
      TabIndex        =   29
      Top             =   7080
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6192
      TabIndex        =   30
      Top             =   7080
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12105
      Begin VB.CheckBox chkN10 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkNegocio 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Negocio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkFiltroNombre 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkN9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkN8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkN7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkN6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkN5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkEjecutivo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkEmprendedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkDirector 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGerente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar N1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "CONSUMIDOR"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   3360
         TabIndex        =   27
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar CI/RUC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   315
         Left            =   8760
         TabIndex        =   16
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   315
         Left            =   8760
         TabIndex        =   18
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   315
         Left            =   8760
         TabIndex        =   20
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   315
         Left            =   8760
         TabIndex        =   22
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   315
         Left            =   8760
         TabIndex        =   24
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmbN10 
         Height          =   315
         Left            =   8760
         TabIndex        =   26
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3960
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   12060
      _cx             =   21272
      _cy             =   6985
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
      Cols            =   80
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClienteModACT.frx":030A
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
      FrozenRows      =   0
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
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   5760
      Picture         =   "frmClienteModACT.frx":0CB6
      ToolTipText     =   "Elimina una Fila"
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmClienteModACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private Tipo As String
Private Tipo2 As String
Public TIPOUsu As String
Private esMail As New vbSendMail.clsSendMail
Private Sub IniDato()
    Tipo = "Cliente"
    Tipo2 = "el Cliente"
    Me.Caption = Tipo
End Sub

Private Sub chkEjecutivo_Click()
    If chkEjecutivo.Value = 1 Then
        cmbEjecutivo.Enabled = True
    Else
        cmbEjecutivo.Enabled = False
    End If

End Sub

Private Sub chkEmprendedor_Click()
    If chkEmprendedor.Value = 1 Then
        cmbEmprendedor.Enabled = True
    Else
        cmbEmprendedor.Enabled = False
    End If

End Sub

Private Sub chkGerente_Click()
    If chkGerente.Value = 1 Then
        cmbGerente.Enabled = True
    Else
        cmbGerente.Enabled = False
    End If
End Sub

Private Sub chkDirector_Click()
    If chkDirector.Value = 1 Then
        cmbDirector.Enabled = True
    Else
        cmbDirector.Enabled = False
    End If
End Sub

Private Sub chkN5_Click()
    If chkN5.Value = 1 Then
        cmbN5.Enabled = True
    Else
        cmbN5.Enabled = False
    End If
End Sub

Private Sub chkN6_Click()
    If chkN6.Value = 1 Then
        cmbN6.Enabled = True
    Else
        cmbN6.Enabled = False
    End If
End Sub

Private Sub chkN7_Click()
    If chkN7.Value = 1 Then
        cmbN7.Enabled = True
    Else
        cmbN7.Enabled = False
    End If
End Sub

Private Sub chkN8_Click()
    If chkN8.Value = 1 Then
        cmbN8.Enabled = True
    Else
        cmbN8.Enabled = False
    End If
End Sub

Private Sub chkN9_Click()
    If chkN9.Value = 1 Then
        cmbN9.Enabled = True
    Else
        cmbN9.Enabled = False
    End If
End Sub

Private Sub chkN10_Click()
    If chkN10.Value = 1 Then
        cmbN10.Enabled = True
    Else
        cmbN10.Enabled = False
    End If
End Sub

Private Sub chkNegocio_Click()
    If chkNegocio.Value = 1 Then
        cmbNegocio.Enabled = True
    Else
        cmbNegocio.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    Carga False
End Sub
Private Sub Carga(Optional booTodo As Boolean = True)
    Dim clsAConsulta As New clsConsulta
    Dim clsACombo As New clsConsulta
    Dim clsACombo1 As New clsConsulta
    Dim clsACombo2 As New clsConsulta
    Dim clsACombo3 As New clsConsulta
    Dim clsACombo4 As New clsConsulta
    Dim clsACombo5 As New clsConsulta
    Dim clsACombo6 As New clsConsulta
    Dim clsACombo7 As New clsConsulta
    Dim clsACombo8 As New clsConsulta
    Dim clsACombo9 As New clsConsulta
    Dim clsACombo10 As New clsConsulta
    Dim clsACombo11 As New clsConsulta
    Dim clsACombo12 As New clsConsulta
    Dim clsACombo13 As New clsConsulta
    Dim clsACombo14 As New clsConsulta
    Dim clsACombo15 As New clsConsulta
    Dim clsACombo16 As New clsConsulta
    Dim clsACombo17 As New clsConsulta
    Dim clsACombo18 As New clsConsulta
    Dim clsACombo19 As New clsConsulta
    clsAConsulta.Inicializar AdoConn, AdoConnMaster
    clsACombo.Inicializar AdoConn, AdoConnMaster
    clsACombo1.Inicializar AdoConn, AdoConnMaster
    clsACombo2.Inicializar AdoConn, AdoConnMaster
    clsACombo3.Inicializar AdoConn, AdoConnMaster
    clsACombo4.Inicializar AdoConn, AdoConnMaster
    clsACombo5.Inicializar AdoConn, AdoConnMaster
    clsACombo6.Inicializar AdoConn, AdoConnMaster
    clsACombo7.Inicializar AdoConn, AdoConnMaster
    clsACombo8.Inicializar AdoConn, AdoConnMaster
    clsACombo9.Inicializar AdoConn, AdoConnMaster
    clsACombo10.Inicializar AdoConn, AdoConnMaster
    clsACombo11.Inicializar AdoConn, AdoConnMaster
    clsACombo12.Inicializar AdoConn, AdoConnMaster
    clsACombo13.Inicializar AdoConn, AdoConnMaster
    clsACombo14.Inicializar AdoConn, AdoConnMaster
    clsACombo15.Inicializar AdoConn, AdoConnMaster
    clsACombo16.Inicializar AdoConn, AdoConnMaster
    clsACombo17.Inicializar AdoConn, AdoConnMaster
    clsACombo18.Inicializar AdoConn, AdoConnMaster
    clsACombo19.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT per_codigo,per_cm,per_rcm,'' as red, " & _
             " per_tipo,per_apellido,per_nombre,COALESCE(per_sexo,'') as per_sex," & _
             " COALESCE(est_civ_codigo,'') as est_civ_codigo,COALESCE(ori_ing_codigo,'') as ori_ing_codigo,cat_p_codigo,fid_codigo,can_codigo,per_ruc,ciu_codigo," & _
             " dis_pol_codigo,per_codigo_postal,zon_codigo," & _
             " per_direccion_act,per_ubicacion_act,per_telf_act,per_fax_act,per_celular_act,per_email_act,per_fechacumplea,per_direccion2,for_ent_codigo," & _
             " per_credito,per_dcto,for_pag_codigo,for_pag_codigo_imp,per_pagare,tip_gar_codigo,per_garantiasolidariareal,per_nombregarante,gar_aut_codigo,per_codigo_resp,CONCAT(ven_codigo,'') as vend,tip_ped_codigo," & _
             " COALESCE(per_codigo_ref,''),COALESCE(per_codigo_ref2,''),COALESCE(per_codigo_ref3,''),COALESCE(per_codigo_ref4,''),COALESCE(" & _
             " per_codigo_ref5,''),COALESCE(per_codigo_ref6,''),COALESCE(per_codigo_ref7,''),COALESCE(per_codigo_ref8,''),COALESCE(per_codigo_ref9,''),COALESCE(per_codigo_ref10,'')," & _
             " per_observacion,'' as datos,per_fac_flete,per_especial,per_bloqueado,per_bloqueado_g," & _
             " per_sec_publico,per_siniva,per_inactivo,per_perdesde," & _
             " per_es_gz,per_es_di,per_es_em,per_es_ee," & _
             " per_es_n5,per_es_n6,per_es_n7,per_es_n8,per_es_n9,per_es_n10," & _
             " sac_codigo,cob_codigo,per_aplica_nc, per_aplica_ret,per_observacion_mod,per_fechamod,per_usumod,per_fechaing,per_usuing, '0' as modi " & _
             " FROM persona" & _
             " WHERE persona.emp_codigo ='" & strEmpresa & "'" & _
             " AND persona.cat_p_tipo='C'"
    If chkFiltroCodigo.Value = 1 Then
        strSql = strSql & "AND  per_ruc LIKE  '%" & txtCodigo.Text & "%'"
    End If
    If chkFiltroNombre.Value = 1 Then
        strSql = strSql & " AND CONCAT(per_apellido,' ',per_nombre) LIKE '%" & txtNombre.Text & "%' "
    End If
    If chkNegocio.Value = 1 Then
        strSql = strSql & " AND tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "' "
    End If
    If chkGerente.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref LIKE '" & cmbGerente.BoundText & "' "
    End If
    If chkDirector.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref2 LIKE '" & cmbDirector.BoundText & "' "
    End If
    If chkEmprendedor.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref3 LIKE '" & cmbEmprendedor.BoundText & "' "
    End If
    If chkEjecutivo.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref4 LIKE '" & cmbEjecutivo.BoundText & "' "
    End If
    If chkN5.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref5 LIKE '" & cmbN5.BoundText & "' "
    End If
    If chkN6.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref6 LIKE '" & cmbN6.BoundText & "' "
    End If
    If chkN7.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref7 LIKE '" & cmbN7.BoundText & "' "
    End If
    If chkN8.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref8 LIKE '" & cmbN8.BoundText & "' "
    End If
    If chkN9.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref9 LIKE '" & cmbN9.BoundText & "' "
    End If
    If chkN10.Value = 1 Then
        strSql = strSql & " AND per_codigo_ref10 LIKE '" & cmbN10.BoundText & "' "
    End If
    strSql = strSql & " ORDER BY CONCAT(per_apellido,' ',per_nombre) "
    clsAConsulta.Ejecutar strSql
    'Set VSFG.DataSource = Nothing
    Set VSFG.DataSource = clsAConsulta.adorec_Def.DataSource
    'If booTodo = True Then
        Set VSFG.CellButtonPicture = imgBtnUp
        VSFG.ColComboList(4) = "..."
        If VSFG.Rows > 1 Then
            VSFG.Cell(flexcpPicture, 1, 4, VSFG.Rows - 1, 4) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, 1, 4, VSFG.Rows - 1, 4) = flexPicAlignRightCenter
        End If
        'crea combo de estado civil
        strSql = " SELECT ' ' as est_civ_codigo, '  -  ' AS est_civ_nombre UNION " & _
                 " SELECT est_civ_codigo, est_civ_nombre" & _
                     " FROM estado_civil " & _
                     " ORDER BY est_civ_nombre"
         clsACombo.Ejecutar strSql
        VSFG.ColComboList(9) = VSFG.BuildComboList(clsACombo.adorec_Def, " *est_civ_nombre", "est_civ_codigo")
        'crea combo de origen de ingresos
        strSql = " SELECT ' ' as ori_ing_codigo, '  -  ' AS ori_ing_nombre UNION " & _
                 " SELECT ori_ing_codigo, ori_ing_nombre" & _
                     " FROM origen_ingresos " & _
                     " ORDER BY ori_ing_nombre"
         clsACombo.Ejecutar strSql
        VSFG.ColComboList(10) = VSFG.BuildComboList(clsACombo.adorec_Def, " *ori_ing_nombre", "ori_ing_codigo")
        'crea combo de categoria
        strSql = " SELECT cat_p_codigo, cat_p_nombre" & _
                     " FROM categoria_p " & _
                     " WHERE cat_p_tipo='C' " & _
                     " AND emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY cat_p_nombre"
         clsACombo.Ejecutar strSql
        VSFG.ColComboList(11) = VSFG.BuildComboList(clsACombo.adorec_Def, " *cat_p_nombre", "cat_p_codigo")
        
        'crea combo de tipo garantia
        strSql = " SELECT ' ' as tip_gar_codigo, '  -  ' AS tip_gar_nombre UNION " & _
                 " SELECT tip_gar_codigo, tip_gar_nombre" & _
                     " FROM tipo_garantia " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY tip_gar_nombre"
         clsACombo.Ejecutar strSql
        VSFG.ColComboList(33) = VSFG.BuildComboList(clsACombo.adorec_Def, " *tip_gar_nombre", "tip_gar_codigo")
        
        'crea combo de autiruzacion de garantia
        strSql = " SELECT ' ' as gar_aut_codigo, '  -  ' AS gar_aut_nombre UNION " & _
                 " SELECT gar_aut_codigo, gar_aut_nombre" & _
                     " FROM garantia_autorizada " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY gar_aut_nombre"
         clsACombo.Ejecutar strSql
        VSFG.ColComboList(36) = VSFG.BuildComboList(clsACombo.adorec_Def, " *gar_aut_nombre", "gar_aut_codigo")
        
        
        'crea combo de fidelizacion
        strSql = " SELECT fid_codigo, fid_nombre" & _
                     " FROM fidelizacion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY fid_nombre"
         clsACombo1.Ejecutar strSql
        VSFG.ColComboList(12) = VSFG.BuildComboList(clsACombo1.adorec_Def, " *fid_nombre", "fid_codigo")
        
        'crea combo de canal
        strSql = " SELECT can_codigo, can_nombre" & _
                     " FROM canal " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY can_nombre"
         clsACombo1.Ejecutar strSql
        VSFG.ColComboList(13) = VSFG.BuildComboList(clsACombo1.adorec_Def, " *can_nombre", "can_codigo")
        'crea combo de ciudad
        strSql = " SELECT ciu_codigo,pai_nombre,ciu_nombre" & _
                     " FROM ciudad INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo" & _
                     " ORDER BY pai_nombre,ciu_nombre"
         clsACombo2.Ejecutar strSql
        VSFG.ColComboList(15) = VSFG.BuildComboList(clsACombo2.adorec_Def, " pai_nombre, *ciu_nombre", "ciu_codigo")
        'crea combo de distribucion politica
        strSql = " SELECT dis_pol_codigo,dis_pol_nombre" & _
                     " FROM distribucion_politica" & _
                     " ORDER BY dis_pol_codigo,dis_pol_nombre"
         clsACombo2.Ejecutar strSql
        VSFG.ColComboList(16) = VSFG.BuildComboList(clsACombo2.adorec_Def, " dis_pol_codigo, *dis_pol_nombre", "dis_pol_codigo")
        'crea combo de zona
        strSql = " SELECT zon_codigo, zon_nombre" & _
                     " FROM zona " & _
                     " ORDER BY zon_nombre"
         clsACombo3.Ejecutar strSql
        VSFG.ColComboList(18) = VSFG.BuildComboList(clsACombo3.adorec_Def, " *zon_nombre", "zon_codigo")
        'crea combo de forma de entrega
        strSql = " SELECT for_ent_codigo, for_ent_nombre " & _
                     " FROM forma_entrega " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " ORDER BY for_ent_nombre "
        clsACombo4.Ejecutar strSql
        VSFG.ColComboList(27) = VSFG.BuildComboList(clsACombo4.adorec_Def, " *for_ent_nombre", "for_ent_codigo")
        'crea combo de forma de pago
        strSql = " SELECT for_pag_codigo, for_pag_nombre " & _
                     " FROM forma_pago " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " ORDER BY for_pag_nombre "
        clsACombo5.Ejecutar strSql
        VSFG.ColComboList(30) = VSFG.BuildComboList(clsACombo5.adorec_Def, " *for_pag_nombre", "for_pag_codigo")
        VSFG.ColComboList(31) = VSFG.BuildComboList(clsACombo5.adorec_Def, " *for_pag_nombre", "for_pag_codigo")
    
    
        'crea combo de tipo negocio
        strSql = " SELECT tip_ped_codigo, tip_ped_nombre " & _
                 " FROM tipo_pedido " & _
                 " ORDER BY tip_ped_nombre "
        clsACombo6.Ejecutar strSql
        VSFG.ColComboList(39) = VSFG.BuildComboList(clsACombo6.adorec_Def, " *tip_ped_nombre", "tip_ped_codigo")
        'crea combo de vendedor
        strSql = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as ven " & _
                     " FROM vendedor " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " ORDER BY CONCAT(ven_apellido,' ',ven_nombre) "
        clsACombo7.Ejecutar strSql
        VSFG.ColComboList(38) = VSFG.BuildComboList(clsACombo7.adorec_Def, " *ven", "ven_codigo")
        
        Set VSFG.CellButtonPicture = imgBtnUp
        VSFG.ColComboList(51) = "..."
        If VSFG.Rows > 1 Then
            VSFG.Cell(flexcpPicture, 1, 51, VSFG.Rows - 1, 51) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, 1, 51, VSFG.Rows - 1, 51) = flexPicAlignRightCenter
        End If
        
        'crea combo de sac
        strSql = " SELECT sac_codigo, CONCAT(sac_apellido,' ',sac_nombre) as sacn " & _
                     " FROM sac " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " ORDER BY CONCAT(sac_apellido,' ',sac_nombre) "
        clsACombo17.Ejecutar strSql
        VSFG.ColComboList(70) = VSFG.BuildComboList(clsACombo17.adorec_Def, " *sacn", "sac_codigo")
        
        'crea combo de cobrador
        strSql = " SELECT cob_codigo, CONCAT(cob_apellido,' ',cob_nombre) as cobn " & _
                     " FROM cobrador " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " ORDER BY CONCAT(cob_apellido,' ',cob_nombre) "
        clsACombo18.Ejecutar strSql
        VSFG.ColComboList(71) = VSFG.BuildComboList(clsACombo18.adorec_Def, " *cobn", "cob_codigo")
        
        
    'End If
    
    'combo de deudores
    
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND for_pag_codigo NOT IN ('CONT','EFE','CONCR')" & _
             " ORDER BY nombre "
    clsACombo8.Ejecutar strSql
    VSFG.ColComboList(37) = VSFG.BuildComboList(clsACombo8.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de gerente de zona
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_gz=1 " & _
             " ORDER BY nombre "
    clsACombo8.Ejecutar strSql
    VSFG.ColComboList(40) = VSFG.BuildComboList(clsACombo8.adorec_Def, " *nombre", "per_codigo")
    'crea combo de director
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_di=1 " & _
             " ORDER BY nombre "
    clsACombo9.Ejecutar strSql
    VSFG.ColComboList(41) = VSFG.BuildComboList(clsACombo9.adorec_Def, " *nombre", "per_codigo")
    'crea combo de emprendedor
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_em=1 " & _
             " ORDER BY nombre "
    clsACombo10.Ejecutar strSql
    VSFG.ColComboList(42) = VSFG.BuildComboList(clsACombo10.adorec_Def, " *nombre", "per_codigo")
    'crea combo de emprendedor
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_ee=1 " & _
             " ORDER BY nombre "
    clsACombo11.Ejecutar strSql
    VSFG.ColComboList(43) = VSFG.BuildComboList(clsACombo11.adorec_Def, " *nombre", "per_codigo")
    'crea combo de n5
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n5=1 " & _
             " ORDER BY nombre "
    clsACombo12.Ejecutar strSql
    VSFG.ColComboList(44) = VSFG.BuildComboList(clsACombo12.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de n6
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n6=1 " & _
             " ORDER BY nombre "
    clsACombo13.Ejecutar strSql
    VSFG.ColComboList(45) = VSFG.BuildComboList(clsACombo13.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de n7
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n7=1 " & _
             " ORDER BY nombre "
    clsACombo14.Ejecutar strSql
    VSFG.ColComboList(46) = VSFG.BuildComboList(clsACombo14.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de n8
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n8=1 " & _
             " ORDER BY nombre "
    clsACombo15.Ejecutar strSql
    VSFG.ColComboList(47) = VSFG.BuildComboList(clsACombo15.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de n9
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n9=1 " & _
             " ORDER BY nombre "
    clsACombo16.Ejecutar strSql
    VSFG.ColComboList(48) = VSFG.BuildComboList(clsACombo16.adorec_Def, " *nombre", "per_codigo")
    
    'crea combo de n10
    strSql = " SELECT ' ' as per_codigo, '  -  ' AS nombre UNION " & _
             " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',tip_ped_nombre,')') as nombre " & _
             " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C' AND per_es_n10=1 " & _
             " ORDER BY nombre "
    clsACombo19.Ejecutar strSql
    VSFG.ColComboList(49) = VSFG.BuildComboList(clsACombo19.adorec_Def, " *nombre", "per_codigo")
    
    ucrtVSFG.PonerNum
End Sub

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim control As Long 'control de que esten llenos los datos

    Dim emailCliente As String
    Dim emailPapacliente As String
    Dim emailN1cliente As String
    Dim ClienteNombre As String
      
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    control = 0 'inicializa control en 0
    Dim entra As Integer
    entra = 0
    For i = 1 To VSFG.Rows - 1
        'update
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            strSql = " SELECT count(per_ruc) " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 11))) & "' " & _
                     " AND tip_ped_codigo='" & VSFG.TextMatrix(i, 28) & "' " & _
                     " AND per_codigo!='" & VSFG.TextMatrix(i, 1) & "' "
            clsCon_Def.Ejecutar strSql
            
            If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                If MsgBox("El CI/RUC " & UCase(Trim(VSFG.TextMatrix(i, 14))) & " ya existe, desea continuar?", vbQuestion + vbYesNo, "Modificar") = vbNo Then
                    entra = 1
                    control = 1
                Else
                    entra = 0
                End If
            End If
            
            If entra = 0 Then
                strSql = " UPDATE persona " & _
                         " SET per_observacion_mod = CONCAT(COALESCE(per_observacion_mod,''),IF(per_bloqueado!='" & Abs(FormatoD0(VSFG.TextMatrix(i, 54))) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - CLIENTE ',IF(per_bloqueado=0,'BLOQUEADO CARTERA','DESBLOQUEADO CARTERA'),'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(per_bloqueado_g!='" & Abs(FormatoD0(VSFG.TextMatrix(i, 55))) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - CLIENTE ',IF(per_bloqueado=0,'BLOQUEADO GERENCIA','DESBLOQUEADO GERENCIA'),'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(for_pag_codigo!='" & VSFG.TextMatrix(i, 30) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - FORMA DE PAGO ANTERIOR ',for_pag_codigo,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(for_pag_codigo_imp!='" & VSFG.TextMatrix(i, 31) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - FORMA DE PAGO IMPRESION ANTERIOR ',for_pag_codigo_imp,'" & vbNewLine & "'), " & _
                                                                          " ''), "
                                                  strSql = strSql & " IF(per_codigo_resp!='" & VSFG.TextMatrix(i, 37) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - DEUDOR ANTERIOR ',per_codigo_resp,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(per_credito!='" & FormatoD2(VSFG.TextMatrix(i, 28)) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - CUPO DE CREDITO ANTERIOR ',per_credito,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(per_dcto!='" & FormatoD2(VSFG.TextMatrix(i, 29)) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - DCTO ANTERIOR ',per_dcto,'" & vbNewLine & "'), " & _
                                                                          " ''), "
                                                  strSql = strSql & " IF(per_pagare!='" & FormatoD2(VSFG.TextMatrix(i, 32)) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - PAGARE ANTERIOR ',per_pagare,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(tip_gar_codigo!='" & VSFG.TextMatrix(i, 33) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - TIPO GARANTIA ANTERIOR ',tip_gar_codigo,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(per_garantiasolidariareal!='" & UCase(VSFG.TextMatrix(i, 34)) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - GARANTIA SOLIDARIAREAL ANTERIOR ',per_garantiasolidariareal,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(per_nombregarante!='" & UCase(VSFG.TextMatrix(i, 35)) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - GARANTE ANTERIOR ',per_nombregarante,'" & vbNewLine & "'), " & _
                                                                          " ''), " & _
                                                                    " IF(gar_aut_codigo!='" & VSFG.TextMatrix(i, 36) & "'," & _
                                                                          " CONCAT(CURRENT_TIMESTAMP,' - " & strUsuario & "',' - AUTORIZACION POR ANTERIOR ',gar_aut_codigo,'" & vbNewLine & "'), " & _
                                                                          " '')) "
                strSql = strSql & " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C'"
                clsCon_Def.Ejecutar strSql, "M"
            
                strSql = " UPDATE persona " & _
                     " SET per_cm='" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "'," & _
                     " per_rcm='" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "'," & _
                     " per_tipo='" & VSFG.TextMatrix(i, 5) & "'," & _
                     " per_apellido='" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                     " per_nombre='" & UCase(VSFG.TextMatrix(i, 7)) & "'," & _
                     " per_sexo='" & UCase(Trim(VSFG.TextMatrix(i, 8))) & "'," & _
                     " est_civ_codigo='" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "'," & _
                     " ori_ing_codigo='" & UCase(Trim(VSFG.TextMatrix(i, 10))) & "'," & _
                     " cat_p_codigo='" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
                     " fid_codigo='" & UCase(VSFG.TextMatrix(i, 12)) & "'," & _
                     " can_codigo='" & UCase(VSFG.TextMatrix(i, 13)) & "'," & _
                     " per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 14))) & "'," & _
                     " ciu_codigo='" & UCase(VSFG.TextMatrix(i, 15)) & "'," & _
                     " dis_pol_codigo='" & UCase(VSFG.TextMatrix(i, 16)) & "'," & _
                     " per_codigo_postal='" & UCase(VSFG.TextMatrix(i, 17)) & "',"
                strSql = strSql & " zon_codigo='" & UCase(VSFG.TextMatrix(i, 18)) & "'," & _
                     " per_direccion_act='" & UCase(VSFG.TextMatrix(i, 19)) & "'," & _
                     " per_ubicacion_act='" & UCase(VSFG.TextMatrix(i, 20)) & "'," & _
                     " per_telf_act='" & UCase(VSFG.TextMatrix(i, 21)) & "'," & _
                     " per_fax_act='" & UCase(VSFG.TextMatrix(i, 22)) & "'," & _
                     " per_celular_act='" & UCase(VSFG.TextMatrix(i, 23)) & "'," & _
                     " per_email_act='" & VSFG.TextMatrix(i, 24) & "'," & _
                     " per_fechacumplea='" & VSFG.TextMatrix(i, 25) & "'," & _
                     " per_direccion2='" & UCase(VSFG.TextMatrix(i, 26)) & "'," & _
                     " for_ent_codigo='" & VSFG.TextMatrix(i, 27) & "'," & _
                     " per_credito='" & FormatoD2(VSFG.TextMatrix(i, 28)) & "'," & _
                     " per_dcto='" & FormatoD4(VSFG.TextMatrix(i, 29)) & "'," & _
                     " for_pag_codigo='" & VSFG.TextMatrix(i, 30) & "'," & _
                     " for_pag_codigo_imp='" & VSFG.TextMatrix(i, 31) & "',"
                strSql = strSql & " per_pagare='" & FormatoD2(VSFG.TextMatrix(i, 32)) & "'," & _
                     " tip_gar_codigo='" & Trim(Replace(VSFG.TextMatrix(i, 33), "-", "")) & "'," & _
                     " per_garantiasolidariareal='" & UCase(VSFG.TextMatrix(i, 34)) & "'," & _
                     " per_nombregarante='" & Trim(UCase(VSFG.TextMatrix(i, 35))) & "'," & _
                     " gar_aut_codigo='" & Trim(Replace(VSFG.TextMatrix(i, 36), "-", "")) & "'," & _
                     " per_codigo_resp='" & Trim(Replace(VSFG.TextMatrix(i, 37), "-", "")) & "',"

                strSql = strSql & " ven_codigo='" & VSFG.TextMatrix(i, 38) & "'," & _
                     " tip_ped_codigo='" & VSFG.TextMatrix(i, 39) & "'," & _
                     " per_codigo_ref='" & Trim(Replace(VSFG.TextMatrix(i, 40), "-", "")) & "'," & _
                     " per_codigo_ref2='" & Trim(Replace(VSFG.TextMatrix(i, 41), "-", "")) & "'," & _
                     " per_codigo_ref3='" & Trim(Replace(VSFG.TextMatrix(i, 42), "-", "")) & "'," & _
                     " per_codigo_ref4='" & Trim(Replace(VSFG.TextMatrix(i, 43), "-", "")) & "'," & _
                     " per_codigo_ref5='" & Trim(Replace(VSFG.TextMatrix(i, 44), "-", "")) & "'," & _
                     " per_codigo_ref6='" & Trim(Replace(VSFG.TextMatrix(i, 45), "-", "")) & "'," & _
                     " per_codigo_ref7='" & Trim(Replace(VSFG.TextMatrix(i, 46), "-", "")) & "'," & _
                     " per_codigo_ref8='" & Trim(Replace(VSFG.TextMatrix(i, 47), "-", "")) & "'," & _
                     " per_codigo_ref9='" & Trim(Replace(VSFG.TextMatrix(i, 48), "-", "")) & "'," & _
                     " per_codigo_ref10='" & Trim(Replace(VSFG.TextMatrix(i, 49), "-", "")) & "',"
                strSql = strSql & " per_observacion='" & UCase(VSFG.TextMatrix(i, 50)) & "'," & _
                     " per_fac_flete='" & Abs(FormatoD0(VSFG.TextMatrix(i, 52))) & "'," & _
                     " per_especial='" & Abs(FormatoD0(VSFG.TextMatrix(i, 53))) & "'," & _
                     " per_bloqueado='" & Abs(FormatoD0(VSFG.TextMatrix(i, 54))) & "'," & _
                     " per_bloqueado_g='" & Abs(FormatoD0(VSFG.TextMatrix(i, 55))) & "'," & _
                     " per_sec_publico='" & Abs(FormatoD0(VSFG.TextMatrix(i, 56))) & "'," & _
                     " per_siniva='" & Abs(FormatoD0(VSFG.TextMatrix(i, 57))) & "'," & _
                     " per_inactivo='" & Abs(FormatoD0(VSFG.TextMatrix(i, 58))) & "'," & _
                     " per_perdesde='" & VSFG.TextMatrix(i, 59) & "'," & _
                     " per_es_gz='" & Abs(FormatoD0(VSFG.TextMatrix(i, 60))) & "'," & _
                     " per_es_di='" & Abs(FormatoD0(VSFG.TextMatrix(i, 61))) & "'," & _
                     " per_es_em='" & Abs(FormatoD0(VSFG.TextMatrix(i, 62))) & "'," & _
                     " per_es_ee='" & Abs(FormatoD0(VSFG.TextMatrix(i, 63))) & "'," & _
                     " per_es_n5='" & Abs(FormatoD0(VSFG.TextMatrix(i, 64))) & "'," & _
                     " per_es_n6='" & Abs(FormatoD0(VSFG.TextMatrix(i, 65))) & "'," & _
                     " per_es_n7='" & Abs(FormatoD0(VSFG.TextMatrix(i, 66))) & "'," & _
                     " per_es_n8='" & Abs(FormatoD0(VSFG.TextMatrix(i, 67))) & "'," & _
                     " per_es_n9='" & Abs(FormatoD0(VSFG.TextMatrix(i, 68))) & "'," & _
                     " per_es_n10='" & Abs(FormatoD0(VSFG.TextMatrix(i, 69))) & "',"
                strSql = strSql & " sac_codigo='" & VSFG.TextMatrix(i, 70) & "'," & _
                     " cob_codigo='" & VSFG.TextMatrix(i, 71) & "'," & _
                     " per_aplica_nc='" & Abs(FormatoD0(VSFG.TextMatrix(i, 72))) & "'," & _
                     " per_aplica_ret='" & Abs(FormatoD0(VSFG.TextMatrix(i, 73))) & "'," & _
                     " per_fechamod=CURRENT_TIMESTAMP," & _
                     " per_usumod='" & strUsuario & "' "
                strSql = strSql & " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C'"
                clsCon_Def.Ejecutar strSql, "M"
            End If
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            'controla que este lleno los datos
            If VSFG.TextMatrix(i, 5) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el tipo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 6) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Nombre o Apellido", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 8) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Sexo", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 9) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Estado Civil", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 10) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Origen de Ingresos", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 11) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Categoria", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 12) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta categoria de fidelidad", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 13) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Canal", vbInformation, "Ingreso"
                control = 1
            ElseIf Trim(VSFG.TextMatrix(i, 14)) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el CI/RUC", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 15) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Ciudad", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 16) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Parroquia", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 17) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Código Postal", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 18) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Zona", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 30) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Forma de Pago", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 31) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta la Forma de Pago Impresion", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 38) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Vendedor", vbInformation, "Ingreso"
                control = 1
            ElseIf VSFG.TextMatrix(i, 39) = "" Then
                MsgBox "No puede ingresar " & Tipo2 & " falta el Tipo de Negocio", vbInformation, "Ingreso"
                control = 1
            Else
                strSql = " SELECT count(per_ruc) " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND tip_ped_codigo='" & VSFG.TextMatrix(i, 39) & "' " & _
                         " AND per_ruc='" & UCase(Trim(VSFG.TextMatrix(i, 14))) & "' "
                clsCon_Def.Ejecutar strSql
                
                If FormatoD0(clsCon_Def.adorec_Def(0)) > 0 Then
                    MsgBox "No puede ingresar " & Tipo2 & " el CI/RUC ya existe", vbInformation, "Ingreso"
                    control = 1
                Else
                    strSql = " SELECT CONCAT('C',LPAD(ROUND(COALESCE(MAX(REPLACE(per_codigo,'C','0')+0),0)+1,0),6,'0')) as cod " & _
                             " FROM persona " & _
                             " WHERE cat_p_tipo='C'" & _
                             " AND emp_codigo='" & strEmpresa & "'" & _
                             " GROUP BY emp_codigo"
                    clsCon_Def.Ejecutar strSql
                    VSFG.TextMatrix(i, 1) = clsCon_Def.adorec_Def("cod")
                    'controla que no se repita el código
                        strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido," & _
                                 " per_nombre,per_sexo,est_civ_codigo,ori_ing_codigo,cat_p_codigo,fid_codigo,can_codigo,per_ruc,ciu_codigo," & _
                                 " dis_pol_codigo,per_codigo_postal,zon_codigo," & _
                                 " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea," & _
                                 " per_direccion2,for_ent_codigo,per_credito,per_dcto,for_pag_codigo,for_pag_codigo_imp," & _
                                 " per_pagare,tip_gar_codigo,per_garantiasolidariareal,per_nombregarante,gar_aut_codigo," & _
                                 " per_codigo_resp,ven_codigo,tip_ped_codigo," & _
                                 " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                                 " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9,per_codigo_ref10," & _
                                 " per_observacion," & _
                                 " per_fac_flete,per_especial,per_bloqueado,per_bloqueado_g,per_sec_publico,per_siniva, " & _
                                 " per_inactivo,per_perdesde,per_es_gz,per_es_di,per_es_em,per_es_ee," & _
                                 " per_es_n5,per_es_n6,per_es_n7,per_es_n8,per_es_n9,per_es_n10," & _
                                 " sac_codigo,cob_codigo,per_aplica_nc,per_aplica_ret ," & _
                                 " per_fechamod,per_usumod,per_fechaing,per_usuing) "
                        strSql = strSql & " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 5) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & Trim(UCase(VSFG.TextMatrix(i, 8))) & "','" & Trim(UCase(VSFG.TextMatrix(i, 9))) & "','" & Trim(UCase(VSFG.TextMatrix(i, 10))) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "','" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 14))) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & UCase(VSFG.TextMatrix(i, 17)) & "','" & UCase(VSFG.TextMatrix(i, 18)) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & UCase(VSFG.TextMatrix(i, 20)) & "','" & UCase(VSFG.TextMatrix(i, 21)) & "','" & UCase(VSFG.TextMatrix(i, 22)) & "','" & UCase(VSFG.TextMatrix(i, 23)) & "','" & VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 26)) & "','" & VSFG.TextMatrix(i, 27) & "','" & FormatoD2(VSFG.TextMatrix(i, 28)) & "','" & FormatoD4(VSFG.TextMatrix(i, 29)) & "','" & VSFG.TextMatrix(i, 30) & "','" & VSFG.TextMatrix(i, 31) & "'," & _
                                 " '" & FormatoD2(VSFG.TextMatrix(i, 32)) & "','" & Trim(Replace(VSFG.TextMatrix(i, 33), "-", "")) & "','" & UCase(VSFG.TextMatrix(i, 34)) & "','" & UCase(VSFG.TextMatrix(i, 35)) & "','" & Trim(Replace(VSFG.TextMatrix(i, 36), "-", "")) & "'," & _
                                 " '" & Trim(Replace(VSFG.TextMatrix(i, 37), "-", "")) & "','" & VSFG.TextMatrix(i, 38) & "','" & VSFG.TextMatrix(i, 39) & "'," & _
                                 " '" & Trim(Replace(VSFG.TextMatrix(i, 40), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 41), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 42), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 43), "-", "")) & "'," & _
                                 " '" & Trim(Replace(VSFG.TextMatrix(i, 44), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 45), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 46), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 47), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 48), "-", "")) & "','" & Trim(Replace(VSFG.TextMatrix(i, 49), "-", "")) & "'," & _
                                 " '" & UCase(VSFG.TextMatrix(i, 50)) & "'," & _
                                 " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 52))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 53))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 54))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 55))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 56))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 57))) & "'," & _
                                 " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 58))) & "','" & VSFG.TextMatrix(i, 59) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 60))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 61))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 62))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 63))) & "'," & _
                                 " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 64))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 65))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 66))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 67))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 68))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 69))) & "'," & _
                                 " '" & VSFG.TextMatrix(i, 70) & "','" & VSFG.TextMatrix(i, 71) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 72))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 73))) & "'," & _
                                 " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
                        clsCon_Def.Ejecutar strSql, "M"
'                        If MsgBox("Desea registrar contactos al cliente" & vbNewLine & UCase(VSFG.TextMatrix(i, 3)) & " " & UCase(VSFG.TextMatrix(i, 4)) & "?", vbQuestion + vbYesNo, "Clientes") = vbYes Then
'                            frmContacto.CodPer = VSFG.TextMatrix(i, 1)
'                            frmContacto.Top = frmClienteMod.Top
'                            frmContacto.Show 1
'                        End If

                        'ENVIO DE CORREOS
                        
                        strSql = " SELECT COALESCE(persona.per_email,'') as per_email,CONCAT(COALESCE(persona.per_apellido,''),' ',COALESCE(persona.per_nombre,'')) as cli, " & _
                                 " COALESCE(IF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,N9.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,N8.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,N7.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,N6.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,N5.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,N4.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,N3.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,N2.per_email," & _
                                 " IF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,N1.per_email,''))))))))),'') as emailpapa," & _
                                 " COALESCE(N1.per_email,'') as emailn1"
                        strSql = strSql & " FROM persona " & _
                                 " LEFT JOIN persona as N1 ON N1.emp_codigo=persona.emp_codigo " & _
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
                                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                                 " WHERE persona.per_codigo='" & VSFG.TextMatrix(i, 1) & "' "
                        clsCon_Def.Ejecutar strSql
                        If clsCon_Def.adorec_Def.RecordCount > 0 Then
                            emailCliente = clsCon_Def.adorec_Def("per_email")
                            emailPapacliente = clsCon_Def.adorec_Def("emailpapa")
                            emailN1cliente = clsCon_Def.adorec_Def("emailn1")
                            ClienteNombre = clsCon_Def.adorec_Def("cli")
                        End If
                        
                        If Trim(emailCliente) & Trim(emailPapacliente) & Trim(emailN1cliente) <> "" Then
                            If Trim(emailPapacliente) <> "" Then
                                emailCliente = emailCliente & ";" & Trim(emailPapacliente)
                            End If
                            If Trim(emailN1cliente) <> "" Then
                                emailCliente = emailCliente & ";" & Trim(emailN1cliente)
                            End If
                        
                            EnviarMail NombreComercial & " Servicio al Cliente", CorreoServicioAlCliente, ClienteNombre, emailCliente, "", "Bienvenid@", _
                                "Estimado(a) Señor(a)" & vbNewLine & _
                                ClienteNombre & vbNewLine & _
                                "Le saludamos de R&B importadores, sus catálogos JSN y VPC, para darle la bienvenida a nuestra empresa y agradecerle por preferir nuestros productos." & vbNewLine & vbNewLine & _
                                "Conoce nuestro plan de crecimiento?  Le invitamos a visitar nuestra página web en la sección Venta directa / conoce el Multinivel." & vbNewLine & vbNewLine & _
                                "Si tiene cualquier inquietud, no dude en comunicarse con nosotros al 1800-CATALOGOS (228 256), www.rbimportadores.com o mediante nuestro Facebook: https://www.facebook.com/JsnEcuador" & vbNewLine & vbNewLine & _
                                "Gracias por preferirnos,  estamos para servirle!" & vbNewLine & _
                                "Servicio al Cliente" & vbNewLine & _
                                NombreComercial
                        End If

                End If
             End If
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        Carga False
    End If
    
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim Cedula As String
    If OldCol = 14 Then
        If Abs(Val(VSFG.TextMatrix(OldRow, VSFG.Cols - 1))) = 2 Or Abs(Val(VSFG.TextMatrix(OldRow, VSFG.Cols - 1))) = 3 Then
            Cedula = Trim(VSFG.TextMatrix(OldRow, OldCol))
            If VerificaCedula(Left(Cedula, 10)) = False Then
                Cancel = True
            Else
                VSFG.TextMatrix(OldRow, OldCol) = Cedula
            End If
        End If
    End If
End Sub

Private Sub VSFG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 51 Then
        frmContacto.CodPer = VSFG.TextMatrix(Row, 1)
        frmContacto.Top = frmClienteMod.Top
        frmContacto.Show 1
    ElseIf Col = 4 Then
        frmRed.CodPer = VSFG.TextMatrix(Row, 1)
        frmRed.Linea = Row
        frmRed.Top = frmClienteMod.Top
        frmRed.Show 1
    End If
    
End Sub

Private Sub VSFG_DblClick()
    Dim i As Long
    Set DAT = New frmDatos
    If VSFG.Row >= 1 Then
        DAT.Show
        DAT.VSFG.Rows = VSFG.Cols
        For i = 1 To VSFG.Cols - 1
            DAT.VSFG.TextMatrix(i, 0) = VSFG.TextMatrix(0, i)
            DAT.VSFG.Cell(flexcpText, i, 1) = VSFG.Cell(flexcpTextDisplay, VSFG.Row, i)
            If VSFG.ColComboList(i) <> "" Then
                DAT.VSFG.TextMatrix(i, 2) = VSFG.ColComboList(i)
                DAT.VSFG.Cell(flexcpText, i, 3) = VSFG.Cell(flexcpText, VSFG.Row, i)
            End If
        Next i
        DAT.VSFG.Cell(flexcpBackColor, 1, 1, DAT.VSFG.Rows - 1, 1) = VSFG.Cell(flexcpBackColor, VSFG.Row, VSFG.Col)
        DAT.VSFG.RowHidden(DAT.VSFG.Rows - 1) = True
        Set DAT.VSFGOrigen = VSFG
        DAT.VSFGOrigen.Tag = VSFG.Row
        DAT.Caption = Tipo
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        If Col <> 50 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If TIPOUsu = "AC" Or TIPOUsu = "SA" Then
            VSFG.TextMatrix(Row, 28) = 0
            VSFG.TextMatrix(Row, 29) = 0
            'VSFG.TextMatrix(Row, 29) = "CONT"
            'VSFG.TextMatrix(Row, 30) = "CONT"
            'VSFG.TextMatrix(Row, 11) = "EJE"
            VSFG.TextMatrix(Row, 12) = "000"
            VSFG.TextMatrix(Row, 72) = 0
            VSFG.TextMatrix(Row, 73) = 0
            If Col = 11 Or Col = 12 Or Col = 13 Or (26 <= Col And Col <= 49) Or (Col >= 53) Then
                Cancel = True
            End If
        ElseIf TIPOUsu = "PV" Or TIPOUsu = "PR" Or TIPOUsu = "PVG" Then
            VSFG.TextMatrix(Row, 28) = 0
            VSFG.TextMatrix(Row, 29) = 0
            VSFG.TextMatrix(Row, 30) = "CONT"
            VSFG.TextMatrix(Row, 31) = "CONT"
            If TIPOUsu = "PV" Then
                VSFG.TextMatrix(Row, 11) = "PVP"
                VSFG.TextMatrix(Row, 39) = "PDV"
            ElseIf TIPOUsu = "PVG" Then
                VSFG.TextMatrix(Row, 11) = "PVP"
                VSFG.TextMatrix(Row, 39) = "GPE"
            Else
                VSFG.TextMatrix(Row, 11) = "PRO"
                VSFG.TextMatrix(Row, 39) = "PRO"
            End If
            VSFG.TextMatrix(Row, 12) = "000"
            VSFG.TextMatrix(Row, 72) = 0
            VSFG.TextMatrix(Row, 73) = 0
            If Col = 2 Or Col = 11 Or Col = 12 Or Col = 28 Or Col = 29 Or Col = 30 Or Col = 39 Or (Col >= 53 And (Col <> 70 And Col <> 71)) Then
                Cancel = True
            End If
        Else
            If Col >= VSFG.Cols - 4 Then
                Cancel = True
            End If
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If TIPOUsu = "SA" Then
            If Col = 1 Or Col = 11 Or Col = 12 Or Col = 13 Or (Col >= 28 And Col <= 43) Or (Col >= 53) Or Col >= VSFG.Cols - 6 Then
                Cancel = True
            End If
        ElseIf TIPOUsu = "PV" Or TIPOUsu = "PR" Or TIPOUsu = "PVG" Then
            If Col = 1 Or Col = 3 Or Col = 11 Or Col = 12 Or Col = 13 Or Col = 39 Or (Col >= 28 And Col <= 43) Or (Col >= 53) Or Col >= VSFG.Cols - 6 Then
                Cancel = True
            End If
        ElseIf TIPOUsu = "CA" Then
            If Not (Col = 9 Or Col = 10 Or Col = 15 Or Col = 16 Or Col = 19 Or Col = 20 Or (21 <= Col And Col <= 24) Or Col = 28 Or (30 <= Col And Col <= 37) Or Col = 50 Or Col = 54 Or Col = 71) Then
                Cancel = True
            End If
        ElseIf TIPOUsu = "GE" Then
            If Not (Col = 9 Or Col = 10 Or Col = 15 Or Col = 16 Or Col = 19 Or Col = 20 Or (21 <= Col And Col <= 24) Or Col = 28 Or (30 <= Col And Col <= 37) Or Col = 50 Or Col = 55) Then
                Cancel = True
            End If
        Else
            If Col = 1 Or (Col = 59 And VSFG.TextMatrix(Row, 59) <> "") Or Col >= VSFG.Cols - 6 Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub chkFiltroNombre_Click()
    If chkFiltroNombre.Value = 1 Then
        txtNombre.Enabled = True
    Else
        txtNombre.Enabled = False
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.Value = 1 Then
        txtCodigo.Enabled = True
    Else
        txtCodigo.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    On Error Resume Next
    Unload frmContacto
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
    Dim clsAComFiltro As New clsConsulta
    Dim clsAComFiltro1 As New clsConsulta
    Dim clsAComFiltro2 As New clsConsulta
    Dim clsAComFiltro3 As New clsConsulta
    Dim clsAComFiltro4 As New clsConsulta
    Dim clsAComFiltro5 As New clsConsulta
    Dim clsAComFiltro6 As New clsConsulta
    Dim clsAComFiltro7 As New clsConsulta
    Dim clsAComFiltro8 As New clsConsulta
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro1.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro2.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro3.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro4.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro5.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro6.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro7.Inicializar AdoConn, AdoConnMaster
    clsAComFiltro8.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    If TIPOUsu = "AC" Or TIPOUsu = "JS" _
        Or TIPOUsu = "JC" Or TIPOUsu = "GE" Then
        ucrtVSFG.Inicializar False, False, True, False, False, False, False, False, True
    ElseIf TIPOUsu = "CA" Then
        ucrtVSFG.Inicializar False, False, True, False, False, True, False, True, True, "190"
    ElseIf TIPOUsu = "SU" Or TIPOUsu = "SA" Or TIPOUsu = "PV" Or TIPOUsu = "PVG" Or TIPOUsu = "PR" Then
        ucrtVSFG.Inicializar False, False
    End If
    IniDato
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    
    chkNegocio.Enabled = True
    cmbNegocio.Enabled = True
    
    If TIPOUsu = "PV" Then
        chkNegocio.Value = 1
        cmbNegocio.BoundText = "PDV"
        chkNegocio.Enabled = False
        cmbNegocio.Enabled = False
    ElseIf TIPOUsu = "PVG" Then
        chkNegocio.Value = 1
        cmbNegocio.BoundText = "GPE"
        chkNegocio.Enabled = False
        cmbNegocio.Enabled = False
    ElseIf TIPOUsu = "PR" Then
        chkNegocio.Value = 1
        cmbNegocio.BoundText = "PRO"
        chkNegocio.Enabled = False
        cmbNegocio.Enabled = False
    End If
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_gz=1" & _
             " ORDER BY nombre "
    clsAComFiltro.Ejecutar strSql
    Set cmbGerente.RowSource = clsAComFiltro.adorec_Def
    cmbGerente.BoundColumn = "codigo"
    cmbGerente.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_di=1" & _
             " ORDER BY nombre "
    clsAComFiltro1.Ejecutar strSql
    Set cmbDirector.RowSource = clsAComFiltro1.adorec_Def
    cmbDirector.BoundColumn = "codigo"
    cmbDirector.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_em=1" & _
             " ORDER BY nombre "
    clsAComFiltro2.Ejecutar strSql
    Set cmbEmprendedor.RowSource = clsAComFiltro2.adorec_Def
    cmbEmprendedor.BoundColumn = "codigo"
    cmbEmprendedor.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_ee=1" & _
             " ORDER BY nombre "
    clsAComFiltro3.Ejecutar strSql
    Set cmbEjecutivo.RowSource = clsAComFiltro3.adorec_Def
    cmbEjecutivo.BoundColumn = "codigo"
    cmbEjecutivo.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n5=1" & _
             " ORDER BY nombre "
    clsAComFiltro4.Ejecutar strSql
    Set cmbN5.RowSource = clsAComFiltro4.adorec_Def
    cmbN5.BoundColumn = "codigo"
    cmbN5.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n6=1" & _
             " ORDER BY nombre "
    clsAComFiltro5.Ejecutar strSql
    Set cmbN6.RowSource = clsAComFiltro5.adorec_Def
    cmbN6.BoundColumn = "codigo"
    cmbN6.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n7=1" & _
             " ORDER BY nombre "
    clsAComFiltro6.Ejecutar strSql
    Set cmbN7.RowSource = clsAComFiltro6.adorec_Def
    cmbN7.BoundColumn = "codigo"
    cmbN7.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n8=1" & _
             " ORDER BY nombre "
    clsAComFiltro7.Ejecutar strSql
    Set cmbN8.RowSource = clsAComFiltro7.adorec_Def
    cmbN8.BoundColumn = "codigo"
    cmbN8.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n9=1" & _
             " ORDER BY nombre "
    clsAComFiltro8.Ejecutar strSql
    Set cmbN9.RowSource = clsAComFiltro8.adorec_Def
    cmbN9.BoundColumn = "codigo"
    cmbN9.ListField = "nombre"
    
    strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
             " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
             " AND p1.cat_p_tipo='C'" & _
             " AND p1.per_es_n10=1" & _
             " ORDER BY nombre "
    clsAComFiltro8.Ejecutar strSql
    Set cmbN10.RowSource = clsAComFiltro8.adorec_Def
    cmbN10.BoundColumn = "codigo"
    cmbN10.ListField = "nombre"
  
    Carga
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
        If VSFG.TextMatrix(Row, 25) = "" Then
            VSFG.TextMatrix(Row, 25) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 59) = "" Then
            VSFG.TextMatrix(Row, 59) = HoyDia
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
        If VSFG.TextMatrix(Row, 59) = "" Then
            VSFG.TextMatrix(Row, 59) = HoyDia
        End If
        If VSFG.TextMatrix(Row, 25) = "" Then
            VSFG.TextMatrix(Row, 25) = HoyDia
        End If
    End If
    If Col = 25 Then
        If Not IsDate(VSFG.TextMatrix(Row, 25)) Then
            VSFG.TextMatrix(Row, 25) = HoyDia
        End If
    End If
    
    If Col = 25 Then
        If Not IsDate(VSFG.TextMatrix(Row, 25)) Then
            VSFG.TextMatrix(Row, 25) = HoyDia
        End If
    End If
    
    If Col = 24 And VSFG.TextMatrix(Row, Col) <> "" Then
        
        If revisarEmail(VSFG.TextMatrix(Row, Col)) = False Then
            MsgBox "El email no tiene un formato valido", vbInformation, "Email"
            VSFG.TextMatrix(Row, Col) = ""
        End If
    End If
    
    If Col = 61 Then
        If Not IsDate(VSFG.TextMatrix(Row, 59)) Then
            VSFG.TextMatrix(Row, 59) = HoyDia
        End If
    End If
    If Col = 14 Then
        If VSFG.TextMatrix(Row, 14) <> "" And Not IsNumeric(VSFG.TextMatrix(Row, 14)) Then
            If MsgBox("Está ingresando en el campo CI/RUC valores no numéricos, desea continuar?", vbQuestion + vbYesNo, "CI/RUC") = vbNo Then
                VSFG.TextMatrix(Row, 14) = ""
            End If
        End If
    End If
End Sub

Private Sub VSFG_EnterCell()
    If VSFG.Col = 59 And VSFG.TextMatrix(VSFG.Row, VSFG.Col) = "" And (Val(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1)) = -2) Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Col) = HoyDia
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    If TIPOUsu = "SU" Or TIPOUsu = "PR" Or TIPOUsu = "CA" Then
        ucrtVSFG.Editar KeyAscii
    End If
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
