VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmContabilizarAdq 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilizar Adquisiciones"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmContabilizarAdq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10785
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Contabilizar Adquisiciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5880
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Adquisicion"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4575
         Begin VB.ComboBox cmbAñoF 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":030A
            Left            =   1800
            List            =   "frmContabilizarAdq.frx":036B
            TabIndex        =   3
            Text            =   "AÑO"
            Top             =   600
            Width           =   780
         End
         Begin VB.ComboBox cmbMesF 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":0429
            Left            =   2640
            List            =   "frmContabilizarAdq.frx":0454
            TabIndex        =   4
            Text            =   "MES"
            Top             =   600
            Width           =   780
         End
         Begin VB.ComboBox cmbDiaF 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":0494
            Left            =   3480
            List            =   "frmContabilizarAdq.frx":04F5
            TabIndex        =   5
            Text            =   "DIA"
            Top             =   600
            Width           =   780
         End
         Begin VB.ComboBox cmbAñoI 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":056C
            Left            =   1800
            List            =   "frmContabilizarAdq.frx":05CD
            TabIndex        =   0
            Text            =   "AÑO"
            Top             =   240
            Width           =   780
         End
         Begin VB.ComboBox cmbMesI 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":068B
            Left            =   2640
            List            =   "frmContabilizarAdq.frx":06B6
            TabIndex        =   1
            Text            =   "MES"
            Top             =   240
            Width           =   780
         End
         Begin VB.ComboBox cmbDiaI 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":06F6
            Left            =   3480
            List            =   "frmContabilizarAdq.frx":0757
            TabIndex        =   2
            Text            =   "DIA"
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Fin:"
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
            TabIndex        =   23
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inicio:"
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
            TabIndex        =   22
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Asiento"
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
         Height          =   1095
         Left            =   4920
         TabIndex        =   19
         Top             =   360
         Width           =   4335
         Begin VB.ComboBox cmbAñoA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":07CE
            Left            =   1320
            List            =   "frmContabilizarAdq.frx":082F
            TabIndex        =   6
            Text            =   "AÑO"
            Top             =   360
            Width           =   780
         End
         Begin VB.ComboBox cmbMesA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":08ED
            Left            =   2160
            List            =   "frmContabilizarAdq.frx":0918
            TabIndex        =   7
            Text            =   "MES"
            Top             =   360
            Width           =   780
         End
         Begin VB.ComboBox cmbDiaA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmContabilizarAdq.frx":0958
            Left            =   3000
            List            =   "frmContabilizarAdq.frx":09B9
            TabIndex        =   8
            Text            =   "DIA"
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Left            =   600
            TabIndex        =   20
            Top             =   420
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGdb 
         Height          =   2295
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   5535
         _cx             =   9763
         _cy             =   4048
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
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmContabilizarAdq.frx":0A30
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGAdquisicion 
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   10335
         _cx             =   18230
         _cy             =   2990
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
         BackColorSel    =   8388608
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmContabilizarAdq.frx":0ADC
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
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Height          =   210
         Left            =   5400
         TabIndex        =   25
         Top             =   3270
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Height          =   210
         Left            =   2520
         TabIndex        =   24
         Top             =   5910
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   3642
      TabIndex        =   16
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   5442
      TabIndex        =   17
      Top             =   6480
      Width           =   1700
   End
End
Attribute VB_Name = "frmContabilizarAdq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para Contabilizar de Adquisiciones                                    #
'#  frmContabilizarAdq V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el asiento adquisición                                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  adquisicion: Esta tabla contiene los datos de la adquisicion                #
'#  aisento: tabla donde se almace los asientoa                                 #
'#  tipo de aisento: donde se guardan los datos tipos de asiento                #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'


Private clsAdq As New clsConsulta
Private clsAsi As New clsConsulta
Private clsPro As New clsConsulta
Private clsAct As New clsConsulta
Private clsSum As New clsConsulta
Private clsMaxAsi As New clsConsulta
Private clsSql As New clsConsulta
Dim strSql As String
Dim j As Integer
Dim ban As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsAdq = Nothing
    Set clsAsi = Nothing
    Set clsPro = Nothing
    Set clsAct = Nothing
    Set clsSum = Nothing
    Set clsMaxAsi = Nothing
    Set clsSql = Nothing
End Sub
Private Sub PonerNumeros(Optional conBot As Boolean = True)
    For i = 1 To (VSFGAdquisicion.Rows - 1)
        VSFGAdquisicion.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub DebeHaber()
    Dim count As Integer
    count = 0
    
    VSFGdb.Clear 1
    VSFGdb.Rows = 1
    For k = 1 To (VSFGAdquisicion.Rows - 1)
        If VSFGAdquisicion.TextMatrix(k, 1) = -1 Then
               
        strSql = " SELECT par_texto,par_nombre,'0.00' as debe,adq_subtotal " & _
                 " FROM ((adquisicion " & _
                 " INNER JOIN  persona " & _
                 " ON adquisicion.per_codigo = persona.per_codigo " & _
                 " AND adquisicion.emp_codigo = persona.emp_codigo ) " & _
                 " INNER JOIN parametro " & _
                 " ON persona.emp_codigo = parametro.emp_codigo) " & _
                 " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                 " AND cat_p_tipo = 'P' " & _
                 "  AND par_codigo = 'CXP' " & _
                 " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                 " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "

        clsPro.Ejecutar (strSql)
        j = clsPro.adorec_Def.RecordCount
            If (clsPro.adorec_Def.RecordCount > 0) Then
            clsPro.adorec_Def.MoveFirst
            'Set VSFGdb.DataSource = clsAct.adorec_Def.DataSource
             ' For i = 1 To j
                    VSFGdb.Rows = VSFGdb.Rows + 1
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsPro.adorec_Def("par_texto")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsPro.adorec_Def("par_nombre")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsPro.adorec_Def("debe")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsPro.adorec_Def("adq_subtotal")
            ' Next i
            End If
        
        strSql = " SELECT tip_act_ctaconta,cta_nombre,act_fij_valor,'0.00'as ha " & _
                " FROM ((((adquisicion " & _
                " INNER JOIN  det_adquisicion_af " & _
                " ON adquisicion.adq_codigo = det_adquisicion_af.adq_codigo " & _
                " AND adquisicion.emp_codigo = det_adquisicion_af.emp_codigo ) " & _
                " INNER JOIN activo_fijo " & _
                " ON det_adquisicion_af.emp_codigo = activo_fijo.emp_codigo " & _
                " AND activo_fijo.act_fij_codigo = det_adquisicion_af.act_fij_codigo ) " & _
                " INNER JOIN tipo_activo " & _
                " ON tipo_activo.tip_act_codigo = activo_fijo.tip_act_codigo " & _
                " AND tipo_activo.emp_codigo = activo_fijo.emp_codigo ) " & _
                " INNER JOIN ctaconta " & _
                " ON tipo_activo.tip_act_ctaconta = ctaconta.cta_codigo " & _
                " AND tipo_activo.emp_codigo = ctaconta.emp_codigo ) " & _
                " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "
        clsAct.Ejecutar (strSql)

        If (clsAct.adorec_Def.RecordCount > 0) Then
            
            For m = 1 To clsAct.adorec_Def.RecordCount
                VSFGdb.Rows = VSFGdb.Rows + 1
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsAct.adorec_Def("tip_act_ctaconta")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsAct.adorec_Def("cta_nombre")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsAct.adorec_Def("act_fij_valor")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsAct.adorec_Def("ha")
                clsAct.adorec_Def.MoveNext
            Next m
        End If

       strSql = " SELECT tip_sum_ctaconta,cta_nombre,(det_adq_su_cantidad*sum_ultimo_precio) as deb,'0.00'as hab " & _
                " FROM ((((adquisicion " & _
                " INNER JOIN  det_adquisicion_su " & _
                " ON adquisicion.adq_codigo = det_adquisicion_su.adq_codigo " & _
                " AND adquisicion.emp_codigo = det_adquisicion_su.emp_codigo ) " & _
                " INNER JOIN suministro " & _
                " ON det_adquisicion_su.emp_codigo = suministro.emp_codigo " & _
                " AND suministro.sum_codigo = det_adquisicion_su.sum_codigo ) " & _
                " INNER JOIN tipo_suministro " & _
                " ON tipo_suministro.tip_sum_codigo = suministro.tip_sum_codigo " & _
                " AND tipo_suministro.emp_codigo = suministro.emp_codigo ) " & _
                " INNER JOIN ctaconta " & _
                " ON tipo_suministro.tip_sum_ctaconta = ctaconta.cta_codigo " & _
                " AND tipo_suministro.emp_codigo = ctaconta.emp_codigo ) " & _
                " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "
        clsSum.Ejecutar (strSql)

        If (clsSum.adorec_Def.RecordCount > 0) Then
            For n = 1 To clsSum.adorec_Def.RecordCount
                VSFGdb.Rows = VSFGdb.Rows + 1
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsSum.adorec_Def("tip_sum_ctaconta")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsSum.adorec_Def("cta_nombre")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsSum.adorec_Def("deb")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsSum.adorec_Def("hab")
                clsSum.adorec_Def.MoveNext
            Next n
        End If
    End If
            If VSFGAdquisicion.TextMatrix(k, 1) = "-1" Then
                count = count + 1
            End If
   Next k
   'abilita el boton aceptar
        If count <= 0 Then
                cmdAceptar.Enabled = False
            Else
                cmdAceptar.Enabled = True
        End If
End Sub
Private Sub CalTotalDebeHaber()

   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    'Calcula total debe
    For i = 1 To VSFGdb.Rows - 1
        SumaDebe = SumaDebe + Val(VSFGdb.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    'Calcula total haber
    For i = 1 To VSFGdb.Rows - 1
        SumaHaber = SumaHaber + Val(VSFGdb.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")

End Sub
Private Sub Rango_Fecha()
   'Ejectua el Selct con el rango de fechas deseado
    fi = Format(cmbAñoI.Text + "-" + LCase(cmbMesI.Text) + "-" + cmbDiaI.Text, "yyyy-mm-dd")
    ff = Format(cmbAñoF.Text + "-" + LCase(cmbMesF.Text) + "-" + cmbDiaF.Text, "yyyy-mm-dd")
    'If cmbDiaA.Tag <> "A" Then
        'Verifican si las fecha ingresadas son correctas
        If (IsDate(fi)) = False Then
            MsgBox "La fecha de Inicio no es correcta", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
            'cmbAñoI.SetFocus
            Exit Sub
        End If
        If (IsDate(ff)) = False Then
            MsgBox "La fecha de Fin no es correcta", vbExclamation, "SisAdmi - Contabilizar Adquisiciones "
            'cmbAñoF.SetFocus
            Exit Sub
        End If
'    Else
'        Exit Sub
'    End If

    If ff >= fi Then
    'llenar flexgrid
        strSql = " SELECT '0'as sel,adq_codigo, adq_numdoc,adq_fecha, concat(per_apellido,' ',per_nombre)as proveedor,adq_subtotal,adq_impuesto,'' as tot,concat(substring(adq_observacion,1,10),'...') as obs " & _
                 " FROM adquisicion  INNER JOIN  persona  ON adquisicion.per_codigo = persona.per_codigo " & _
                 " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
                 " AND adq_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND adq_asentada = '0' " & _
                 " ORDER BY adq_codigo "
        clsAdq.Ejecutar (strSql)
        
        If (clsAdq.adorec_Def.RecordCount > 0) Then
            TxtSubTotal.Text = 0
            TxtIva.Text = 0
            TxtTotal.Text = 0
            ban = 0
            Set VSFGAdquisicion.DataSource = clsAdq.adorec_Def.DataSource
            VSFGAdquisicion.ColDataType(1) = flexDTBoolean
            ban = 1
            Call PonerNumeros
            Call Cal_Total
        Else
            MsgBox "No hay Adquisiciones ingresadas en el sistema", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
            Call limpiarFxGD
            TxtSubTotal.Text = ""
            TxtIva.Text = ""
            TxtTotal.Text = ""
        End If
    Else
        MsgBox "La Fecha Fin es mayor que la Fecha Inicio, Verifiquela por Favor!", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
        Exit Sub
    End If
      
End Sub
Private Sub Cal_Total()
   'Calcula totales del grid de adquisicion
    Dim Subtotal As Double
    Dim IVA As Double
    Dim Total As Double
    Subtotal = 0
    IVA = 0
    Total = 0
    For i = 1 To VSFGAdquisicion.Rows - 1
        VSFGAdquisicion.TextMatrix(i, 8) = (Val(VSFGAdquisicion.TextMatrix(i, 6)) + Val(VSFGAdquisicion.TextMatrix(i, 7)))
    Next i
    For i = 1 To VSFGAdquisicion.Rows - 1
        If VSFGAdquisicion.TextMatrix(i, 1) = -1 Then
            Subtotal = Subtotal + (Val(VSFGAdquisicion.TextMatrix(i, 6)))
            IVA = IVA + (Val(VSFGAdquisicion.TextMatrix(i, 7)))
            Total = Total + (Val(VSFGAdquisicion.TextMatrix(i, 8)))
         End If
    Next i
    TxtSubTotal.Text = FormatoD2(Subtotal)
    TxtIva.Text = FormatoD2(IVA)
    TxtTotal.Text = FormatoD2(Total)
End Sub
Private Sub limpiar()
    VSFGAdquisicion.Clear 1
    VSFGAdquisicion.Rows = 2
    VSFGdb.Clear 1
    VSFGdb.Rows = 2
    TxtSubTotal = 0
    TxtIva = 0
    TxtTotal = 0
    txtTotalHaber = 0
    txtTotalDebe = 0
End Sub

Private Sub cmdAceptar_Click()
    'Comprueba que todos los datos esten ingresados
    fa = Format(cmbAñoA.Text + "-" + cmbMesA.Text + "-" + cmbDiaA.Text, "yyyy-mm-dd")
    If (IsDate(fa) = False) Then
        MsgBox "La fecha de Asiento no es válida", vbInformation, "SisAdmi-Comprobante Aquisición"
        cmbAñoA.SetFocus
        Exit Sub
    End If
    'verifica que el debe y el haber esten cuadrados
'    If txtTotalDebe.Text <> txtTotalHaber.Text Then
'        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "SisAdmi-Comprobante Aquisición"
'    End If
    'Compacta la matriz
    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el grid debe haber
    a = VSFGdb.Rows - 1
    For i = 1 To a
        For j = i + 1 To a
            If VSFGdb.TextMatrix(i, 1) = VSFGdb.TextMatrix(j, 1) Then
                VSFGdb.TextMatrix(i, 3) = Val(VSFGdb.TextMatrix(i, 3)) + Val(VSFGdb.TextMatrix(j, 3))
                VSFGdb.TextMatrix(i, 4) = Val(VSFGdb.TextMatrix(i, 4)) + Val(VSFGdb.TextMatrix(j, 4))
                VSFGdb.RemoveItem j
                a = a - 1
                j = j - 1
            End If
            If j >= a Then
                Exit For
            End If
        Next j
    Next i
        'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If VSFGdb.TextMatrix(1, 1) = "" Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "SisAdmi-Comprobante Aquisición"
            Exit Sub
            Else
            
            Dim NumMes As String
            NumMes = cmbMesA.ListIndex + 1
            If (Len(NumMes) = 1) Then
                NumMes = "0" & NumMes
            End If
            
            'Busca el código máximo de la tabla asiento
            strSql = " Select max(SUBSTRING(asi_numasiento,7,10)) as numAs " & _
                    " From asiento " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND asi_numasiento LIKE '" & cmbAñoA.Text & NumMes & "%'" & _
                    " GROUP BY emp_codigo"
            clsSql.Ejecutar strSql
            If Not IsNull(clsSql.adorec_Def("numAs")) Then
                maximo = clsSql.adorec_Def("numAs") + 1
            Else
                maximo = 1
            End If
            While (Len(maximo) < 4)
                maximo = "0" & maximo
            Wend
            strMaximo = cmbAñoA.Text & NumMes & maximo
            
            Dim Des As String
            'Actualiza el campo adq_asentada poniendo "1"
            For i = 1 To (VSFGAdquisicion.Rows - 1)
                If VSFGAdquisicion.TextMatrix(i, 1) = -1 Then
                    strSql = " UPDATE adquisicion " & _
                        " SET adq_asentada = '1'" & _
                        " WHERE emp_codigo = '" & strEmpresa & "' " & _
                        " AND adq_codigo = '" & VSFGAdquisicion.TextMatrix(i, 2) & "'"
                    clsSql.Ejecutar strSql, "M"
                Des = Des & VSFGAdquisicion.TextMatrix(i, 2) & "(" & VSFGAdquisicion.TextMatrix(i, 3) & ")/ "
                End If
            Next i
            'Inserta los datos en el asiento del egreso
            strSql = " INSERT INTO asiento (asi_numasiento, emp_codigo, asi_fecha, asi_revisado, asi_mayorizado, asi_totaldebe, asi_totalhaber, asi_descripcion, asi_fechamod, asi_usumod) " & _
                     " VALUES ('" & strMaximo & "','" & strEmpresa & "', '" & fa & "', '0','0', '" & FormatoD2(txtTotalDebe) & "', '" & FormatoD2(txtTotalHaber) & "', '" & "ASIENTO DE ADQUISICION(ES):" & Des & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsSql.Ejecutar strSql, "M"

            With VSFGdb
                For i = 1 To .Rows - 1
                    'Ingresa el detalle del asiento del comprobante
                    If .TextMatrix(i, 1) = "" Then
                        Exit For
                    Else
                        strSql = " INSERT INTO det_asiento (emp_codigo, asi_numasiento, cta_codigo, det_asi_debe, det_asi_haber, det_asi_fechamod, det_asi_usumod) " & _
                                 " VALUES ('" & strEmpresa & "','" & strMaximo & "','" & .TextMatrix(i, 1) & "','" & Replace(Val(.TextMatrix(i, 3)), ",", ".") & "', '" & Replace(Val(.TextMatrix(i, 4)), ",", ".") & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                        clsSql.Ejecutar strSql, "M"
                    End If
                Next i
            End With
            MsgBox " Los datos han sido ingresado", vbInformation, "SisAdmi - Asientos"
        End If

    Call limpiar
    Call Rango_Fecha
End Sub
Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim NumMes As String
    NumMes = cmbMesA.ListIndex + 1
    If (Len(NumMes) = 1) Then
        NumMes = "0" & NumMes
    End If
    
    'Busca el código máximo de la tabla asiento
    strSql = " Select max(SUBSTRING(asi_numasiento,7,10)) as numAs " & _
            " From asiento " & _
            " WHERE emp_codigo='" & strEmpresa & "' " & _
            " AND asi_numasiento LIKE '" & cmbAñoA.Text & NumMes & "%'" & _
            " GROUP BY emp_codigo"
    clsSql.Ejecutar strSql
    If Not IsNull(clsSql.adorec_Def("numAs")) Then
        maximo = clsSql.adorec_Def("numAs") + 1
    Else
        maximo = 1
    End If
    While (Len(maximo) < 4)
        maximo = "0" & maximo
    Wend
    strMaximo = cmbAñoA.Text & NumMes & maximo
    MsgBox strMaximo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Detecta cuando se ha dado un enter para enviar un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
Private Sub Form_Load()
'Inicializa las clases para hacer distintas consultas
    clsAdq.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsPro.Inicializar AdoConn, AdoConnMaster
    clsSum.Inicializar AdoConn, AdoConnMaster
    clsAct.Inicializar AdoConn, AdoConnMaster
    clsMaxAsi.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    d = CStr(Day(HoyDia))
    m = Month(HoyDia)
    Y = CStr(Year(HoyDia))
    cmbDiaA.Tag = "A"
    cmbDiaA.Text = d
    cmbAñoA.Text = Y
    cmbDiaI.Text = d
    cmbAñoI.Text = Y
    cmbDiaF.Text = d
    cmbAñoF.Text = Y
    'Selecciona el mes actual
    For var = 0 To 11
        If (cmbMesA.ItemData(var) = m) Then
            cmbMesA.ListIndex = var
            'cmbMesA.Text = cmbMesA.List(var)
            cmbMesI.Text = cmbMesI.List(var - 1)
            If (cmbMesI.Text = cmbMesI.List(-1) Or cmbMesI.Text = "") Then
             cmbMesI.Text = cmbMesI.List(11)
             cmbAñoI.Text = Y - 1
             End If
            cmbMesF.Text = cmbMesF.List(var)
            Exit For
        End If
    Next var
    cmbDiaA.Tag = "B"
    Call limpiarFxGD
    Call Rango_Fecha
    'boton aceptar desactivado
    cmdAceptar.Enabled = False

End Sub
Private Sub cmbDiaI_LostFocus()
    Call Rango_Fecha
End Sub
Private Sub cmbMesI_LostFocus()
    Call Rango_Fecha
End Sub
Private Sub cmbAñoI_LostFocus()
    Call Rango_Fecha
End Sub
Private Sub cmbDiaF_LostFocus()
    Call Rango_Fecha
End Sub

Private Sub cmbMesF_LostFocus()
    Call Rango_Fecha
End Sub
Private Sub cmbAñoF_LostFocus()
    Call Rango_Fecha
End Sub
Private Sub TxtSubTotal_Change()
    TxtSubTotal = FormatoD2(TxtSubTotal)
End Sub
Private Sub TxtIva_Change()
    TxtIva = FormatoD2(TxtIva)
End Sub
Private Sub txtTotal_Change()
    TxtTotal = FormatoD2(TxtTotal)
End Sub
Private Sub txtTotalDebe_Change()
    txtTotalDebe = FormatoD2(txtTotalDebe)
End Sub
Private Sub txtTotalHaber_Change()
 txtTotalHaber = FormatoD2(txtTotalHaber)
End Sub
Private Sub VSFGdb_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 And Col = 2 And Col = 3 And Col = 4 Then
         Cancel = True
    End If
End Sub
Private Sub VSFGAdquisicion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Or Col = 8 Or Col = 9 Then
    Cancel = True
    VSFGAdquisicion.Col = 1
  End If
End Sub
Private Sub VSFGAdquisicion_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'este es el check box
    If ban = 1 Then
        If Col = 1 And Row > 0 Then
            If VSFGAdquisicion.TextMatrix(Row, 1) = "-1" Then
                VSFGAdquisicion.Select Row, 1, Row, 9
                VSFGAdquisicion.FillStyle = flexFillRepeat
                VSFGAdquisicion.CellBackColor = &HC0FFFF
                VSFGAdquisicion.Select Row, 9
            ElseIf VSFGAdquisicion.TextMatrix(Row, 1) = "0" Then
              VSFGAdquisicion.Select Row, 1, Row, 9
              VSFGAdquisicion.FillStyle = flexFillRepeat
              VSFGAdquisicion.CellBackColor = &HFFFFFF
              VSFGAdquisicion.Select Row, 9
            End If
        Call DebeHaber
        Call Cal_Total
        Call CalTotalDebeHaber
        End If
    End If
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Integer
    VSFGAdquisicion.Rows = 1
    VSFGAdquisicion.Clear 1
End Sub
