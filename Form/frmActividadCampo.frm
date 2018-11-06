VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmActividadCampo 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actividad de Campo"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15810
   Icon            =   "frmActividadCampo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   15810
   Begin VSFlex8Ctl.VSFlexGrid VSFGExportar2 
      Height          =   1200
      Left            =   8280
      TabIndex        =   12
      Top             =   5160
      Width           =   7380
      _cx             =   13017
      _cy             =   2117
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
      GridColor       =   0
      GridColorFixed  =   0
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmActividadCampo.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFGExportar 
      Height          =   1200
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   11340
      _cx             =   20002
      _cy             =   2117
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
      GridColor       =   0
      GridColorFixed  =   0
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
      Cols            =   33
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmActividadCampo.frx":03FD
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Parametros"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   4455
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
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
      Begin MSDataListLib.DataCombo cmbCampania 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGFechas 
         Height          =   1200
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   5340
         _cx             =   9419
         _cy             =   2117
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmActividadCampo.frx":080D
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaña:"
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
         TabIndex        =   8
         Top             =   645
         Width           =   720
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
         TabIndex        =   6
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2390
      TabIndex        =   1
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4190
      TabIndex        =   0
      Top             =   6480
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2880
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   15540
      _cx             =   27411
      _cy             =   5080
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
      Cols            =   42
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmActividadCampo.frx":089B
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
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmActividadCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Private TipoRed As Integer
Private Destino As String

Private Sub cmbCampania_Validate(Cancel As Boolean)
    strSql = " SELECT TOP 3 concat(cam_anio,'-',cam_mes),cam_nombre, " & _
             " cam_fecha_fac_inicial,cam_fecha_fac_final" & _
             " FROM campaniafecha " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND concat(cam_anio,'-',cam_mes)<='" & cmbCampania.BoundText & "' " & _
             " ORDER BY CONCAT(cam_anio,cam_mes) DESC"
    clsCon_Def.Ejecutar strSql
        Set VSFGFechas.DataSource = clsCon_Def.adorec_Def.DataSource
End Sub

Private Sub cmdProcesar_Click()
    CargarInforme
    'GenerarExcel1
    'GenerarExcel2
End Sub

Private Sub GenerarExcel2(Linea As Long, Lider As String, codigoLider As String)
    Dim i As Long
    Dim LiderARevisar As String
    
    Dim fechaINI As String
    Dim fechaFIN As String
    Dim fechaIniIni As String
    Dim fechaFinIni As String
    Dim fechaIniMit As String
    Dim fechaFinMit As String
    Dim fechaIniFin As String
    Dim fechaFinFin As String
    Dim anioIni As String
    Dim anioMit As String
    Dim anioFin As String
    Dim mesIni As String
    Dim mesMit As String
    Dim mesFin As String
    
    
    fechaINI = VSFGFechas.TextMatrix(3, 2)
    fechaIniIni = VSFGFechas.TextMatrix(3, 2)
    fechaFinIni = VSFGFechas.TextMatrix(3, 3)
    fechaIniMit = VSFGFechas.TextMatrix(2, 2)
    fechaFinMit = VSFGFechas.TextMatrix(2, 3)
    fechaIniFin = VSFGFechas.TextMatrix(1, 2)
    fechaFIN = VSFGFechas.TextMatrix(1, 3)
    fechaFinFin = VSFGFechas.TextMatrix(1, 3)
    
    
    anioIni = Left(VSFGFechas.TextMatrix(3, 0), 4)
    anioMit = Left(VSFGFechas.TextMatrix(2, 0), 4)
    anioFin = Left(VSFGFechas.TextMatrix(1, 0), 4)
    mesIni = Right(VSFGFechas.TextMatrix(3, 0), 2)
    mesMit = Right(VSFGFechas.TextMatrix(2, 0), 2)
    mesFin = Right(VSFGFechas.TextMatrix(1, 0), 2)

    
        LiderARevisar = codigoLider
        strSql = " SELECT " & _
                 " IIF(SUM(tv.tvA)>0,'ACTIVO'," & _
                 "  IIF(SUM(tv.tvA)<=0 AND SUM(tv.tvI1)>0,'INACTIVO C-1'," & _
                 "   IIF(SUM(tv.tvA)<=0 AND SUM(tv.tvI1)<=0,'INACTIVO C-2',''" & _
                 " ))) as ejeIn2,CONCAT(per.per_apellido,'',per.per_nombre) as cli,per.per_ruc,per.per_telf,per.per_fax,per.per_celular,per.per_email,per.per_direccion " & _
                 " FROM ("
            strSql = strSql & " SELECT p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde," & _
                    " SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvI2," & _
                    " SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvI1," & _
                    " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvA," & _
                    " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvFac," & _
                    " 0 as tvNc" & _
                    " FROM persona p INNER JOIN egreso e" & _
                    " ON p.emp_codigo=e.emp_codigo" & _
                    " AND p.per_codigo=e.per_codigo" & _
                    " AND e.tip_egr_codigo='FAC'" & _
                    " AND e.egr_anulado=0" & _
                    " AND e.egr_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
                    " INNER JOIN det_egreso de" & _
                    " ON e.emp_codigo=de.emp_codigo" & _
                    " AND e.egr_codigo=de.egr_codigo" & _
                    " AND e.tip_egr_codigo=de.tip_egr_codigo" & _
                    " INNER JOIN producto pr" & _
                    " ON de.emp_codigo=pr.emp_codigo" & _
                    " AND de.prd_codigo=pr.prd_codigo" & _
                    " INNER JOIN ("
            strSql = strSql & " SELECT TOP 3 emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                        " FROM campaniafecha " & _
                        " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                        " ORDER BY CONCAT(cam_anio,cam_mes) DESC " & _
                    " ) as c"
            strSql = strSql & " ON e.emp_codigo=c.emp_codigo" & _
                    " AND e.egr_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
                    " LEFT JOIN producto_comision pc ON e.emp_codigo=pc.emp_codigo " & _
                    " AND de.prd_codigo=pc.prd_codigo " & _
                    " AND c.cam_anio=pc.cam_anio " & _
                    " AND c.cam_mes=pc.cam_mes " & _
                    " AND p.tip_ped_codigo=pc.tip_ped_codigo "

            strSql = strSql & " WHERE p.emp_codigo='" & strEmpresa & "'" & _
                    " AND p.per_inactivo=0 " & _
                    " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
                    " AND ((p.per_codigo_ref='" & LiderARevisar & "' AND p.per_codigo_ref2='' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref2='" & LiderARevisar & "' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref3='" & LiderARevisar & "' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref4='" & LiderARevisar & "' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref5='" & LiderARevisar & "' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref6='" & LiderARevisar & "' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref7='" & LiderARevisar & "' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref8='" & LiderARevisar & "' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref9='" & LiderARevisar & "' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref10='" & LiderARevisar & "' ))" & _
                    " GROUP BY p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde" & _
                    " UNION"
            strSql = strSql & " SELECT p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde," & _
                    " -1*SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvI2," & _
                    " -1*SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvI1," & _
                    " -1*SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvA," & _
                    " 0 as tvFac," & _
                    " -1*SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0.00)) as tvNc" & _
                    " FROM persona p INNER JOIN ingreso i" & _
                    " ON p.emp_codigo=i.emp_codigo" & _
                    " AND p.per_codigo=i.per_codigo" & _
                    " AND i.tip_ing_codigo='DCL'" & _
                    " AND i.ing_anulado=0" & _
                    " AND i.ing_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
                    " INNER JOIN det_ingreso di" & _
                    " ON i.emp_codigo=di.emp_codigo" & _
                    " AND i.ing_codigo=di.ing_codigo" & _
                    " AND i.tip_ing_codigo=di.tip_ing_codigo" & _
                    " INNER JOIN producto pr" & _
                    " ON di.emp_codigo=pr.emp_codigo" & _
                    " AND di.prd_codigo=pr.prd_codigo" & _
                    " INNER JOIN ("
            strSql = strSql & " SELECT TOP 3 emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                        " FROM campaniafecha " & _
                        " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                        " ORDER BY CONCAT(cam_anio,cam_mes) DESC " & _
                    " ) as c"
            strSql = strSql & " ON i.emp_codigo=c.emp_codigo" & _
                    " AND i.ing_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
                    " LEFT JOIN producto_comision pc ON i.emp_codigo=pc.emp_codigo " & _
                    " AND di.prd_codigo=pc.prd_codigo " & _
                    " AND c.cam_anio=pc.cam_anio " & _
                    " AND c.cam_mes=pc.cam_mes " & _
                    " AND p.tip_ped_codigo=pc.tip_ped_codigo "

            strSql = strSql & " WHERE p.emp_codigo= '" & strEmpresa & "'" & _
                    " AND p.per_inactivo=0 " & _
                    " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
                    " AND ((p.per_codigo_ref='" & LiderARevisar & "' AND p.per_codigo_ref2='' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref2='" & LiderARevisar & "' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref3='" & LiderARevisar & "' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref4='" & LiderARevisar & "' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref5='" & LiderARevisar & "' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref6='" & LiderARevisar & "' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref7='" & LiderARevisar & "' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref8='" & LiderARevisar & "' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref9='" & LiderARevisar & "' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref10='" & LiderARevisar & "' ))" & _
                    " GROUP BY p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde" & _
                    " UNION"
            strSql = strSql & " SELECT p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde," & _
                    " SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',ing_dcto,0)) as tvI2," & _
                    " SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',ing_dcto,0)) as tvI1," & _
                    " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',ing_dcto,0)) as tvA," & _
                    " 0 as tvFac," & _
                    " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',ing_dcto,0)) as tvNc" & _
                    " FROM persona p INNER JOIN ingreso i" & _
                    " ON p.emp_codigo=i.emp_codigo" & _
                    " AND p.per_codigo=i.per_codigo" & _
                    " AND i.tip_ing_codigo='DCL'" & _
                    " AND i.ing_anulado=0" & _
                    " AND i.ing_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
                    " LEFT JOIN det_ingreso di" & _
                    " ON i.emp_codigo=di.emp_codigo" & _
                    " AND i.ing_codigo=di.ing_codigo" & _
                    " AND i.tip_ing_codigo=di.tip_ing_codigo" & _
                    " INNER JOIN ("
            strSql = strSql & " SELECT emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                        " FROM campaniafecha " & _
                        " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                        "  " & _
                    " ) as c"
            strSql = strSql & " ON i.emp_codigo=c.emp_codigo" & _
                    " AND i.ing_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
                    " WHERE di.emp_codigo is null" & _
                    " AND p.emp_codigo= '" & strEmpresa & "'" & _
                    " AND p.per_inactivo=0 " & _
                    " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
                    " AND ((p.per_codigo_ref='" & LiderARevisar & "' AND p.per_codigo_ref2='' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref2='" & LiderARevisar & "' AND p.per_codigo_ref3='' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref3='" & LiderARevisar & "' AND p.per_codigo_ref4='' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref4='" & LiderARevisar & "' AND p.per_codigo_ref5='' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref5='" & LiderARevisar & "' AND p.per_codigo_ref6='' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref6='" & LiderARevisar & "' AND p.per_codigo_ref7='' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref7='" & LiderARevisar & "' AND p.per_codigo_ref8='' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref8='" & LiderARevisar & "' AND p.per_codigo_ref9='' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref9='" & LiderARevisar & "' AND p.per_codigo_ref10='')" & _
                    " OR (p.per_codigo_ref10='" & LiderARevisar & "' ))" & _
                    " GROUP BY p.emp_codigo," & _
                    " p.per_codigo,p.per_perdesde) tv INNER JOIN persona per" & _
                    " ON tv.emp_codigo=per.emp_codigo AND  tv.per_codigo=per.per_codigo"
        strSql = strSql & " GROUP BY per.per_codigo,per.per_apellido,per.per_nombre,per.per_ruc,per.per_telf,per.per_fax,per.per_celular,per.per_email,per.per_direccion ORDER BY ejeIn2 DESC, CONCAT(per.per_apellido,' ',per.per_nombre) "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            Set VSFGExportar2.DataSource = clsCon_Def.adorec_Def.DataSource
            VSFGExportar2.SaveGrid Destino & "\" & Format(Linea, "0000") & Replace(Lider, " ", "_") & "_DET.xls", flexFileExcel, flexXLSaveFixedCells
            strSql = " SELECT per_email FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND cat_p_tipo='C' AND per_codigo='" & LiderARevisar & "'"
            clsCon_Def.Ejecutar strSql
            'EnviarMail NombreComercial & " Seguimiento", CorreoServicioAlCliente, Lider, Trim(clsCon_Def.adorec_Def("per_email")), "", "Reporte Campo y Actividad " & cmbCampania.Text, _
                            "Estimad@" & vbNewLine & _
                            ClienteFactura & vbNewLine & _
                            "Adjunto encontrarás el Reporte de Campo y Actividad donde estan las variables de venta." & vbNewLine & _
                            "Si tiene alguna novedad se tratará en la llamada de confirmación telefónica que realizaremos el día de hoy." & vbNewLine & _
                            "Saludos Cordiales" & vbNewLine & _
                            "Servicio al Cliente" & vbNewLine & _
                            NombreComercial, Destino & "\" & Format(Linea, "0000") & Replace(Lider, " ", "_") & ".xls;" & Destino & "\" & Format(Linea, "0000") & Replace(Lider, " ", "_") & "_DET.xls"
        End If
    
End Sub

Private Sub GenerarExcel1(ini As Long, fin As Long, Lider As String, codigoLider As String)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    If Lider <> "" Then
        VSFGExportar.Clear flexClearScrollable, flexClearText
        VSFGExportar.Rows = 1
        k = 1
        For i = ini To fin
            VSFGExportar.AddItem ""
            For j = 0 To VSFG.Cols - 11
                VSFGExportar.TextMatrix(k, j) = VSFG.TextMatrix(i, j)
                If j = 14 Or j = 16 Or j = 17 Or j = 20 Or j = 23 Then '%Recluta
                    VSFGExportar.Cell(flexcpBackColor, k, j, k, j) = VSFG.Cell(flexcpBackColor, i, j, i, j)
                End If
            Next j
            k = k + 1
        Next i
        VSFGExportar.SaveGrid Destino & "\" & Format(ini, "0000") & Replace(Lider, " ", "_") & ".xls", flexFileExcel, flexXLSaveFixedCells
        GenerarExcel2 ini, Lider, codigoLider
    End If
End Sub

Private Sub CargarInforme()
    Dim clscon As New clsConsulta
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim c6 As String
    Dim c7 As String
    Dim c8 As String
    Dim c9 As String
    Dim c10 As String
    Dim c1a As String
    Dim c2a As String
    Dim c3a As String
    Dim c4a As String
    Dim c5a As String
    Dim c6a As String
    Dim c7a As String
    Dim c8a As String
    Dim c9a As String
    Dim c10a As String
    Dim c1i As Long
    Dim c2i As Long
    Dim c3i As Long
    Dim c4i As Long
    Dim c5i As Long
    Dim c6i As Long
    Dim c7i As Long
    Dim c8i As Long
    Dim c9i As Long
    Dim c10i As Long

    Dim fechaINI As String
    Dim fechaFIN As String
    Dim fechaIniIni As String
    Dim fechaFinIni As String
    Dim fechaIniMit As String
    Dim fechaFinMit As String
    Dim fechaIniFin As String
    Dim fechaFinFin As String
    Dim anioIni As String
    Dim anioMit As String
    Dim anioFin As String
    Dim mesIni As String
    Dim mesMit As String
    Dim mesFin As String
    
    clscon.Inicializar AdoConn, AdoConnMaster
    
    fechaINI = VSFGFechas.TextMatrix(3, 2)
    fechaIniIni = VSFGFechas.TextMatrix(3, 2)
    fechaFinIni = VSFGFechas.TextMatrix(3, 3)
    fechaIniMit = VSFGFechas.TextMatrix(2, 2)
    fechaFinMit = VSFGFechas.TextMatrix(2, 3)
    fechaIniFin = VSFGFechas.TextMatrix(1, 2)
    fechaFIN = VSFGFechas.TextMatrix(1, 3)
    fechaFinFin = VSFGFechas.TextMatrix(1, 3)
    
    
    anioIni = Left(VSFGFechas.TextMatrix(3, 0), 4)
    anioMit = Left(VSFGFechas.TextMatrix(2, 0), 4)
    anioFin = Left(VSFGFechas.TextMatrix(1, 0), 4)
    mesIni = Right(VSFGFechas.TextMatrix(3, 0), 2)
    mesMit = Right(VSFGFechas.TextMatrix(2, 0), 2)
    mesFin = Right(VSFGFechas.TextMatrix(1, 0), 2)
    
    strSql = " SELECT CONCAT(COALESCE(n1.per_apellido,''),' ',COALESCE(n1.per_nombre,'')) as nn1," & _
             " CONCAT(IIF(n1.per_codigo=n2.per_codigo,0,1),COALESCE(n2.per_apellido,''),' ',COALESCE(n2.per_nombre,'')) as nn2," & _
             " CONCAT(IIF(n2.per_codigo=n3.per_codigo,0,1),COALESCE(n3.per_apellido,''),' ',COALESCE(n3.per_nombre,'')) as nn3," & _
             " CONCAT(IIF(n3.per_codigo=n4.per_codigo,0,1),COALESCE(n4.per_apellido,''),' ',COALESCE(n4.per_nombre,'')) as nn4," & _
             " CONCAT(IIF(n4.per_codigo=n5.per_codigo,0,1),COALESCE(n5.per_apellido,''),' ',COALESCE(n5.per_nombre,'')) as nn5," & _
             " CONCAT(IIF(n5.per_codigo=n6.per_codigo,0,1),COALESCE(n6.per_apellido,''),' ',COALESCE(n6.per_nombre,'')) as nn6," & _
             " CONCAT(IIF(n6.per_codigo=n7.per_codigo,0,1),COALESCE(n7.per_apellido,''),' ',COALESCE(n7.per_nombre,'')) as nn7," & _
             " CONCAT(IIF(n7.per_codigo=n8.per_codigo,0,1),COALESCE(n8.per_apellido,''),' ',COALESCE(n8.per_nombre,'')) as nn8," & _
             " CONCAT(IIF(n8.per_codigo=n9.per_codigo,0,1),COALESCE(n9.per_apellido,''),' ',COALESCE(n9.per_nombre,'')) as nn9," & _
             " CONCAT(IIF(n9.per_codigo=n10.per_codigo,0,1),COALESCE(n10.per_apellido,''),' ',COALESCE(n10.per_nombre,'')) as nn10," & _
             " sum(COALESCE(act.ejeAct,0)) as ejeAct," & _
             " sum(COALESCE(act.ejeIn1,0)) as ejeIn1," & _
             " sum(COALESCE(act.ejeAct,0))+sum(COALESCE(act.ejeIn1,0)) as StencilCampana," & _
             " COALESCE(tEjeNuevo.ejeNuevo,0) as ejeNuevo," & _
             " IIF(sum(COALESCE(act.ejeAct,0))+sum(COALESCE(act.ejeIn1,0))!=0,COALESCE(tEjeNuevo.ejeNuevo,0)/(sum(COALESCE(act.ejeAct,0))+sum(COALESCE(act.ejeIn1,0)))*100.00,0) as PRecluta," & _
             " sum(COALESCE(act.ejeIn2,0)) as ejeIn2," & _
             " COALESCE(tEjeNuevo.ejeNuevo,0)-sum(COALESCE(act.ejeIn2,0)) as Capitaliza," & _
             " IIF(sum(COALESCE(act.ejeAct,0))+sum(COALESCE(act.ejeIn1,0))!=0,sum(COALESCE(act.ejeAct,0))/(sum(COALESCE(act.ejeAct,0))+sum(COALESCE(act.ejeIn1,0)))*100.00,0) as PActividad," & _
             " sum(COALESCE(act.ejeActAnt,0)) as ejeActAnt," & _
             " SUM(act.ejeConse) as ejeConsecu,IIF(sum(COALESCE(act.ejeActAnt,0))!=0,SUM(act.ejeConse)/sum(COALESCE(act.ejeActAnt,0))*100.00,0) as PConsecu," & _
             " COALESCE(tEjeNuevoAnt.ejeNuevoAnt,0) as ejeNuevoAnt,"
    strSql = strSql & " sum(act.ejeNuevoConsecutivo) as ejeNuevoConsecutivo," & _
             " IIF(COALESCE(tEjeNuevoAnt.ejeNuevoAnt,0)!=0,sum(act.ejeNuevoConsecutivo)/COALESCE(tEjeNuevoAnt.ejeNuevoAnt,0)*100.00,0) as PRetencion, " & _
             " COALESCE(totFac,0) as nFac, ROUND(sum(COALESCE(act.tvFac,0)),2) as tFac," & _
             " IIF(COALESCE(totFac,0)!=0,ROUND(sum(COALESCE(act.tvFac,0)),2)/COALESCE(totFac,0),0) as op,  " & _
             " COALESCE(totNc,0) as nNc,ROUND(sum(COALESCE(act.tvNc,0)),2) as tNc, " & _
             " IIF(COALESCE(totNc,0)!=0,ROUND(sum(COALESCE(act.tvNc,0)),2)/COALESCE(totNc,0),0) as od," & _
             " IIF(sum(COALESCE(act.tvFac,0))!=0,ROUND(sum(COALESCE(act.tvNc,0)),2)/ROUND(sum(COALESCE(act.tvFac,0)),2)*100.00,0) as PDev , sum(act.tvA) as ventaNeta,"
    strSql = strSql & " n1.per_codigo,n2.per_codigo,n3.per_codigo," & _
             " n4.per_codigo,n5.per_codigo,n6.per_codigo,n7.per_codigo,n8.per_codigo," & _
             " n9.per_codigo,n10.per_codigo "
    strSql = strSql & " FROM (" & _
                " SELECT emp_codigo,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
                " per_codigo_ref9,per_codigo_ref10, count(per_codigo) as totEje" & _
                " FROM persona" & _
                " WHERE emp_codigo='" & strEmpresa & "'" & _
                " AND tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
                " AND persona.per_inactivo=0 " & _
                " GROUP BY emp_codigo,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
                " per_codigo_ref9,per_codigo_ref10"
    strSql = strSql & " ) te LEFT JOIN (" & _
        " SELECT tv.emp_codigo,tv.per_codigo_ref,tv.per_codigo_ref2,tv.per_codigo_ref3,tv.per_codigo_ref4," & _
        " tv.per_codigo_ref5,tv.per_codigo_ref6,tv.per_codigo_ref7,tv.per_codigo_ref8," & _
        " tv.per_codigo_ref9,tv.per_codigo_ref10,tv.per_codigo," & _
        " IIF(SUM(tv.tvA)>0,1,0) as ejeAct," & _
        " IIF(SUM(tv.tvI1)>0,1,0) as ejeActAnt," & _
        " IIF(SUM(tv.tvA)<=0 AND SUM(tv.tvI1)>0,1,0) as ejeIn1," & _
        " IIF(SUM(tv.tvA)<=0 AND SUM(tv.tvI1)<=0,1,0) as ejeIn2," & _
        " IIF(SUM(tv.tvA)>0 AND SUM(tv.tvI1)>0,1,0) as ejeConse," & _
        " IIF(tv.per_perdesde between '" & fechaIniMit & "' AND '" & fechaFinMit & "' AND SUM(tv.tvA)>0,1,0) as ejeNuevoConsecutivo," & _
        " SUM(tv.tvA) as tvA," & _
        " SUM(tv.tvFac) as tvFac," & _
        " SUM(tv.tvNc) as tvNc" & _
        " FROM ("
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,p.per_codigo,p.per_perdesde," & _
            " SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvI2," & _
            " SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvI1," & _
            " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvA," & _
            " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_egr_cantidad*det_egr_precio-det_egr_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvFac," & _
            " 0 as tvNc" & _
            " FROM persona p INNER JOIN egreso e" & _
            " ON p.emp_codigo=e.emp_codigo" & _
            " AND p.per_codigo=e.per_codigo" & _
            " AND e.tip_egr_codigo='FAC'" & _
            " AND e.egr_anulado=0" & _
            " AND e.egr_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
            " INNER JOIN det_egreso de" & _
            " ON e.emp_codigo=de.emp_codigo" & _
            " AND e.egr_codigo=de.egr_codigo" & _
            " AND e.tip_egr_codigo=de.tip_egr_codigo"
    strSql = strSql & " INNER JOIN producto pr " & _
            " ON de.emp_codigo=pr.emp_codigo" & _
            " AND de.prd_codigo=pr.prd_codigo" & _
            " INNER JOIN ("
    strSql = strSql & " SELECT TOP 3 emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                " FROM campaniafecha " & _
                " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                " ORDER BY CONCAT(cam_anio,cam_mes) DESC " & _
            " ) as c"
    strSql = strSql & " ON e.emp_codigo=c.emp_codigo" & _
            " AND e.egr_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
            " LEFT JOIN producto_comision pc ON e.emp_codigo=pc.emp_codigo " & _
            " AND de.prd_codigo=pc.prd_codigo " & _
            " AND pc.cam_anio=c.cam_anio " & _
            " AND pc.cam_mes=c.cam_mes " & _
            " AND p.tip_ped_codigo=pc.tip_ped_codigo "

            strSql = strSql & " WHERE p.emp_codigo='" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " GROUP BY p.emp_codigo," & _
            " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
            " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
            " per_codigo_ref9,per_codigo_ref10,p.per_codigo,p.per_perdesde" & _
            " UNION"
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,p.per_codigo,p.per_perdesde," & _
            " -1*SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvI2," & _
            " -1*SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvI1," & _
            " -1*SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvA," & _
            " 0 as tvFac," & _
            " -1*SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',(det_ing_cantidad*det_ing_precio-det_ing_dcto)*COALESCE(pc.pro_com_comision,100.00)/100.00,0)) as tvNc" & _
            " FROM persona p INNER JOIN ingreso i" & _
            " ON p.emp_codigo=i.emp_codigo" & _
            " AND p.per_codigo=i.per_codigo" & _
            " AND i.tip_ing_codigo='DCL'" & _
            " AND i.ing_anulado=0" & _
            " AND i.ing_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
            " INNER JOIN det_ingreso di" & _
            " ON i.emp_codigo=di.emp_codigo" & _
            " AND i.ing_codigo=di.ing_codigo" & _
            " AND i.tip_ing_codigo=di.tip_ing_codigo"
    strSql = strSql & " INNER JOIN producto pr" & _
            " ON di.emp_codigo=pr.emp_codigo" & _
            " AND di.prd_codigo=pr.prd_codigo" & _
            " INNER JOIN ("
    strSql = strSql & " SELECT TOP 3 emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                " FROM campaniafecha " & _
                " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                " ORDER BY CONCAT(cam_anio,cam_mes) DESC " & _
            " ) as c"
    strSql = strSql & " ON i.emp_codigo=c.emp_codigo" & _
            " AND i.ing_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
            " LEFT JOIN producto_comision pc ON i.emp_codigo=pc.emp_codigo " & _
            " AND di.prd_codigo=pc.prd_codigo " & _
            " AND pc.cam_anio=c.cam_anio " & _
            " AND pc.cam_mes=c.cam_mes " & _
            " AND p.tip_ped_codigo=pc.tip_ped_codigo "

            strSql = strSql & " WHERE p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " GROUP BY p.emp_codigo," & _
            " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
            " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
            " per_codigo_ref9,per_codigo_ref10,p.per_codigo,p.per_perdesde" & _
            " UNION"
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,p.per_codigo,p.per_perdesde," & _
            " SUM(IIF(c.cam_anio='" & anioIni & "' AND c.cam_mes='" & mesIni & "',ing_dcto,0)) as tvI2," & _
            " SUM(IIF(c.cam_anio='" & anioMit & "' AND c.cam_mes='" & mesMit & "',ing_dcto,0)) as tvI1," & _
            " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',ing_dcto,0)) as tvA," & _
            " 0 as tvFac," & _
            " SUM(IIF(c.cam_anio='" & anioFin & "' AND c.cam_mes='" & mesFin & "',ing_dcto,0)) as tvNc" & _
            " FROM persona p INNER JOIN ingreso i" & _
            " ON p.emp_codigo=i.emp_codigo" & _
            " AND p.per_codigo=i.per_codigo" & _
            " AND i.tip_ing_codigo='DCL'" & _
            " AND i.ing_anulado=0" & _
            " AND i.ing_fecha between '" & fechaINI & "' and '" & fechaFIN & "'" & _
            " LEFT JOIN det_ingreso di" & _
            " ON i.emp_codigo=di.emp_codigo" & _
            " AND i.ing_codigo=di.ing_codigo" & _
            " AND i.tip_ing_codigo=di.tip_ing_codigo" & _
            " INNER JOIN ("
    strSql = strSql & " SELECT TOP 3 emp_codigo,cam_anio,cam_mes,cam_fecha_fac_inicial,cam_fecha_fac_final" & _
                " FROM campaniafecha " & _
                " WHERE CONCAT(cam_anio,cam_mes)<=CONCAT('" & anioFin & "','" & mesFin & "')" & _
                " ORDER BY CONCAT(cam_anio,cam_mes) DESC " & _
            " ) as c"
    strSql = strSql & " ON i.emp_codigo=c.emp_codigo" & _
            " AND i.ing_fecha between cam_fecha_fac_inicial and cam_fecha_fac_final" & _
            " WHERE di.emp_codigo is null" & _
            " AND p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " GROUP BY p.emp_codigo," & _
            " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
            " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8," & _
            " per_codigo_ref9,per_codigo_ref10,p.per_codigo,p.per_perdesde) tv" & _
            " GROUP BY tv.emp_codigo,tv.per_codigo_ref,tv.per_codigo_ref2,tv.per_codigo_ref3,tv.per_codigo_ref4," & _
            " tv.per_codigo_ref5,tv.per_codigo_ref6,tv.per_codigo_ref7,tv.per_codigo_ref8," & _
            " tv.per_codigo_ref9,tv.per_codigo_ref10,tv.per_codigo,tv.per_perdesde" & _
        ") act"
    strSql = strSql & " ON te.emp_codigo=act.emp_codigo" & _
        " AND te.per_codigo_ref=act.per_codigo_ref" & _
        " AND te.per_codigo_ref2=act.per_codigo_ref2" & _
        " AND te.per_codigo_ref3=act.per_codigo_ref3" & _
        " AND te.per_codigo_ref4=act.per_codigo_ref4" & _
        " AND te.per_codigo_ref5=act.per_codigo_ref5" & _
        " AND te.per_codigo_ref6=act.per_codigo_ref6" & _
        " AND te.per_codigo_ref7=act.per_codigo_ref7" & _
        " AND te.per_codigo_ref8=act.per_codigo_ref8" & _
        " AND te.per_codigo_ref9=act.per_codigo_ref9" & _
        " AND te.per_codigo_ref10=act.per_codigo_ref10" & _
        " LEFT JOIN ("
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,COUNT(DISTINCT e.egr_codigo) as totFac " & _
            " FROM egreso e INNER JOIN persona p " & _
            " ON e.emp_codigo=p.emp_codigo " & _
            " AND e.per_codigo=p.per_codigo " & _
            " AND p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " WHERE e.egr_fecha between '" & fechaIniFin & "' AND '" & fechaFinFin & "'" & _
            " AND e.tip_egr_codigo='FAC' " & _
            " AND e.egr_anulado=0 " & _
            " GROUP BY p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10" & _
        ") tfac "
    strSql = strSql & " ON te.emp_codigo=tfac.emp_codigo" & _
        " AND te.per_codigo_ref=tfac.per_codigo_ref" & _
        " AND te.per_codigo_ref2=tfac.per_codigo_ref2" & _
        " AND te.per_codigo_ref3=tfac.per_codigo_ref3" & _
        " AND te.per_codigo_ref4=tfac.per_codigo_ref4" & _
        " AND te.per_codigo_ref5=tfac.per_codigo_ref5" & _
        " AND te.per_codigo_ref6=tfac.per_codigo_ref6" & _
        " AND te.per_codigo_ref7=tfac.per_codigo_ref7" & _
        " AND te.per_codigo_ref8=tfac.per_codigo_ref8" & _
        " AND te.per_codigo_ref9=tfac.per_codigo_ref9" & _
        " AND te.per_codigo_ref10=tfac.per_codigo_ref10" & _
        " LEFT JOIN ("
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,COUNT(DISTINCT e.ing_codigo) as totNc " & _
            " FROM ingreso e INNER JOIN persona p " & _
            " ON e.emp_codigo=p.emp_codigo " & _
            " AND e.per_codigo=p.per_codigo " & _
            " AND p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " WHERE e.ing_fecha between '" & fechaIniFin & "' AND '" & fechaFinFin & "'" & _
            " AND e.tip_ing_codigo='DCL' " & _
            " AND e.ing_anulado=0 " & _
            " GROUP BY p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10" & _
        ") tnc "
    strSql = strSql & " ON te.emp_codigo=tnc.emp_codigo" & _
        " AND te.per_codigo_ref=tnc.per_codigo_ref" & _
        " AND te.per_codigo_ref2=tnc.per_codigo_ref2" & _
        " AND te.per_codigo_ref3=tnc.per_codigo_ref3" & _
        " AND te.per_codigo_ref4=tnc.per_codigo_ref4" & _
        " AND te.per_codigo_ref5=tnc.per_codigo_ref5" & _
        " AND te.per_codigo_ref6=tnc.per_codigo_ref6" & _
        " AND te.per_codigo_ref7=tnc.per_codigo_ref7" & _
        " AND te.per_codigo_ref8=tnc.per_codigo_ref8" & _
        " AND te.per_codigo_ref9=tnc.per_codigo_ref9" & _
        " AND te.per_codigo_ref10=tnc.per_codigo_ref10" & _
        " LEFT JOIN ("
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,COUNT(DISTINCT p.per_codigo) as ejeNuevo " & _
            " FROM persona p " & _
            " WHERE p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " AND p.per_perdesde between '" & fechaIniFin & "' AND '" & fechaFinFin & "'" & _
            " GROUP BY p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10" & _
        ") tEjeNuevo "
    strSql = strSql & " ON te.emp_codigo=tEjeNuevo.emp_codigo" & _
        " AND te.per_codigo_ref=tEjeNuevo.per_codigo_ref" & _
        " AND te.per_codigo_ref2=tEjeNuevo.per_codigo_ref2" & _
        " AND te.per_codigo_ref3=tEjeNuevo.per_codigo_ref3" & _
        " AND te.per_codigo_ref4=tEjeNuevo.per_codigo_ref4" & _
        " AND te.per_codigo_ref5=tEjeNuevo.per_codigo_ref5" & _
        " AND te.per_codigo_ref6=tEjeNuevo.per_codigo_ref6" & _
        " AND te.per_codigo_ref7=tEjeNuevo.per_codigo_ref7" & _
        " AND te.per_codigo_ref8=tEjeNuevo.per_codigo_ref8" & _
        " AND te.per_codigo_ref9=tEjeNuevo.per_codigo_ref9" & _
        " AND te.per_codigo_ref10=tEjeNuevo.per_codigo_ref10" & _
        " LEFT JOIN ("
    strSql = strSql & " SELECT p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10,COUNT(DISTINCT p.per_codigo) as ejeNuevoAnt " & _
            " FROM persona p " & _
            " WHERE p.emp_codigo= '" & strEmpresa & "'" & _
            " AND p.per_inactivo=0 " & _
            " AND p.tip_ped_codigo LIKE '" & cmbNegocio.BoundText & "'" & _
            " AND p.per_perdesde between '" & fechaIniMit & "' AND '" & fechaFinMit & "'" & _
            " GROUP BY p.emp_codigo," & _
            " p.per_codigo_ref,p.per_codigo_ref2,p.per_codigo_ref3,p.per_codigo_ref4," & _
            " p.per_codigo_ref5,p.per_codigo_ref6,p.per_codigo_ref7,p.per_codigo_ref8," & _
            " p.per_codigo_ref9,p.per_codigo_ref10" & _
        ") tEjeNuevoAnt "
    strSql = strSql & " ON te.emp_codigo=tEjeNuevoAnt.emp_codigo" & _
        " AND te.per_codigo_ref=tEjeNuevoAnt.per_codigo_ref" & _
        " AND te.per_codigo_ref2=tEjeNuevoAnt.per_codigo_ref2" & _
        " AND te.per_codigo_ref3=tEjeNuevoAnt.per_codigo_ref3" & _
        " AND te.per_codigo_ref4=tEjeNuevoAnt.per_codigo_ref4" & _
        " AND te.per_codigo_ref5=tEjeNuevoAnt.per_codigo_ref5" & _
        " AND te.per_codigo_ref6=tEjeNuevoAnt.per_codigo_ref6" & _
        " AND te.per_codigo_ref7=tEjeNuevoAnt.per_codigo_ref7" & _
        " AND te.per_codigo_ref8=tEjeNuevoAnt.per_codigo_ref8" & _
        " AND te.per_codigo_ref9=tEjeNuevoAnt.per_codigo_ref9" & _
        " AND te.per_codigo_ref10=tEjeNuevoAnt.per_codigo_ref10"
    strSql = strSql & " LEFT JOIN persona n1 ON te.emp_codigo=n1.emp_codigo AND te.per_codigo_ref=n1.per_codigo AND n1.per_es_gz=1" & _
        " LEFT JOIN persona n2 ON te.emp_codigo=n2.emp_codigo AND te.per_codigo_ref2=n2.per_codigo AND n2.per_es_di=1" & _
        " LEFT JOIN persona n3 ON te.emp_codigo=n3.emp_codigo AND te.per_codigo_ref3=n3.per_codigo AND n3.per_es_em=1" & _
        " LEFT JOIN persona n4 ON te.emp_codigo=n4.emp_codigo AND te.per_codigo_ref4=n4.per_codigo AND n4.per_es_ee=1" & _
        " LEFT JOIN persona n5 ON te.emp_codigo=n5.emp_codigo AND te.per_codigo_ref5=n5.per_codigo AND n5.per_es_n5=1" & _
        " LEFT JOIN persona n6 ON te.emp_codigo=n6.emp_codigo AND te.per_codigo_ref6=n6.per_codigo AND n6.per_es_n6=1 " & _
        " LEFT JOIN persona n7 ON te.emp_codigo=n7.emp_codigo AND te.per_codigo_ref7=n7.per_codigo AND n7.per_es_n7=1" & _
        " LEFT JOIN persona n8 ON te.emp_codigo=n8.emp_codigo AND te.per_codigo_ref8=n8.per_codigo AND n8.per_es_n8=1" & _
        " LEFT JOIN persona n9 ON te.emp_codigo=n9.emp_codigo AND te.per_codigo_ref9=n9.per_codigo AND n9.per_es_n9=1" & _
        " LEFT JOIN persona n10 ON te.emp_codigo=n10.emp_codigo AND te.per_codigo_ref10=n10.per_codigo AND n10.per_es_n10=1 " & _
        " group by te.emp_codigo,te.per_codigo_ref,te.per_codigo_ref2,te.per_codigo_ref3,te.per_codigo_ref4," & _
        " te.per_codigo_ref5,te.per_codigo_ref6,te.per_codigo_ref7,te.per_codigo_ref8," & _
        " te.per_codigo_ref9,te.per_codigo_ref10," & _
        " n1.per_apellido,n1.per_nombre,n1.per_codigo,n2.per_apellido,n2.per_nombre,n2.per_codigo," & _
        " n3.per_apellido,n3.per_nombre,n3.per_codigo,n4.per_apellido,n4.per_nombre,n4.per_codigo," & _
        " n5.per_apellido,n5.per_nombre,n5.per_codigo,n6.per_apellido,n6.per_nombre,n6.per_codigo," & _
        " n7.per_apellido,n7.per_nombre,n7.per_codigo,n8.per_apellido,n8.per_nombre,n8.per_codigo," & _
        " n9.per_apellido,n9.per_nombre,n9.per_codigo,n10.per_apellido,n10.per_nombre,n10.per_codigo," & _
        " tEjeNuevo.ejeNuevo,tEjeNuevoAnt.ejeNuevoAnt,totFac,totNc" & _
        " order by nn1, nn2, nn3, nn4, nn5, nn6, nn7, nn8, nn9, nn10"
    clscon.Ejecutar strSql
    Destino = Buscar_Carpeta(Me.hwnd, "Carpetas a Subir")
    Dim i As Long
    Dim j As Long
    VSFG.Rows = 1
    i = 1
    If clscon.adorec_Def.RecordCount > 0 Then
        While Not clscon.adorec_Def.EOF
            VSFG.AddItem ""
            For j = 0 To clscon.adorec_Def.Fields.count - 1
                VSFG.TextMatrix(i, j) = IIf(IsNull(clscon.adorec_Def(j)), 0, clscon.adorec_Def(j))
                If 0 < j And j < 10 Then
                    VSFG.TextMatrix(i, j) = Right(VSFG.TextMatrix(i, j), Len(VSFG.TextMatrix(i, j)) - 1)
                End If
                If j = 14 Then '%Recluta
                    If FormatoD2(VSFG.TextMatrix(i, j)) < 7 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbRed
                    ElseIf FormatoD2(VSFG.TextMatrix(i, j)) < 10 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbYellow
                    Else
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbGreen
                    End If
                    
                ElseIf j = 16 Then 'Capilaliza
                    If FormatoD2(VSFG.TextMatrix(i, j)) < 0 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbRed
                    ElseIf FormatoD2(VSFG.TextMatrix(i, j)) < 5 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbYellow
                    Else
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbGreen
                    End If
                
                ElseIf j = 17 Then '%Actividad
                    If FormatoD2(VSFG.TextMatrix(i, j)) < 70 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbRed
                    ElseIf FormatoD2(VSFG.TextMatrix(i, j)) < 80 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbYellow
                    Else
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbGreen
                    End If
                
                ElseIf j = 20 Then '%consecutividad
                    If FormatoD2(VSFG.TextMatrix(i, j)) < 75 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbRed
                    ElseIf FormatoD2(VSFG.TextMatrix(i, j)) < 84 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbYellow
                    Else
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbGreen
                    End If
                    
                
                ElseIf j = 23 Then '%retencionnuevos
                    If FormatoD2(VSFG.TextMatrix(i, j)) < 70 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbRed
                    ElseIf FormatoD2(VSFG.TextMatrix(i, j)) < 80 Then
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbYellow
                    Else
                        VSFG.Cell(flexcpBackColor, i, j, i, j) = vbGreen
                    End If
                
                End If
                VSFG.ShowCell i, j
            Next j
            VSFG.ShowCell i, 15
            i = i + 1
            clscon.adorec_Def.MoveNext
            c10 = VSFG.TextMatrix(i - 1, VSFG.Cols - 1)
            c9 = VSFG.TextMatrix(i - 1, VSFG.Cols - 2)
            c8 = VSFG.TextMatrix(i - 1, VSFG.Cols - 3)
            c7 = VSFG.TextMatrix(i - 1, VSFG.Cols - 4)
            c6 = VSFG.TextMatrix(i - 1, VSFG.Cols - 5)
            c5 = VSFG.TextMatrix(i - 1, VSFG.Cols - 6)
            c4 = VSFG.TextMatrix(i - 1, VSFG.Cols - 7)
            c3 = VSFG.TextMatrix(i - 1, VSFG.Cols - 8)
            c2 = VSFG.TextMatrix(i - 1, VSFG.Cols - 9)
            c1 = VSFG.TextMatrix(i - 1, VSFG.Cols - 10)
            
            If c10 <> c10a Then
                If c10i <> 0 And c10a <> "0" Then
                    GenerarExcel1 c10i, i - 2, VSFG.TextMatrix(c10i, 9), c10a
                End If
                c10i = i - 1
            End If
            If c9 <> c9a Then
                If c9i <> 0 And c9a <> "0" Then
                    GenerarExcel1 c9i, i - 2, VSFG.TextMatrix(c9i, 8), c9a
                End If
                c9i = i - 1
                c10i = i - 1
            End If
            If c8 <> c8a Then
                If c8i <> 0 And c8a <> "0" Then
                    GenerarExcel1 c8i, i - 2, VSFG.TextMatrix(c8i, 7), c8a
                End If
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c7 <> c7a Then
                If c7i <> 0 And c7a <> "0" Then
                    GenerarExcel1 c7i, i - 2, VSFG.TextMatrix(c7i, 6), c7a
                End If
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c6 <> c6a Then
                If c6i <> 0 And c6a <> "0" Then
                    GenerarExcel1 c6i, i - 2, VSFG.TextMatrix(c6i, 5), c6a
                End If
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c5 <> c5a Then
                If c5i <> 0 And c5a <> "0" Then
                    GenerarExcel1 c5i, i - 2, VSFG.TextMatrix(c5i, 4), c5a
                End If
                c5i = i - 1
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c4 <> c4a Then
                If c4i <> 0 And c4a <> "0" Then
                    GenerarExcel1 c4i, i - 2, VSFG.TextMatrix(c4i, 3), c4a
                End If
                c4i = i - 1
                c5i = i - 1
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c3 <> c3a Then
                If c3i <> 0 And c3a <> "0" Then
                    GenerarExcel1 c3i, i - 2, VSFG.TextMatrix(c3i, 2), c3a
                End If
                c3i = i - 1
                c4i = i - 1
                c5i = i - 1
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c2 <> c2a Then
                If c2i <> 0 And c2a <> "0" Then
                    GenerarExcel1 c2i, i - 2, VSFG.TextMatrix(c2i, 1), c2a
                End If
                c2i = i - 1
                c3i = i - 1
                c4i = i - 1
                c5i = i - 1
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            If c1 <> c1a Then
                If c1i <> 0 And c1a <> "0" Then
                    GenerarExcel1 c1i, i - 2, VSFG.TextMatrix(c1i, 0), c1a
                End If
                c1i = i - 1
                c2i = i - 1
                c3i = i - 1
                c4i = i - 1
                c5i = i - 1
                c6i = i - 1
                c7i = i - 1
                c8i = i - 1
                c9i = i - 1
                c10i = i - 1
            End If
            
            c10a = c10
            c9a = c9
            c8a = c8
            c7a = c7
            c6a = c6
            c5a = c5
            c4a = c4
            c3a = c3
            c2a = c2
            c1a = c1
            
        Wend
        'ULTIMO EN BLANCO
        VSFG.AddItem ""
        i = i + 1
        c10 = VSFG.TextMatrix(i - 1, VSFG.Cols - 1)
        c9 = VSFG.TextMatrix(i - 1, VSFG.Cols - 2)
        c8 = VSFG.TextMatrix(i - 1, VSFG.Cols - 3)
        c7 = VSFG.TextMatrix(i - 1, VSFG.Cols - 4)
        c6 = VSFG.TextMatrix(i - 1, VSFG.Cols - 5)
        c5 = VSFG.TextMatrix(i - 1, VSFG.Cols - 6)
        c4 = VSFG.TextMatrix(i - 1, VSFG.Cols - 7)
        c3 = VSFG.TextMatrix(i - 1, VSFG.Cols - 8)
        c2 = VSFG.TextMatrix(i - 1, VSFG.Cols - 9)
        c1 = VSFG.TextMatrix(i - 1, VSFG.Cols - 10)
        If c10 <> c10a Then
            If c10i <> 0 And c10a <> "0" Then
                GenerarExcel1 c10i, i - 2, VSFG.TextMatrix(c10i, 9), c10a
            End If
            c10i = i - 1
        End If
        If c9 <> c9a Then
            If c9i <> 0 And c9a <> "0" Then
                GenerarExcel1 c9i, i - 2, VSFG.TextMatrix(c9i, 8), c9a
            End If
            c9i = i - 1
            c10i = i - 1
        End If
        If c8 <> c8a Then
            If c8i <> 0 And c8a <> "0" Then
                GenerarExcel1 c8i, i - 2, VSFG.TextMatrix(c8i, 7), c8a
            End If
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c7 <> c7a Then
            If c7i <> 0 And c7a <> "0" Then
                GenerarExcel1 c7i, i - 2, VSFG.TextMatrix(c7i, 6), c7a
            End If
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c6 <> c6a Then
            If c6i <> 0 And c6a <> "0" Then
                GenerarExcel1 c6i, i - 2, VSFG.TextMatrix(c6i, 5), c6a
            End If
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c5 <> c5a Then
            If c5i <> 0 And c5a <> "0" Then
                GenerarExcel1 c5i, i - 2, VSFG.TextMatrix(c5i, 4), c5a
            End If
            c5i = i - 1
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c4 <> c4a Then
            If c4i <> 0 And c4a <> "0" Then
                GenerarExcel1 c4i, i - 2, VSFG.TextMatrix(c4i, 3), c4a
            End If
            c4i = i - 1
            c5i = i - 1
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c3 <> c3a Then
            If c3i <> 0 And c3a <> "0" Then
                GenerarExcel1 c3i, i - 2, VSFG.TextMatrix(c3i, 2), c3a
            End If
            c3i = i - 1
            c4i = i - 1
            c5i = i - 1
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c2 <> c2a Then
            If c2i <> 0 And c2a <> "0" Then
                GenerarExcel1 c2i, i - 2, VSFG.TextMatrix(c2i, 1), c2a
            End If
            c2i = i - 1
            c3i = i - 1
            c4i = i - 1
            c5i = i - 1
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        If c1 <> c1a Then
            If c1i <> 0 And c1a <> "0" Then
                GenerarExcel1 c1i, i - 2, VSFG.TextMatrix(c1i, 0), c1a
            End If
            c1i = i - 1
            c2i = i - 1
            c3i = i - 1
            c4i = i - 1
            c5i = i - 1
            c6i = i - 1
            c7i = i - 1
            c8i = i - 1
            c9i = i - 1
            c10i = i - 1
        End If
        
        c10a = c10
        c9a = c9
        c8a = c8
        c7a = c7
        c6a = c6
        c5a = c5
        c4a = c4
        c3a = c3
        c2a = c2
        c1a = c1
    End If
    'Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
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
    'Tipo de negocios
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
    End If
    
    strSql = " SELECT concat(cam_anio,'-',cam_mes) as cam_codigo, cam_nombre " & _
             " FROM campaniafecha " & _
             " ORDER BY cam_nombre DESC "
    clsCon_Def.Ejecutar strSql
    Set cmbCampania.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbCampania.ListField = "cam_nombre"
    cmbCampania.BoundColumn = "cam_codigo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
