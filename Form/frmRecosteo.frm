VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecosteo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recosteo"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecosteo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6795
   Begin VB.TextBox txtPausa 
      Height          =   315
      Left            =   600
      TabIndex        =   11
      Text            =   "aaaa-mm-dd hh:mm:ss"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtSaltar 
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Text            =   "0"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPausa 
      Caption         =   "&Pausa"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfgKardex 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
      _cx             =   10610
      _cy             =   8916
      Appearance      =   0
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4935
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3375
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDesde 
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   67436547
      CurrentDate     =   37463
   End
   Begin MSComCtl2.DTPicker dtpHasta 
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
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   67436547
      CurrentDate     =   37463
   End
   Begin NEED2.uctrVSFG ucrtVSFG 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4695
      _extentx        =   8281
      _extenty        =   661
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Saltar:"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label lblPorcentaje 
      Caption         =   "Porcentaje:"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmRecosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As New clsConsulta
Private clsCos As New clsCostear

Private Sub cmbAceptar_Click()
    Dim i As Long
    Dim strIni As String
    Dim strFin As String
    clsCos.Inicializar AdoConn, AdoConnMaster
    strIni = Ahora
    If MsgBox("Recosteo Total", vbYesNo) = vbNo Then
        For i = 1 To vsfgKardex.Rows - 1
            clsCos.Recostear Format(dtpDesde.Value, "yyyy-mm-dd"), Format(dtpHasta.Value, "yyyy-mm-dd"), vsfgKardex.TextMatrix(i, 0), True, , txtPausa.Text
            vsfgKardex.Select i, 0, i, 0
            vsfgKardex.ShowCell i, 0
            
            vsfgKardex.Refresh
        Next i
    Else
        clsCos.Recostear Format(dtpDesde.Value, "yyyy-mm-dd"), Format(dtpHasta.Value, "yyyy-mm-dd"), , True, Val(txtSaltar.Text), txtPausa.Text
    End If
    strFin = Ahora
    MsgBox "FIN" & vbNewLine & strIni & vbNewLine & strFin
    'Recostear Format(dtpDesde.Value, "yyyy-mm-dd"), Format(dtpHasta.Value, "yyyy-mm-dd")
End Sub

Private Sub cmdPausa_Click()
    clsCos.booPausa = True
End Sub

Private Sub Command1_Click()
    CD.ShowOpen
    vsfgKardex.LoadGrid CD.FileName, flexFileExcel
    'vsfgKardex.LoadGrid CD.FileName, flexFileTabText
    
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

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpDesde.Value = HoyDia
    dtpHasta.Value = HoyDia
    Set ucrtVSFG.VSFGControl = vsfgKardex
    ucrtVSFG.Inicializar
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    ElseIf KeyCode = vbKeyF5 Then
        clsCos.booPausa = True
    End If
End Sub

Private Sub Recostear(desde As String, hasta As String, Optional produc As String = "%")
    Dim strSql As String
    Dim strPrd As String
    Dim tdoc As String
    Dim Doc As String
    Dim docfecha As String
    Dim prd_codigo As String
    Dim egr As Long
    Dim ing As Long
    Dim exi As Long
    Dim d As String
    Dim prec As Double
    Dim PC As Double
    Dim nn As Long
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCos.CreaKardex produc, desde, hasta
    strSql = " SELECT tdoc,doc,docfecha,prd_codigo,egr,ing,exi,d,prec,PC," & _
             " if(tdoc='AAU','A'," & _
             " if(tdoc='BAU','B'," & _
             " if(tdoc='IGR','C'," & _
             " if(tdoc='EGR','D'," & _
             " if(tdoc='COM','E'," & _
             " if(tdoc='IIM','F'," & _
             " if(tdoc='GRE','G'," & _
             " if(tdoc='DRE','H'," & _
             " if(tdoc='NRP','I'," & _
             " if(tdoc='DPR','J'," & _
             " if(tdoc='ETR','K'," & _
             " if(tdoc='ITR','L'," & _
             " if(tdoc='FAC','M'," & _
             " if(tdoc='DCL','N'," & _
             " if(tdoc='DPV','O',' '))))))))))))))) as dd " & _
             " FROM kardex " & _
             " ORDER BY prd_codigo,docfecha,dd,d,tdoc,doc "
    clsCon_Def.Ejecutar strSql
    strPrd = ""
    nn = 0
    lblPorcentaje.Caption = "Porcentaje: 0 %"
    While Not clsCon_Def.adorec_Def.EOF
        tdoc = clsCon_Def.adorec_Def("tdoc")
        Doc = clsCon_Def.adorec_Def("doc")
        docfecha = clsCon_Def.adorec_Def("docfecha")
        prd_codigo = clsCon_Def.adorec_Def("prd_codigo")
        egr = clsCon_Def.adorec_Def("egr")
        ing = clsCon_Def.adorec_Def("ing")
        d = clsCon_Def.adorec_Def("d")
        prec = clsCon_Def.adorec_Def("prec")
        PC = clsCon_Def.adorec_Def("PC")
        If strPrd <> prd_codigo Then
            exi = clsCon_Def.adorec_Def("exi")
            strPrd = prd_codigo
            PC = UltimoPPP(prd_codigo, desde)
        End If
        exi = exi + ing - egr
        vsfgKardex.AddItem prd_codigo & vbTab & tdoc & vbTab & _
                              docfecha & vbTab & Doc & vbTab & ing & vbTab & _
                              egr & vbTab & exi & vbTab & prec & vbTab & PC & vbTab & _
                              (egr + ing) * prec & vbTab & exi * PC
        clsCon_Def.adorec_Def.MoveNext
        nn = nn + 1
        lblPorcentaje.Caption = "Porcentaje: " & Round(nn / clsCon_Def.adorec_Def.RecordCount * 100, 2) & " %"
        lblPorcentaje.Refresh
    Wend
    strSql = " DROP TABLE kardex "
    clsCon_Def.Ejecutar strSql
End Sub
Private Function CalculaPPP(tdoc As String, Doc As String, docfecha As String, prd_codigo As String, egr As Long, ing As Long, exi As Long, d As String, prec As Double, PC As Double, PPP As Double) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    If d = 1 Then
        If tdoc = "COM" Or tdoc = "IIM" Then 'Compras e Importaciones
            PPP = Round((PPP * exi + ing * prec) / (exi + ing), 8)
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "AAU" Then ' Altas de Audit
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     ",det_ing_precio='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DCL" Then 'Devoluciones Cliente - Notas de Credito
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DRE" Then ' Devoluciones de Guias de Remision
            prec = CostoGuiaCli(prd_codigo, Doc)
            If exi + ing <> 0 Then
                PPP = Round((PPP * exi + ing * prec) / (exi + ing), 8)
            End If
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     ",det_ing_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "IGR" Then ' Guias de Proveedor
            PPP = UbicaProximaAdqui(prd_codigo, docfecha, exi, PPP)
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        End If
        CalculaPPP = PPP
        clsAux.Ejecutar strSql, "M"
    ElseIf d = 2 Then
        If tdoc = "FAC" Then
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DPV" Then
            If exi - ing <> 0 Then
                PPP = Round((PPP * exi - egr * prec) / (exi - ing), 8)
            End If
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "BAU" Then
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     ",det_egr_precio='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "GRE" Then
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "EGR" Then
            prec = CostoGuiaPro(prd_codigo, Doc)
            If exi - ing <> 0 Then
                PPP = Round((PPP * exi - egr * prec) / (exi - ing), 8)
            End If
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     ",det_egr_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        End If
        CalculaPPP = PPP
        clsAux.Ejecutar strSql, "M"
    Else
        CalculaPPP = PPP
    End If
    Set clsAux = Nothing
End Function
Private Sub ActualizaCosto(prd_codigo As String, PPP As Double)
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " UPDATE producto " & _
             " SET prd_costo='" & PPP & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND prd_codigo='" & prd_codigo & "' "
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Sub
Private Function UbicaProximaAdqui(prd_codigo As String, doc_fecha As String, exi As Long, PPP As Double) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT det_ing_precio,det_ing_cantidad " & _
             " FROM ingreso INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " AND ingreso.ing_codigo=det_ingreso.ing_codigo" & _
             " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo" & _
             " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ing_fecha>='" & doc_fecha & "' " & _
             " AND (ingreso.tip_ing_codigo='COM' OR ingreso.tip_ing_codigo='IIM') " & _
             " ORDER BY ing_fecha LIMIT 0,1 "
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        PPP = Round((PPP * exi + clsAux.adorec_Def("det_ing_cantidad") * clsAux.adorec_Def("det_ing_precio")) / (exi + clsAux.adorec_Def("det_ing_cantidad")), 8)
    End If
    UbicaProximaAdqui = PPP
    Set clsAux = Nothing
End Function
Private Function CostoGuiaCli(prd_codigo As String, Doc As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ing_factura " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND tip_ing_codigo='DRE'"
    clsAux.Ejecutar strSql
    strSql = " SELECT det_egr_costo " & _
             " FROM egreso INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
             " AND egreso.egr_codigo=det_egreso.egr_codigo" & _
             " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo" & _
             " AND det_egreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_codigo='" & clsAux.adorec_Def("ing_factura") & "'" & _
             " AND egreso.tip_egr_codigo='GRE' "
    clsAux.Ejecutar strSql
    CostoGuiaCli = clsAux.adorec_Def("det_egr_costo")
    strSql = " UPDATE det_ingreso " & _
             " SET det_ing_precio='" & clsAux.adorec_Def("det_egr_costo") & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND tip_ing_codigo='DRE' " & _
             " AND prd_codigo='" & prd_codigo & "'"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Function
Private Function CostoGuiaPro(prd_codigo As String, Doc As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT egr_factura " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & Doc & "'" & _
             " AND tip_egr_codigo='EGR'"
    clsAux.Ejecutar strSql
    strSql = " SELECT det_ing_costo " & _
             " FROM ingreso INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " AND ingreso.ing_codigo=det_ingreso.ing_codigo" & _
             " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo" & _
             " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.ing_codigo='" & clsAux.adorec_Def("egr_factura") & "'" & _
             " AND ingreso.tip_ing_codigo='IGR' "
    clsAux.Ejecutar strSql
    CostoGuiaPro = clsAux.adorec_Def("det_ing_costo")
    strSql = " UPDATE det_egreso " & _
             " SET det_egr_precio='" & clsAux.adorec_Def("det_ing_costo") & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & Doc & "'" & _
             " AND tip_egr_codigo='EGR' " & _
             " AND prd_codigo='" & prd_codigo & "'"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Function
Private Function UltimoPPP(prd_codigo As String, desde As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    Dim ing_fecha As String
    Dim ing_costo As Double
    Dim egr_fecha As String
    Dim egr_costo As Double
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ing_fecha,det_ing_costo " & _
             " FROM det_ingreso INNER JOIN ingreso ON ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
             " WHERE det_ing_costo!=0 " & _
             " AND det_ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND prd_codigo='" & prd_codigo & "' " & _
             " AND ing_fecha < '" & desde & "' " & _
             " ORDER BY ing_fecha DESC, ingreso.tip_ing_codigo DESC, ingreso.ing_codigo DESC LIMIT 0,1"
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        ing_fecha = clsAux.adorec_Def("ing_fecha")
        ing_costo = clsAux.adorec_Def("det_ing_costo")
    Else
        ing_costo = 0
    End If
    strSql = " SELECT egr_fecha,det_egr_costo " & _
             " FROM det_egreso INNER JOIN egreso ON egreso.egr_codigo=det_egreso.egr_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
             " WHERE det_egr_costo!=0 " & _
             " AND det_egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND prd_codigo='" & prd_codigo & "' " & _
             " AND egr_fecha < '" & desde & "' " & _
             " AND egreso.tip_egr_codigo!='EGR' AND egreso.tip_egr_codigo!='GRE' " & _
             " AND if(egreso.tip_egr_codigo='FAC',egr_observacion NOT LIKE 'FACTURA ANULA%',1=1) " & _
             " ORDER BY egr_fecha DESC, egreso.tip_egr_codigo DESC, egreso.egr_codigo DESC LIMIT 0,1 "
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        egr_fecha = clsAux.adorec_Def("egr_fecha")
        egr_costo = clsAux.adorec_Def("det_egr_costo")
    Else
        egr_costo = 0
    End If
    Set clsAux = Nothing
    If ing_costo > 0 And egr_costo > 0 Then
        If ing_fecha > egr_fecha Then
            UltimoPPP = ing_costo
        Else
            UltimoPPP = egr_costo
        End If
    Else
        If ing_costo > 0 Then
            UltimoPPP = ing_costo
        ElseIf egr_costo > 0 Then
            UltimoPPP = egr_costo
        Else
            UltimoPPP = 0
        End If
    End If
End Function
Private Sub CreaKardex(desde As String, hasta As String, prd_codigo As String)
    Dim strSql As String
    strSql = " CREATE TEMPORARY TABLE ProdMovAUX " & _
             " SELECT DISTINCT CONCAT(COALESCE(prd_codigo,' '),'                    ') as prd_codigo " & _
             " FROM (egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo = det_egreso.tip_egr_codigo " & _
             " AND egreso.emp_codigo = det_egreso.emp_codigo AND egreso.egr_codigo = det_egreso.egr_codigo) " & _
             " WHERE egr_fecha BETWEEN '" & desde & "' AND '" & hasta & "'" & _
             " AND egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND if(egreso.tip_egr_codigo='FAC',egr_observacion NOT LIKE 'FACTURA ANULADA%',1) "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO ProdMovAUX " & _
             " SELECT DISTINCT CONCAT(COALESCE(prd_codigo,' '),'                    ') as prd_codigo " & _
             " FROM (ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo " & _
             " AND ingreso.emp_codigo = det_ingreso.emp_codigo AND ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
             " WHERE ing_fecha BETWEEN '" & desde & "' AND '" & hasta & "'" & _
             " AND ingreso.emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql
    strSql = " CREATE TEMPORARY TABLE ProdMov (prd_codigo varchar(20) NOT NULL DEFAULT '', PRIMARY KEY(prd_codigo))"
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO ProdMov " & _
             " SELECT DISTINCT TRIM(prd_codigo) as prd_codigo " & _
             " FROM ProdMovAUX " & _
             " WHERE prd_codigo NOT LIKE 'PR-%' AND prd_codigo NOT LIKE 'MDO%' " & _
             " AND prd_codigo LIKE '" & prd_codigo & "' " & _
             " ORDER BY prd_codigo "
    clsCon_Def.Ejecutar strSql
    strSql = " DROP TABLE ProdMovAUX "
    clsCon_Def.Ejecutar strSql
    
    strSql = " CREATE TEMPORARY TABLE AuxiExis " & _
             " SELECT det_egreso.dep_codigo, det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) " & _
             " AS egr, 0 AS ing, 0 AS exi, egreso.emp_codigo, egreso.egr_factura as factura" & _
             " FROM (egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo = det_egreso.tip_egr_codigo " & _
             " AND egreso.emp_codigo = det_egreso.emp_codigo AND egreso.egr_codigo = det_egreso.egr_codigo) " & _
             " INNER JOIN ProdMov ON det_egreso.prd_codigo=ProdMov.prd_codigo " & _
             " WHERE egreso.egr_fecha >= '" & desde & "' " & _
             " AND egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.tip_egr_codigo!='ETR' " & _
             " AND if(egreso.tip_egr_codigo='FAC',egr_observacion NOT LIKE 'FACTURA ANULADA%',1) " & _
             " GROUP BY det_egreso.prd_codigo "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO AuxiExis " & _
             " SELECT det_ingreso.dep_codigo,det_ingreso.prd_codigo, 0 AS egr, " & _
             " SUM(det_ingreso.det_ing_cantidad) AS ing, 0 AS exi, ingreso.emp_codigo, ingreso.ing_factura as factura" & _
             " FROM (ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo " & _
             " AND ingreso.emp_codigo = det_ingreso.emp_codigo AND ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
             " INNER JOIN ProdMov ON det_ingreso.prd_codigo=ProdMov.prd_codigo " & _
             " WHERE ingreso.ing_fecha >= '" & desde & "' " & _
             " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.tip_ing_codigo!='ITR' " & _
             " GROUP BY det_ingreso.prd_codigo "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO AuxiExis " & _
             " SELECT '' as dep_codigo,existencia.prd_codigo, 0 AS egr, 0 AS ing, " & _
             " SUM(existencia.exi_cantidad) AS exi, existencia.emp_codigo, '' as factura" & _
             " FROM existencia " & _
             " INNER JOIN ProdMov ON existencia.prd_codigo=ProdMov.prd_codigo " & _
             " WHERE existencia.emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY existencia.prd_codigo "
    clsCon_Def.Ejecutar strSql
    strSql = " CREATE TEMPORARY TABLE kardex (  tdoc char(24) NOT NULL default '', doc decimal(6,0) NOT NULL default '0'," & _
             " docfecha datetime NOT NULL,lin_codigo char(3) default NULL, mar_codigo char(3) default NULL, " & _
             " prd_nombre varchar(50) NOT NULL default '', prd_codigo varchar(20) NOT NULL default '', " & _
             " egr decimal(9,0) NOT NULL default '0', ing decimal(9,0) NOT NULL default '0',exi decimal(9,0) default NULL, " & _
             " prd_baja smallint(5) NOT NULL default '0',dep_codigo char(3) NOT NULL ,d int(1) NOT NULL default '0', " & _
             " factura char(30) default NULL,per_codigo varchar(8) NOT NULL default '',prec decimal(14,4) NOT NULL default '0.0000', " & _
             " PC decimal(14,4) default '0.0000') "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO kardex " & _
             " SELECT 'SALDO INICIAL A LA FECHA' as tdoc, 0 as doc, '" & desde & "' as docfecha, lin_codigo, mar_codigo, producto.prd_nombre, producto.prd_codigo, 0 AS egr, " & _
             " 0 AS ing, (SUM(AuxiExis.exi) + SUM(AuxiExis.egr) - SUM(AuxiExis.ing)) AS exi, producto.prd_baja, dep_codigo,0 as d, factura,'                    ' as per_codigo,0.0000 as prec,prd_costo as PC " & _
             " FROM (ProdMov INNER JOIN producto ON ProdMov.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN AuxiExis ON producto.prd_codigo = AuxiExis.prd_codigo AND producto.emp_codigo = AuxiExis.emp_codigo) " & _
             " WHERE producto.emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY lin_codigo, mar_codigo, producto.prd_nombre, producto.prd_codigo " & _
             " ORDER BY lin_codigo, mar_codigo, producto.prd_nombre "
    clsCon_Def.Ejecutar strSql
    strSql = " DROP TABLE AuxiExis"
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO kardex " & _
             " SELECT ingreso.tip_ing_codigo as tdoc,ingreso.ing_codigo as doc,ingreso.ing_fecha as docfecha, lin_codigo, mar_codigo, producto.prd_nombre, det_ingreso.prd_codigo, 0 AS egr, " & _
             " det_ingreso.det_ing_cantidad AS ing, 0 AS exi, producto.prd_baja,det_ingreso.dep_codigo,1 as d, ingreso.ing_factura as factura, ingreso.per_codigo,det_ing_precio,det_ing_costo " & _
             " FROM ((ProdMov INNER JOIN producto ON ProdMov.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN det_ingreso ON producto.prd_codigo = det_ingreso.prd_codigo AND producto.emp_codigo = det_ingreso.emp_codigo) " & _
             " INNER JOIN ingreso ON det_ingreso.ing_codigo=ingreso.ing_codigo AND det_ingreso.emp_codigo = ingreso.emp_codigo AND det_ingreso.tip_ing_codigo = ingreso.tip_ing_codigo) " & _
             " WHERE producto.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.ing_fecha BETWEEN '" & desde & "' AND '" & hasta & "' " & _
             " AND ingreso.tip_ing_codigo!='ITR' " & _
             " ORDER BY lin_codigo, mar_codigo, producto.prd_nombre,tdoc,ingreso.ing_fecha "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO kardex " & _
             " SELECT egreso.tip_egr_codigo as tdoc,egreso.egr_codigo as doc,egreso.egr_fecha as docfecha, lin_codigo, mar_codigo, producto.prd_nombre, det_egreso.prd_codigo, det_egreso.det_egr_cantidad AS egr, " & _
             " 0 AS ing, 0 AS exi, producto.prd_baja,det_egreso.dep_codigo,2 as d, egreso.egr_factura as factura,egreso.per_codigo,det_egr_precio,det_egr_costo " & _
             " FROM ((ProdMov INNER JOIN producto ON ProdMov.prd_codigo=producto.prd_codigo " & _
             " INNER JOIN det_egreso ON producto.prd_codigo = det_egreso.prd_codigo AND producto.emp_codigo = det_egreso.emp_codigo) " & _
             " INNER JOIN egreso ON det_egreso.egr_codigo=egreso.egr_codigo AND det_egreso.emp_codigo = egreso.emp_codigo AND det_egreso.tip_egr_codigo = egreso.tip_egr_codigo) " & _
             " WHERE producto.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_fecha BETWEEN '" & desde & "' AND '" & hasta & "' " & _
             " AND egreso.tip_egr_codigo!='ETR' " & _
             " AND if(egreso.tip_egr_codigo='FAC',egr_observacion NOT LIKE 'FACTURA ANULADA%',1) " & _
             " ORDER BY lin_codigo, mar_codigo, producto.prd_nombre,tdoc,egreso.egr_fecha "
    clsCon_Def.Ejecutar strSql
End Sub
