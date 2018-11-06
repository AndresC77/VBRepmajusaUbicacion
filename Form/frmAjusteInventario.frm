VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAjusteInventario 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustar Inventario según Conteos"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmAjusteInventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   10785
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   98
      TabIndex        =   2
      Top             =   120
      Width           =   10455
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Value           =   42064.9507986111
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGDet 
         Height          =   4335
         Left            =   240
         TabIndex        =   10
         Top             =   3840
         Width           =   9975
         _cx             =   86721723
         _cy             =   86711774
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAjusteInventario.frx":030A
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
         ExplorerBar     =   5
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
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar Todos"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdCargarTodos 
         Caption         =   "Cargar todos los Productos"
         Height          =   375
         Left            =   5745
         TabIndex        =   8
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CommandButton cmdCargarSoloConteo 
         Caption         =   "Cargar solo Productos contados"
         Height          =   375
         Left            =   2385
         TabIndex        =   7
         Top             =   3240
         Width           =   2655
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   9975
         _cx             =   86721723
         _cy             =   86708176
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         AllowUserResizing=   0
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
         FormatString    =   $"frmAjusteInventario.frx":03EE
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
         ExplorerBar     =   5
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
      Begin VB.TextBox txtObs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   8505
         Width           =   10110
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Ajuste:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label lblObserv 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   8280
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdAjustar 
      Caption         =   "&Ajustar"
      Height          =   375
      Left            =   3896
      TabIndex        =   0
      Top             =   9405
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5433
      TabIndex        =   1
      Top             =   9405
      Width           =   1455
   End
End
Attribute VB_Name = "frmAjusteInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para visualizar los ingresos de mercadería realizados por concepto de #
'#  Importaciones,  esta forma es solo de visualización, no permite la edición. #
'#  frmVerIngImp V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los ingresos de mercadería a una determinada  emp-   #
'#  presa por concepto de Importaciones.                                        #
'#  En esta ventana solo se puede visuallizar cualesquiera de los ingresos por  #
'#  este concepto pero no se puede realizar ningún cambio.                      #
'#  Se puede escoger el número de documento o ingresar dicho número en el combo #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se consulta los egresos realizados de tipo     #
'#                 INI.                                                         #
'#    persona    : En esta tabla se consulta los datos del proveedor al que se  #
'#                 le adquirió la mercadería y se importó.                      #
'#    det_ingreso: En esta tabla se consulta los detalles del ingreso.          #
'#    producto   : En esta tabla se consulta el nombre del producto.            #
'#    deposito   : En esta tabla se consulta el nombre del depósito.            #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#    limpiarFxGD() : Permite borrar el flexgrid utilizado para cuando se       #
'#                    realiza un cambio de documento.                           #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsCon_det As New clsConsulta
Private Opcion As String
Private Sub CargaAjustes(strPrd As String, strInv As String)
    Dim strSQL As String
    strSQL = " DROP TABLE IF EXISTS Conteo "
    clsCon_det.Ejecutar strSQL
    strSQL = " CREATE TABLE Conteo ( " & _
             " tipo char(1) NOT NULL default ''," & _
             " emp_codigo char(3) NOT NULL default ''," & _
             " dep_codigo char(3) NOT NULL default ''," & _
             " prd_codigo varchar(40) NOT NULL default ''," & _
             " conteo decimal(14,0) default '0'," & _
             " existencia decimal(14,0) default '0'," & _
             " cod_conteo varchar(255) NOT NULL default ''," & _
             " KEY pri (tipo,emp_codigo,dep_codigo,prd_codigo))"
    clsCon_det.Ejecutar strSQL
    strSQL = " INSERT INTO Conteo " & _
             " SELECT 'C' as c,inventario.emp_codigo,dep_codigo,det_inventario.prd_codigo,sum(det_inv_cantidad),'0',GROUP_CONCAT(inventario.inv_codigo) " & _
             " FROM inventario INNER JOIN det_inventario ON inventario.emp_codigo=det_inventario.emp_codigo " & _
             " AND inventario.inv_codigo=det_inventario.inv_codigo " & _
             " INNER JOIN producto on det_inventario.emp_codigo=producto.emp_codigo" & _
             " and det_inventario.prd_codigo=producto.prd_codigo " & _
             " WHERE inventario.emp_codigo='" & strEmpresa & "' AND" & strInv & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " ORDER BY prd_codigo "
    clsCon_det.Ejecutar strSQL
    strSQL = " INSERT INTO Conteo " & _
             " SELECT 'E' as c,existencia.emp_codigo,dep_codigo,existencia.prd_codigo,'0',sum(exi_cantidad),NULL " & _
             " FROM existencia " & _
             " INNER JOIN producto on existencia.emp_codigo=producto.emp_codigo" & _
             " and existencia.prd_codigo=producto.prd_codigo " & _
             " WHERE existencia.emp_codigo='" & strEmpresa & "' AND" & strPrd & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " ORDER BY prd_codigo "
    clsCon_det.Ejecutar strSQL
    strSQL = " INSERT INTO Conteo " & _
             " SELECT 'E' as c,ingreso.emp_codigo,dep_codigo,det_ingreso.prd_codigo,'0',-1 * sum(det_ing_cantidad),NULL " & _
             " FROM ingreso inner join det_ingreso " & _
             " ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
             " AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
             " AND ingreso.ing_anulado=0 " & _
             " AND ingreso.ing_fecha>='2016-10-30' " & _
             " INNER JOIN producto on det_ingreso.emp_codigo=producto.emp_codigo" & _
             " and det_ingreso.prd_codigo=producto.prd_codigo " & _
             " " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' AND" & strPrd & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " ORDER BY prd_codigo "
    clsCon_det.Ejecutar strSQL
    strSQL = " INSERT INTO Conteo " & _
             " SELECT 'E' as c,egreso.emp_codigo,dep_codigo,det_egreso.prd_codigo,'0',sum(det_egr_cantidad),NULL " & _
             " FROM egreso inner join det_egreso " & _
             " ON egreso.emp_codigo=det_egreso.emp_codigo " & _
             " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
             " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
             " AND egreso.egr_anulado=0 " & _
             " AND egreso.egr_fecha>='2016-10-30'" & _
             " INNER JOIN producto on det_egreso.emp_codigo=producto.emp_codigo" & _
             " and det_egreso.prd_codigo=producto.prd_codigo " & _
             " " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' AND" & strPrd & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " ORDER BY prd_codigo "
    clsCon_det.Ejecutar strSQL
    strSQL = " SELECT dep_codigo,Conteo.prd_codigo,prd_nombre,COALESCE(sum(conteo),0) as con,COALESCE(sum(existencia),0) as exi,COALESCE(sum(conteo),0) - COALESCE(sum(existencia),0) as dif,GROUP_CONCAT(cod_conteo) as conteo " & _
             " FROM Conteo INNER JOIN producto ON Conteo.emp_codigo=producto.emp_codigo AND Conteo.prd_codigo=producto.prd_codigo " & _
             " WHERE Conteo.emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY Conteo.emp_codigo,dep_codigo,prd_codigo " & _
             " HAVING dif != 0 " & _
             " ORDER BY prd_codigo,dep_codigo "
    clsCon_det.Ejecutar strSQL
    Set VSFGDet.DataSource = clsCon_det.adorec_Def.DataSource
End Sub

Private Sub cmdAjustar_Click()
    Dim strSQL As String
    Dim strInv As String
    Dim strPrd As String
    Dim AAU As Long
    Dim BAU As Long
    Dim Fecha As String
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Fecha = Format(dtpFecha.Value, "yyyy-mm-dd")
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    clsIngreso.NuevoIng False, "AAU", False, "001", "001", , , , Fecha, , , txtObs.Text
    clsEgreso.NuevoEgr False, "BAU", False, "001", "001", , , , Fecha, , , txtObs.Text
    
    strSQL = " SELECT prd_codigo,dep_codigo,COALESCE(sum(conteo),0) - COALESCE(sum(existencia),0) as dif " & _
             " FROM Conteo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " HAVING dif > 0 " & _
             " ORDER BY prd_codigo,dep_codigo "
    clsCon_det.Ejecutar strSQL
    While Not clsCon_det.adorec_Def.EOF
        clsIngreso.NuevoDetIng clsCon_det.adorec_Def("prd_codigo"), clsCon_det.adorec_Def("dep_codigo"), clsCon_det.adorec_Def("dif")
        clsCon_det.adorec_Def.MoveNext
    Wend
    InicializarContenedorRecurrente
    strSQL = " SELECT prd_codigo,dep_codigo,COALESCE(sum(existencia),0) - COALESCE(sum(conteo),0) as dif " & _
             " FROM Conteo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " GROUP BY emp_codigo,dep_codigo,prd_codigo " & _
             " HAVING dif > 0 " & _
             " ORDER BY prd_codigo,dep_codigo "
    clsCon_det.Ejecutar strSQL
    While Not clsCon_det.adorec_Def.EOF
        clsEgreso.NuevoDetEgr clsCon_det.adorec_Def("prd_codigo"), clsCon_det.adorec_Def("dep_codigo"), clsCon_det.adorec_Def("dif")
        clsCon_det.adorec_Def.MoveNext
    Wend
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            strInv = strInv & " OR inventario.inv_codigo='" & VSFG.TextMatrix(i, 1) & "'"
        End If
    Next i
    strInv = " (1=0" & strInv & " )"
    If Opcion = "SOLO" Then
        strSQL = " SELECT DISTINCT prd_codigo,dep_codigo " & _
                 " FROM inventario INNER JOIN det_inventario ON inventario.emp_codigo=det_inventario.emp_codigo AND inventario.inv_codigo=det_inventario.inv_codigo " & _
                 " WHERE inventario.emp_codigo='" & strEmpresa & "' "
        strSQL = strSQL & " AND " & strInv
        clsCon_det.Ejecutar strSQL
        While Not clsCon_det.adorec_Def.EOF
            strPrd = strPrd & " OR (producto.prd_codigo='" & clsCon_det.adorec_Def("prd_codigo") & "' AND dep_codigo='" & clsCon_det.adorec_Def("dep_codigo") & "')"
            clsCon_det.adorec_Def.MoveNext
        Wend
        strPrd = " (1=0" & strPrd & " )"
    Else
        strPrd = " producto.prd_codigo LIKE '%' "
    End If
    strSQL = " UPDATE inventario SET inv_estado=1 " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    strSQL = strSQL & " AND " & strInv
    clsCon_det.Ejecutar strSQL, "M"
    
    MsgBox "Es necesario que se haga un reconteo para actualizar los dias en bodega de las existencias", vbInformation, "Ajuste"
    Unload Me
End Sub

Private Sub cmdCargarSoloConteo_Click()
    Dim strPrd As String
    Dim i As Long
    Dim strSQL As String
    Dim strInv As String
    strSQL = " SELECT DISTINCT prd_codigo " & _
             " FROM inventario INNER JOIN det_inventario ON inventario.emp_codigo=det_inventario.emp_codigo AND inventario.inv_codigo=det_inventario.inv_codigo " & _
             " WHERE inventario.emp_codigo='" & strEmpresa & "' "
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            strInv = strInv & " OR inventario.inv_codigo='" & VSFG.TextMatrix(i, 1) & "'"
        End If
    Next i
    strInv = " (1=0" & strInv & " )"
    strSQL = strSQL & " AND " & strInv
    clsCon_det.Ejecutar strSQL
    While Not clsCon_det.adorec_Def.EOF
        strPrd = strPrd & " OR producto.prd_codigo='" & clsCon_det.adorec_Def("prd_codigo") & "'"
        clsCon_det.adorec_Def.MoveNext
    Wend
    strPrd = " (1=0" & strPrd & " )"
    Opcion = "SOLO"
    CargaAjustes strPrd, strInv
End Sub

Private Sub cmdCargarTodos_Click()
    Dim strPrd As String
    Dim i As Long
    Dim strSQL As String
    Dim strInv As String
    For i = 1 To VSFG.Rows - 1
        If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
            strInv = strInv & " OR inventario.inv_codigo='" & VSFG.TextMatrix(i, 1) & "'"
        End If
    Next i
    strInv = " (1=0" & strInv & " )"
    strPrd = " producto.prd_codigo LIKE '%' "
    Opcion = "TODO"
    CargaAjustes strPrd, strInv
End Sub

Private Sub cmdSeleccionar_Click()
    Dim i As Long
    If cmdSeleccionar.Caption = "Seleccionar Todos" Then
        For i = 1 To VSFG.Rows - 1
            VSFG.TextMatrix(i, 0) = 1
        Next i
        cmdSeleccionar.Caption = "Quitar Selección"
    Else
        For i = 1 To VSFG.Rows - 1
            VSFG.TextMatrix(i, 0) = 0
        Next i
        cmdSeleccionar.Caption = "Seleccionar Todos"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim strSQL As String
    On Error Resume Next
    strSQL = " DROP TABLE IF EXISTS Conteo "
    clsCon_det.Ejecutar strSQL
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_det = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
  clsConsu.Actualizar
  Set VSFG.DataSource = clsConsu.adorec_Def.DataSource

End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 200
    Set ucrtVSFG.VSFGControl = VSFGDet
    ucrtVSFG.Inicializar False, False, False, True, True, True, False, False, False
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    'Descompone la fecha actual  en día, mes y año
    
    dtpFecha.Value = HoyDia
    'Ejecuta un SQL contra la base de datosCONCAT(a.ing_codigo) AS ing_codigo
    strSQL = " SELECT inv_estado,inv_codigo,inv_fecha,inv_observacion " & _
             " FROM inventario " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND inv_estado=0 " & _
             " ORDER BY inv_codigo"
    clsConsu.Ejecutar (strSQL)
    Set VSFG.DataSource = clsConsu.adorec_Def.DataSource
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores
        
    If clsConsu.adorec_Def.EOF Then
        MsgBox "No existen Conteos sin procesar", vbInformation, "SisAdmi"
        'Unload Me
    End If
    
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 0 Or Col > 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 And Row > 0 Then
        If Abs(VSFG.TextMatrix(Row, 0)) = 1 Then
            VSFG.Select Row, 0, Row, 3
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HC0FFFF
            VSFG.Select Row, 0
        ElseIf Abs(VSFG.TextMatrix(Row, 0)) = 0 Then
            VSFG.Select Row, 0, Row, 3
            VSFG.FillStyle = flexFillRepeat
            VSFG.CellBackColor = &HFFFFFF
            VSFG.Select Row, 0
        End If
    End If
End Sub
