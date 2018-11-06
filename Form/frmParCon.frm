VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmParCon 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Contables"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "frmParCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4760
      TabIndex        =   9
      Top             =   5520
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2960
      TabIndex        =   8
      Top             =   5520
      Width           =   1700
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _cx             =   3201
      _cy             =   1085
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmParCon.frx":030A
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
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar CtaConta...el %$$% botón que buscaba!!"
      Height          =   855
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Parámetro"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   380
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9135
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2895
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   8640
         _cx             =   15240
         _cy             =   5106
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParCon.frx":0371
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmParCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsSql1 As New clsConsulta
Private clsCta As New clsConsulta
Private strSql As String
Private Hacer As Boolean

Private Tipo As String
Private Tipo2 As String
Private Sub IniDato()
    Tipo = "Parámetros Contables"
    Tipo2 = "Parámetros Contables"
    Me.Caption = Tipo
End Sub


    
    


Private Sub cmbAceptar_Click()
    Dim i As Long, j As Long, k As Long
    Dim control As Long 'control de que esten llenos los datos
    If VSFG.Rows > 1 Then
        VSFG.Select 1, VSFG.Cols - 1
        VSFG.Sort = flexSortGenericDescending
    End If
    control = 0
    
    For i = 1 To VSFG.Rows - 1
        'update
        
        
        
        
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
            If dcmbTipo.BoundText <> "INVENTARIOS" Then
                'For j = 1 To VSFG.Rows - 1
                    For j = 1 To VSFG.Rows - 1
                        If VSFG.TextMatrix(i, 3) = VSFG.TextMatrix(j, 3) And i <> j Then
                            MsgBox "No puede poner la misma cuenta contable dos veces " & vbNewLine & "(" & VSFG.TextMatrix(i, 3) & " - " & VSFG.TextMatrix(i, 4) & ").", vbInformation, "Contabilidad "
                            Exit Sub
                        End If
                    Next j
                'Next i
            End If
    
        strSql = " UPDATE parametro_contable SET par_con_descripcion='" & UCase(VSFG.TextMatrix(i, 2)) & "', par_con_cta_codigo ='" & VSFG.TextMatrix(i, 3) & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND par_con_tipo='" & dcmbTipo & "' AND par_con_codigo='" & VSFG.TextMatrix(i, 1) & "'"
        clsSql.Ejecutar strSql, "M"
   
       
        'insert
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
        'controla que este lleno los datos
            
        'delete
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
           
           
        ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
            Exit For
        End If
    Next i
    If control = 0 Then
        dcmbTipo_Change
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    Set clsSql1 = Nothing
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 1 Then
            Cancel = True
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 1 Then
            Cancel = True
        End If
    End If
End Sub


Private Sub CmdCerrar_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
    GrabarPlan "1", "ACTIVO", 1, 1, "G"
    GrabarPlan "1.1", "ACTIVO CORRIENTE", 2, 1, "G"
    GrabarPlan "1.1.01", "DISPONIBLE", 3, 1, "G"
    GrabarPlan "1.1.01.01", "CAJA Y FONDOS FIJOS", 4, 1, "G"
    GrabarPlan "1.1.01.01.001", "Cobranza a Depositar", 5, 0, "G"
    GrabarPlan "1.1.01.01.002", "Fondo Fijo Scz", 5, 0, "G"
    GrabarPlan "1.1.01.01.003", "Fondo Fijo LPZ", 5, 0, "G"
    GrabarPlan "1.1.01.01.004", "Fondo Fijo CBBA", 5, 0, "G"
    GrabarPlan "1.1.01.02", "BANCOS", 4, 1, "G"
    GrabarPlan "1.1.01.02.001", "Banco Union Bs.", 5, 0, "G"
    GrabarPlan "1.1.01.02.002", "Banco Union $us.", 5, 0, "G"
    GrabarPlan "1.1.01.02.003", "Banco de Credito Bs.", 5, 0, "G"
    GrabarPlan "1.1.01.02.004", "Banco de Credito $us.", 5, 0, "G"
    GrabarPlan "1.1.02", "INVENTARIO TEMPORARIOS", 3, 1, "G"
    GrabarPlan "1.1.02.01", "VALORES PUBLICOS", 4, 0, "G"
    GrabarPlan "1.1.02.02", "DEPOSITOS BANCARIOS", 4, 0, "G"
    GrabarPlan "1.1.02.03", "OTRAS", 4, 0, "G"
    GrabarPlan "1.1.03", "CREDITO POR VENTAS", 3, 1, "G"
    GrabarPlan "1.1.03.01", "DEUDORES PLAZA 3EROS.", 4, 1, "G"
    GrabarPlan "1.1.03.01.001", "Deudores por venta local", 5, 0, "G"
    GrabarPlan "1.1.03.02", "DEUDORES EXTERIOR 3EROS.", 4, 0, "G"
    GrabarPlan "1.1.03.03", "DOCUEMENTOS COBRAR 3EROS.", 4, 1, "G"
    GrabarPlan "1.1.03.03.001", "Cheques pos Datados Clientes", 5, 0, "G"
    GrabarPlan "1.1.03.04", "COMPA¥IA VINCULADAS", 4, 1, "G"
    GrabarPlan "1.1.03.04.001", "Cuentas x Cobrar Recalcine Ecuador", 5, 0, "G"
    GrabarPlan "1.1.03.05", "DEUDORES EN GESTION", 4, 1, "G"
    GrabarPlan "1.1.03.05.001", "Cuentas en Ejecucion Clientes", 5, 0, "G"
    GrabarPlan "1.1.03.06", "MENOS PREVISION PARA DEUDORES INCOB.", 4, 1, "G"
    GrabarPlan "1.1.03.06.001", "Prevision para Incobrables", 5, 0, "G"
    GrabarPlan "1.1.04", "OTROS CREDITOS", 3, 1, "G"
    GrabarPlan "1.1.04.01", "ANTICIPOS A PROVEEDORES", 4, 0, "G"
    GrabarPlan "1.1.04.02", "ANTICIPOS DE IMPUESTOS", 4, 1, "G"
    GrabarPlan "1.1.04.02.001", "Anticipos de Impuestos I.T.", 5, 0, "G"
    GrabarPlan "1.1.04.03", "ANTICIPOS Y PRESTAMOS AL PERSONAL", 4, 1, "G"
    GrabarPlan "1.1.04.03.001", "Anticipos de Sueldos", 5, 0, "G"
    GrabarPlan "1.1.04.03.002", "Prestamos al Personal", 5, 0, "G"
    GrabarPlan "1.1.04.04", "ANTICIPOS PARA GASTOS", 4, 0, "G"
    GrabarPlan "1.1.04.05", "DEPOSITOS EN GARANTIA", 4, 1, "G"
    GrabarPlan "1.1.04.05.001", "Deposito Boletas de Garantia (licitacion", 5, 0, "G"
    GrabarPlan "1.1.04.06", "CREDITO FISCAL", 4, 1, "G"
    GrabarPlan "1.1.04.06.001", "Credito Fiscal", 5, 0, "G"
    GrabarPlan "1.1.04.07", "CUENTAS DE DIRECTORES Y ACCIONISTAS", 4, 0, "G"
    GrabarPlan "1.1.04.08", "PRESTAMOS A 3EROS", 4, 0, "G"
    GrabarPlan "1.1.04.09", "PRESTAMOS A COMPA¥IAS VINCULADAS", 4, 0, "G"
    GrabarPlan "1.1.04.10", "GASTOS DIFERIDOS", 4, 1, "G"
    GrabarPlan "1.1.04.10.001", "Adelanto Desp. Aduanero", 5, 0, "G"
    GrabarPlan "1.1.04.10.002", "Seguro Pag. Adelantado BISA", 5, 0, "G"
    GrabarPlan "1.1.04.10.003", "Orpan", 5, 0, "G"
    GrabarPlan "1.1.04.10.004", "Amortizacion Orpan", 5, 0, "G"
    GrabarPlan "1.1.04.10.005", "Adelantos a Rendir", 5, 0, "G"
    GrabarPlan "1.1.04.11", "DIVERSOS", 4, 1, "G"
    GrabarPlan "1.1.04.11.001", "Otras Cuentas a Cobrar", 5, 0, "G"
    GrabarPlan "1.1.05", "BIENES DE CAMBIO     ", 3, 1, "G"
    GrabarPlan "1.1.05.01", "MERCADERIA DE REVENTA", 4, 1, "G"
    GrabarPlan "1.1.05.01.001", "Inventario de Mercaderia", 5, 0, "G"
    GrabarPlan "1.1.05.01.002", "Existencia de Productos en Poder 3eros", 5, 0, "G"
    GrabarPlan "1.1.05.02", "PRODUCTOS TERMINADOS", 4, 0, "G"
    GrabarPlan "1.1.05.03", "GRANELES", 4, 0, "G"
    GrabarPlan "1.1.05.04", "PRODUCTOS EN PROCESO", 4, 0, "G"
    GrabarPlan "1.1.05.05", "MATERIA PRIMA Y EXCIPIENTES", 4, 0, "G"
    GrabarPlan "1.1.05.06", "MATERIALES Y SUMINISTROS", 4, 0, "G"
    GrabarPlan "1.1.05.07", "MUESTRAS MEDICAS", 4, 1, "G"
    GrabarPlan "1.1.05.07.001", "Muestras Medicas", 5, 0, "G"
    GrabarPlan "1.1.05.08", "MATERIAL PROMOCIONAL", 4, 1, "G"
    GrabarPlan "1.1.05.08.001", "Literatura y Material Promocional", 5, 0, "G"
    GrabarPlan "1.1.05.09", "IMPORTACIONES EN TRAMITE", 4, 1, "G"
    GrabarPlan "1.1.05.09.001", "Mercaderia en Trans.Originales", 5, 0, "G"
    GrabarPlan "1.1.05.09.002", "Mercaderia en Trans.Muestra Medica", 5, 0, "G"
    GrabarPlan "1.1.05.10", "MENOS PREVISION POR DESVALORIZACION Y OB", 4, 1, "G"
    GrabarPlan "1.1.05.10.001", "Existencia de Productos Vencidos/Da¤ados", 5, 0, "G"
    GrabarPlan "1.2", "ACTIVO NO CORRIENTE", 2, 1, "G"
    GrabarPlan "1.2.01", "CREDITOS POR VENTA", 3, 0, "G"
    GrabarPlan "1.2.02", "OTROS CREDITOS", 3, 1, "G"
    GrabarPlan "1.2.02.11", "DIVERSOS", 4, 1, "G"
    GrabarPlan "1.2.02.11.001", "Depositos en Garantia", 5, 0, "G"
    GrabarPlan "1.2.02.11.002", "Cotas $us.", 5, 0, "G"
    GrabarPlan "1.2.02.11.003", "Comteco $us.", 5, 0, "G"
    GrabarPlan "1.2.02.11.004", "Cotel $us.", 5, 0, "G"
    GrabarPlan "1.2.03", "BIENES DE CAMBIO NO CORRIENTE", 3, 0, "G"
    GrabarPlan "1.2.04", "INVERSIONES A LARGO PLAZO", 3, 1, "G"
    GrabarPlan "1.2.04.01", "VALORES PUBLICOS", 4, 0, "G"
    GrabarPlan "1.2.04.02", "DEPOSITOS BANCARIOS", 4, 0, "G"
    GrabarPlan "1.2.04.03", "INVERSIONES EN COMPA¤IAS VINCULADAS", 4, 0, "G"
    GrabarPlan "1.2.04.04", "OTRAS INVERSIONES", 4, 0, "G"
    GrabarPlan "1.2.05", "BIENES DE USO", 3, 1, "G"
    GrabarPlan "1.2.05.01", "VALORES BRUTOS", 4, 1, "G"
    GrabarPlan "1.2.05.01.001", "Muebles y Utiles", 5, 0, "G"
    GrabarPlan "1.2.05.01.002", "Vehiculo", 5, 0, "G"
    GrabarPlan "1.2.05.01.003", "Equipo de Informatica", 5, 0, "G"
    GrabarPlan "1.2.05.02", "MENOS AMORTIZACIONES ACUMULADAS", 4, 1, "G"
    GrabarPlan "1.2.05.02.001", "Amortiz.Acumulada Muebles y Utiles", 5, 0, "G"
    GrabarPlan "1.2.05.02.002", "Amortiz.Acumulada Vehiculo", 5, 0, "G"
    GrabarPlan "1.2.05.02.003", "Amortiz.Acumulada Equipos Informatica", 5, 0, "G"
    GrabarPlan "1.2.06", "INTANGIBLES", 3, 1, "G"
    GrabarPlan "1.2.06.01", "VALORES BRUTOS", 4, 0, "G"
    GrabarPlan "1.2.06.02", "MENOS AMORTIZACIONES ACUMULADAS         ", 4, 0, "G"
    GrabarPlan "2", "P A S I V O", 1, 1, "G"
    GrabarPlan "2.1", "PASIVO CORRIENTE", 2, 1, "G"
    GrabarPlan "2.1.01", "DEUDAS COMERCIALES", 3, 1, "G"
    GrabarPlan "2.1.01.01", "PROVEEDORES LOCALES 3EROS.", 4, 0, "G"
    GrabarPlan "2.1.01.02", "PROVEEDORES DEL EXTERIOR 3EROS", 4, 0, "G"
    GrabarPlan "2.1.01.03", "COMPA¤IAS VINCULADAS", 4, 1, "G"
    GrabarPlan "2.1.01.03.001", "Biotech S.A. Sucursal LUGANO", 5, 0, "G"
    GrabarPlan "2.1.01.04", "DOCUMENTOS A PAGAR A 3EROS", 4, 0, "G"
    GrabarPlan "2.1.01.05", "ANTICIPOS DE CLIENTES", 4, 0, "G"
    GrabarPlan "2.1.02", "DEUDAS FINANCIERAS", 3, 1, "G"
    GrabarPlan "2.1.02.01", "PRESTAMOS O SOBREGIROS BANCARIOS", 4, 0, "G"
    GrabarPlan "2.1.02.02", "OTROS PRESTAMOS DE 3EROS", 4, 0, "G"
    GrabarPlan "2.1.02.03", "OTROS PRESTAMOS DE COMPA¤IAS VINCULADAS", 4, 0, "G"
    GrabarPlan "2.1.03", "DEUDAS DIVERSAS", 3, 1, "G"
    GrabarPlan "2.1.03.01", "DIVIDENDOS A PAGAR", 4, 0, "G"
    GrabarPlan "2.1.03.02", "RETRIBUCIONES Y CARGAS SOCIALES A PAGAR", 4, 1, "G"
    GrabarPlan "2.1.03.02.001", "Remuneraciones a Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.02.002", "AFPs.", 5, 0, "G"
    GrabarPlan "2.1.03.02.003", "Caja Petrolera", 5, 0, "G"
    GrabarPlan "2.1.03.02.004", "INFOCAL", 5, 0, "G"
    GrabarPlan "2.1.03.03", "ACREDDORES FISCALES", 4, 1, "G"
    GrabarPlan "2.1.03.03.001", "Debito Fiscal", 5, 0, "G"
    GrabarPlan "2.1.03.03.002", "Imptos.S/Transacciones por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.003", "Imptos.RC-IVA a Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.004", "Retenciones RC-IVA por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.005", "Retenciones I.T. por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.006", "Retenciones I.U.E. por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.007", "Impto. S/Util.por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.03.008", "Devolucion iva muestras medicas", 5, 0, "G"
    GrabarPlan "2.1.03.04", "OBLIGACIONES POR LEASING", 4, 0, "G"
    GrabarPlan "2.1.03.05", "OTRAS DEUDAS", 4, 1, "G"
    GrabarPlan "2.1.03.05.001", "Seguros a Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.002", "VISA-TARJETAS x Pag.", 5, 0, "G"
    GrabarPlan "2.1.03.05.003", "Telef.Agua y Luz por Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.004", "Orpan x Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.005", "Honorarios a Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.006", "Otras deudas a Pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.007", "Auditoria por pagar", 5, 0, "G"
    GrabarPlan "2.1.03.05.008", "Clouse up por pagar", 5, 0, "G"
    GrabarPlan "2.1.03.06", "PROVISIONES PARA BENEFICIOS Y CARGAS SOC", 4, 1, "G"
    GrabarPlan "2.1.03.06.001", "ProviSion Aguinaldo", 5, 0, "G"
    GrabarPlan "2.1.03.07", "OTRAS PROVISIONES", 4, 1, "G"
    GrabarPlan "2.1.03.07.001", "Prov.Gastos Fondo Fijo", 5, 0, "G"
    GrabarPlan "2.1.03.07.002", "Provision Gastos de Importacion", 5, 0, "G"
    GrabarPlan "2.1.03.08", "OTRAS PREVISIONES", 4, 0, "G"
    GrabarPlan "2.2", "PASIVO NO CORRIENTE", 2, 1, "G"
    GrabarPlan "2.2.01", "DEUDAS COMERCIALES", 3, 0, "G"
    GrabarPlan "2.2.02", "DEUDAS FINANCIERAS", 3, 0, "G"
    GrabarPlan "2.2.03", "DEUDAS DIVERSAS", 3, 1, "G"
    GrabarPlan "2.2.03.08", "OTRAS PREVISIONES", 4, 1, "G"
    GrabarPlan "2.2.03.08.001", "Prevision Indemnizacion", 5, 0, "G"
    GrabarPlan "2.2.03.08.002", "Prevision Para Futuras Contigencias", 5, 0, "G"
    GrabarPlan "3", "P A T R I M O N I O", 1, 1, "G"
    GrabarPlan "3.1", "PATRIMONIO NETO", 2, 1, "G"
    GrabarPlan "3.1.01", "PATRIMONIO NETO", 3, 1, "G"
    GrabarPlan "3.1.01.01", "CAPITAL SOCIAL", 4, 1, "G"
    GrabarPlan "3.1.01.01.001", "Capital Autorizado", 5, 0, "G"
    GrabarPlan "3.1.01.01.002", "Acciones por Emitir", 5, 0, "G"
    GrabarPlan "3.1.01.02", "APORTES A CAPITALIZAR", 4, 1, "G"
    GrabarPlan "3.1.01.02.001", "Aportes a Capitalizar", 5, 0, "G"
    GrabarPlan "3.1.01.03", "RESERVAS", 4, 1, "G"
    GrabarPlan "3.1.01.03.001", "Reserva Legal y Estatuitaria", 5, 0, "G"
    GrabarPlan "3.1.01.04", "AJUSTES AL PATRIMONIO", 4, 1, "G"
    GrabarPlan "3.1.01.04.001", "Ajuste Global al Patrimonio", 5, 0, "G"
    GrabarPlan "3.1.01.05", "RESULTADO ACUMULADO", 4, 1, "G"
    GrabarPlan "3.1.01.05.001", "Resultado Acum.Gestiones Anteriores", 5, 0, "G"
    GrabarPlan "3.1.01.06", "DISTRIBUCION DE UTILIDADES ANTICIPADAS", 4, 0, "G"
    GrabarPlan "3.1.01.07", "RESULTADO DEL EJERCICIO", 4, 1, "G"
    GrabarPlan "3.1.01.07.001", "Resultado del Ejercicio", 5, 0, "G"
    GrabarPlan "4", "INGRESOS", 1, 1, "PYG"
    GrabarPlan "4.1", "INGRESOS DE EXPLOTACION", 2, 1, "PYG"
    GrabarPlan "4.1.01", "INGRESOS TERCEROS VTAS.NETAS", 3, 1, "PYG"
    GrabarPlan "4.1.01.01", "VENTAS BRUTAS", 4, 1, "PYG"
    GrabarPlan "4.1.01.01.001", "Ventas Comerciales SCZ", 5, 0, "PYG"
    GrabarPlan "4.1.01.01.002", "Ventas Comerciales LPZ", 5, 0, "PYG"
    GrabarPlan "4.1.01.01.003", "Ventas Comerciales CBB", 5, 0, "PYG"
    GrabarPlan "4.1.01.02", "VENTAS BRUTAS LICITACION", 4, 1, "PYG"
    GrabarPlan "4.1.01.02.001", "Ventas Licitacion SCZ", 5, 0, "PYG"
    GrabarPlan "4.1.01.02.002", "Ventas Licitacion LPZ", 5, 0, "PYG"
    GrabarPlan "4.1.01.02.003", "Ventas Licitacion CBBA", 5, 0, "PYG"
    GrabarPlan "4.1.02", "BONIFICACIONES Y DESCUENTOS", 3, 1, "PYG"
    GrabarPlan "4.1.02.01", "BONIFICACIONES", 4, 1, "PYG"
    GrabarPlan "4.1.02.01.001", "Bonificaciones SCZ.", 5, 0, "PYG"
    GrabarPlan "4.1.02.01.002", "Bonificaciones LPZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.01.003", "Bonificaciones SCZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.02", "DESCUENTOS", 4, 1, "PYG"
    GrabarPlan "4.1.02.02.001", "Descuentos SCZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.02.002", "Descuentos LPZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.02.003", "Descuentos CBBA", 5, 0, "PYG"
    GrabarPlan "4.1.02.03", "DESCUENTOS LICITACION", 4, 1, "PYG"
    GrabarPlan "4.1.02.03.001", "Descuentos Licitacion SCZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.03.002", "Descuentos Licitacion LPZ", 5, 0, "PYG"
    GrabarPlan "4.1.02.03.003", "Descuentos Licitacion CBBA", 5, 0, "PYG"
    GrabarPlan "4.1.03", "REINTEGROS Y DEVOLUCIONES", 3, 1, "PYG"
    GrabarPlan "4.1.03.01", "REINTEGROS EXPORTACION", 4, 1, "PYG"
    GrabarPlan "4.1.03.01.001", "Reintegros Exportacion", 5, 0, "PYG"
    GrabarPlan "4.1.03.02", "DEVOLUCIONES", 4, 1, "PYG"
    GrabarPlan "4.1.03.02.001", "Devoluciones", 5, 0, "PYG"
    GrabarPlan "4.1.04", "INGRESOS DE OTRAS SOC-OTROS", 3, 0, "PYG"
    GrabarPlan "4.2", "OTROS INGRESOS", 2, 1, "PYG"
    GrabarPlan "4.2.01", "VARIACION EXISTENCIA PROD.TERM", 3, 1, "PYG"
    GrabarPlan "4.2.01.01", "VARIACION EXISTENCIA PROD.TERM.", 4, 1, "PYG"
    GrabarPlan "4.2.01.01.001", "Variacion Existencia Prod.Terminados", 5, 0, "PYG"
    GrabarPlan "4.2.02", "VARIACION EXIST.PROD.PROCESO", 3, 0, "PYG"
    GrabarPlan "4.2.03", "CONTRAPARTIDA MM IMPORTADAS", 3, 1, "PYG"
    GrabarPlan "4.2.03.01", "CONTRAPARTIDA MM IMPORTADAS", 4, 1, "PYG"
    GrabarPlan "4.2.03.01.001", "Contrapartida MM Importadas", 5, 0, "PYG"
    GrabarPlan "4.2.04", "CONTRAPARTIDA INV.INICIAL MM", 3, 0, "PYG"
    GrabarPlan "4.3", "INGRESOS EXTRAORDINARIOS", 2, 1, "PYG"
    GrabarPlan "4.3.01", "GANANCIA DE CAMBIO", 3, 0, "PYG"
    GrabarPlan "4.3.02", "INTERESES GANADOS", 3, 0, "PYG"
    GrabarPlan "4.3.03", "OTROS INGRESOS EXTRAORDINARIOS", 3, 0, "PYG"
    GrabarPlan "4.3.04", "GANACIA DE GEST. ANT.                   ", 3, 0, "PYG"
    GrabarPlan "5", "GASTOS", 1, 1, "PYG"
    GrabarPlan "5.1", "GASTOS", 2, 1, "PYG"
    GrabarPlan "5.1.01", "MARGEN II", 3, 1, "PYG"
    GrabarPlan "5.1.01.01", "REGALIAS", 4, 0, "PYG"
    GrabarPlan "5.1.01.02", "GASTOS DIRECTOS DE PROMOCION Y VENTA", 4, 1, "PYG"
    GrabarPlan "5.1.01.02.001", "Congresos Medicos", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.002", "Convenciones y Jornadas Medicas", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.003", "Gastos de Capacitacion Consultores", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.004", "Lanzamientos", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.005", "Literaturas", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.006", "Muestra Medica distribuidores", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.007", "Otros Gastos Promocionales", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.008", "Otros Materiales Promocionales", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.009", "Publicidad Medios", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.010", "Cajas y Tickets", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.011", "Gastos de Exportacion", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.012", "Gastos de Licitaciones", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.013", "Otros Gastos directos de ventas", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.014", "Registros de Marca", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.015", "Registro Sanitario", 5, 0, "PYG"
    GrabarPlan "5.1.01.02.016", "Productos Originales Distribuidos", 5, 0, "PYG"
    GrabarPlan "5.1.02", "MARGEN III", 3, 1, "PYG"
    GrabarPlan "5.1.02.01", "GASTOS INDIRECTOS DE PROMOCION Y VENTAS", 4, 1, "PYG"
    GrabarPlan "5.1.02.01.001", "Casa del Medico", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.002", "Congresos Medicos", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.003", "Convenciones y Jornadas Medicas", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.004", "Convenciones y jornadas consultores", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.005", "Gastos de Capacitacion Consultores", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.006", "Gastos de Representacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.007", "Gastos de Viaje al exterior", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.008", "Gastos de Viaje en el Pais", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.009", "Gastos divicion Internacional", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.010", "Lanzamientos", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.011", "Literaturas", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.012", "Servicios Promocionales", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.013", "Otros Gastos Promocionales", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.014", "Otros Materiales Promocionales", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.015", "Publicidad Medios Institucional", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.016", "Gastos de Exportacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.017", "Gastos de Licitaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.018", "Gastos de Registros", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.019", "Listas de Precios", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.020", "Material de Embalaje", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.021", "Multas y Recargos", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.022", "Otros Gastos Indirectos de Ventas", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.023", "TADT-NEORAXIS", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.024", "COSMETICOS", 5, 0, "PYG"
    GrabarPlan "5.1.02.01.025", "Vacunas", 5, 0, "PYG"
    GrabarPlan "5.1.02.02", "GASTOS ESTRUCTURA DE PROMOCION Y VENTAS", 4, 1, "PYG"
    GrabarPlan "5.1.02.02.001", "Agua", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.002", "Alquileres inmuebles", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.003", "Amortizaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.004", "Asociaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.005", "Auditoria de Casa Matriz", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.006", "Cafeteria", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.007", "Couriers", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.008", "Energia Electrica", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.009", "Fotocopias", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.010", "Gastos de Computacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.011", "Gastos de Seleccion de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.012", "Honorarios Profesionales", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.013", "Impuestos, Tasas y Contribuciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.014", "Informacion de Mercado (IMS, Close UP)", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.015", "Material de Computacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.016", "Material de Emalaje", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.017", "Material de Limpieza", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.018", "Otros Alquileres", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.019", "Papeleria y Utiles de Oficina", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.020", "Reaparaciones y Mantenimiento Inmuebles", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.021", "Reaparacion y Mantenimiento Instalacione", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.022", "Reaparacion y Mantenimiento Maquinarias", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.023", "Reparacion y Mantenimiento Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.024", "Reparacion y Mantenimiento Equipos", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.025", "Servicio de Vigilancia", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.026", "Servicios Varios", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.027", "Suscripciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.02.028", "Telefono y Fax", 5, 0, "PYG"
    GrabarPlan "5.1.02.03", "GASTOS EXTERNOS DE PROMOCION Y VENTAS", 4, 1, "PYG"
    GrabarPlan "5.1.02.03.001", "Sueldos", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.002", "Comisiones", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.003", "Premios", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.004", "Gratificaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.005", "Horas extras", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.006", "Decimotercer Salario", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.007", "Decimocuarto Salario", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.008", "Vacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.009", "Aportes Sociales", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.010", "Alimentacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.011", "Uniformes", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.012", "Portafolios", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.013", "Gastos de Radicacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.014", "Seguro Medico", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.015", "Seguro de Trabajo", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.016", "Otros Gastos de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.017", "Gastos de Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.018", "Gastos de Gira", 5, 0, "PYG"
    GrabarPlan "5.1.02.03.019", "Viaticos", 5, 0, "PYG"
    GrabarPlan "5.1.02.04", "GASTOS INTERNOS DE PROMOCION Y VENTAS", 4, 1, "PYG"
    GrabarPlan "5.1.02.04.001", "Alimentacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.002", "Aportes Sociales", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.003", "Decimocuarto Salario", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.004", "Decimotercer Salario", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.005", "Gastos de Capacitacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.006", "Gastos de Gira", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.007", "Gastos de Radicacion", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.008", "Gastos de Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.009", "Gratificaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.010", "Horas Extras", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.011", "Otros Gastos de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.012", "Portafolios", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.013", "Seguro de Trabajo", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.014", "Seguro Medico", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.015", "Sueldos", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.016", "Uniformes", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.017", "Vacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.02.04.018", "Viacticos", 5, 0, "PYG"
    GrabarPlan "5.1.03", "MARGEN DE COMERCIALIZACION", 3, 1, "PYG"
    GrabarPlan "5.1.03.01", "GASTOS SECCION MEDICA", 4, 1, "PYG"
    GrabarPlan "5.1.03.01.001", "Alimentacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.002", "Aportes Sociales", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.003", "Decimocuarto Salario", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.004", "Decimotercer Salario", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.005", "Gastos de Capacitacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.006", "Gastos de Gira", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.007", "Gastos de Radicacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.008", "Gastos de Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.009", "Gratificaciones", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.010", "Honorarios Profesionales", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.012", "Horas Extras", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.013", "Otros Gastos de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.014", "Portafolios", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.015", "Seguro de Trabajo", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.016", "Seguro Medico", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.017", "Sueldos", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.018", "Vacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.03.01.019", "Viaticos", 5, 0, "PYG"
    GrabarPlan "5.1.03.02", "GASTOS SECCION TECNICA", 4, 1, "PYG"
    GrabarPlan "5.1.03.02.001", "Alimentacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.002", "Aportes Sociales", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.003", "Decimocuarto Salario", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.004", "Decimotercer Salario", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.005", "Gastso de Capacitacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.006", "Gastos de Gira", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.007", "Gastos de Radicacion", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.008", "Gastos de Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.009", "Gratificacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.010", "Honorarios Profesionales", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.011", "Horas Extras", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.012", "Otros Gastos de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.013", "Portafolios", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.014", "Seguro de Trabajo", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.015", "Seguro Medico", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.016", "Sueldos", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.017", "Vacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.03.02.018", "Viaticos", 5, 0, "PYG"
    GrabarPlan "5.1.04", "RESULTADOS DE EXPLOTACION", 3, 1, "PYG"
    GrabarPlan "5.1.04.01", "GASTOS DE ADMINISTRACION", 4, 1, "PYG"
    GrabarPlan "5.1.04.01.001", "Agua", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.002", "Alimentacion", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.003", "Alquileres Inmuebles", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.004", "Amortizaciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.005", "Aportes Sociales", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.006", "Asociaciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.007", "Auditoria de Casa Matriz", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.008", "Auditoria Externa", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.009", "Cafeteria", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.010", "Couriers", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.011", "Decimocuarto Salario", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.012", "Decimotercer Salario", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.013", "Energia Electrica", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.014", "Fletes", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.015", "Fotocopias", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.016", "Gastos de Capacitacion", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.017", "Gastos de Computacion", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.018", "Gastos de Gira", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.019", "Gastos de Radicacion", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.020", "Gastos de Seleccion de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.021", "Gastos de Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.022", "Gratificaciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.023", "Honorarios Profesionales", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.024", "Horas Extras", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.025", "Impuestos, Tasas y Constribuciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.026", "Material de Computacion", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.027", "Material de Embalaje", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.028", "Material de Limpieza", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.029", "Otros Alquileres", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.030", "Otros Gastos de Personal", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.031", "Papeleria y Utiles de Oficina", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.032", "Portafolios", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.033", "Reparacion y Mantenimiento Inmuebles", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.034", "Reparacion y Mantenimiento Instalaciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.035", "Reparacion y Mantenimiento Maquinaria", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.036", "Reparacion y Mantenimiento Vehiculos", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.037", "Reparacion y Mantenimiento Equipos", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.038", "Seguro de Trabajo", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.039", "Seguro Medico", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.040", "Servicio de Vigilancia", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.041", "Servicios Varios", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.042", "Sueldos", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.043", "Uniformes", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.044", "Suscripciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.045", "Telefonos y Fax", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.046", "Vacaciones", 5, 0, "PYG"
    GrabarPlan "5.1.04.01.047", "Viaticos", 5, 0, "PYG"
    GrabarPlan "5.1.04.02", "PRESENTACIONES VENCIDAS", 4, 1, "PYG"
    GrabarPlan "5.1.04.02.001", "Presentaciones Vencidas", 5, 0, "PYG"
    GrabarPlan "5.1.04.03", "MERMAS Y CASTIGOS", 4, 1, "PYG"
    GrabarPlan "5.1.04.03.001", "Mermas y Castigos", 5, 0, "PYG"
    GrabarPlan "5.1.04.04", "OTROS INGRESOS", 4, 1, "PYG"
    GrabarPlan "5.1.04.04.001", "Ajuste de Ingresos de Ejercicios Anterio", 5, 0, "PYG"
    GrabarPlan "5.1.04.04.002", "Recuperacion de Seguro", 5, 0, "PYG"
    GrabarPlan "5.1.04.04.003", "Ganancia por Venta de Activo Fijo", 5, 0, "PYG"
    GrabarPlan "5.1.04.04.004", "Otros Ingresos", 5, 0, "PYG"
    GrabarPlan "5.1.04.05", "OTROS EGRESOS", 4, 1, "PYG"
    GrabarPlan "5.1.04.05.001", "Ajuste de Egresos de Ejercicios Anterior", 5, 0, "PYG"
    GrabarPlan "5.1.04.05.002", "Indemnizacion por Despido", 5, 0, "PYG"
    GrabarPlan "5.1.04.05.003", "Otros Egresos", 5, 0, "PYG"
    GrabarPlan "5.1.04.05.004", "Perdida por Deudores Incobrables", 5, 0, "PYG"
    GrabarPlan "5.1.04.05.005", "Perdida por Siniestro", 5, 0, "PYG"
    GrabarPlan "5.1.04.05.006", "Perdida por Venta de Activos Fijos", 5, 0, "PYG"
    GrabarPlan "5.1.05", "RESULTADO ANTES DE IMPUESTOS", 3, 1, "PYG"
    GrabarPlan "5.1.05.01", "INTERESES GANADOS", 4, 1, "PYG"
    GrabarPlan "5.1.05.01.001", "Intereses Ganados", 5, 0, "PYG"
    GrabarPlan "5.1.05.02", "DESCUENTOS OBTENIDOS", 4, 1, "PYG"
    GrabarPlan "5.1.05.02.001", "Descuentos Obtenidos", 5, 0, "PYG"
    GrabarPlan "5.1.05.03", "INTERESES PERDIDOS", 4, 1, "PYG"
    GrabarPlan "5.1.05.03.001", "Intereses Perdidos", 5, 0, "PYG"
    GrabarPlan "5.1.05.04", "DESCUENTOS OTORGADOS", 4, 1, "PYG"
    GrabarPlan "5.1.05.04.001", "Descuentos Otorgados", 5, 0, "PYG"
    GrabarPlan "5.1.05.05", "DIFERENCIA DE CAMBIO", 4, 1, "PYG"
    GrabarPlan "5.1.05.05.001", "Diferencia de Cambio", 5, 0, "PYG"
    GrabarPlan "5.1.05.06", "RESULTADO POR CONVERSION", 4, 1, "PYG"
    GrabarPlan "5.1.05.06.001", "Resultado por Conversion", 5, 0, "PYG"
    GrabarPlan "5.1.06", "RESULTADOS ANTES DE IMPUESTO", 3, 1, "PYG"
    GrabarPlan "5.1.06.01", "IMPUESTO A LA RENTA", 4, 1, "PYG"
    GrabarPlan "5.1.06.01.001", "Impuestos a la Renta", 5, 0, "PYG"
    GrabarPlan "5.2", "CONSUMO", 2, 1, "PYG"
    GrabarPlan "5.2.01", "CONSUMO DE PRODUCTO", 3, 1, "PYG"
    GrabarPlan "5.2.01.01", "CONSUMO DE PRODUCTOS TERMINADOS", 4, 1, "PYG"
    GrabarPlan "5.2.01.01.001", "Consumo Prod.Term.Importados", 5, 0, "PYG"
    GrabarPlan "5.2.01.01.002", "Consumo de M.M. Importadas", 5, 0, "PYG"

    'MsgBox "Somos felices porque está huevada terminó"
    Command1.Enabled = False
End Sub

Private Sub GrabarPlan(codCuenta As String, NomCuenta As String, Nivel As Integer, SubCta As Integer, Interviene As String)
    strSql = " SELECT cta_codigo FROM ctaconta" & _
        " WHERE emp_codigo='" & strEmpresa & "' AND cta_codigo='" & codCuenta & "'"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.EOF = True Then
        strSql = "INSERT INTO ctaconta " & _
                "(cta_codigo, emp_codigo, cta_nombre,cta_descripcion, cta_nivel, cta_interviene, cta_subcta, cta_fechamod, cta_usumod) " & _
                "VALUES ('" & codCuenta & "','" & strEmpresa & "','" & Trim(UCase(NomCuenta)) & "','','" & Nivel & "','" & Interviene & "','" & SubCta & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
        clsSql.Ejecutar strSql, "M"
    End If
End Sub

Private Sub dcmbTipo_Change()

    strSql = " SELECT par_con_codigo, par_con_descripcion, par_con_cta_codigo, cta_nombre, '0' as modi FROM parametro_contable" & _
        " LEFT JOIN ctaconta ON ctaconta.cta_codigo=parametro_contable.par_con_cta_codigo AND ctaconta.emp_codigo=parametro_contable.emp_codigo" & _
        " WHERE parametro_contable.emp_codigo='" & strEmpresa & "' AND par_con_tipo='" & dcmbTipo & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.EOF = False Then
        Hacer = False
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        ucrtVSFG.PonerNum
        dcmbTipo.ListField = "par_con_tipo"
        Hacer = True
    End If
    
    strSql = " SELECT cta_codigo, cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY cta_codigo "
    clsCta.Ejecutar strSql
    
    VSFG.ColComboList(3) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")

End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
   
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsSql1.Inicializar AdoConn, AdoConnMaster
    clsCta.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False
    IniDato

    strSql = " SELECT distinct par_con_tipo FROM parametro_contable" & _
         " WHERE emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.EOF = False Then
        Set dcmbTipo.RowSource = clsSql.adorec_Def.DataSource
        dcmbTipo.ListField = "par_con_tipo"
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Hacer = False Then Exit Sub
    If Row > 0 Then
        With VSFG
            If .TextMatrix(Row, Col) <> "" Then
                If Col = 3 Then
                     clsCta.Filtrar ("cta_codigo = '" & .TextMatrix(Row, 3) & "'")
                        .TextMatrix(Row, 4) = clsCta.adorec_Def("cta_nombre")
                     clsCta.QuitarFiltro
                 End If
             End If
        End With
       
    End If
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
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub
