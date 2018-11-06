VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_cargaCuentaxPagar 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas x Pagar"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin MSDataListLib.DataCombo dtcm_CuentaContable 
         Height          =   315
         Left            =   8280
         TabIndex        =   15
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6495
         TabIndex        =   11
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   3720
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos de la cuenta"
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
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   12495
         Begin VB.TextBox TxtObservacion 
            Height          =   735
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   3480
            Width           =   10935
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFG 
            Height          =   2895
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   12015
            _cx             =   21193
            _cy             =   5106
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   1900
            ExtendLastCol   =   0   'False
            FormatString    =   $"frm_cargaCuentaxPagar.frx":0000
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observación:"
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
            Left            =   135
            TabIndex        =   3
            Top             =   3480
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo dcmbTipoDoc 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   4680
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo dcmbSustento 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Sustento:"
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
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Gasto:"
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
         Index           =   1
         Left            =   6480
         TabIndex        =   12
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Doc."
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
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm_cargaCuentaxPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsPersona As New clsConsulta
Private clssql As New clsConsulta
Private clsAsi As New clsConsulta
Private clsparametro As New clsConsulta
Dim strSql As String
Private Var_NumCuenta As Long
Private clsCon_Def As New clsConsulta
Dim i_flag As Integer
Dim fechaEmision As String
Dim i_cont As Integer
 
 Dim cuentacontagasto As String
 Dim cuentacontaiva As String
 Dim cuentacontaventa As String
 Dim cuentacontair As String
 Dim cuentacontaretiva As String
 
 Dim d_imprenta As Double
 Dim d_retiva As Double
 Dim d_comisiones As Double
 Dim d_iva As Double
 Dim d_clienteprov As Double
 
 Dim CentroCosto As String
 

Private Sub cmdAplicar_Click()
 Dim clsAsiento As New clsContable
 Dim TotalCuenta As Double
 Dim ban As Boolean
 Dim Total_debe As Double
 Dim CuentaXPagar As String
 Dim i_flag As Integer
 Dim i_porce_iva As Integer
 Dim i_porce_renta As Integer
 
 i_flag = 0
 
 If (txtArchivo.Text <> "" And dcmbSustento.Text <> "" And dcmbTipoDoc.Text <> "" And dtcm_CuentaContable.Text <> "") Then
    i_flag = 1
 End If
 
 TotalCuenta = 0
 ban = True
      
  If (ban = True And i_flag = 1) Then
    For i = 1 To VSFG.Rows - 1
     If (VSFG.TextMatrix(i, 0) <> "") Then
       If (VSFG.TextMatrix(i, 17) = 1) Then
        'Variables del grid
        Total_debe = VSFG.TextMatrix(i, 7)
        
        'Busca el código máximo de la tabla asiento
        clsAsiento.Inicializar AdoConn, AdoConnMaster
        clsAsiento.NuevoAsiento "D", VSFG.TextMatrix(i, 8), 0, 0, FormatoD2(Total_debe), _
        "Persona: " & VSFG.TextMatrix(i, 15) & vbNewLine & _
        dcmbTipoDoc.Text & ": " & VSFG.TextMatrix(i, 1) & " " & VSFG.TextMatrix(i, 2) & " Aut: " & VSFG.TextMatrix(i, 3) & vbNewLine & _
        VSFG.TextMatrix(i, 14)
                
        strMaximo = clsAsiento.NumAsiento
     
        'Ingreso de datos en cuenta_p_c
        Set clsIngCuentas = New clsConsulta
        clsIngCuentas.Inicializar AdoConn, AdoConnMaster
        Dim clsCxX As New clsCtaXx
        clsCxX.Inicializar AdoConn, AdoConnMaster
        clsCxX.NuevaCta "P", dcmbTipoDoc.BoundText, dcmbSustento.BoundText, VSFG.TextMatrix(i, 8), VSFG.TextMatrix(i, 8), VSFG.TextMatrix(i, 16), VSFG.TextMatrix(i, 14), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 2), VSFG.TextMatrix(i, 3), Format(VSFG.TextMatrix(i, 4), "mm/yyyy"), 0, VSFG.TextMatrix(i, 5), 0, VSFG.TextMatrix(i, 5), 2, Val(VSFG.TextMatrix(i, 6)), 0, 0, 0, 0, VSFG.TextMatrix(i, 7), clsAsiento.NumAsiento
        CuentaXPagar = clsCxX.strNoCta
     
        'Cuentas y valores a insertar
        cuentacontagasto = dtcm_CuentaContable.BoundText
        d_imprenta = FormatoD2(VSFG.TextMatrix(i, 5)) * FormatoD2(VSFG.TextMatrix(i, 21)) / 100
        d_iva = FormatoD2(VSFG.TextMatrix(i, 6)) * FormatoD2(VSFG.TextMatrix(i, 23)) / 100
        d_clienteprov = (FormatoD2(VSFG.TextMatrix(i, 7)) - d_iva - d_imprenta)
        d_comisiones = (FormatoD2(VSFG.TextMatrix(i, 7)) - FormatoD2(VSFG.TextMatrix(i, 6)))
        
        'Inserta el detalle de los asientos
        clsAsiento.NuevoDetAsiento VSFG.TextMatrix(i, 19), "", FormatoD2(0), FormatoD2(d_clienteprov)
        If (VSFG.TextMatrix(i, 6) <> "") Then
            clsAsiento.NuevoDetAsiento "1.01.08.01.001", "", FormatoD2(VSFG.TextMatrix(i, 6)), FormatoD2(0)
        End If
        clsAsiento.NuevoDetAsiento cuentacontagasto, "", FormatoD2(d_comisiones), FormatoD2(0)
        clsAsiento.NuevoDetAsiento VSFG.TextMatrix(i, 20), "", FormatoD2(0), FormatoD2(d_imprenta)
        If (VSFG.TextMatrix(i, 6) <> "") Then
            clsAsiento.NuevoDetAsiento VSFG.TextMatrix(i, 22), "", FormatoD2(0), FormatoD2(d_iva)
        End If
        
        Dim Tipo As String
        Dim TotalRet As Double
        Dim IR As Integer
        Dim RETIVA As Integer
        TotalRet = d_iva + d_imprenta
        Tipo = "P"
        'Insertar valores del Comprobante de retención
        strSql = " INSERT INTO comprobante_retencion (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,com_ret_fecha,com_ret_serie,com_ret_numero,com_ret_autorizacion,com_ret_total,com_ret_fechamod,com_ret_usumod) " & _
        " VALUES ('" & strEmpresa & "', '" & CuentaXPagar & "','" & Tipo & "','" & Format((VSFG.TextMatrix(i, 8)), "yyyy-mm-dd") & "','" & VSFG.TextMatrix(i, 9) & "','" & VSFG.TextMatrix(i, 10) & "','" & VSFG.TextMatrix(i, 11) & "', " & _
        TotalRet & ", CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clssql.Ejecutar strSql, "M"
    
        'Insetar valores del Detalle de la retención
    
        'Primer registro del detalle de la retención
        If (VSFG.TextMatrix(i, 13) <> "") Then
            strSql = " INSERT INTO det_comp_ret (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_com_ret_fechamod,det_com_ret_usumod) " & _
            " VALUES ('" & strEmpresa & "','" & CuentaXPagar & "','" & Tipo & "', '" & VSFG.TextMatrix(i, 13) & "', " & _
            FormatoD2(VSFG.TextMatrix(i, 6)) & "," & FormatoD2(VSFG.TextMatrix(i, 23)) & ", CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clssql.Ejecutar strSql, "M"
        End If
        'Segundo registro del detalle de la retención
        If (VSFG.TextMatrix(i, 12) <> "") Then
            strSql = " INSERT INTO det_comp_ret (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_com_ret_fechamod,det_com_ret_usumod) " & _
            " VALUES ('" & strEmpresa & "','" & CuentaXPagar & "','" & Tipo & "', '" & VSFG.TextMatrix(i, 12) & "', " & _
            FormatoD2(VSFG.TextMatrix(i, 5)) & "," & FormatoD2(VSFG.TextMatrix(i, 21)) & ", CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clssql.Ejecutar strSql, "M"
        End If
      End If
     End If
    Next i
        MsgBox ("Los datos han sido registrados correctamente")
        txtArchivo.Text = ""
        dtcm_CuentaContable.Text = ""
        dcmbSustento.Text = ""
        dcmbTipoDoc.Text = ""
        VSFG.Clear 1
 Else
        MsgBox ("Faltan datos en el formulario.")
 End If
 
 
End Sub

Private Sub cmdExplorar_Click()
   
    Dim sDir As String
    Dim i As Long
    Dim i_len As Integer
    Dim i_leng As Integer
    Dim i_cont As Integer
    sDir = CurDir

    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
        
        'Lee archivo para cargar las cuentas por pagar
        VSFG.Clear 1
        VSFG.LoadGrid txtArchivo.Text, flexFileTabText
        
        VSFG.TextMatrix(0, 0) = "RUC"
        VSFG.TextMatrix(0, 1) = "Serie"
        VSFG.TextMatrix(0, 2) = "Número Factura"
        VSFG.TextMatrix(0, 3) = "Autorización"
        VSFG.TextMatrix(0, 4) = "Fecha de Caducidad"
        VSFG.TextMatrix(0, 5) = "Subtotal"
        VSFG.TextMatrix(0, 6) = "IVA"
        VSFG.TextMatrix(0, 7) = "Total"
        VSFG.TextMatrix(0, 8) = "Fecha Factura"
        VSFG.TextMatrix(0, 9) = "Serie Retención"
        VSFG.TextMatrix(0, 10) = "Número Retención"
        VSFG.TextMatrix(0, 11) = "Autorización Ret."
        VSFG.TextMatrix(0, 12) = "Ret. Renta"
        VSFG.TextMatrix(0, 13) = "% Ret. IVA"
        VSFG.TextMatrix(0, 14) = "Concepto"
        VSFG.TextMatrix(0, 15) = "Proveedor"
        VSFG.TextMatrix(0, 16) = "Codigo Proveedor"
        VSFG.TextMatrix(0, 17) = "Estado"
        VSFG.TextMatrix(0, 18) = "Error Descripcion"
        VSFG.TextMatrix(0, 19) = "Cuenta Contable Proveedor"
        VSFG.TextMatrix(0, 20) = "Cuenta Retencion"
        VSFG.TextMatrix(0, 21) = "%Ret.Renta"
        VSFG.TextMatrix(0, 22) = "Cuenta IVA Retencion"
        VSFG.TextMatrix(0, 23) = "%Ret.IVA"
        
        For i = 1 To VSFG.Rows - 1
         If (VSFG.TextMatrix(i, 0) <> "") Then
           i_cont = i_cont + 1
            strSql = " SELECT per_codigo, per_ruc, CONCAT(per_apellido,per_nombre) as Proveedor" & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND per_ruc='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 15) = clsCon_Def.adorec_Def("Proveedor")
                VSFG.TextMatrix(i, 16) = clsCon_Def.adorec_Def("per_codigo")
                VSFG.TextMatrix(i, 17) = 1
            Else
                VSFG.Cell(flexcpBackColor, i, 0, i, 0) = &HC0C0FF
                VSFG.TextMatrix(i, 17) = 0
            End If
            
            'Carga la cuenta contable del RUC
            strSql = " SELECT c.cat_p_ctaconta, p.per_ruc" & _
                     " FROM categoria_p  c inner join persona p on c.cat_p_codigo = p.cat_p_codigo " & _
                     " WHERE p.emp_codigo='" & strEmpresa & "'" & _
                     " AND p.per_ruc='" & VSFG.TextMatrix(i, 0) & "' "
            clsCon_Def.Ejecutar strSql
            
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 19) = clsCon_Def.adorec_Def("cat_p_ctaconta")
            End If
            
            'Carga la cuenta de la retencion
            If (VSFG.TextMatrix(i, 12) <> "") Then
            strSql = " SELECT ret_codigo, ret_ctaconta, ret_porcentaje" & _
                     " FROM retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND ret_codigo='" & VSFG.TextMatrix(i, 12) & "' "
            clsCon_Def.Ejecutar strSql
            
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 20) = clsCon_Def.adorec_Def("ret_ctaconta")
                VSFG.TextMatrix(i, 21) = clsCon_Def.adorec_Def("ret_porcentaje")
            End If
            End If
            
            If (VSFG.TextMatrix(i, 13) <> "") Then
            strSql = " SELECT ret_codigo, ret_ctaconta, ret_porcentaje" & _
                     " FROM retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND ret_codigo='" & VSFG.TextMatrix(i, 13) & "' "
            clsCon_Def.Ejecutar strSql
            
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(i, 22) = clsCon_Def.adorec_Def("ret_ctaconta")
                VSFG.TextMatrix(i, 23) = clsCon_Def.adorec_Def("ret_porcentaje")
            End If
            End If
            
           'Completar el numero de la serie de la factura
            If (VSFG.TextMatrix(i, 1) <> "") Then
                i_len = (Len(VSFG.TextMatrix(i, 1))) - 6
                If (i_len > 0 Or i_len < 0) Then
                        If (i_len < 0) Then
                            For j = 1 To i_len * -1
                              VSFG.TextMatrix(i, 1) = "0" & VSFG.TextMatrix(i, 1)
                            Next j
                        Else
                            VSFG.Cell(flexcpBackColor, i, 1, i, 1) = &HC0C0FF
                        End If
                End If
            Else
                VSFG.Cell(flexcpBackColor, i, 1, i, 1) = &HC0C0FF
            End If
            
           'Completar el numero de la serie de la retencion
            If (VSFG.TextMatrix(i, 9) <> "") Then
                i_leng = (Len(VSFG.TextMatrix(i, 9))) - 6
                If (i_leng > 0 Or i_leng < 0) Then
                        If (i_leng < 0) Then
                            For j = 1 To i_leng * -1
                              VSFG.TextMatrix(i, 9) = "0" & VSFG.TextMatrix(i, 9)
                            Next j
                        Else
                            VSFG.Cell(flexcpBackColor, i, 9, i, 9) = &HC0C0FF
                        End If
                End If
            Else
                VSFG.Cell(flexcpBackColor, i, 9, i, 9) = &HC0C0FF
            End If
            
            
        End If
        Next i
        Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    
    Set clsCon_Def = New clsConsulta
    
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    clsAsi.Inicializar AdoConn, AdoConnMaster
    
    clsparametro.Inicializar AdoConn, AdoConnMaster
    
    clssql.Inicializar AdoConn, AdoConnMaster
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0

    'Llena combos de fecha
'    dtpFecha.value = Format(HoyDia, "yyyy-mm-dd")
'    dtpCaduca = Format(HoyDia, "mm/yyyy")
'
    'Consulta para saber los tipos de documentos
    strSql = " SELECT cod_sus_com_codigo, cod_sus_com_nombre " & _
             " FROM codigo_sustento_comprobante " & _
             " ORDER BY cod_sus_com_codigo "
    clsAsi.Ejecutar strSql

    Set dcmbSustento.RowSource = clsAsi.adorec_Def.DataSource
    dcmbSustento.ListField = "cod_sus_com_nombre"
    dcmbSustento.BoundColumn = "cod_sus_com_codigo"
    
    
    'Consulta para saber los tipos de documentos
    strSql = " SELECT tip_doc_cue_codigo, tip_doc_cue_descripcion " & _
             " FROM tipo_doc_cuenta "
    clsAsi.Ejecutar strSql

    Set dcmbTipoDoc.RowSource = clsAsi.adorec_Def.DataSource
    dcmbTipoDoc.ListField = "tip_doc_cue_descripcion"
    dcmbTipoDoc.BoundColumn = "tip_doc_cue_codigo"
    
    'Consulta para llenar el combo de las cuentas contables
     strSql = " SELECT cta_codigo, CONCAT(cta_codigo,'     ',cta_nombre) as Descripcion" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsAsi.Ejecutar strSql
     Set dtcm_CuentaContable.RowSource = clsAsi.adorec_Def.DataSource
     dtcm_CuentaContable.ListField = "Descripcion"
     dtcm_CuentaContable.BoundColumn = "cta_codigo"

End Sub
