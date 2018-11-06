VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmCobrosPagos 
   BackColor       =   &H00BAA892&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros"
   ClientHeight    =   8280
   ClientLeft      =   4680
   ClientTop       =   1110
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7770
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3240
      TabIndex        =   28
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame FrmCobrosPagos 
      BackColor       =   &H00BAA892&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   3975
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   7335
      Begin VB.TextBox TxtSaldo 
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox TxtTotalAbonos 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TxtDocumento 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtObservacion 
         Height          =   735
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1440
         Width           =   3495
      End
      Begin VB.ComboBox CmbAno 
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
         ItemData        =   "frmCobrosPagos.frx":0000
         Left            =   1800
         List            =   "frmCobrosPagos.frx":0061
         TabIndex        =   8
         Text            =   "AÑO"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CmbMes 
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
         ItemData        =   "frmCobrosPagos.frx":011F
         Left            =   2760
         List            =   "frmCobrosPagos.frx":014A
         TabIndex        =   9
         Text            =   "MES"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CmbDia 
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
         ItemData        =   "frmCobrosPagos.frx":018A
         Left            =   3600
         List            =   "frmCobrosPagos.frx":0201
         TabIndex        =   10
         Text            =   "DIA"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtMonto 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFGValores 
         Height          =   1215
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   6855
         _cx             =   12091
         _cy             =   2143
         _ConvInfo       =   1
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCobrosPagos.frx":0278
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
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Abonos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label LblFechaCP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de cobro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame FrmTitulo 
      BackColor       =   &H00BAA892&
      Caption         =   "Cuenta por Cobrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   7335
      Begin VB.TextBox TxtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   920
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DCmbCodCuenta 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LblDescripcion 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   855
         Left            =   4200
         TabIndex        =   19
         Top             =   470
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código (Fecha):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAA892&
      Caption         =   "Cliente / Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   7335
      Begin VB.OptionButton OptProveedores 
         BackColor       =   &H00BAA892&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptCliente 
         BackColor       =   &H00BAA892&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DCmbNomPersona 
         Height          =   315
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCmbPersona 
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmCobrosPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsPersona As New clsConsulta
Private clsCuenta As New clsConsulta
Private clsPago As New clsConsulta
Private clsSQL As New clsConsulta
Private SumaAbonos As Double
Private SumaSaldo As Double


Private Sub Command1_Click()
    LLena_Cuenta
End Sub

Private Sub CmdGrabar_Click()
    Dim Maximo As Integer
    Dim Var_Error As Integer
    Dim Estado_Pagado As Integer
    Estado_Pagado = 0
    Maximo = 0
    Var_Error = 0
    FechaCobroPago = CmbAno.Text + "-" + cmbMes.Text + "-" + cmbDia.Text
    FechaCobroPago = Format(FechaCobroPago, "yyyy-mm-dd")
    FechaHoy = Format(Date, "yyyy-mm-dd")
    
    'Controlar que esten llenos campos
    If Not IsDate(FechaCobroPago) Then
        MsgBox "Fecha NO válida", vbCritical
        Var_Error = 1
    End If
    If Not IsNumeric(TxtMonto.Text) Or Val(TxtMonto.Text) <= 0 Then
        MsgBox "Cantidad ingresada en el monto NO válida", vbCritical
        TxtMonto.SetFocus
        Var_Error = 1
    End If
    'Validar que el monto no sea mayor que el total
    Calcula_Total
    If (SumaAbonos + Val(TxtMonto.Text)) > Val(txtValor.Text) Then
        MsgBox "El monto sobrepasa el valor del saldo de la cuenta", vbCritical
        TxtMonto.SetFocus
    ElseIf Var_Error = 0 Then
        'Busca el código máximo de la tabla pago de esa cuenta

        strSql = " Select max(pag_codigo) as MaxPago " & _
                " FROM pago " & _
                " WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_codigo=" & DCmbCodCuenta.BoundText & " AND cue_p_c_tipo='" & Me.Tag & "' "
        clsSQL.Ejecutar (strSql)
        
        If Not clsSQL.adorec_Def.EOF Then
            Maximo = clsSQL.adorec_Def("MaxPago") + 1
        Else
            Maximo = Maximo + 1
        End If
                    
        'Ingreso de datos en la tabla pago
        strSql = " INSERT INTO pago (pag_codigo, emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_fecha, pag_monto, pag_no_doc, pag_observacion, pag_fechamod, pag_usumod) " & _
                " VALUES (" & Maximo & ",'" & strEmpresa & "', '" & DCmbCodCuenta.BoundText & "', '" & Me.Tag & "','" & FechaCobroPago & "', '" & Replace(TxtMonto.Text, ",", ".") & "', '" & txtdocumento.Text & "', '" & TxtObservacion.Text & "', CURRENT_TIMESTAMP, substring_index(USER(),'@',1))"
        clsSQL.Ejecutar (strSql)
                
        'Genera los asientos de los cobros y pagos
        strSql = " Select max(SUBSTRING(asi_numasiento,1,11)) as numAs " & _
                " From asiento " & _
                " WHERE emp_codigo='" & strEmpresa & "' "
        clsSQL.Ejecutar strSql
        If Not IsNull(clsSQL.adorec_Def("numas")) Then
            Maximo = clsSQL.adorec_Def("numAS") + 1
            strmaximo = Space(11 - Len(str(Maximo))) & Trim(str(Maximo))
        Else
            Maximo = 1
            strmaximo = Space(11 - Len(str(Maximo))) & Trim(str(Maximo))
        End If

        strSql = " INSERT INTO asiento (asi_numasiento, emp_codigo, asi_fecha, asi_revisado, asi_mayorizado, asi_totaldebe, asi_totalhaber, asi_descripcion, asi_fechamod, asi_usumod) " & _
                 " VALUES ('" & strmaximo & "','" & strEmpresa & "', '" & ff & "', '0','0', '" & Replace(txtTotalDebe, ",", ".") & "', '" & Replace(txtTotalHaber, ",", ".") & "', '" & txtDescripcion.Text & "', CURRENT_TIMESTAMP, substring_index(USER(),'@',1))"
        clsSQL.Ejecutar strSql
        

        strSql = " INSERT INTO det_asiento " & _
                 " SELECT emp_codigo,'" & strmaximo & "',cta_codigo,det_com_egr_debe,det_com_egr_haber,CURRENT_TIMESTAMP, substring_index(USER(),'@',1) " & _
                 " FROM det_comp_egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND com_egr_codigo='" & txtCodigo & "'"
        clsSQL.Ejecutar strSql

        
        TxtMonto.Text = ""
        txtdocumento.Text = ""
        TxtObservacion.Text = ""
        Llena_Fecha
        Llena_CobrosPagos
        
        'Update de la tabla cuenta_p_c y cambia estado a pagado
         If SumaSaldo = 0 Then
            Estado_Pagado = 1
         End If
         strSql = " UPDATE cuenta_p_c SET cue_p_c_pagado=" & Estado_Pagado & ", cue_p_c_fechapago='" & FechaCobroPago & "', cue_p_c_fechamod=CURRENT_TIMESTAMP, cue_p_c_usumod=substring_index(USER(),'@',1) " & _
                " WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_codigo=" & DCmbCodCuenta.BoundText & " AND cue_p_c_tipo='" & Me.Tag & "'"
         clsSQL.Ejecutar (strSql)
         
        MsgBox " Los datos han sido ingresados", vbInformation, "Ingresos"
        
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub DcmbCodCuenta_Change()
    If Not DCmbCodCuenta.Text = "" Then
        Llena_Descripcion
        Llena_CobrosPagos
    End If
    Llena_CobrosPagos
End Sub

Private Sub DCmbNomPersona_Change()
    If DCmbPersona.Tag <> "A" Then
        If DCmbNomPersona.MatchedWithList = True Then
            DCmbPersona.Text = DCmbNomPersona.BoundText
        End If
    End If
End Sub
Private Sub DCmbNomPersona_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    DCmbPersona.Text = DCmbNomPersona.BoundText
End Sub

Private Sub DCmbNomPersona_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        DCmbPersona.Text = DCmbNomPersona.BoundText
    End If
End Sub

Private Sub DCmbPersona_Change()
        If clsPersona.adorec_Def.RecordCount > 0 Then
            clsPersona.adorec_Def.MoveFirst
        End If
        clsPersona.adorec_Def.Find "per_codigo = '" & DCmbPersona & "'", , adSearchForward
        DCmbPersona.Tag = "A"
        If clsPersona.adorec_Def.EOF = True Then
            DCmbNomPersona = ""
            DCmbNomPersona.BoundText = ""
        Else
            DCmbNomPersona.Text = clsPersona.adorec_Def("nomb")
            DCmbNomPersona.BoundText = DCmbPersona.Text
        End If
        DCmbPersona.Tag = ""
        LLena_Cuenta
End Sub

Private Sub Form_Activate()
    If Me.Tag = "C" Then
        LblTitulo.Caption = "Cobros"
        FrmTitulo.Caption = "Cuenta por Cobrar"
        FrmCobrosPagos.Caption = "Cobros"
        LblFechaCP.Caption = "Fecha de Cobro"
        Me.Caption = "Cobros"
    ElseIf Me.Tag = "P" Then
        LblTitulo.Caption = "Pagos"
        FrmTitulo.Caption = "Cuenta por Pagar"
        FrmCobrosPagos.Caption = "Pagos"
        LblFechaCP.Caption = "Fecha de Pago"
        Me.Caption = "Pagos"
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    clsPersona.Inicializar AdoConn
    clsCuenta.Inicializar AdoConn
    clsPago.Inicializar AdoConn
    clsSQL.Inicializar AdoConn
    Llena_Fecha
    Llena_Cliente
    
End Sub
Private Sub Llena_Fecha()
    'Llena combo de fecha
    d = CStr(Day(Date))
    m = Month(Date)
    y = CStr(Year(Date))
    cmbDia.Text = d
    CmbAno.Text = y
    For var = 1 To 12
        If cmbMes.ItemData(var) = m Then
            cmbMes.Text = cmbMes.List(var)
            Exit For
        End If
    Next var
End Sub
Private Sub Llena_Cliente()
        
        'Llena Combo de clientes
        DCmbPersona.Text = ""
        DCmbNomPersona.Text = ""
        lblDescripcion.Caption = ""
        txtValor.Text = ""
        
        strSql = " SELECT per_codigo, CONCAT(per_nombre,' ',per_apellido) as nomb " & _
             " From persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='C'" & _
             " ORDER BY per_codigo"
        clsPersona.Ejecutar (strSql)
        
        Set DCmbPersona.RowSource = clsPersona.adorec_Def.DataSource
        DCmbPersona.ListField = "per_codigo"
        
        Set DCmbNomPersona.RowSource = clsPersona.adorec_Def.DataSource
        DCmbNomPersona.ListField = "nomb"
        DCmbNomPersona.BoundColumn = "per_codigo"
End Sub

Private Sub Llena_Proveedor()
        
        'Llena Combo de proveedores
        DCmbPersona.Text = ""
        DCmbNomPersona.Text = ""
        lblDescripcion.Caption = ""
        txtValor.Text = ""
        
        strSql = " SELECT per_codigo, CONCAT(per_nombre,' ',per_apellido) as nomb " & _
             " From persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='P'" & _
             " ORDER BY per_codigo"
        clsPersona.Ejecutar (strSql)
        
        Set DCmbPersona.RowSource = clsPersona.adorec_Def.DataSource
        DCmbPersona.ListField = "per_codigo"
        
        Set DCmbNomPersona.RowSource = clsPersona.adorec_Def.DataSource
        DCmbNomPersona.ListField = "nomb"
        DCmbNomPersona.BoundColumn = "per_codigo"

End Sub

Private Sub LLena_Cuenta()
        
        'LLena Combo de Codigo de Cuenta
        DCmbCodCuenta.Text = ""
        txtValor.Text = ""
        lblDescripcion.Caption = ""
        SumaAbonos = 0
        SumaSaldo = 0
        TxtTotalAbonos.Text = Format(SumaAbonos, "##0.00")
        TxtSaldo.Text = Format(SumaSaldo, "##0.00")
        
            strSql = " SELECT cue_p_c_codigo,cue_p_c_descripcion,cue_p_c_valor, CONCAT(cue_p_c_codigo,' (',SUBSTRING(cue_p_c_fechaemision,1,10),')') as cuen " & _
             " From cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND per_codigo='" & DCmbPersona.Text & "' AND cue_p_c_tipo='" & Me.Tag & "' AND cue_p_c_pagado=0" & _
             " ORDER BY cue_p_c_codigo"
        
        clsCuenta.Ejecutar (strSql)
        Set DCmbCodCuenta.RowSource = clsCuenta.adorec_Def.DataSource
        DCmbCodCuenta.ListField = "cuen"
        DCmbCodCuenta.BoundColumn = "cue_p_c_codigo"
        
End Sub

Private Sub Llena_Descripcion()
    If clsCuenta.adorec_Def.RecordCount > 0 Then
            clsCuenta.adorec_Def.MoveFirst
        End If
        clsCuenta.adorec_Def.Find "cue_p_c_codigo = '" & DCmbCodCuenta.BoundText & "'", , adSearchForward
        If clsCuenta.adorec_Def.EOF = True Then
            lblDescripcion.Caption = ""
            txtValor.Text = ""
        Else
            lblDescripcion.Caption = clsCuenta.adorec_Def("cue_p_c_descripcion")
            txtValor.Text = clsCuenta.adorec_Def("cue_p_c_valor")
        End If
End Sub
Private Sub Llena_CobrosPagos()
    'LLena grid
        Dim Codigo_Cuenta As Integer
        Codigo_Cuenta = 0
        If Not DCmbCodCuenta.BoundText = "" Then
            Codigo_Cuenta = Val(DCmbCodCuenta.BoundText)
        End If
            strSql = " SELECT pag_codigo,pag_fecha,pag_monto,pag_no_doc,pag_observacion" & _
             " From pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_tipo='" & Me.Tag & "' AND cue_p_c_codigo=" & Codigo_Cuenta & " " & _
             " ORDER BY pag_codigo"
        clsPago.Ejecutar (strSql)
        clsPago.Actualizar
        'MsgBox clsPago.adorec_Def.RecordCount
        Set VSFGValores.DataSource = clsPago.adorec_Def.DataSource
        VSFGValores.Refresh
        Calcula_Total
        TxtTotalAbonos.Text = Format(SumaAbonos, "##0.00")
        TxtSaldo.Text = Format(SumaSaldo, "##0.00")
End Sub
Private Sub Calcula_Total()
        'Calcula totales
    SumaAbonos = 0
    SumaSaldo = 0
    'Calcula total abonos
    For i = 1 To VSFGValores.Rows - 1
        SumaAbonos = SumaAbonos + Val(VSFGValores.TextMatrix(i, 3))
    Next i
    SumaSaldo = Val(txtValor.Text) - SumaAbonos
End Sub

Private Sub OptCliente_Click()
    Llena_Cliente
End Sub

Private Sub Optproveedores_Click()
    Llena_Proveedor
End Sub

Private Sub TxtMonto_LostFocus()
    If Not IsNumeric(TxtMonto.Text) Then
        MsgBox "Cantidad ingresada en monto NO válida", vbCritical
    End If
End Sub
'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

