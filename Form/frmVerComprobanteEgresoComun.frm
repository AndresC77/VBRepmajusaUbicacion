VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerComprobanteEgresoComun 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Comprobante de Egresos"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmVerComprobanteEgresoComun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9270
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "Imprimir Cheque"
      Height          =   375
      Left            =   5524
      TabIndex        =   37
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   128
      TabIndex        =   19
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   4680
         Width           =   7335
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   7320
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   7320
         Width           =   1935
      End
      Begin VB.TextBox txtFechac 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFechach 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.Frame fmeCliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente/Proveedor"
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
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   8775
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtTelefono 
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtDireccion 
            Height          =   765
            Left            =   6240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtRuc 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtBeneficiario 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtNombre 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
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
            Left            =   5040
            TabIndex        =   31
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblTelefono 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono/fax:"
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
            Left            =   5040
            TabIndex        =   30
            Top             =   397
            Width           =   960
         End
         Begin VB.Label lbldireccion 
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
            Left            =   5040
            TabIndex        =   29
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label lblruc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruc:"
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
            Left            =   360
            TabIndex        =   28
            Top             =   1080
            Width           =   330
         End
         Begin VB.Label lblnombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
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
            Left            =   360
            TabIndex        =   27
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblBeneficiario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beneficiario:"
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
            Left            =   360
            TabIndex        =   26
            Top             =   397
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   8775
         Begin VB.TextBox txtValor 
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
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCheque 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtBanco 
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor del cheque:"
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
            Left            =   4680
            TabIndex        =   24
            Top             =   637
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria:"
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
            Left            =   4680
            TabIndex        =   23
            Top             =   277
            Width           =   1245
         End
         Begin VB.Label lblfecha 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cheque:"
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
            Top             =   637
            Width           =   885
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
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
            TabIndex        =   21
            Top             =   277
            Width           =   510
         End
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Width           =   8760
         _cx             =   15452
         _cy             =   2778
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
         FormatString    =   $"frmVerComprobanteEgresoComun.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label lbltotal 
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
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   7335
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   1560
         Picture         =   "frmVerComprobanteEgresoComun.frx":03E4
         ToolTipText     =   "Elimina una Fila"
         Top             =   7320
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   1800
         Picture         =   "frmVerComprobanteEgresoComun.frx":051A
         Top             =   7320
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del comprobante:"
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
         Height          =   285
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Cheque:"
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
         Height          =   285
         Left            =   4920
         TabIndex        =   33
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   4320
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   2172
      TabIndex        =   17
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3852
      TabIndex        =   18
      Top             =   7920
      Width           =   1575
   End
End
Attribute VB_Name = "frmVerComprobanteEgresoComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso del comprobante de egresos comunes                         #
'#  frmComprobanteEgresoComun V1.0                                              #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el comprobante de egresos comunes                     #
'#  Permite ingresar los datos de egresos comunes y sus detalles                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  COMP_EGRESO: Esta tabla almacena los datos del comprobante                  #
'#  PERSONA: donde se guardan los datos de los benficiarios de los comprobantes #
'#  DET_COMP_EGRESO: Guarda los detalles del comprobante de Egreso              #
'#  RET_COMP_EGRESO: Guarda las retenciones que puede tener el comprobante      #
'#  CTA_BANCO: consulta los datos del numero de cuenta y el último cheque       #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'


Private clsCom As New clsConsulta
Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsEgr As New clsConsulta
Private strSQL As String
Private intDato As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCom = Nothing
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsEgr = Nothing
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub limpia()
    On Error Resume Next
        txtFechac = ""
        txtFechach = ""
        txtBeneficiario = ""
        txtNombre = ""
        txtTelefono = ""
        txtEmail = ""
        txtDireccion = ""
        txtRuc = ""
        txtBanco = ""
        txtCheque = ""
        txtCuenta = ""
        txtValor = ""
        txtDescripcion = ""
        VSFG.Clear 1
        VSFG.Row = 2
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdReimprimir_Click()
    Dim CompEgr As New frmReporte
    CompEgr.strReporte = "rptComprobanteEgreso"
    CompEgr.strNumero = dcmbCodigo
    CompEgr.Show
    Dim Cheque As New frmReporte
    Cheque.strReporte = "rptCheque"
    Cheque.strNumero = dcmbCodigo
    Cheque.Show
End Sub

Private Sub dcmbCodigo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub CalcuTotal()
   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    
    'Calcula total debe
    
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    
    'Calcula total haber
    
    For i = 1 To VSFG.Rows - 1
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")
    
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdNuevo_Click()
    frmComprobanteEgresoComun.Show
End Sub

Private Sub dcmbCodigo_Change()
If dcmbCodigo = "" Then
    limpia
Else
  'Realiza la consulta para poner valores en la forma

    strSQL = " SELECT comp_egreso.cta_ban_numero, comp_egreso.ban_codigo, banco.ban_nombre, comp_egreso.per_codigo, CONCAT(per_apellido,' ', per_nombre) as per_nombre, CONCAT(per_telf,' / ',per_fax) as per_telf,per_email,per_direccion,per_ruc, " & _
             " com_egr_fecha, com_egr_descripcion, com_egr_ch_num, com_egr_ch_fecha, com_egr_ch_estado, com_egr_ch_valor" & _
             " FROM (((comp_egreso INNER JOIN cta_banco ON comp_egreso.cta_ban_numero=cta_banco.cta_ban_numero " & _
             "                                          AND comp_egreso.emp_codigo=cta_banco.emp_codigo " & _
             "                                          AND comp_egreso.ban_codigo=cta_banco.ban_codigo)" & _
             "                     INNER JOIN banco ON cta_banco.ban_codigo=banco.ban_codigo)" & _
             "                     LEFT JOIN persona ON comp_egreso.per_codigo= persona.per_codigo " & _
             "                                        AND comp_egreso.emp_codigo= persona.emp_codigo) " & _
             " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' " & _
             " AND com_egr_codigo = '" & dcmbCodigo & "' "
    clsCom.Ejecutar strSQL
 
    If clsCom.adorec_Def.EOF = False Then
        txtFechac = clsCom.adorec_Def("com_egr_fecha")
        txtFechach = clsCom.adorec_Def("com_egr_ch_fecha")
        txtBeneficiario = clsCom.adorec_Def("per_codigo")
        txtNombre = clsCom.adorec_Def("per_nombre")
        txtTelefono = clsCom.adorec_Def("per_telf")
    '    txtEmail = clsCom.adorec_Def("per_email")
     '   txtDireccion = clsCom.adorec_Def("per_direccion")
       '  txtRuc = clsCom.adorec_Def("per_ruc")
      '  txtBanco = clsCom.adorec_Def("ban_nombre")
      '  txtCheque = clsCom.adorec_Def("com_egr_ch_num")
      '  txtCuenta = clsCom.adorec_Def("cta_ban_numero")
       ' txtValor = clsCom.adorec_Def("com_egr_ch_valor")
      ' txtDescripcion = clsCom.adorec_Def("com_egr_descripcion")
    
    strSQL = " SELECT distinct det_asiento.cta_codigo,ctaconta.cta_nombre ,det_asi_debe, det_asi_haber,cen_cos_nombre " & _
             " FROM ((( comp_egreso INNER JOIN det_asiento ON comp_egreso.asi_numasiento=det_asiento.asi_numasiento " & _
             "                                                 AND comp_egreso.emp_codigo=det_asiento.emp_codigo) " & _
             "                      INNER JOIN ctaconta ON det_asiento.cta_codigo= ctaconta.cta_codigo " & _
             "                                          AND det_asiento.emp_codigo= ctaconta.emp_codigo)" & _
             "                      INNER JOIN cta_banco ON comp_egreso.cta_ban_numero=cta_banco.cta_ban_numero " & _
             "                                           AND comp_egreso.ban_codigo=cta_banco.ban_codigo " & _
             "                                           AND comp_egreso.emp_codigo=cta_banco.emp_codigo) " & _
             "                      LEFT JOIN centro_costo ON det_asiento.cen_cos_codigo= centro_costo.cen_cos_codigo " & _
             "                                          AND det_asiento.emp_codigo= centro_costo.emp_codigo " & _
             " WHERE comp_egreso.com_egr_codigo = '" & dcmbCodigo & "' " & _
             " AND comp_egreso.emp_codigo = '" & strEmpresa & "'"
            '  ORDER BY if(det_asiento.cta_codigo=cta_banco.cta_ban_ctaconta,0,1)
    clsDet.Ejecutar strSQL
    If clsDet.adorec_Def.EOF = False Then
        Set VSFG.DataSource = clsDet.adorec_Def.DataSource
        PonerBotones
        CalcuTotal
    Else
        VSFG.Clear 1
    End If
Else
   limpia
   On Error Resume Next
   VSFG.Clear 1
   VSFG.Row = 2
End If
End If
 
End Sub

Private Sub dcmbCodigo_GotFocus()
 Seleccionar_Contenido
End Sub

Private Sub Form_Activate()
Dim strComparar As String
    
'On Error GoTo errhandler
    dcmbCodigo.SetFocus
    
     strSQL = " SELECT com_egr_codigo " & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo= '" & strEmpresa & "' " & _
             " ORDER BY com_egr_codigo"
    clsEgr.Ejecutar strSQL
    If clsEgr.adorec_Def.EOF = False Then
        Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
        dcmbCodigo.ListField = "com_egr_codigo"
        dcmbCodigo.Text = clsEgr.adorec_Def("com_egr_codigo")
    End If
Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
'
End Sub

'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa las clases para hacer distintas consultas
    
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsCom.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
  
'    Consulta para poner el codigo del comprobante en el combo
    
    strSQL = " SELECT com_egr_codigo " & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo= '" & strEmpresa & "' " & _
             " ORDER BY com_egr_codigo"
    clsEgr.Ejecutar strSQL
    If clsEgr.adorec_Def.EOF = False Then
        Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
        dcmbCodigo.ListField = "com_egr_codigo"
        dcmbCodigo.Text = clsEgr.adorec_Def("com_egr_codigo")
    End If

        Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select

End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

    ' only interesetd in left button
    If Button <> 1 Then Exit Sub

    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub

    If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub

    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        If r > 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Egresos Comunes"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
    
        'Recorro el FlexGrid para poner números a las filas
    
        If respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            PonerBotones
            CalcuTotal
        Else
            VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 1 Then
        If Col = 1 Then
            Cancel = True
        End If
        If Col = 2 Then
            Cancel = True
        End If
        If Col = 3 Then
            Cancel = True
        End If
        If Col = 4 Then
            Cancel = True
        End If
    End If
End Sub
