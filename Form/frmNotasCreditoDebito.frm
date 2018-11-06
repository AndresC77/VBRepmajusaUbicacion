VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNotaCreditoDebito 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Crédito y Débito"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmNotasCreditoDebito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8985
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Height          =   7335
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtDescripcion 
         Height          =   525
         Left            =   765
         TabIndex        =   13
         Top             =   4800
         Width           =   7335
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   6960
         Width           =   1275
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   6960
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Height          =   735
         Left            =   765
         TabIndex        =   31
         Top             =   120
         Width           =   7455
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   285
            Left            =   1680
            TabIndex        =   38
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   0
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha comprobante:"
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
            Left            =   120
            TabIndex        =   33
            Top             =   270
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
            Left            =   4680
            TabIndex        =   32
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Nota de Crédito/Débito"
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
         Height          =   1455
         Left            =   765
         TabIndex        =   28
         Top             =   960
         Width           =   7455
         Begin VB.TextBox txtDescripciont 
            Enabled         =   0   'False
            Height          =   765
            Left            =   2880
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   600
            Width           =   4335
         End
         Begin VB.OptionButton optDebito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Débito"
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
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton optCredito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Crédito"
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
            Left            =   480
            TabIndex        =   1
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcmbTipo 
            Height          =   315
            Left            =   2880
            TabIndex        =   37
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblBeneficiario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de nota:"
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
            Left            =   1920
            TabIndex        =   30
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lbldescripcion1 
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
            Left            =   1920
            TabIndex        =   29
            Top             =   615
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos Bancarios"
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
         Left            =   765
         TabIndex        =   20
         Top             =   2520
         Width           =   7455
         Begin VB.TextBox txtr 
            Height          =   285
            Left            =   5490
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtp 
            Height          =   285
            Left            =   5490
            TabIndex        =   12
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtd 
            Height          =   285
            Left            =   5490
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
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
            Left            =   1800
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtPrevisto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtDisponible 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtSaldoReal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtdocumento 
            Height          =   285
            Left            =   1770
            TabIndex        =   5
            Top             =   1080
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   1770
            TabIndex        =   3
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1770
            TabIndex        =   4
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de la nota:"
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
            TabIndex        =   27
            Top             =   1477
            Width           =   1185
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Previsto:"
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
            Left            =   4200
            TabIndex        =   26
            Top             =   1335
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Disponible:"
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
            Left            =   4200
            TabIndex        =   25
            Top             =   855
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Real:"
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
            Left            =   4200
            TabIndex        =   24
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   23
            Top             =   772
            Width           =   1245
         End
         Begin VB.Label lbldocumento 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número de Documento:"
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
            TabIndex        =   22
            Top             =   1117
            Width           =   1680
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
            Left            =   120
            TabIndex        =   21
            Top             =   412
            Width           =   510
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   5400
         Width           =   8520
         _cx             =   15028
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
         FormatString    =   $"frmNotasCreditoDebito.frx":030A
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
      Begin VB.Label lblDescripcion 
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
         Left            =   765
         TabIndex        =   35
         Top             =   4560
         Width           =   900
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
         Left            =   2640
         TabIndex        =   34
         Top             =   6975
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   1320
         Picture         =   "frmNotasCreditoDebito.frx":03E4
         ToolTipText     =   "Elimina una Fila"
         Top             =   6960
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   1560
         Picture         =   "frmNotasCreditoDebito.frx":051A
         Top             =   6960
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2355
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4005
      TabIndex        =   18
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmNotaCreditoDebito"
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

Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private clsCom As New clsConsulta
Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsEgr As New clsConsulta
Private clsNota As New clsConsulta
Private clsTip As New clsConsulta
Private strSQL As String

Dim ff As Variant
Dim d As Variant
Dim m As String
Dim n As Integer
Private intDato As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
    Set clsCom = Nothing
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsEgr = Nothing
    Set clsNota = Nothing
    Set clsTip = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub
Private Sub saldodisponible()
    
    'Calcula el saldo disponible de la cuenta bancaria
    
    strSQL = " SELECT  sum(com_egr_ch_valor) as valor" & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND com_egr_ch_estado = 'GIRADO' AND cta_ban_numero = '" & dcmbCuenta.Text & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'  AND com_egr_ch_fecha <= CURRENT_TIMESTAMP " & _
             " GROUP BY cta_ban_numero "
    clsSql.Ejecutar strSQL
    
    If Not IsNull(clsSql.adorec_Def("valor")) And clsSql.adorec_Def.EOF = False Then
        Valor = clsSql.adorec_Def("valor")
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
    Else
        Valor = 0
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
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

Private Sub cmdAceptar_Click()

   Dim strTNota As String
   Dim ff As String
   Dim d As String
'Comprueba que todos los datos esten ingresados
    ff = Format(dtpFecha.Value, "yyyy-mm-dd")
    d = Format(HoyDia, "yyyy-MM-dd")

    Fecha = DateDiff("d", d, ff)
       
    dia = Mid(d, 9, 2)
    Mes = Mid(d, 6, 2)
    Año = Mid(d, 1, 4)
    m = Mid(ff, 6, 2)
    If (IsDate(ff) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Egresos Comunes"
        Exit Sub
    End If

    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd

    a = VSFG.Rows - 1
'    For i = 1 To a
'        For j = i + 1 To a
'            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) Then
'                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
'                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
'                VSFG.RemoveItem j
'                a = a - 1
'                j = j - 1
'            End If
'            If j >= a Then
'                Exit For
'            End If
'        Next j
'    Next i

    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbExclamation, "Comprobante de Egreso Común"
        Exit Sub
    End If
    'Verificar que todos los datos se han llenado para ingresar en la base de datos
    If Fecha > 0 Or txtCodigo = "" Or VSFG.TextMatrix(1, 1) = "" Or txtDescripcion = "" Or txtDocumento = "" Or dcmbTipo = "" Then
        If Fecha > 0 Then
            MsgBox "La fecha ingresado debe ser posterior a la fecha actual", vbExclamation, "Comprobante de Egreso Común"
            
        Else
            MsgBox "No están ingresados todos los datos", vbInformation, "Ingreso"
            
        End If
        Exit Sub
    End If
       
    If Me.Tag = "N" Then
        Dim clsAsiento As New clsContable
        clsAsiento.Inicializar AdoConn, AdoConnMaster
        clsAsiento.NuevoAsiento "D", ff, 0, 0, FormatoD2(txtTotalDebe.Text), txtDescripcion.Text, True
        'Ingreso de datos en nota_d_c
        If n = 0 Then
            strSQL = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha,not_d_c_descripcion, not_d_c_monto,asi_numasiento, not_d_c_fechamod, not_d_c_usumod) " & _
                     " VALUES ('D','" & txtCodigo.Text & "', '" & dcmbCuenta.BoundText & "', '" & dcmbBanco.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtDocumento.Text & "','" & ff & "','" & txtDescripcion.Text & "','" & txtValor.Text & "','" & clsAsiento.NumAsiento & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsSql.Ejecutar (strSQL), "M"
            
            
        ElseIf n = 1 Then
            strSQL = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto,asi_numasiento, not_d_c_fechamod, not_d_c_usumod) " & _
                     " VALUES ('C','" & txtCodigo.Text & "', '" & dcmbCuenta.BoundText & "', '" & dcmbBanco.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtDocumento.Text & "','" & ff & "','" & txtDescripcion.Text & "','" & txtValor.Text & "','" & clsAsiento.NumAsiento & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
            clsSql.Ejecutar (strSQL), "M"
        
        End If
            
        
            'ingreso de datos en el la tabla det_nota_d_c
        
            With VSFG
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                        clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                    End If
                Next i
            End With
        
        MsgBox " Los datos han sido ingresados", vbInformation, "Ingresos"
        Dim frmAsien As New frmReporte
        frmAsien.strAsiento = clsAsiento.NumAsiento
        frmAsien.strReporte = "rptAsiento"
        frmAsien.Show

    End If
        'Actualiza los valores de los saldos
         strSQL = " UPDATE cta_banco " & _
                  " SET cta_ban_saldoreal= '" & txtSaldoReal & "',cta_ban_saldoprevisto= '" & txtPrevisto & "',cta_ban_saldodisponible= '" & txtDisponible & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                  " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
         clsSql.Ejecutar (strSQL), "M"
  Unload Me
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub dcmbBanco_Change()
dcmbCuenta = ""
dcmbBanco.Tag = dcmbBanco.BoundText
    
    strSQL = " SELECT cta_ban_numero, cta_ban_ch_ultimo as ban, cta_ban_ctaconta,cta_ban_saldoreal,cta_ban_saldodisponible,cta_ban_saldoprevisto" & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSQL
    If clsCtb.adorec_Def.EOF = False Then
        Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
        dcmbCuenta.ListField = ("cta_ban_numero")
     
    Else
        Set dcmbCuenta.RowSource = Nothing
        dcmbCuenta.Text = ""
        dcmbCuenta.BoundText = ""
    End If


    
End Sub


Private Sub dcmbTipo_Change()
 If dcmbTipo = "" Then
    txtDescripciont = ""
    Exit Sub
 End If
 
 
 If n = 1 Then
    strSQL = " SELECT CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'C' " & _
             " AND  tip_not_codigo = '" & dcmbTipo.BoundText & "' "
    clsTip.Ejecutar strSQL

    If clsTip.adorec_Def.EOF = False Then
        txtDescripciont = clsTip.adorec_Def("descripcion")
    Else
        txtDescripciont.Text = ""
    End If
End If

If n = 0 Then
    strSQL = " SELECT CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'D' " & _
             " AND  tip_not_codigo = '" & dcmbTipo.BoundText & "' "
    clsTip.Ejecutar strSQL

    If clsTip.adorec_Def.EOF = False Then
        txtDescripciont = clsTip.adorec_Def("descripcion")
    Else
        txtDescripciont.Text = ""
    End If
End If
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
    
    On Error GoTo errhandler
    'Inicializa las clases para hacer distintas consultas
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsCom.Inicializar AdoConn, AdoConnMaster
    clsNota.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster
    clsTip.Inicializar AdoConn, AdoConnMaster
    
    'Pone la fecha actual en los combos
    
    dtpFecha.Value = HoyDia
    optCredito.Value = True
    
    If optCredito.Value = True Then
        n = 1
        Me.Caption = " Nota de Crédito"
        strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
                 " FROM tipo_nota " & _
                 " WHERE tip_not_d_c = 'C'" & _
                 " ORDER BY tip_not_codigo"
        clsTip.Ejecutar strSQL
        
    ElseIf optDebito.Value = True Then
        
        n = 0
        Me.Caption = " Nota de Débito"
        strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
                 " FROM tipo_nota " & _
                 " WHERE tip_not_d_c = 'D'" & _
                 " ORDER BY tip_not_codigo"
        clsTip.Ejecutar strSQL
    
    End If
    
    If clsTip.adorec_Def.EOF = False Then
        Set dcmbTipo.RowSource = clsTip.adorec_Def.DataSource
        dcmbTipo.ListField = "tip_not_nombre"
        dcmbTipo.BoundColumn = "tip_not_codigo"
        dcmbTipo.Text = clsTip.adorec_Def("tip_not_nombre")
        txtDescripciont.Text = clsTip.adorec_Def("descripcion")
    End If
    
'      Consulta para sacar los bancos existentes en el combo
    strSQL = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL
    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    End If
    
    'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsCta.Ejecutar strSQL

     VSFG.ColComboList(5) = VSFG.BuildComboList(clsCta.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
    'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSQL

     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")
    
    
    txtp = 0
    txtD = 0
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



Private Sub Optdebito_Click()
    Frame2.Caption = "Tipo de nota de Débito"
    n = 0
    Me.Caption = "Nota de Débito"
    
    dcmbTipo.Text = ""
    dcmbTipo.BoundText = ""
    txtDescripciont.Text = ""
    txtSaldoReal = 0
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtPrevisto = 0
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
    txtDisponible = 0
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtp = 0
    txtD = 0
    txtValor = 0
    txtValor.Text = FormatoD2(txtValor.Text)
    txtTotalDebe = 0
    txtTotalDebe.Text = FormatoD2(txtTotalDebe.Text)
    txtTotalHaber = 0
    txtTotalHaber.Text = FormatoD2(txtTotalHaber.Text)
    txtDocumento = ""
    txtDescripcion = ""
    p = 4
    a = VSFG.Rows - 1
    
    For i = 2 To a
        If VSFG.Rows - 1 = 1 Then
            Exit For
        End If
        VSFG.RemoveItem i
        i = i - 1
        a = a - 1
    Next i
    
    For i = 1 To p
        VSFG.TextMatrix(1, i) = ""
    Next i
    dcmbBanco = ""

    strSQL = " SELECT COALESCE(max(not_d_c_codigo),0) as num " & _
             " FROM nota_d_c" & _
             " WHERE emp_codigo = '" & strEmpresa & _
             "' AND tip_not_d_c='D'" & _
             " GROUP BY emp_codigo"
    clsTip.Ejecutar strSQL
    If clsTip.adorec_Def.EOF Then
        txtCodigo.Text = 1
    Else
        txtCodigo.Text = clsTip.adorec_Def("num") + 1
    End If

    strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'D'" & _
             " ORDER BY tip_not_codigo"
    clsTip.Ejecutar strSQL
    
    Set dcmbTipo.RowSource = clsTip.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
    dcmbTipo.Text = clsTip.adorec_Def("tip_not_nombre")
    If clsTip.adorec_Def.EOF = False Then
        txtDescripciont.Text = clsTip.adorec_Def("descripcion")
    Else
        dcmbTipo.Text = ""
        dcmbTipo.BoundText = ""
        txtDescripciont = ""
    End If
End Sub



Private Sub optcredito_Click()
    Frame2.Caption = "Tipo de nota de Crédito"
    n = 1
    
    Me.Caption = " Nota de Crédito"
    
    dcmbTipo.Text = ""
    dcmbTipo.BoundText = ""
    txtDescripciont.Text = ""
    txtSaldoReal = 0
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtPrevisto = 0
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
    txtDisponible = 0
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtp = 0
    txtD = 0
    txtValor = 0
    txtValor.Text = FormatoD2(txtValor.Text)
    txtTotalDebe = 0
    txtTotalDebe.Text = FormatoD2(txtTotalDebe.Text)
    txtTotalHaber = 0
    txtTotalHaber.Text = FormatoD2(txtTotalHaber.Text)
    txtDocumento = ""
    txtDescripcion = ""
    p = 4
    a = VSFG.Rows - 1
    dcmbBanco = ""
    For i = 2 To a
         If VSFG.Rows - 1 = 1 Then
            Exit For
        End If
        VSFG.RemoveItem i
        i = i - 1
        a = a - 1
    Next i

    For i = 1 To p
        VSFG.TextMatrix(1, i) = ""
    Next i

    strSQL = " SELECT COALESCE(max(not_d_c_codigo),0) as num " & _
             " FROM nota_d_c" & _
             " WHERE emp_codigo = '" & strEmpresa & _
             "' AND tip_not_d_c='C'" & _
             " GROUP BY emp_codigo"
    clsTip.Ejecutar strSQL
    If clsTip.adorec_Def.EOF Then
        txtCodigo.Text = 1
    Else
        txtCodigo.Text = clsTip.adorec_Def("num") + 1
    End If

    strSQL = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'C'" & _
             " ORDER BY tip_not_codigo"
    clsTip.Ejecutar strSQL
    Set dcmbTipo.RowSource = clsTip.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
    If clsTip.adorec_Def.EOF = False Then
    dcmbTipo.Text = clsTip.adorec_Def("tip_not_nombre")
    
        txtDescripciont.Text = clsTip.adorec_Def("descripcion")
    Else
        dcmbTipo.Text = ""
        dcmbTipo.BoundText = ""
        txtDescripcion = ""
    End If
End Sub



Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub


Private Sub txtdocumento_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_LostFocus()
    If dcmbBanco <> "" Or dcmbCuenta <> "" Then
        If n = 0 Then
            VSFG.TextMatrix(1, 4) = txtValor.Text
            txtValor.Text = FormatoD2(txtValor.Text)
            CalcuTotal
        ElseIf n = 1 Then
            VSFG.TextMatrix(1, 3) = txtValor.Text
            txtValor.Text = FormatoD2(txtValor.Text)
            CalcuTotal
        End If
    Else
        MsgBox "ingrese una cuenta bancaria", vbInformation, "Valor"
        txtValor.Text = 0
        txtValor.Text = FormatoD2(txtValor.Text)
    End If
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
' Verifica si el dato uçingresado es numérico
    If IsNumeric(txtValor.Text) = False Then
        MsgBox "Solo se permiten valores numéricos", vbOKOnly + vbInformation, "ERROR"
        Cancel = True
    Else
        ' Pone dos decimales al valor
        txtValor.Text = FormatoD2(txtValor.Text)
        Cancel = False
    End If

End Sub
'
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
'
Private Sub dcmbCuenta_Change()
    If Trim(dcmbCuenta) <> "" Then
        strSQL = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre,cta_ban_ch_ultimo as cheque,cta_ban_saldoreal,cta_ban_saldoprevisto " & _
                 " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo " & _
                 "                                    AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
        clsctc.Ejecutar strSQL
        If Not clsctc.adorec_Def.EOF Then
            txtSaldoReal.Text = clsctc.adorec_Def("cta_ban_saldoreal")
            txtPrevisto.Text = clsctc.adorec_Def("cta_ban_saldoprevisto")
            txtr = clsctc.adorec_Def("cta_ban_saldoreal")
            txtp = clsctc.adorec_Def("cta_ban_saldoprevisto")
            If clsctc.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(1, 1) = clsctc.adorec_Def("cta_ban_ctaconta")
                VSFG.TextMatrix(1, 2) = clsctc.adorec_Def("cta_nombre")
            End If
            saldodisponible
        End If
    Else
        txtSaldoReal = 0
        txtPrevisto = 0
        txtDisponible = 0
        txtp = 0
        txtD = 0
        txtValor = 0
        txtDescripcion = ""
        p = 4
        a = VSFG.Rows - 1

        For i = 2 To a
              If VSFG.Rows - 1 = 1 Then
            Exit For
        End If
        VSFG.RemoveItem i
        i = i - 1
        a = a - 1
        Next i

        For i = 1 To p
            VSFG.TextMatrix(1, i) = ""
        Next i

    End If
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
    txtValor.Text = FormatoD2(txtValor.Text)
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes

    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            VSFG.AddItem ""
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
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

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
            MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
            VSFG.TextMatrix(Row, 3) = ""
            VSFG.TextMatrix(Row, 4) = ""


        ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe

        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = intDato
        End If

        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
    CalcuTotal
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)

    'Calcula los saldos previsto y disponible
    d = Format(HoyDia, "yyyy-MM-dd")
    dia = Mid(d, 9, 2)
    Mes = Mid(d, 6, 2)
    Año = Mid(d, 1, 4)
    ff = Format(dtpFecha.Value, "yyyy-mm-dd")
    m = Mid(ff, 6, 2)

    If n = 0 Then

        If Format(ff, "dd") <= dia Or m <= Mes Or Format(ff, "yyyy") <= Año Then
            txtDisponible.Text = Val(txtD) - Val(VSFG.TextMatrix(1, 4))
            txtDisponible.Text = FormatoD2(txtDisponible.Text)
            txtPrevisto.Text = Val(txtp) - Val(VSFG.TextMatrix(1, 4))
            txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
            txtSaldoReal.Text = Val(txtr) - Val(VSFG.TextMatrix(1, 4))
            txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
        End If

    ElseIf n = 1 Then
        If Format(ff, "dd") <= dia Or m <= Mes Or Format(ff, "yyyy") <= Año Then
            txtDisponible.Text = Val(txtD) + Val(VSFG.TextMatrix(1, 3))
            txtDisponible.Text = FormatoD2(txtDisponible.Text)
            txtPrevisto.Text = txtp + Val(VSFG.TextMatrix(1, 3))
            txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
            txtSaldoReal.Text = Val(txtr) + Val(VSFG.TextMatrix(1, 3))
            txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)

        End If
    End If

    
' filtra el nombre y codigo de cuenta para los combos del greed
If Row > 1 Then
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                     .TextMatrix(Row, 2) = .TextMatrix(Row, 1)
             End If

             If Col = 2 Then
                     .TextMatrix(Row, 1) = .TextMatrix(Row, 2)
             End If
         End If
    End With
End If
End Sub

Private Sub VSFG_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            frmSelecCtaConta.Show
            Set frmSelecCtaConta.objEscribir = VSFG
        End If
    End If
End Sub
