VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerNotasCreditoDebito 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Crédito y Débito"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerNotasCreditoDebito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8100
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Height          =   5055
      Left            =   143
      TabIndex        =   15
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2025
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   4665
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   4665
         Width           =   1935
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
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   1320
         TabIndex        =   1
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1425
         Width           =   1215
      End
      Begin VB.TextBox txtdocumento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1425
         Width           =   2055
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   705
         Width           =   1815
      End
      Begin VB.TextBox txtTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   705
         Width           =   2295
      End
      Begin VB.TextBox txtBanco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1065
         Width           =   2295
      End
      Begin VB.TextBox txtCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1065
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   3105
         Width           =   7320
         _cx             =   12912
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
         FormatString    =   $"frmVerNotasCreditoDebito.frx":030A
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
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   6120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Doc:"
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
         Left            =   5520
         TabIndex        =   26
         Top             =   270
         Width           =   615
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
         Left            =   240
         TabIndex        =   24
         Top             =   1785
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
         TabIndex        =   23
         Top             =   4680
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   1800
         Picture         =   "frmVerNotasCreditoDebito.frx":03D3
         ToolTipText     =   "Elimina una Fila"
         Top             =   4665
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   2040
         Picture         =   "frmVerNotasCreditoDebito.frx":0509
         Top             =   4665
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
         Left            =   4080
         TabIndex        =   22
         Top             =   705
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
         Left            =   3000
         TabIndex        =   21
         Top             =   270
         Width           =   615
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
         Left            =   4080
         TabIndex        =   20
         Top             =   1455
         Width           =   1185
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
         Left            =   4080
         TabIndex        =   19
         Top             =   1095
         Width           =   1245
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de nota"
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
         TabIndex        =   18
         Top             =   735
         Width           =   885
      End
      Begin VB.Label lbldocumento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Documento:"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1455
         Width           =   1365
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
         TabIndex        =   16
         Top             =   1095
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4103
      TabIndex        =   14
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2423
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmVerNotasCreditoDebito"
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

Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsNota As New clsConsulta
Private strSQL As String
Dim n As String
Private intDato As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsNota = Nothing
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub Limpiar()
    'dcmbCodigo.Text = ""
    'dcmbCodigo.BoundText = ""
    'dcmbCodigo.ListField = ""
    txtTipo = ""
    txtBanco = ""
    txtFecha = ""
    txtCuenta = ""
    txtDocumento = ""
    txtValor = ""
    txtValor.Text = FormatoD2(txtValor.Text)
    txtDescripcion = ""
    txtTotalDebe = 0
    txtTotalDebe.Text = FormatoD2(txtTotalDebe.Text)
    txtTotalHaber = 0
    txtTotalHaber.Text = FormatoD2(txtTotalHaber.Text)
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
    VSFG.Clear 1
    'For i = 1 To p
    '    VSFG.TextMatrix(1, i) = ""
    'Next i
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

Private Sub cmdModificar_Click()
    frmNotaCreditoDebito.Show
    frmNotaCreditoDebito.Tag = "M"
    
    If optCredito.Value = True Then
        frmNotaCreditoDebito.optCredito.Value = True
    ElseIf optDebito.Value = True Then
        frmNotaCreditoDebito.optDebito.Value = True
    End If
    
    frmNotaCreditoDebito.txtCodigo = dcmbCodigo
    frmNotaCreditoDebito.dcmbBanco.Text = txtBanco
    frmNotaCreditoDebito.dcmbBanco.BoundText = txtBanco.Tag
    frmNotaCreditoDebito.dcmbCuenta.Text = txtCuenta
    frmNotaCreditoDebito.dcmbCuenta.BoundText = txtCuenta
    frmNotaCreditoDebito.txtDocumento = txtDocumento
    frmNotaCreditoDebito.txtValor = txtValor
    frmNotaCreditoDebito.txtDescripcion = txtDescripcion
    frmNotaCreditoDebito.dcmbTipo.Text = txtTipo
    frmNotaCreditoDebito.dcmbTipo.BoundText = txtTipo.Tag
    

    
End Sub

Private Sub cmdNuevo_Click()
    frmNotaCreditoDebito.Show
    frmNotaCreditoDebito.Tag = "N"
        strSQL = " SELECT COALESCE(max(not_d_c_codigo),0) as num " & _
             " FROM nota_d_c" & _
             " WHERE emp_codigo = '" & strEmpresa & _
             "' AND tip_not_d_c='C'" & _
             " GROUP BY emp_codigo"
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.EOF Then
        frmNotaCreditoDebito.txtCodigo.Text = 1
    Else
        frmNotaCreditoDebito.txtCodigo.Text = clsSql.adorec_Def("num") + 1
    End If
  End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
    End If

End Sub

Private Sub dcmbNombre_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub


Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub Form_Activate()
    If optCredito.Value = True Then
        optcredito_Click
    ElseIf optDebito.Value = True Then
        Optdebito_Click
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
    
    clsNota.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    
    n = 1 'valor del option
    Me.Caption = " Nota de Crédito"
    
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
    n = 0
    Me.Caption = "Nota de Débito"
    dcmbCodigo = ""
    Limpiar
    
    strSQL = " SELECT (nota_d_c.not_d_c_codigo) as not_d_c_codigo, nota_d_c.cta_ban_numero, nota_d_c.ban_codigo, banco.ban_nombre,nota_d_c.tip_not_codigo,tipo_nota.tip_not_nombre," & _
             " not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto " & _
             " FROM ((nota_d_c INNER JOIN banco ON nota_d_c.ban_codigo = banco.ban_codigo) " & _
             "                 INNER JOIN tipo_nota ON nota_d_c.tip_not_codigo= tipo_nota.tip_not_codigo  " & _
             "                                      AND nota_d_c.tip_not_d_c=tipo_nota.tip_not_d_c) " & _
             " WHERE nota_d_c.tip_not_d_c = 'D' AND emp_codigo= '" & strEmpresa & "' " & _
             " ORDER BY not_d_c_numero,LEFT(nota_d_c.not_d_c_codigo,'000000 ')"
    clsNota.Ejecutar strSQL
    If clsNota.adorec_Def.EOF = False Then
        Set dcmbCodigo.RowSource = clsNota.adorec_Def.DataSource
        Set dcmbNombre.RowSource = clsNota.adorec_Def.DataSource
    Else
        Set dcmbCodigo.RowSource = Nothing
        Set dcmbNombre.RowSource = Nothing
    End If
    dcmbCodigo.ListField = "not_d_c_codigo"
    dcmbNombre.ListField = "not_d_c_numero"
    dcmbNombre.BoundColumn = "not_d_c_codigo"
End Sub

Private Sub optcredito_Click()
    n = 1
    Me.Caption = "Nota de Crédito"
    dcmbCodigo = ""
    Limpiar
  
    strSQL = " SELECT (nota_d_c.not_d_c_codigo) as not_d_c_codigo, nota_d_c.cta_ban_numero, nota_d_c.ban_codigo, banco.ban_nombre,nota_d_c.tip_not_codigo,tipo_nota.tip_not_nombre," & _
             " not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto " & _
             " FROM ((nota_d_c INNER JOIN banco ON nota_d_c.ban_codigo = banco.ban_codigo) " & _
             "                 INNER JOIN tipo_nota ON nota_d_c.tip_not_codigo= tipo_nota.tip_not_codigo " & _
             "                                      AND nota_d_c.tip_not_d_c=tipo_nota.tip_not_d_c) " & _
             " WHERE nota_d_c.tip_not_d_c = 'C' AND emp_codigo= '" & strEmpresa & "' " & _
             " ORDER BY not_d_c_numero,LEFT(nota_d_c.not_d_c_codigo,'000000 ')"
    clsNota.Ejecutar strSQL
    If clsNota.adorec_Def.EOF = False Then
        Set dcmbCodigo.RowSource = clsNota.adorec_Def.DataSource
        Set dcmbNombre.RowSource = clsNota.adorec_Def.DataSource
    Else
        Set dcmbCodigo.RowSource = Nothing
        Set dcmbNombre.RowSource = Nothing
    End If
    dcmbCodigo.ListField = "not_d_c_codigo"
    dcmbNombre.ListField = "not_d_c_numero"
    dcmbNombre.BoundColumn = "not_d_c_codigo"
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
Private Sub dcmbCodigo_Change()
    Dim strTNota As String
    'variable que guarda el valor del option, C si es Crédito y D si es débito
    n = 1
    strTNota = "C"
    If optDebito.Value = True Then
        n = 0
        strTNota = "D"
    End If
    If clsNota.adorec_Def.RecordCount > 0 Then
        clsNota.adorec_Def.MoveFirst
    End If
    If dcmbCodigo <> "" Then
    
        clsNota.adorec_Def.Find ("not_d_c_codigo = '" & dcmbCodigo.Text & "'")
        
        dcmbCodigo.Tag = "A"
        If clsNota.adorec_Def.EOF = False Then
                    
            txtTipo = clsNota.adorec_Def("tip_not_nombre")
            txtTipo.Tag = clsNota.adorec_Def("tip_not_codigo")
            txtBanco = clsNota.adorec_Def("ban_nombre")
            txtBanco.Tag = clsNota.adorec_Def("ban_codigo")
            txtFecha = clsNota.adorec_Def("not_d_c_fecha")
            txtCuenta = clsNota.adorec_Def("cta_ban_numero")
            txtDocumento = clsNota.adorec_Def("not_d_c_numero")
            txtValor = clsNota.adorec_Def("not_d_c_monto")
            txtValor.Text = FormatoD2(txtValor.Text)
            txtDescripcion = clsNota.adorec_Def("not_d_c_descripcion")
        
        strSQL = " SELECT distinct det_asiento.cta_codigo, cta_nombre, det_asi_debe, det_asi_haber " & _
                 " FROM (((nota_d_c INNER JOIN det_asiento ON nota_d_c.asi_numasiento=det_asiento.asi_numasiento " & _
                 "                                          AND nota_d_c.emp_codigo=det_asiento.emp_codigo " & _
                 " " & _
                 ") " & _
                 "                  INNER JOIN cta_banco ON nota_d_c.cta_ban_numero=cta_banco.cta_ban_numero " & _
                 "                                       AND nota_d_c.ban_codigo=cta_banco.ban_codigo " & _
                 "                                       AND nota_d_c.emp_codigo=cta_banco.emp_codigo) " & _
                 "                  INNER JOIN ctaconta ON det_asiento.cta_codigo = ctaconta.cta_codigo " & _
                 "                                      AND det_asiento.emp_codigo = ctaconta.emp_codigo) " & _
                 " WHERE nota_d_c.not_d_c_codigo = '" & dcmbCodigo.BoundText & "' " & _
                 " AND nota_d_c.tip_not_d_c='" & strTNota & "'" & _
                 " AND ctaconta.emp_codigo = '" & strEmpresa & "' "
                 'order by if(det_asiento.cta_codigo=cta_banco.cta_ban_ctaconta,0,1)
        clsDet.Ejecutar strSQL
    
        Set VSFG.DataSource = clsDet.adorec_Def.DataSource
        PonerBotones
        CalcuTotal
        Else
            Limpiar
        End If
        dcmbCodigo.Tag = ""
    Else
        Limpiar
    End If
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
