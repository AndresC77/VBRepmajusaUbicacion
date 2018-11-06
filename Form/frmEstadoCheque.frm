VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEstadoCheque 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cheques"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmEstadoCheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8025
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Estado Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   105
      TabIndex        =   19
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox cmbAño 
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
         ItemData        =   "frmEstadoCheque.frx":030A
         Left            =   5295
         List            =   "frmEstadoCheque.frx":036B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   705
         Width           =   780
      End
      Begin VB.ComboBox cmbMes 
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
         ItemData        =   "frmEstadoCheque.frx":0429
         Left            =   6090
         List            =   "frmEstadoCheque.frx":0454
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   705
         Width           =   780
      End
      Begin VB.ComboBox cmbDia 
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
         ItemData        =   "frmEstadoCheque.frx":0494
         Left            =   6885
         List            =   "frmEstadoCheque.frx":04F5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   705
         Width           =   780
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   4440
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   4920
         Width           =   1815
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   5400
         Width           =   1815
      End
      Begin VB.ComboBox cmbdiach 
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
         ItemData        =   "frmEstadoCheque.frx":056C
         Left            =   6915
         List            =   "frmEstadoCheque.frx":05CD
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1065
         Width           =   780
      End
      Begin VB.ComboBox cmbmesch 
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
         ItemData        =   "frmEstadoCheque.frx":0644
         Left            =   6120
         List            =   "frmEstadoCheque.frx":066F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1065
         Width           =   780
      End
      Begin VB.ComboBox Cmbañoch 
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
         ItemData        =   "frmEstadoCheque.frx":06AF
         Left            =   5280
         List            =   "frmEstadoCheque.frx":0710
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1065
         Width           =   780
      End
      Begin VB.CheckBox chkfechas 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Todas las Fechas"
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
         Left            =   4560
         TabIndex        =   3
         Top             =   390
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cmbestado 
         Height          =   315
         ItemData        =   "frmEstadoCheque.frx":07CE
         Left            =   1560
         List            =   "frmEstadoCheque.frx":07D0
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox real 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox disponible 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox previsto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcmbBanco 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2895
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   7575
         _cx             =   13361
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
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEstadoCheque.frx":07D2
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
         ExplorerBar     =   1
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
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
         TabIndex        =   27
         Top             =   405
         Width           =   510
      End
      Begin VB.Label lbldescripcion1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del cheque:"
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
         TabIndex        =   26
         Top             =   1125
         Width           =   1380
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   1920
         Picture         =   "frmEstadoCheque.frx":0953
         ToolTipText     =   "Elimina una Fila"
         Top             =   4920
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   2160
         Picture         =   "frmEstadoCheque.frx":0A89
         Top             =   4920
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   4605
         TabIndex        =   25
         Top             =   765
         Width           =   510
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
         TabIndex        =   24
         Top             =   765
         Width           =   1245
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
         Left            =   3480
         TabIndex        =   23
         Top             =   4455
         Width           =   975
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
         Left            =   3480
         TabIndex        =   22
         Top             =   4935
         Width           =   1455
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
         Left            =   3480
         TabIndex        =   21
         Top             =   5415
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   4605
         TabIndex        =   20
         Top             =   1125
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2385
      TabIndex        =   17
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4065
      TabIndex        =   18
      Top             =   6240
      Width           =   1575
   End
End
Attribute VB_Name = "frmEstadoCheque"
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
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private clsCom As New clsConsulta
Private clsSql As New clsConsulta
Private strSql As String
Dim m As String
Dim p As String
Dim posanterior As String
Dim posactual As String
Dim filact As String
Dim filant As String
Private intDato As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
    Set clsCom = Nothing
    Set clsSql = Nothing
End Sub

Private Sub Saldos()
    Dim estado As String
    With VSFG
        txtSaldoReal = real
        txtDisponible = disponible
        txtPrevisto = previsto
        
        For i = 1 To (.Rows - 1)
            
            If .TextMatrix(i, 5) <> .TextMatrix(i, 9) Then
                estado = .TextMatrix(i, 5)
            Else
                estado = ""
            End If
            Select Case estado
                Case "GIRADO"
                             txtSaldoReal = real
                             txtDisponible = disponible
                             txtPrevisto = previsto
                Case "COBRADO"
                             txtSaldoReal = Val(txtSaldoReal) - Val(.TextMatrix(i, 4))
                             txtDisponible = disponible
                             txtPrevisto = previsto
                Case "ANULADO", "PROTESTADO"
                             txtDisponible = Val(txtDisponible) + Val(.TextMatrix(i, 4))
                             txtPrevisto = Val(txtPrevisto) + Val(.TextMatrix(i, 4))
                             txtSaldoReal = real
                Case ""
                                
            End Select
        Next i
    End With
End Sub

Private Sub saldodisponible()
    
    'Calcula el saldo disponible de la cuenta bancaria
    
    strSql = " SELECT  sum(com_egr_ch_valor) as valor, com_egr_ch_fecha" & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND com_egr_ch_estado = 'GIRADO' AND cta_ban_numero = '" & dcmbCuenta.Text & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'  AND com_egr_ch_fecha <= CURRENT_TIMESTAMP " & _
             " GROUP BY cta_ban_numero "
    clsSql.Ejecutar strSql
    
    If Not IsNull(clsSql.adorec_Def("valor")) And clsSql.adorec_Def.EOF = False Then
        valor = clsSql.adorec_Def("valor")
        disp = Val(txtSaldoReal) - Val(valor)
        txtDisponible = disp
        disponible = disp
    Else
        valor = 0
        disp = Val(txtSaldoReal) - Val(valor)
        txtDisponible = disponible
        disponible = disp
    End If
End Sub
Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub chkFechas_Click()
    cmbestado_Validate False
End Sub


Private Sub Cmbaño_Click()
cmbestado_Validate False
End Sub



Private Sub Cmbañoch_Click()
cmbestado_Validate False
End Sub


Private Sub cmbdia_Click()
cmbestado_Validate False
End Sub



Private Sub cmbdiach_Click()
cmbestado_Validate False
End Sub

'Private Sub cmbestado_Change()

Private Sub cmbestado_Validate(Cancel As Boolean)
Dim ff As Variant
Dim ffch As Variant

On Error Resume Next

ff = Format(cmbAño.Text + "-" + cmbMes.Text + "-" + cmbDia.Text, "yyyy-mm-dd")

ffch = Format(Cmbañoch.Text + "-" + cmbmesch.Text + "-" + cmbdiach.Text, "yyyy-mm-dd")

If dcmbBanco <> "" And dcmbCuenta <> "" Then
If UCase(cmbestado) = "TODOS" Then
    cmbestado = "TODO"
End If
If UCase(cmbestado) = "GIRADOS" Then
    cmbestado = "GIRADO"
End If
If UCase(cmbestado) = "COBRADOS" Then
    cmbestado = "COBRADO"
End If
If UCase(cmbestado) = "ANULADOS" Then
    cmbestado = "ANULADO"
End If
If UCase(cmbestado) = "PROTESTADOS" Then
    cmbestado = "PROTESTADO"
End If



    If UCase(cmbestado) = "TODO" Then
        strSql = " SELECT distinct com_egr_ch_fecha,com_egr_ch_num , CONCAT(per_apellido,' ',per_nombre) as persona,com_egr_ch_valor, com_egr_ch_estado, com_egr_codigo, com_egr_fecha, CONCAT(SUBSTRING(com_egr_descripcion,1,15),'...') as descripcion, com_egr_ch_estado " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo = persona.per_codigo " & _
                 " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND cta_ban_numero = '" & dcmbCuenta.BoundText & "'"
    End If
    
    If UCase(cmbestado) = "GIRADO" Or UCase(cmbestado) = "COBRADO" Or UCase(cmbestado) = "ANULADO" Or UCase(cmbestado) = "PROTESTADO" Then
        strSql = " SELECT distinct com_egr_ch_fecha, com_egr_ch_num, CONCAT(per_apellido,' ',per_nombre) as persona,com_egr_ch_valor, com_egr_ch_estado, com_egr_codigo, com_egr_fecha, CONCAT(SUBSTRING(com_egr_descripcion,1,15),'...') as descripcion, com_egr_ch_estado " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo = persona.per_codigo " & _
                 " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND cta_ban_numero = '" & dcmbCuenta.BoundText & "' AND com_egr_ch_estado = '" & cmbestado.Text & "'"
                 
    End If
    If chkFechas.value = 0 Then
        strSql = strSql & " AND comp_egreso.com_egr_ch_fecha BETWEEN '" & ff & "' AND '" & ffch & "' "
    End If
    'strSql = strSql & " ORDER BY LPAD(com_egr_ch_num,15,' ') "
    clsCom.Ejecutar strSql
    If clsCom.adorec_Def.RecordCount > 0 Then
        Set VSFG.DataSource = clsCom.adorec_Def.DataSource
        VSFG.Col = 5
        PonerBotones
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
        a = VSFG.Rows - 1
        p = 8
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
End If
 If Trim(dcmbCuenta) <> "" Then
        strSql = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre,cta_ban_ch_ultimo as cheque,cta_ban_saldoreal,cta_ban_saldoprevisto " & _
                 " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo " & _
                 "                                    AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
        clsctc.Ejecutar strSql
        If Not clsctc.adorec_Def.EOF Then
            txtSaldoReal.Text = clsctc.adorec_Def("cta_ban_saldoreal")
            real = clsctc.adorec_Def("cta_ban_saldoreal")
            txtPrevisto.Text = clsctc.adorec_Def("cta_ban_saldoprevisto")
            previsto = clsctc.adorec_Def("cta_ban_saldoprevisto")
            saldodisponible
        End If
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
End If
End Sub

Private Sub cmbmes_Click()
cmbestado_Validate False
End Sub

Private Sub cmbmesch_Click()
cmbestado_Validate False
End Sub

Private Sub cmdAceptar_Click()
    
   
    'Actualiza los saldos en la tabla cta_banco
    strSql = " UPDATE cta_banco " & _
             " SET cta_ban_saldoreal= '" & txtSaldoReal & "',cta_ban_saldoprevisto= '" & txtPrevisto & "',cta_ban_saldodisponible= '" & txtDisponible & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
             " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
    clsSql.Ejecutar (strSql), "M"
    'Actualiza el estado del cheque en la tabla comp_egreso
    With VSFG
        For i = 1 To .Rows - 1
           strSql = " UPDATE comp_egreso " & _
                    " SET com_egr_ch_estado= '" & .TextMatrix(i, 5) & "', com_egr_fechamod = CURRENT_TIMESTAMP, com_egr_usumod= '" & strUsuario & "'" & _
                    " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "' AND com_egr_codigo = '" & .TextMatrix(i, 6) & "' "
           clsSql.Ejecutar strSql, "M"
        Next i
    End With
    
    MsgBox "El Estado de los Cheques seleccionados ha sido modificado", vbInformation, "Estado de Cheques"
    
    cmbestado_Validate False
    
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBanco_Change()
dcmbCuenta = ""
cmbestado.Text = ""
    strSql = " SELECT cta_ban_numero, cta_ban_ctaconta,cta_ban_saldoreal,cta_ban_saldodisponible,cta_ban_saldoprevisto" & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSql
    If clsCtb.adorec_Def.EOF = False Then
        Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
        dcmbCuenta.ListField = ("cta_ban_numero")
    Else
        Set dcmbCuenta.RowSource = Nothing
        dcmbCuenta = ""
    End If
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa las clases para hacer distintas consultas
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsCom.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster

    'Pone valores al combo de estado
    cmbestado.AddItem "TODO"
    cmbestado.AddItem "GIRADO"
    cmbestado.AddItem "COBRADO"
    cmbestado.AddItem "ANULADO"
    cmbestado.AddItem "PROTESTADO"
    cmbestado.Text = ""
    
    cmbestado.Enabled = False
    
    VSFG.TextMatrix(1, 5) = ""
    VSFG.Col = 5
   'Pone la fecha actual en los combos
    d = CStr(Day(HoyDia))
    mm = Month(HoyDia)
    m = Month(HoyDia)
    Y = CStr(Year(HoyDia))
    cmbDia.Text = d
    cmbAño.Text = Y
    cmbdiach.Text = d
    Cmbañoch.Text = Y
    cmbMes.Text = Format(HoyDia, "mmm")
    cmbmesch.Text = Format(HoyDia, "mmm")

'    Consulta para sacar los bancos existentes en el combo
    strSql = " SELECT ban_codigo, ban_nombre " & _
             " FROM banco " & _
             " ORDER BY ban_codigo"
    clsBan.Ejecutar strSql
    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    End If
    
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
'    End Sub
'errhandler:
'    Select Case Err.Number
'        Case 1046
'            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
'                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
'                   " database, so Please Choose a database."
'        Case Else
'            MsgBox "[" & Err.Number & "] " & Err.Description
'    End Select

End Sub
Private Sub dcmbCuenta_Change()

    cmbestado.Text = ""
   
    a = VSFG.Rows - 1
    p = 8
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

    If Trim(dcmbCuenta) <> "" Then
        strSql = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre,cta_ban_ch_ultimo as cheque,cta_ban_saldoreal,cta_ban_saldoprevisto " & _
                 " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo " & _
                 "                                    AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
        clsctc.Ejecutar strSql
        If Not clsctc.adorec_Def.EOF Then
            txtSaldoReal.Text = clsctc.adorec_Def("cta_ban_saldoreal")
            real = clsctc.adorec_Def("cta_ban_saldoreal")
            txtPrevisto.Text = clsctc.adorec_Def("cta_ban_saldoprevisto")
            previsto = clsctc.adorec_Def("cta_ban_saldoprevisto")
            saldodisponible
        End If
    Else
        txtSaldoReal = 0
        txtPrevisto = 0
        txtDisponible = 0
        real = 0
        dispoible = 0
        previsto = 0
        n = 9
        a = VSFG.Rows - 1
        For i = 2 To a
          If VSFG.Rows - 1 = 1 Then
              Exit For
          End If
          VSFG.RemoveItem i
          i = i - 1
          a = a - 1
        Next i

        For i = 1 To n
            VSFG.TextMatrix(1, i) = ""
        Next i

    End If
    cmbestado.Enabled = True
    
    txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)

End Sub


Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
 Saldos
 txtSaldoReal.Text = FormatoD2(txtSaldoReal.Text)
    txtDisponible.Text = FormatoD2(txtDisponible.Text)
    txtPrevisto.Text = FormatoD2(txtPrevisto.Text)
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol <> 5 Then
        VSFG.Col = 5
    End If
    
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   
        If Col <> 5 Then
            Cancel = True
        End If
        
        If VSFG.TextMatrix(Row, 9) = "COBRADO" Then
            Cancel = True
        End If

        If VSFG.TextMatrix(Row, 9) = "ANULADO" Then
            Cancel = True
        End If

        If VSFG.TextMatrix(Row, 9) = "PROTESTADO" Then
            Cancel = True
        End If
              
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_GotFocus()
    VSFG.Col = 5
End Sub
