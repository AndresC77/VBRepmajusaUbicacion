VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelEgresoComun 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egresos Comunes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "FrmSelEgresoComun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7815
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Egresos Comunes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
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
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   4800
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
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   7320
         _cx             =   12912
         _cy             =   3625
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmSelEgresoComun.frx":030A
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
      Begin MSDataListLib.DataCombo dcmbDescripcion 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   405
         Width           =   540
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
         TabIndex        =   19
         Top             =   720
         Width           =   510
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
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblcuentaban 
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
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   1245
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
         Left            =   2760
         TabIndex        =   16
         Top             =   4815
         Width           =   855
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
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   1560
         Picture         =   "FrmSelEgresoComun.frx":03C0
         ToolTipText     =   "Elimina una Fila"
         Top             =   4800
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   1920
         Picture         =   "FrmSelEgresoComun.frx":04F6
         Top             =   4800
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelEgresoComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de consulta para egresos comunes                                      #
'#  frmSelEgresoComun V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar egresos comunes de la empresa.                       #
'#  Permite visualizar los datos de egresos comunes y sus detalles, no se puede #
'#  modficar, solo se puede eliminar el egreso comun junto con su detalle       #
'#  llama a la forma frmEgresoComun para modificar y crear nuevos de nuevos     #
'#  egresos comunes                                                             #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  EGRESO_COMUN: En esta tabla se almacenan las cabeceras de los egresos       #
'#  det_egreso_comun: donde se guardan los datos de el detalle del egreso       #
'#  BANCO: De donde se consultan los bancos existentes                          #
'#  Cta_banco: de donde se consultan los números de cuentas bancarias y         #
'#             su cuenta contable                                               #
'#  CTACONTA: Para consultar las cuentas cosntables y sus nombres               #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsEgr As New clsConsulta
Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsCtb As New clsConsulta
Private strSQL As String
Private intDato As Variant
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsEgr = Nothing
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsCtb = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
    Next i
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

Private Sub cmdEliminar_Click()
If (MsgBox("Esta seguro de eliminar este egreso común?", vbYesNo, "Egreso Común")) = vbYes Then
 
 'Primero se elimina el detalle del egreso
        strSQL = " DELETE " & _
                 " FROM det_egreso_comun " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND egr_com_codigo='" & dcmbCodigo.Text & "'"
        clsDet.Ejecutar (strSQL), "M"
 
 'Luego se elimina la cabecera del egreso

        strSQL = " DELETE " & _
                 " FROM egreso_comun" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND egr_com_codigo = '" & dcmbCodigo.Text & "'"
        clsEgr.Ejecutar (strSQL), "M"
   
   'Consulta para actualizar los combos
   
    strSQL = " SELECT egr_com_codigo, egr_com_descripcion, CONCAT(SUBSTRING(egr_com_descripcion,1,20),'...') as descripcion,egr_com_fecha, egreso_comun.ban_codigo, cta_ban_numero,ban_nombre " & _
             " FROM egreso_comun INNER JOIN banco ON egreso_comun.ban_codigo=banco.ban_codigo" & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY egr_com_codigo"
    clsEgr.Ejecutar strSQL
    
    If Not clsEgr.adorec_Def.EOF Then
        Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
        dcmbCodigo.ListField = "egr_com_codigo"
        Set dcmbDescripcion.RowSource = clsEgr.adorec_Def.DataSource
        dcmbDescripcion.ListField = "descripcion"
        dcmbDescripcion.BoundColumn = "egr_com_codigo"
        dcmbCodigo = clsEgr.adorec_Def("egr_com_codigo")
    Else
        dcmbCodigo = ""
    End If
Else
    Exit Sub
End If
End Sub

Private Sub cmdModificar_Click()
    frmEgresoComun.Tag = "M"
    frmEgresoComun.Show
    frmEgresoComun.dcmbBanco.Text = Me.txtBanco.Text
    frmEgresoComun.dcmbBanco.Tag = Me.txtBanco.Tag
    frmEgresoComun.dcmbCuenta = Me.txtCuenta
    frmEgresoComun.txtCodigo = Me.dcmbCodigo
    frmEgresoComun.txtDescripcion = Me.txtDescripcion
    frmEgresoComun.txtCodigo.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdNuevo_Click()
     frmEgresoComun.Tag = "N"
    frmEgresoComun.Show
End Sub

Private Sub dcmbCodigo_Change()
    Dim strComparar As String
    On Error GoTo errhandler
    If clsEgr.adorec_Def.RecordCount > 0 Then
        clsEgr.adorec_Def.MoveFirst
    End If
    strComparar = "egr_com_codigo = '" & dcmbCodigo.Text & "'"
    clsEgr.adorec_Def.Find strComparar
    dcmbCodigo.Tag = "A"
    If clsEgr.adorec_Def.EOF = False Then
        'pone valores en los combos
        dcmbDescripcion.Text = clsEgr.adorec_Def("descripcion")
        dcmbDescripcion.BoundText = clsEgr.adorec_Def("egr_com_codigo")
        txtCuenta.Text = clsEgr.adorec_Def("cta_ban_numero")
        txtBanco.Text = clsEgr.adorec_Def("ban_nombre")
        txtBanco.Tag = clsEgr.adorec_Def("ban_codigo")
        txtDescripcion.Text = clsEgr.adorec_Def("egr_com_descripcion")
        txtFecha.Text = clsEgr.adorec_Def("egr_com_fecha")
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True

     For i = 1 To VSFG.Rows - 1
      'pone valores en el grid
          strSQL = " SELECT distinct det_egreso_comun.cta_codigo,ctaconta.cta_nombre ,det_egr_com_debe, det_egr_com_haber " & _
                   " FROM ((( egreso_comun INNER JOIN det_egreso_comun ON egreso_comun.egr_com_codigo=det_egreso_comun.egr_com_codigo " & _
                   "                                                   AND egreso_comun.emp_codigo=det_egreso_comun.emp_codigo) " & _
                   "                       INNER JOIN ctaconta ON det_egreso_comun.cta_codigo= ctaconta.cta_codigo " & _
                   "                                           AND det_egreso_comun.emp_codigo= ctaconta.emp_codigo) " & _
                   "                       INNER JOIN cta_banco ON egreso_comun.cta_ban_numero=cta_banco.cta_ban_numero  " & _
                   "                                            AND egreso_comun.ban_codigo=cta_banco.ban_codigo " & _
                   "                                            AND egreso_comun.emp_codigo=cta_banco.emp_codigo) " & _
                   " WHERE egreso_comun.egr_com_codigo = '" & dcmbCodigo & "' AND egreso_comun.emp_codigo = '" & strEmpresa & "'" & _
                   " ORDER BY iff(det_egreso_comun.cta_codigo=cta_banco.cta_ban_ctaconta,0,1)"
          clsDet.Ejecutar strSQL
      Set VSFG.DataSource = clsDet.adorec_Def.DataSource
      PonerBotones
      Next i
      CalcuTotal
   Else
        If clsEgr.adorec_Def.RecordCount = 0 Then
            Set dcmbCodigo.RowSource = Nothing
            Set dcmbDescripcion.RowSource = Nothing
        End If
      dcmbDescripcion.Text = ""
      txtCuenta.Text = ""
      txtBanco.Text = ""
      txtBanco.Tag = ""
      txtDescripcion.Text = ""
      txtFecha = ""
      txtTotalDebe.Text = 0
      txtTotalHaber.Text = 0
      Set VSFG.DataSource = Nothing
      VSFG.Clear flexClearScrollable
      VSFG.Rows = 2
      cmdModificar.Enabled = False
      cmdEliminar.Enabled = False
  End If
  dcmbCodigo.Tag = ""
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

Private Sub dcmbDescripcion_Change()
    
    If dcmbCodigo.Tag <> "A" Then
        If dcmbDescripcion.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbDescripcion.BoundText
        End If
    End If
End Sub

Private Sub dcmbDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbDescripcion.BoundText
    End If
End Sub

Private Sub dcmbdescripcion_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    dcmbCodigo.Text = dcmbDescripcion.BoundText
End Sub

Private Sub Form_Activate()
 Dim strComparar As String
    
    'On Error GoTo errhandler
        'Realiza la consulta para saber los egresos comunes del sistema
 
    
    strSQL = " SELECT egr_com_codigo, egr_com_descripcion, CONCAT(SUBSTRING(egr_com_descripcion,1,20),'...') as descripcion,egr_com_fecha, egreso_comun.ban_codigo, cta_ban_numero,ban_nombre " & _
             " FROM egreso_comun INNER JOIN banco ON egreso_comun.ban_codigo=banco.ban_codigo " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY egr_com_codigo"
    clsEgr.Ejecutar strSQL
    If Not clsEgr.adorec_Def.EOF Then
        Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
        dcmbCodigo.ListField = "egr_com_codigo"
        dcmbCodigo.Text = clsEgr.adorec_Def("egr_com_codigo")
        Set dcmbDescripcion.RowSource = clsEgr.adorec_Def.DataSource
        dcmbDescripcion.ListField = "descripcion"
        dcmbDescripcion.BoundColumn = "egr_com_codigo"
        banco = "banco_codigo"
        dcmbCodigo_Change
    Else
        Set dcmbCodigo.RowSource = Nothing
        dcmbCodigo = ""
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
    
    VSFG.Editable = flexEDNone
    
    'Inicializa las clases para hacer distintas consultas
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    'Realiza la consulta para saber los códigos de los egresos comunes

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

Private Sub txtTotalDebe_Change()
    txtTotalDebe = FormatoD2(txtTotalDebe)
End Sub

Private Sub txtTotalHaber_Change()
    txtTotalHaber = FormatoD2(txtTotalHaber)
End Sub
