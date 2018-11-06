VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmValidarArchivoPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación de Pagos Automáticos"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidarArchivoPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   13395
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Aplicación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13200
      Begin VB.CheckBox chkTodos 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Todos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtTotalACobrar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   11280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CommandButton cmdConsultaCartera 
         Caption         =   "Consulta Cartera a Pagar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   2
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   5760
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   12975
         _cx             =   22886
         _cy             =   7858
         Appearance      =   1
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmValidarArchivoPagos.frx":030A
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
      Begin NEED2.uctrVSFG ucrtVSFG1 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmValidarArchivoPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private strSql As String


Private Sub chkTodos_Click()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 0) = chkTodos.Value
    Next i
    TotalSel
End Sub

Private Sub cmdConsultaCartera_Click()
    Dim i As Long
    strSql = " SELECT '0' as sel, comp_egreso.com_egr_codigo,CONCAT(per_apellido, ' ',per_nombre) as pro," & _
             " IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, com_egr_ch_valor," & _
             " com_egr_descripcion,COALESCE(ban_codigo_interbancario,'-'),COALESCE(per_cue_tipo,'-'),COALESCE(per_cue_numero,'-'),per_direccion,ciu_nombre,per_telf,per_email,comp_egreso.asi_numasiento " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.emp_codigo=persona.emp_codigo" & _
                 " AND comp_egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON " & _
                 " persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " LEFT JOIN persona_cuenta ON persona.emp_codigo=persona_cuenta.emp_codigo" & _
                 " AND persona.per_codigo=persona_cuenta.per_codigo " & _
                 " LEFT JOIN banco ON " & _
                 " persona_cuenta.ban_codigo=banco.ban_codigo " & _
                 " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND com_egr_proceso_cash=1 AND com_egr_ch_estado!='ANULADO' AND com_egr_ch_valor!=0 " & _
                 " ORDER BY CONCAT(per_apellido, ' ',per_nombre) "
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    txtTotalACobrar.Text = FormatoD2(0)
    TotalSel
End Sub
Private Sub TotalSel()
    Dim i As Long
    txtTotalACobrar.Text = 0
    For i = 1 To VSFG.Rows - 1
        If Abs(FormatoD0(VSFG.TextMatrix(i, 0))) = 1 Then
            txtTotalACobrar.Text = FormatoD2(FormatoD2(txtTotalACobrar.Text) + FormatoD2(VSFG.TextMatrix(i, 5)))
            VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
        Else
            VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbWhite
        End If
    Next i
End Sub

Private Sub cmdVerificar_Click()
    Dim i As Long
    For i = 1 To VSFG.Rows - 1
        If Abs(FormatoD0(VSFG.TextMatrix(i, 0))) = 1 Then
            strSql = " UPDATE comp_egreso  " & _
                     " SET com_egr_proceso_cash=2 " & _
                     " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND comp_egreso.com_egr_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                     " AND com_egr_proceso_cash=1 AND com_egr_ch_valor!=0 "
            clsCon_Def.Ejecutar strSql
        End If
    Next i
    MsgBox "Egresos validados." & vbNewLine & "El resto deberán anular para generar cheque y no volver a generar archivo la proxima vez.", vbInformation
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set ucrtVSFG1.VSFGControl = VSFG
    ucrtVSFG1.Inicializar False, False, False
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
        TotalSel
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub
