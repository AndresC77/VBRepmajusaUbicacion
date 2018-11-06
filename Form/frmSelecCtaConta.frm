VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelecCtaConta 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Cuentas Contables"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelecCtaConta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   5415
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cuentas Contables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Ordenar por"
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
         Height          =   735
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton optCodigo 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Código"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optNombre 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Nombre"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfgLista 
         Height          =   5055
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   4935
         _cx             =   8705
         _cy             =   8916
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelecCtaConta.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
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
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "S&eleccionar"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelecCtaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Categoria de Proveedor, y poder modificar o   #
'#  crear o eliminar categorias                                                 #
'#  frmSelCatProveedor V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las categorias de proveedores que al momento estan   #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  categoria o modificar o eliminar las categorias ya creadas.                 #
'#  Desde esta ventana se llama a la ventana frmCatProveedor en la que se crea  #
'#  y modifica las cateorias                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    categoria_p: En esta tabla se almacenan las nuevas cateorias, se          #
'#                 modifican los datos de las categorias y se eliminan.         #
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
Private clsCon_Def As clsConsulta
Public objEscribir As Object
Private booSel As Boolean
Public Normal As Boolean
Public Normal1 As Boolean
Public UserRow As Long
Public UserCol As Long

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
    booSel = False
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
'    booSel = True
'    If booSel = True Then
'        If Me.Tag = "UN" And vsfgLista.TextMatrix(vsfgLista.Row, 3) = 1 Then
'            MsgBox "Esa cuenta no es de ultimo nivel", vbCritical, "Cuenta Contable"
'        Else
'            objEscribir = vsfgLista.TextMatrix(vsfgLista.Row, 0)
'            Unload Me
'        End If
'    End If

     booSel = True
    If booSel = True Then
        If Me.Tag = "UN" And Abs(vsfgLista.TextMatrix(vsfgLista.Row, 3)) = 1 Then
            MsgBox "Esa cuenta no es de ultimo nivel", vbCritical, "Cuenta Contable"
        Else
            If Normal = False Then
                objEscribir.TextMatrix(UserRow, UserCol) = vsfgLista.TextMatrix(vsfgLista.Row, 0)
            Else
                If Normal1 = True Then
                    objEscribir.BoundText = vsfgLista.TextMatrix(vsfgLista.Row, 0)
                Else
                    objEscribir.Text = vsfgLista.TextMatrix(vsfgLista.Row, 0)
                End If
            End If
            'objEscribir = vsfgLista.TextMatrix(vsfgLista.Row, 0)
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
    vsfgLista_AfterDataRefresh
    booSel = False
End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' Consulta para actualizar los combos
        strSql = " SELECT cta_codigo,cta_nombre,cta_nivel,cta_subcta " & _
                 " FROM ctaconta  " & _
                 " WHERE emp_codigo='" & strEmpresa
        If optCodigo.value = True Then
            strSql = strSql & "' ORDER BY cta_codigo"
        Else
            strSql = strSql & "' ORDER BY cta_nombre"
        End If
        clsCon_Def.Ejecutar (strSql)
        Set vsfgLista.DataSource = clsCon_Def.adorec_Def
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

Private Sub optCodigo_Click()
    ActualizarConsulta
End Sub

Private Sub optNombre_Click()
    ActualizarConsulta
End Sub
Private Sub ActualizarConsulta()
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' Consulta para actualizar los combos
        strSql = " SELECT cta_codigo,cta_nombre,cta_nivel,cta_subcta " & _
                 " FROM ctaconta  " & _
                 " WHERE emp_codigo='" & strEmpresa
        If optCodigo.value = True Then
            strSql = strSql & "' ORDER BY cta_codigo"
        Else
            strSql = strSql & "' ORDER BY cta_nombre"
        End If
        clsCon_Def.Ejecutar (strSql)
        Set vsfgLista.DataSource = clsCon_Def.adorec_Def
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

Private Sub vsfgLista_AfterDataRefresh()
    Dim i As Long
    If Me.Tag = "UN" Then
        For i = 1 To vsfgLista.Rows - 1
            If vsfgLista.TextMatrix(i, 3) = 1 Then
                vsfgLista.Cell(flexcpForeColor, i, 0, i, 2) = &H999999
            Else
                vsfgLista.Cell(flexcpForeColor, i, 0, i, 2) = vbBlack
            End If
        Next i
    End If
End Sub

Private Sub vsfgLista_DblClick()
    cmdSeleccionar_Click
End Sub
