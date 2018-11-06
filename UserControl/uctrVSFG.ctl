VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl uctrVSFG 
   BackColor       =   &H00DDDDDD&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   4950
   Begin VB.Frame frm 
      BackColor       =   &H00DDDDDD&
      Height          =   975
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   6375
      Begin VB.CommandButton cmdAgregar 
         Height          =   375
         Left            =   240
         Picture         =   "uctrVSFG.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Agregar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdEliminar 
         Height          =   375
         Left            =   600
         Picture         =   "uctrVSFG.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Height          =   375
         Left            =   960
         Picture         =   "uctrVSFG.ctx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Modificar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdExcel 
         Height          =   375
         Left            =   1440
         Picture         =   "uctrVSFG.ctx":0B86
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdCopiar 
         Height          =   375
         Left            =   2400
         Picture         =   "uctrVSFG.ctx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Copiar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdCortar 
         Height          =   375
         Left            =   2760
         Picture         =   "uctrVSFG.ctx":1086
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cortar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   375
         Left            =   1920
         Picture         =   "uctrVSFG.ctx":11D0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdPegar 
         Height          =   375
         Left            =   3120
         Picture         =   "uctrVSFG.ctx":131A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Pegar"
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox chkBusqueda 
         BackColor       =   &H00DDDDDD&
         Caption         =   "AutoBusqueda"
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
         Left            =   3600
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog cdArch 
         Left            =   1680
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu menMenu 
      Caption         =   "Menu"
      Begin VB.Menu menAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu menEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu menS1 
         Caption         =   "-"
      End
      Begin VB.Menu menExcel 
         Caption         =   "Exportar Excel"
      End
      Begin VB.Menu menS2 
         Caption         =   "-"
      End
      Begin VB.Menu menImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu menS3 
         Caption         =   "-"
      End
      Begin VB.Menu menCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu menCortar 
         Caption         =   "Cortar"
      End
      Begin VB.Menu menPegar 
         Caption         =   "Pegar"
      End
   End
End
Attribute VB_Name = "uctrVSFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Public VSFGControl As VSFlexGrid
Public VSFGUltimoBoton As String
Private verAgregar As Boolean
Private verEliminar As Boolean
Private verModificar As Boolean
Private verExcel As Boolean
Private verImprimir As Boolean
Private verCopiar As Boolean
Private verCortar As Boolean
Private verPegar As Boolean
Private RestriccionCopiarCortarPegar As String

Private Sub chkBusqueda_Click()
    If chkBusqueda.Value = 0 Then
        VSFGControl.AutoSearch = flexSearchNone
    Else
        VSFGControl.AutoSearch = flexSearchFromTop
    End If
End Sub

Public Sub Agregar()
    cmdAgregar_Click
End Sub

Private Sub cmdAgregar_Click()
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    chkBusqueda.Value = 0
    chkBusqueda_Click
    VSFGControl.GetSelection r1, c1, r2, c2
    If r1 = 0 And r2 = 0 Then
        VSFGControl.AddItem "", c1
        VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = -2
        VSFGControl.Cell(flexcpBackColor, c1, 1, c1, VSFGControl.Cols - 1) = &H80FFFF
    Else
        If r2 > VSFGControl.Rows - 1 Then r2 = VSFGControl.Rows - 1
        For c1 = r1 To r2
            If c1 > VSFGControl.Rows - 1 Then Exit For
            If c1 <> -1 Then
                VSFGControl.AddItem "", c1
                VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = -2
                VSFGControl.Cell(flexcpBackColor, c1, 1, c1, VSFGControl.Cols - 1) = &H80FFFF
            End If
        Next c1
    End If
    PonerNum
    VSFGUltimoBoton = "AGREGAR"
    VSFGControl.SetFocus
End Sub

Private Sub Copiar()
    VSFGControl.Copy
    VSFGUltimoBoton = "COPIAR"
    VSFGControl.SetFocus
End Sub

Private Sub Cortar()
    VSFGControl.Cut
    VSFGUltimoBoton = "CORTAR"
    VSFGControl.SetFocus
End Sub

Private Sub Pegar()
    Dim i As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    If Abs(FormatoD0(VSFGControl.TextMatrix(VSFGControl.Row, VSFGControl.Cols - 1))) > 1 Then
        VSFGControl.Paste
        VSFGControl.GetSelection r1, c1, r2, c2
        For i = r1 To r2
            VSFGControl.Row = i
            VSFGControl.Col = c1
            VSFGControl.Paste
            If VSFGControl.TextMatrix(i, VSFGControl.Cols - 1) = -2 Then
                VSFGControl.TextMatrix(i, VSFGControl.Cols - 1) = 2
            ElseIf VSFGControl.TextMatrix(i, VSFGControl.Cols - 1) = -3 Then
                VSFGControl.TextMatrix(i, VSFGControl.Cols - 1) = 3
            End If
        Next i
        VSFGUltimoBoton = "PEGAR"
        VSFGControl.SetFocus
    End If
End Sub

Private Sub cmdCopiar_Click()
    Editar 3
End Sub

Private Sub cmdCortar_Click()
    Editar 24
End Sub

Public Sub Eliminar()
    cmdEliminar_Click
End Sub

Private Sub cmdEliminar_Click()
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    VSFGControl.GetSelection r1, c1, r2, c2
    If r1 = 0 Then Exit Sub
    c2 = 0
    If r2 > VSFGControl.Rows - 1 Then r2 = VSFGControl.Rows - 1
    For c1 = r1 To r2
        If c1 > VSFGControl.Rows - 1 Then Exit For
        If c1 <> -1 Then
            If VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = 2 Or VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = -2 Then
                VSFGControl.RemoveItem c1
                c2 = 1
                r2 = r2 - 1
                c1 = c1 - 1
                If c1 >= r2 Then Exit For
            Else
                VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = 1
                VSFGControl.Cell(flexcpBackColor, c1, 1, c1, VSFGControl.Cols - 1) = &HC0C0FF
            End If
        End If
    Next c1
    If c2 = 1 Then
        PonerNum
    End If
    VSFGUltimoBoton = "ELIMINAR"
    VSFGControl.SetFocus
End Sub

Private Sub cmdExcel_Click()
    cdArch.DefaultExt = "xls"
    cdArch.Filter = "Archivos Excel (*.xls)|*.xls|Texto (delimitado por tabulaciones)|*.txt"
    cdArch.FilterIndex = 1
    cdArch.FileName = ""
    cdArch.ShowSave
    If cdArch.FileName <> "" Then
        If LCase(Right(cdArch.FileName, 3)) = "txt" Then
            VSFGControl.SaveGrid cdArch.FileName, flexFileTabText, flexXLSaveFixedCells
        Else
            VSFGControl.SaveGrid cdArch.FileName, flexFileExcel, flexXLSaveFixedCells
        End If
    End If
    VSFGUltimoBoton = "EXCEL"
    VSFGControl.SetFocus
End Sub

Private Sub cmdImprimir_Click()
    VSFGControl.PrintGrid "NEED", True
    VSFGUltimoBoton = "IMPRIMIR"
    VSFGControl.SetFocus
End Sub

Public Sub Modificar()
    cmdModificar_Click
End Sub

Private Sub cmdModificar_Click()
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    chkBusqueda.Value = 0
    chkBusqueda_Click
    VSFGControl.GetSelection r1, c1, r2, c2
    If r1 = 0 Then Exit Sub
    If r2 > VSFGControl.Rows - 1 Then r2 = VSFGControl.Rows - 1
    For c1 = r1 To r2
        If c1 > VSFGControl.Rows - 1 Then Exit For
        If c1 <> -1 Then
            If VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = 0 Or VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = 1 Then
                VSFGControl.TextMatrix(c1, VSFGControl.Cols - 1) = -3
                VSFGControl.Cell(flexcpBackColor, c1, 1, c1, VSFGControl.Cols - 1) = &HC0FFC0
            End If
        End If
    Next c1
    VSFGUltimoBoton = "MODIFICAR"
    VSFGControl.SetFocus
End Sub

Private Sub cmdPegar_Click()
    Editar 22
End Sub

Private Sub menAgregar_Click()
    cmdAgregar.SetFocus
    cmdAgregar_Click
End Sub

Private Sub menCopiar_Click()
    cmdCopiar.SetFocus
    cmdCopiar_Click
End Sub

Private Sub menCortar_Click()
    cmdCortar.SetFocus
    cmdCortar_Click
End Sub

Private Sub menEliminar_Click()
    cmdEliminar.SetFocus
    cmdEliminar_Click
End Sub

Private Sub menExcel_Click()
    cmdExcel.SetFocus
    cmdExcel_Click
End Sub

Private Sub menImprimir_Click()
    cmdImprimir.SetFocus
    cmdImprimir_Click
End Sub

Private Sub menModificar_Click()
    cmdModificar.SetFocus
    cmdModificar_Click
End Sub

Private Sub menPegar_Click()
    cmdPegar.SetFocus
    cmdPegar_Click
End Sub

Public Sub VerMenu()
    VSFGControl.Select VSFGControl.MouseRow, VSFGControl.MouseCol
    PopupMenu menMenu
End Sub

Public Sub PonerNum()
    Dim i As Long
    For i = 1 To VSFGControl.Rows - 1
        VSFGControl.TextMatrix(i, 0) = i
    Next i
End Sub

Public Sub Editar(KeyAscii As Integer)
    '0 no hay restriccion
    '9 no puede realizar
    '1 puede copiar un solo campo
    Dim booPuede As Boolean
    booPuede = False
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    If KeyAscii = 3 Then
        If Mid(RestriccionCopiarCortarPegar, 1, 1) = "0" Then
            booPuede = True
        ElseIf Mid(RestriccionCopiarCortarPegar, 1, 1) = "1" Then
            VSFGControl.GetSelection r1, c1, r2, c2
            If r1 = r2 And c1 = c2 Then
                booPuede = True
            End If
        ElseIf Mid(RestriccionCopiarCortarPegar, 1, 1) = "9" Then
            booPuede = True
        End If
        If booPuede = True And verCopiar = True Then
            Copiar
        End If
    ElseIf KeyAscii = 24 Then
        If Mid(RestriccionCopiarCortarPegar, 2, 1) = "0" Then
            booPuede = True
        ElseIf Mid(RestriccionCopiarCortarPegar, 2, 1) = "1" Then
            VSFGControl.GetSelection r1, c1, r2, c2
            If r1 = r2 And c1 = c2 Then
                booPuede = True
            End If
        ElseIf Mid(RestriccionCopiarCortarPegar, 1, 1) = "9" Then
            booPuede = True
        End If
        If booPuede = True And verCortar = True Then
            Cortar
        End If
    ElseIf KeyAscii = 22 Then
        If Mid(RestriccionCopiarCortarPegar, 3, 1) = "0" Then
            booPuede = True
        ElseIf Mid(RestriccionCopiarCortarPegar, 3, 1) = "1" Then
            VSFGControl.GetSelection r1, c1, r2, c2
            If r1 = r2 And c1 = c2 Then
                booPuede = True
            End If
        ElseIf Mid(RestriccionCopiarCortarPegar, 1, 1) = "9" Then
            booPuede = True
        End If
        If booPuede = True And verPegar = True Then
            Pegar
        End If
    End If
End Sub

Public Sub Inicializar(Optional Agregar As Boolean = True, Optional Eliminar As Boolean = True, Optional Modificar As Boolean = True, Optional Excel As Boolean = True, Optional Imprimir As Boolean = True, Optional Copiar As Boolean = True, Optional Cortar As Boolean = True, Optional Pegar As Boolean = True, Optional AutoBusqueda As Boolean = False, Optional RestriccionParaCopiarCortarPegar As String = "000")
    verAgregar = Agregar
    verEliminar = Eliminar
    verModificar = Modificar
    verExcel = Excel
    verImprimir = Imprimir
    verCopiar = Copiar
    verCortar = Cortar
    verPegar = Pegar
    RestriccionCopiarCortarPegar = RestriccionParaCopiarCortarPegar
    If AutoBusqueda = False Then
        chkBusqueda.Value = 0
    Else
        chkBusqueda.Value = 1
    End If
    chkBusqueda_Click
    UbicarBotones
End Sub

Private Sub UbicarBotones()
    cmdAgregar.Left = 240
    cmdEliminar.Left = 240 + 360 * 1
    cmdModificar.Left = 240 + 360 * 2
    cmdExcel.Left = 240 + 360 * 3 + 120 * 1
    cmdImprimir.Left = 240 + 360 * 4 + 120 * 2
    cmdCopiar.Left = 240 + 360 * 5 + 120 * 3
    cmdCortar.Left = 240 + 360 * 6 + 120 * 3
    cmdPegar.Left = 240 + 360 * 7 + 120 * 3
    chkBusqueda.Left = 240 + 360 * 8 + 120 * 4
    cmdAgregar.Visible = verAgregar
    menAgregar.Visible = verAgregar
    If verAgregar = False Then
        cmdEliminar.Left = cmdEliminar.Left - 360
        cmdModificar.Left = cmdModificar.Left - 360
        cmdExcel.Left = cmdExcel.Left - 360
        cmdImprimir.Left = cmdImprimir.Left - 360
        cmdCopiar.Left = cmdCopiar.Left - 360
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdEliminar.Visible = verEliminar
    menEliminar.Visible = verEliminar
    If verEliminar = False Then
        cmdModificar.Left = cmdModificar.Left - 360
        cmdExcel.Left = cmdExcel.Left - 360
        cmdImprimir.Left = cmdImprimir.Left - 360
        cmdCopiar.Left = cmdCopiar.Left - 360
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdModificar.Visible = verModificar
    menModificar.Visible = verModificar
    If verModificar = False Then
        cmdExcel.Left = cmdExcel.Left - 360
        cmdImprimir.Left = cmdImprimir.Left - 360
        cmdCopiar.Left = cmdCopiar.Left - 360
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdExcel.Visible = verExcel
    menExcel.Visible = verExcel
    If verExcel = False Then
        cmdImprimir.Left = cmdImprimir.Left - 360
        cmdCopiar.Left = cmdCopiar.Left - 360
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdImprimir.Visible = verImprimir
    menImprimir.Visible = verImprimir
    If verImprimir = False Then
        cmdCopiar.Left = cmdCopiar.Left - 360
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdCopiar.Visible = verCopiar
    menCopiar.Visible = verCopiar
    If verCopiar = False Then
        cmdCortar.Left = cmdCortar.Left - 360
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdCortar.Visible = verCortar
    menCortar.Visible = verCortar
    If verCortar = False Then
        cmdPegar.Left = cmdPegar.Left - 360
        chkBusqueda.Left = chkBusqueda.Left - 360
    End If
    cmdPegar.Visible = verPegar
    menPegar.Visible = verPegar

End Sub

Public Property Let BackColor(Valor As OLE_COLOR)
    frm.BackColor = Valor
    chkBusqueda.BackColor = Valor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = frm.BackColor
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Me.BackColor = PropBag.ReadProperty("BackColor", &HDDDDDD)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Me.BackColor, &HDDDDDD)

End Sub
