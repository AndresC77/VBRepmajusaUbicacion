VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerConSum 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Consumo Suministro"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmVerConSum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7650
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3875
      TabIndex        =   10
      Top             =   5600
      Width           =   1700
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   2075
      TabIndex        =   9
      Top             =   5600
      Width           =   1700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Height          =   5535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtNomArea 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   615
            Width           =   2220
         End
         Begin VB.TextBox txtDesArea 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   255
            Width           =   2235
         End
         Begin VB.TextBox txtCodArea 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00BAA892&
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
            Left            =   480
            TabIndex        =   20
            Top             =   292
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00BAA892&
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
            Left            =   480
            TabIndex        =   19
            Top             =   667
            Width           =   600
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00BAA892&
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
            Height          =   195
            Left            =   3720
            TabIndex        =   18
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Consumo Suministro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtCon_fecha 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtNum_orden 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dcmbCon_codigo 
            Height          =   330
            Left            =   960
            TabIndex        =   0
            Top             =   352
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label13 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "# Orden:"
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
            Left            =   2640
            TabIndex        =   22
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00BAA892&
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
            TabIndex        =   16
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00BAA892&
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
            Left            =   5040
            TabIndex        =   15
            Top             =   390
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Detalle Consumo Suministro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   7335
         Begin VB.TextBox txtObs 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   2520
            Width           =   7050
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2040
            Width           =   1095
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFDetalle 
            Height          =   1695
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   7095
            _cx             =   12515
            _cy             =   2990
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
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmVerConSum.frx":030A
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
            TabBehavior     =   1
            OwnerDraw       =   0
            Editable        =   1
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones:"
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
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label Label9 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   4920
            TabIndex        =   13
            Top             =   2070
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmVerConSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para visualizar los consumos de suministros realizados por areas      #
'#  esta forma es solo de visualización, no permite la edición.                 #
'#  frmVerConSum V1.0                                                           #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los consumos de Suministros de la empresa            #
'#  En esta ventana solo se puede visuallizar cualesquiera de los consumos por  #
'#  este concepto de gasto pero no se puede realizar ningún cambio.             #
'#  Se puede escoger el número de consumo o ingresar dicho número en el combo   #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    suministro : En esta tabla se consulta el codigo,nombre y descripcion     #
'#                       del suministro.                                        #
'#    det_consumo : En esta tabla se consulta el detalle de consumo             #
'#                        de Suministros.                                       #
'#    area : En esta tabla se consulta el codigo,nombre y descripcion           #
'#                       del area.                                              #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#    limpiarFxGD() : Permite borrar el flexgrid utilizado para cuando se       #
'#                    realiza un cambio de documento.                           #
'#    borrar_datos() : Permite limpiar  los datos para cuando se                #
'#                    realiza un cambio de documento.                           #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/*******************************************************************************/'

Private clsConsu As New clsConsulta
Private clsCon_det As New clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_det = Nothing
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFDetalle.Rows - 1)
        VSFDetalle.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub CalcuTotal()
   'Calcula Totales
    Dim Total As Double
    'Calcula Total
    Total = 0
    For i = 1 To VSFDetalle.Rows - 1
        VSFDetalle.TextMatrix(i, 5) = Val(VSFDetalle.TextMatrix(i, 3)) * Val(VSFDetalle.TextMatrix(i, 4))
        Total = Total + Val(VSFDetalle.TextMatrix(i, 5))
    Next i
    TxtTotal = Format(Total, "##0.00")
End Sub
Private Sub cmdNuevo_Click()
    Me.Hide
    frmConSum.Show
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub dcmbCon_codigo_Change()
    Dim i As Integer
    'Despliego los datos según el dato ingresado o seleccioado en el data combo
    If (dcmbCon_codigo.Text = "") Then
        Call borrar_datos
        Call limpiarFxGD
        Exit Sub
    End If
    
    If (CSng(dcmbCon_codigo.Text) > 99999) Then
        borrar_datos
        Call limpiarFxGD
        Exit Sub
    End If
    
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "Con_codigo = '" & dcmbCon_codigo.Text & "'", , adSearchForward
    If clsConsu.adorec_Def.EOF = False Then
        'Muestra los datos del Area tales como: Nombres, Apellidos, Descripcion.
        txtCon_fecha.Text = Format(clsConsu.adorec_Def("Con_fecha"), "yyyy-mmm-dd")
        txtNum_orden.Text = clsConsu.adorec_Def("con_numdoc")
        txtCodArea.Text = clsConsu.adorec_Def("are_codigo")
        txtNomArea.Text = clsConsu.adorec_Def("are_nombre")
        txtDesArea.Text = clsConsu.adorec_Def("are_descripcion")
        Call limpiarFxGD
        'llenar flexgrid
        strSql = " SELECT det_consumo.sum_codigo, sum_nombre,det_con_cantidad, sum_precio_prom,det_con_precio  " & _
                 " FROM det_consumo INNER JOIN suministro ON  det_consumo.emp_codigo = suministro.emp_codigo" & _
                 " AND det_consumo.sum_codigo = suministro.sum_codigo " & _
                 " WHERE det_consumo.emp_codigo = '" & strEmpresa & "' and det_consumo.Con_codigo =  " & dcmbCon_codigo.Text & " "
        clsCon_det.Ejecutar (strSql)
        
        If (clsCon_det.adorec_Def.RecordCount > 0) Then
            
            TxtTotal.Text = 0
            Set VSFDetalle.DataSource = clsCon_det.adorec_Def.DataSource
            Call PonerBotones
            Call CalcuTotal
            txtObs.Text = clsConsu.adorec_Def("con_observacion")
        Else
            Call borrar_datos
            End If
    Else
        Call limpiarFxGD
        Call borrar_datos
                
    End If
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub
    
Private Sub dcmbCon_codigo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
End If
End Sub

Private Sub Form_Activate()
clsConsu.Actualizar
  If (clsConsu.adorec_Def.RecordCount <> 0) Then
        Set dcmbCon_codigo.RowSource = clsConsu.adorec_Def.DataSource
        dcmbCon_codigo.ListField = "Con_codigo"
    End If

End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsCon_det.Inicializar AdoConn, AdoConnMaster
    
    'Ejecuta un SQL contra la base de datos
    strSql = " SELECT concat(a.con_codigo)as con_codigo, a.con_fecha,a.con_total," & _
             " a.con_observacion,a.con_numdoc,b.are_codigo, b.are_nombre," & _
             " b.are_descripcion" & _
             " FROM consumo a, area b" & _
             " WHERE  a.are_codigo = b.are_codigo and a.emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY a.con_codigo"
    clsConsu.Ejecutar (strSql)
    'Muestra los códigos de los Areaes en el combobox de códigos de Areaes
        
    If (clsConsu.adorec_Def.RecordCount = 0) Then
        MsgBox "No existe consumo de Suministros almacenados en el Sistema", vbInformation, "Sis-Admin"
        Exit Sub
    Else
        Set dcmbCon_codigo.RowSource = clsConsu.adorec_Def.DataSource
        dcmbCon_codigo.ListField = "Con_codigo"
    End If
    
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim X, Y  As Integer
    VSFDetalle.Tag = "N"
    VSFDetalle.Rows = 1
    VSFDetalle.Clear 1
    VSFDetalle.Tag = "T"
    
End Sub

Public Sub borrar_datos()
        txtNum_orden.Text = ""
        txtCon_fecha.Text = ""
        txtCodArea.Text = ""
        txtNomArea.Text = ""
        txtDesArea.Text = ""
        txtTelArea.Text = ""
        TxtTotal.Text = ""
        txtObs.Text = ""

End Sub

Private Sub VSFDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Or Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Then
Cancel = True
End If
End Sub

