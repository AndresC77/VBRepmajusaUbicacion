VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegCostosServ 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Pedidos Enviados a Bodega"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "frmRegCostosServ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11145
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   10815
      Begin VB.CheckBox chkFiltroPersona 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Persona"
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
         Left            =   7440
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroLinea 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Linea"
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
         Left            =   3960
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroMarca 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Marca"
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
         Left            =   3960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Left            =   7800
         TabIndex        =   21
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmRegCostosServ.frx":030A
            Left            =   1320
            List            =   "frmRegCostosServ.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
            Enabled         =   0   'False
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
            Left            =   480
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   480
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106168323
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106168323
            CurrentDate     =   37463
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   840
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo dcmbMarca 
         Height          =   330
         Left            =   3960
         TabIndex        =   23
         Top             =   720
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo dcmbLinea 
         Height          =   330
         Left            =   3960
         TabIndex        =   26
         Top             =   1560
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo dcmbPersona 
         Height          =   330
         Left            =   7440
         TabIndex        =   29
         Top             =   720
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblPersona 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   495
         Width           =   3105
      End
      Begin VB.Label lblLinea 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Líneas"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   1335
         Width           =   3105
      End
      Begin VB.Label lblMarca 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marcas"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   495
         Width           =   3105
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle de Factura"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   10815
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1935
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   10395
         _cx             =   18336
         _cy             =   3413
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegCostosServ.frx":039E
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
      Begin VB.Label LblDetalle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de la Factura Nº"
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
         Left            =   225
         TabIndex        =   7
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Facturas"
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
      Height          =   2775
      Left            =   105
      TabIndex        =   5
      Top             =   2280
      Width           =   10815
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   10545
         _cx             =   18600
         _cy             =   4260
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegCostosServ.frx":0521
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6525
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3165
      TabIndex        =   2
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   4845
      TabIndex        =   3
      Top             =   8160
      Width           =   1455
   End
End
Attribute VB_Name = "frmRegCostosServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private strSql As String
Private clsPedidosV As New clsConsulta
Private clsPed As New clsConsulta
Private clsExiPrd As New clsConsulta
Private clsSql As New clsConsulta
Private clsPedidos As New clsConsulta
Private intDato As Variant
Private codCot As Double, numPeds As Long
Private banTm As Integer
Private FechaI As Variant
Private FechaF As Variant


Private Sub cargar()
    
    strSql = " SELECT DISTINCT egreso.egr_codigo,egr_fecha,per_codigo,egr_subtotal," & _
             " egr_subtotal_o,egr_dcto,egr_impuesto,egr_total,egr_observacion,IF(det_egr_costo=0,1,0) AS por_costear " & _
             " FROM egreso " & _
             " INNER JOIN det_egreso " & _
             " ON egreso.emp_codigo=det_egreso.emp_codigo " & _
             " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
             " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
             " INNER JOIN producto " & _
             " ON producto.emp_codigo=det_egreso.emp_codigo " & _
             " AND producto.prd_codigo=det_egreso.prd_codigo " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.tip_egr_codigo='FAC' " & _
             " AND det_egreso.prd_codigo LIKE 'PR-%' " & _
             " AND egr_anulado=0 "
    
    If chkFiltroFecha.value = 1 Then
        If Option1.value = True Then
            CambiarFecha
            strSql = strSql & " AND egr_fecha BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
        ElseIf Option2.value = True Then
            strSql = strSql & " AND egr_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
        End If
    End If
    
    If chkFiltroMarca.value = 1 Then
        strSql = strSql & " AND mar_codigo = '" & dcmbMarca.BoundText & "'"
    End If
    If chkFiltroLinea.value = 1 Then
        strSql = strSql & " AND lin_codigo = '" & dcmbLinea.BoundText & "'"
    End If
    If chkFiltroPersona.value = 1 Then
        strSql = strSql & " AND per_codigo = '" & dcmbPersona.BoundText & "'"
    End If
    strSql = strSql & " ORDER BY egr_codigo,egr_fecha DESC,per_codigo "
    clsSql.Ejecutar strSql
    Set VSFGPeds.DataSource = clsSql.adorec_Def.DataSource
    
    strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre) as nom " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSql
    VSFGPeds.ColComboList(2) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "*nom", "per_codigo")
    cmdLimpiar_Click
    CargarDetalle
End Sub



Private Sub chkFechas_Click()
     If chkFechas.value = 1 Then
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Fecha2.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1", "yyyy-mm-dd")
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal, "yyyy-mm-dd")
        DiaFinal = DiaFinal - 1
    Wend
    cmdBuscar.Enabled = True
End Sub


Private Sub chkFiltroFecha_Click()
    If chkFiltroFecha.value = 1 Then
        fraFecha.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        If Option1.value = True Then
            lblMes.Enabled = True
            cmbMesI.Enabled = True
        ElseIf Option2.value = True Then
            Fecha1.Enabled = True
            
            Fecha1.Enabled = True
            chkFechas.Enabled = True
            If chkFechas.value = 1 Then
                Fecha2.Enabled = True
            End If
        End If
    Else
        fraFecha.Enabled = False
        
        Fecha2.Enabled = False
        Fecha1.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        
        Option1.Enabled = False
        Option2.Enabled = False
        lblMes.Enabled = False
        cmbMesI.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub chkFiltroLinea_Click()
     If chkFiltroLinea.value = 1 Then
        lblLinea.Enabled = True
        dcmbLinea.Enabled = True
    Else
        lblLinea.Enabled = False
        dcmbLinea.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub chkFiltroPersona_Click()
     If chkFiltroPersona.value = 1 Then
        lblPersona.Enabled = True
        dcmbPersona.Enabled = True
    Else
        lblPersona.Enabled = False
        dcmbPersona.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub chkFiltroMarca_Click()
    If chkFiltroMarca.value = 1 Then
        lblMarca.Enabled = True
        dcmbMarca.Enabled = True
    Else
        lblMarca.Enabled = False
        dcmbMarca.Enabled = False
    End If
    cmdBuscar.Enabled = True
End Sub

Private Sub cmbMesI_Change()
    CambiarFecha
End Sub

Private Sub cmdBuscar_Click()
    cargar
End Sub

Private Sub CmdConfirmar_Click()
    Dim i As Long, Cambio As Integer
    Cambio = 0
    If VSFG.Rows > 1 Then
        For i = 1 To VSFG.Rows - 1
            If FormatoD0(VSFG.TextMatrix(i, 10)) = 1 Then
                strSql = " UPDATE det_egreso SET " & _
                         " det_egr_costo='" & FormatoD4(VSFG.TextMatrix(i, 7)) & "', " & _
                         " det_egr_fechamod=CURRENT_TIMESTAMP, " & _
                         " det_egr_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND tip_egr_codigo='FAC' " & _
                         " AND egr_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                         " AND prd_codigo='" & VSFG.TextMatrix(i, 2) & "' "
                clsSql.Ejecutar strSql
                Cambio = Cambio + 1
            End If
        Next i
        
        If Cambio > 0 Then
            strSql = " UPDATE egreso SET" & _
                     " egr_fechamod=CURRENT_TIMESTAMP, " & _
                     " egr_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='FAC' " & _
                     " AND egr_codigo='" & LblPedido.Caption & "' "
            clsSql.Ejecutar strSql
            VSFGPeds.TextMatrix(VSFGPeds.Row, 9) = 1
            MsgBox "Los datos han sido modificados", vbInformation, "Detalle de Factura"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidosV = Nothing
    Set clsPed = Nothing
    Set clsExiPrd = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdLimpiar_Click()
    'Limpia el contenido del grid de detalles
    VSFG.Clear 1
    VSFG.Rows = 1
    LblPedido = "-"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa los objetos de conexión con la base de datos
    clsSql.Inicializar AdoConn, AdoConnMaster
  
    
    'Carga los marcas
    strSql = " SELECT mar_codigo, COALESCE(mar_nombre,' ') as mar " & _
             " FROM marca " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set dcmbMarca.RowSource = clsSql.adorec_Def.DataSource
    dcmbMarca.ListField = "mar"
    dcmbMarca.BoundColumn = "mar_codigo"
    
    'Carga los marcas
    strSql = " SELECT lin_codigo, COALESCE(lin_nombre,' ') as mar " & _
             " FROM linea " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set dcmbLinea.RowSource = clsSql.adorec_Def.DataSource
    dcmbLinea.ListField = "mar"
    dcmbLinea.BoundColumn = "lin_codigo"
    
    'Carga los clientes
    strSql = " SELECT per_codigo, COALESCE(per_apellido,' ',per_nombre) as mar " & _
             " FROM persona " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set dcmbPersona.RowSource = clsSql.adorec_Def.DataSource
    dcmbPersona.ListField = "mar"
    dcmbPersona.BoundColumn = "per_codigo"

    'Selecciona el mes actual
    Fecha1 = Format(HoyDia, "yyyy-mm-dd")
    Fecha2 = Format(HoyDia, "yyyy-mm-dd")
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    LblPedido.Caption = ""

   
End Sub



Private Sub Option1_Click()
        If Option1.value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Fecha1.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        cmdBuscar.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
     If Option2.value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.value = 1 Then
            
            Fecha2.Enabled = True
        End If
        cmdBuscar.Enabled = True
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 7 Then
        'Verifca que solo se ingresen números en el campo de cantidad a entregar
        If Not IsNumeric(VSFG.TextMatrix(Row, 7)) And VSFG.TextMatrix(Row, 7) <> "" Then
            MsgBox "Ingrese solo números en el costo.", vbInformation, "Costo Unitario"
            VSFG.TextMatrix(Row, 7) = "0"
        Else
            VSFG.TextMatrix(Row, 7) = FormatoD4(VSFG.TextMatrix(Row, 7))
        End If
        VSFG.TextMatrix(Row, 8) = FormatoD4(VSFG.TextMatrix(Row, 7)) * FormatoD4(VSFG.TextMatrix(Row, 4))
        VSFG.TextMatrix(Row, 10) = "1"
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 4 de la cantidad a entregar
    If Col <> 7 Then
        Cancel = True
    Else
        If FormatoD4(VSFG.TextMatrix(Row, 9)) <> 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFGPeds_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        CargarDetalle
    End If
End Sub

Private Sub CargarDetalle()
    'Verifica cuando se da un doble click sobre una fila del grid de pedidos
    If VSFGPeds.Row > 0 Then
        'Consulta el detalle de un pedido específico
        strSql = " SELECT det_egreso.egr_codigo,dep_codigo,producto.prd_codigo,prd_nombre," & _
                 " det_egr_cantidad,det_egr_precio,COALESCE(det_egr_cantidad*det_egr_precio,0)," & _
                 " det_egr_costo,COALESCE(det_egr_cantidad*det_egr_costo,0),det_egr_costo,'0' as modi " & _
                 " FROM det_egreso " & _
                 " INNER JOIN producto " & _
                 " ON producto.emp_codigo=det_egreso.emp_codigo " & _
                 " AND producto.prd_codigo=det_egreso.prd_codigo " & _
                 " WHERE tip_egr_codigo='FAC' " & _
                 " AND det_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & VSFGPeds.TextMatrix(VSFGPeds.Row, 0) & "' " & _
                 " AND det_egreso.prd_codigo LIKE 'PR-%' " & _
                 " ORDER BY dep_codigo,det_egreso.prd_codigo "
        clsSql.Ejecutar (strSql)
        
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        'Muestra el número del pedido a modificar
        LblPedido.Caption = VSFGPeds.TextMatrix(VSFGPeds.Row, 0)
        
        strSql = " SELECT dep_codigo,dep_nombre " & _
                 " FROM deposito " & _
                 " WHERE emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSql
        VSFG.ColComboList(1) = VSFG.BuildComboList(clsSql.adorec_Def, "*dep_codigo", "dep_codigo")
        
    End If

End Sub
Private Sub VSFGPeds_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 9 Then
        If FormatoD0(VSFGPeds.TextMatrix(Row, Col)) = 1 Then
            VSFGPeds.Cell(flexcpBackColor, Row, 0, Row, VSFGPeds.Cols - 1) = &HC0C0FF
        Else
            VSFGPeds.Cell(flexcpBackColor, Row, 0, Row, VSFGPeds.Cols - 1) = &HFFFFFF
        End If
    End If
    
End Sub
