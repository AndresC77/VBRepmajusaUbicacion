VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmV_VerRangoComi3 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comision por Utilidad Bruta sobre Utilidad Bruta"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmV_VerRangoComi3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6855
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Rangos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   1500
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2175
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3435
         _cx             =   6059
         _cy             =   3836
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_VerRangoComi3.frx":030A
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tablas de Comisiones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   143
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VSFlex8Ctl.VSFlexGrid VSFGm 
         Height          =   1455
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   6060
         _cx             =   10689
         _cy             =   2566
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_VerRangoComi3.frx":0394
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
      Caption         =   "Salir"
      Height          =   375
      Left            =   4275
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nueva Tabla"
      Height          =   375
      Left            =   1125
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2715
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frmV_VerRangoComi3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para confirmar el stock en bodega de un pedido ya realizado con ante_ #
'#  rioridad.                                                                   #
'#  frmV_VerPedBod V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Opciones que permite:                                                       #
'#  *   En una lista se despliegan los pedidos y sus detalles para que la       #
'#      persona encargada pueda ver en mayor detalle el mismo y así poder       #
'#      confirmar la cantidad de los productos que se está pidiendo.            #
'#                                                                              #
'#  Procesos internos que maneja:                                               #
'#  *   La lista que muestra los distintos pedidos se refresca automáticamente  #
'#      cada 20 segundos para buscar un nuevo pedido generado.                  #
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los   #
'#      detalles del mismo en un segundo grid.                                  #
'#  *   Se controla que el usuario pida como máximo la cantidad de productos    #
'#      solicitada a la bodega.                                                 #
'#  *   Una vez que el pedido ha sido confirmado su estado pasa a revisado.     #
'#  *   Se pueden ver solo los pedidos que aún no están revisados o los que ya  #
'#      se han revisado el día de hoy.                                          #
'#                                                                              #
'#  Tablas que maneja:                                                          #
'#                                                                              #
'#  persona:                                                                    #
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el  #
'#      pedido que se está confirmando.                                         #
'#  *   También se extrae el nombre del vendedor asignado al pedido.            #
'#  pedido:                                                                     #
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.               #
'#  det_pedido:                                                                 #
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.      #
'#                                                                              #
'################################################################################

Private clsMaestro As New clsConsulta
Private clsDetalle As New clsConsulta
Private clsSql As New clsConsulta
Private strSql As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsMaestro = Nothing
    Set clsDetalle = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Sub cmdModificar_Click()
    frmV_RangoComi3.Tag = "M"
    frmV_RangoComi3.lonCod = VSFGm.TextMatrix(VSFGm.Row, 0)
    frmV_RangoComi3.DTPicker1.value = VSFGm.TextMatrix(VSFGm.Row, 1)
    frmV_RangoComi3.DTPicker2.value = VSFGm.TextMatrix(VSFGm.Row, 2)
    frmV_RangoComi3.txtCosto.Text = VSFGm.TextMatrix(VSFGm.Row, 3)
    frmV_RangoComi3.chkActivo.value = VSFGm.TextMatrix(VSFGm.Row, 4)
    frmV_RangoComi3.Show
End Sub

Private Sub cmdNuevo_Click()
    frmV_RangoComi3.Tag = "N"
    frmV_RangoComi3.Show
End Sub

Private Sub Form_Activate()
    'Consulta todos los pedidos que pasan a bodega para ser revisados
    strSql = " SELECT ran_com_codigo,ran_com_fecha,ran_com_fechafin,ran_com_porcentaje,ran_com_activo " & _
             " FROM rango_comi3 " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " ORDER BY ran_com_activo DESC,ran_com_codigo DESC "
    clsMaestro.Ejecutar strSql
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGm.DataSource = clsMaestro.adorec_Def.DataSource
    If VSFGm.Rows > 1 Then
    VSFGm.Select 1, 0, 1, 3
    VSFGm_AfterSelChange 0, 0, 1, 0
    End If
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
    Me.Top = ((mdiPrincipal.Height - 660 - Me.Height) / 2)
    'Inicializa los objetos de conexión con la base de datos
    clsMaestro.Inicializar AdoConn, AdoConnMaster
    clsDetalle.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        'Verifca que solo se ingresen números en el campo de cantidad a entregar
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en la cantidad.", vbInformation, "Cantidad"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
        'Verifica que no se esté pidiendo más productos de los que hay en existencia
        If Val(VSFG.TextMatrix(Row, 4)) > Val(VSFG.TextMatrix(Row, 5)) Then
            If VSFG.TextMatrix(Row, 5) = 0 Then
                MsgBox "No hay existencia del este producto en la bodega.", vbInformation, "Existencia"
                VSFG.TextMatrix(Row, 4) = 0
            Else
                MsgBox "Solo hay diponible " & VSFG.TextMatrix(Row, 5) & " unidades de este producto en esta bodega.", vbInformation, "Cantidad"
                VSFG.TextMatrix(Row, 4) = VSFG.TextMatrix(Row, 5)
            End If
        End If
        'Verifica que no se pidan más productos de los pedidos
        If Val(VSFG.TextMatrix(Row, 4)) > Val(VSFG.TextMatrix(Row, 3)) Then
            MsgBox "Solo puede entregar " & VSFG.TextMatrix(Row, 3) & " unidades del producto.", vbInformation, "Entregar"
            VSFG.TextMatrix(Row, 4) = VSFG.TextMatrix(Row, 3)
        End If
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 4 de la cantidad a entregar
    If Col <> 4 Then
        Cancel = True
    Else
        'Captura el valor actual de la celda de cantidad a entregar
        intDato = VSFG.TextMatrix(Row, Col)
    End If
End Sub

Private Sub VSFGm_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If VSFGm.Row > 0 Then
        'Consulta el detalle de un pedido específico
        strSql = " SELECT det_com_ran_inferior,det_com_ran_superior,det_com_ran_porcentaje " & _
                 " FROM det_rango_comi3 " & _
                 " Where emp_codigo='" & strEmpresa & "'" & _
                 " AND ran_com_codigo='" & VSFGm.TextMatrix(NewRowSel, 0) & "'" & _
                 " ORDER BY det_com_ran_inferior "
        clsDetalle.Ejecutar strSql
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsDetalle.adorec_Def.DataSource
    End If
End Sub
