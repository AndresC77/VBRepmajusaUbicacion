VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmConExiPrd 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Existencias"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConExiPrd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7470
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Productos"
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
      Left            =   128
      TabIndex        =   2
      Top             =   480
      Width           =   7200
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   6975
         _cx             =   12303
         _cy             =   1085
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConExiPrd.frx":030A
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
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2948
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DcmbBodega 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmConExiPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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
Private clsCon_Def2 As New clsConsulta
Private strSqlPrd As String
Private strSqlPrd2 As String

Private Sub DcmbBodega_Change()
    strSqlPrd = " SELECT producto.prd_codigo, sum(existencia.exi_cantidad) as exi_cantidad, " & _
                     " producto.prd_nombre " & _
                     " FROM producto INNER JOIN existencia " & _
                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND dep_codigo='" & dcmbBodega.BoundText & "' " & _
                     " GROUP BY producto.prd_codigo " & _
                     " ORDER BY producto.prd_codigo "
    strSqlPrd2 = " SELECT producto.prd_codigo, sum(existencia.exi_cantidad) as exi_cantidad, " & _
                     " producto.prd_nombre " & _
                     " FROM producto INNER JOIN existencia " & _
                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND dep_codigo='" & dcmbBodega.BoundText & "' " & _
                     " GROUP BY producto.prd_codigo " & _
                     " ORDER BY producto.prd_nombre "
    clsCon_Def.Ejecutar (strSqlPrd)
    clsCon_Def2.Ejecutar (strSqlPrd2)
    VSFG.ColComboList(0) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*prd_codigo, prd_nombre,exi_cantidad", "prd_codigo")
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsCon_Def2.adorec_Def, "*prd_codigo, prd_nombre,exi_cantidad", "prd_codigo")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Def2 = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
' Actualiza la lista de zonas al volver al formulario
    clsCon_Def.Actualizar
    VSFG.ColComboList(0) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*prd_codigo, prd_nombre,exi_cantidad", "prd_codigo")
    clsCon_Def2.Actualizar
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsCon_Def2.adorec_Def, "prd_codigo,*prd_nombre,exi_cantidad", "prd_codigo")
End Sub

Private Sub Form_Load()
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Def2.Inicializar AdoConn, AdoConnMaster
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    strSqlPrd = " SELECT dep_codigo, dep_nombre " & _
                     " FROM deposito " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " ORDER BY dep_nombre "
    clsCon_Def.Ejecutar (strSqlPrd)
    Set dcmbBodega.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbBodega.ListField = "dep_nombre"
    dcmbBodega.BoundColumn = "dep_codigo"
    dcmbBodega = clsCon_Def.adorec_Def("dep_nombre")
    strSqlPrd = " SELECT producto.prd_codigo, sum(existencia.exi_cantidad) as exi_cantidad, " & _
                     " producto.prd_nombre " & _
                     " FROM producto INNER JOIN existencia " & _
                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND dep_codigo='" & dcmbBodega.BoundText & "' " & _
                     " GROUP BY producto.prd_codigo " & _
                     " ORDER BY producto.prd_codigo "
    strSqlPrd2 = " SELECT producto.prd_codigo, sum(existencia.exi_cantidad) as exi_cantidad, " & _
                     " producto.prd_nombre " & _
                     " FROM producto INNER JOIN existencia " & _
                     " ON producto.prd_codigo=existencia.prd_codigo AND producto.emp_codigo=existencia.emp_codigo " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_baja=0 AND dep_codigo='" & dcmbBodega.BoundText & "' " & _
                     " GROUP BY producto.prd_codigo " & _
                     " ORDER BY producto.prd_nombre "
    clsCon_Def.Ejecutar (strSqlPrd)
    clsCon_Def2.Ejecutar (strSqlPrd2)
    VSFG.ColComboList(0) = VSFG.BuildComboList(clsCon_Def.adorec_Def, "*prd_codigo, prd_nombre,exi_cantidad", "prd_codigo")
    VSFG.ColComboList(1) = VSFG.BuildComboList(clsCon_Def2.adorec_Def, "*prd_codigo, prd_nombre,exi_cantidad", "prd_codigo")
    
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
        clsCon_Def.adorec_Def.MoveFirst
        clsCon_Def.Filtrar "prd_codigo='" & VSFG.TextMatrix(1, 0) & "'"
        If Not clsCon_Def.adorec_Def.EOF Then
            VSFG.TextMatrix(1, 1) = clsCon_Def.adorec_Def("prd_nombre")
            VSFG.TextMatrix(1, 2) = clsCon_Def.adorec_Def("exi_cantidad")
        End If
    ElseIf Col = 1 Then
        clsCon_Def2.adorec_Def.MoveFirst
        clsCon_Def2.Filtrar "prd_codigo='" & VSFG.TextMatrix(1, 1) & "'"
        If Not clsCon_Def2.adorec_Def.EOF Then
            VSFG.TextMatrix(1, 0) = clsCon_Def2.adorec_Def("prd_codigo")
            VSFG.TextMatrix(1, 2) = clsCon_Def2.adorec_Def("exi_cantidad")
        End If
    End If
End Sub
