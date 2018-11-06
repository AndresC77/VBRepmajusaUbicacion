VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVerListaEmbarque 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Embarque"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerListaEmbarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdAsignarPeso 
      Caption         =   "Asignar Peso"
      Height          =   375
      Left            =   6863
      TabIndex        =   34
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecibirLista 
      Caption         =   "Recibir Lista"
      Height          =   375
      Left            =   480
      TabIndex        =   33
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnviarCorreo 
      Caption         =   "Enviar Correo"
      Height          =   375
      Left            =   3743
      TabIndex        =   32
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtPeso 
      Height          =   315
      Left            =   7650
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCambiarOperador 
      Caption         =   "Cambiar Operador"
      Height          =   375
      Left            =   2183
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   623
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimirEtiqueta 
      Caption         =   "Imprimir Etiqueta"
      Height          =   375
      Left            =   5303
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtOperador 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtCliente 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1560
      Width           =   7815
   End
   Begin VB.TextBox txtContenedor 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1170
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   6360
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   2175
      Left            =   90
      TabIndex        =   17
      Top             =   2880
      Width           =   8775
      _cx             =   15478
      _cy             =   3836
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerListaEmbarque.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VB.TextBox TxtObserv 
      Height          =   525
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   10
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox txtGuia 
      Height          =   315
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprimirListado 
      Caption         =   "Imprimir Listado"
      Height          =   375
      Left            =   3743
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6863
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin NEED2.dtpFecha dtpFecha 
      Height          =   315
      Left            =   6930
      TabIndex        =   12
      Top             =   1170
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      Value           =   41836.5404166667
      Enabled         =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFFactura 
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFPedido 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFContenedor 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contenedor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG2 
      Height          =   1095
      Left            =   90
      TabIndex        =   29
      Top             =   5160
      Width           =   8775
      _cx             =   15478
      _cy             =   1931
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerListaEmbarque.frx":0425
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
      ExplorerBar     =   5
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7200
      TabIndex        =   31
      Top             =   1935
      Width           =   405
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenedor:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   1215
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pedidos:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   19
      Top             =   6375
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   1965
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "No.Guia:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4560
      TabIndex        =   14
      Top             =   1935
      Width           =   615
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00BAA892&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   6240
      TabIndex        =   13
      Top             =   1215
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   525
   End
End
Attribute VB_Name = "frmVerListaEmbarque"
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
Private strSql As String
Private clsSql As New clsConsulta


Private Sub LimpiarForm()
    VSFG.Clear flexClearScrollable
    VSFG.Rows = 1
    TxtTotal.Text = 0
End Sub

Private Sub cmdAsignarPeso_Click()
    frmAsignarPeso.Show
End Sub

Private Sub cmdCambiarOperador_Click()
    frmModListaEmbarque.txtContenedor.Text = txtContenedor.Text
    frmModListaEmbarque.cmbCourier.BoundText = txtOperador.Tag
    frmModListaEmbarque.txtGuia.Text = txtGuia.Text
    frmModListaEmbarque.Show vbModal
End Sub

Private Sub cmdConsultar_Click()
    BuscarContenedor txtFContenedor.Text, txtFPedido.Text, txtFFactura.Text
End Sub

Private Sub cmdEnviarCorreo_Click()
    Dim RepEmpaque As New frmReporte
    RepEmpaque.strNumero = txtContenedor.Text
    RepEmpaque.strReporte = "rptListaEmbarque"
    RepEmpaque.Show
    RepEmpaque.Form_Activate
    
    
    RepEmpaque.VSRpt.RenderToFile "LE" & FormatoD0(txtContenedor.Text) & ".pdf", vsrPDF
        'MsgBox "archiv generado"
On Error GoTo errhandler
    If Trim(txtCliente.Tag) <> "" Then
        EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, txtCliente.Text, Trim(txtCliente.Tag), CorreoSupervisorDeTransportes, "Lista de Embarque " & txtContenedor.Text, _
                "Estimad@" & vbNewLine & _
                txtCliente.Text & vbNewLine & _
                "Adjunto encontrará la lista de embarque despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                "Saludos Cordiales" & vbNewLine & _
                "Departamento de Despachos" & vbNewLine & _
                NombreComercial, "LE" & FormatoD0(txtContenedor.Text) & ".pdf"
        Kill "LE" & FormatoD0(txtContenedor.Text) & ".pdf"
    Else
        EnviarMail NombreComercial & " Despachos", CorreoSupervisorDeTransportes, "Supervidor de Transportes", CorreoSupervisorDeTransportes, "", "Lista de Embarque " & txtContenedor.Text & " Lider sin mails", _
                "Estimad@" & vbNewLine & _
                "El lider: " & txtCliente.Text & vbNewLine & _
                "No recibio la Lista de Embarque adjunta, despachada el " & Format(dtpFecha.Value, "yyyy-mm-dd") & "." & vbNewLine & _
                "Ya que no tiene ingresado Email" & vbNewLine & _
                "Saludos Cordiales" & vbNewLine & _
                "Departamento de Despachos" & _
                NombreComercial, "LE" & FormatoD0(txtContenedor.Text) & ".pdf"
        Kill "LE" & FormatoD0(txtContenedor.Text) & ".pdf"
    End If
    Unload RepEmpaque
    
    'Unload Me
    Exit Sub
errhandler:
    MsgBox "[" & Err.Number & "] " & Err.Description

    Unload RepEmpaque
    
    'Unload Me
End Sub

Private Sub cmdImprimirEtiqueta_Click()
    Dim RepStk As New frmReporte
    Dim RepStk2 As New frmReporte
    
    If ImpresoraEtiqueta = "" Then
        RepStk.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = RepStk.VSPrint.Device
        GuardarImpresoras
    End If
    RepStk.VSPrint.Device = ImpresoraEtiqueta
    RepStk.VSPrint.PaperWidth = 7669.292
    RepStk.VSPrint.PaperHeight = 3150.039
'    RepStk.VSPrint.PaperWidth = 7669.292
'    RepStk.VSPrint.PaperHeight = 8885.039
    RepStk.strNumero = txtContenedor.Text
    RepStk.strReporte = "rptSTKGuia"
    RepStk.Show
    RepStk.Form_Activate
'    RepStk.VSPrint.PrintDoc
'    Unload RepStk
    
''    RepStk2.VSPrint.Device = ImpresoraEtiqueta
''    RepStk2.VSPrint.PaperWidth = 7669.292
''    RepStk2.VSPrint.PaperHeight = 3150.039
'''    RepStk.VSPrint.PaperWidth = 7669.292
'''    RepStk.VSPrint.PaperHeight = 8885.039
''    RepStk2.strNumero = txtContenedor.Text
''    RepStk2.strReporte = "rptSTKGuiaBlanco"
''    RepStk2.Show
''    RepStk2.Form_Activate
''    RepStk2.VSPrint.PrintDoc
'    'Unload RepStk2
End Sub

Private Sub cmdImprimirListado_Click()
    Dim RepEmpaque As New frmReporte
    If MsgBox("Imprimir en A4?", vbQuestion + vbYesNo, "Lista de Embarque") = vbYes Then
        RepEmpaque.strNumero = txtContenedor.Text
        RepEmpaque.strReporte = "rptListaEmbarque"
        RepEmpaque.Show
        RepEmpaque.Form_Activate
'        RepEmpaque.VSPrint.PrintDoc
'        Unload RepEmpaque
    Else
        frmImpresionDirecta.strReporte = "rptListaEmbarque"
        frmImpresionDirecta.strNumero = txtContenedor.Text
        frmImpresionDirecta.Show
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.cmdImprimir_Click
        frmImpresionDirecta.CmdCerrar_Click
    End If
    
End Sub

Private Sub cmdNuevo_Click()
    frmListaEmbarque.Show
    Unload Me
End Sub

Private Sub cmdRecibirLista_Click()
    frmRecepcionListaEmbarque.Show
    frmRecepcionListaEmbarque.txtContenedor = txtContenedor
    frmRecepcionListaEmbarque.dtpFecha.Value = dtpFecha.Value
    frmRecepcionListaEmbarque.txtCliente = txtCliente
    frmRecepcionListaEmbarque.txtOperador = txtOperador
    frmRecepcionListaEmbarque.txtGuia = txtGuia
    frmRecepcionListaEmbarque.txtPeso = txtPeso
    frmRecepcionListaEmbarque.TxtObserv = TxtObserv
    
    strSql = " SELECT pedido.ped_codigo,egreso.egr_codigo,egr_fecha," & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,pd.per_direccion2,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet, " & _
                 " det_con_estado,COALESCE(det_con_fecha,'') as fecha" & _
                 " FROM contenedor INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo " & _
                 " INNER JOIN egreso ON det_contenedor.emp_codigo=egreso.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo='FAC' AND egreso.egr_anulado=0 " & _
                 " INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo='FAC' AND pedido.ped_estado=2 " & _
                 " INNER JOIN persona pd ON egreso.emp_codigo=pd.emp_codigo " & _
                 " AND egreso.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo " & _
                 " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & txtContenedor.Text & "' "
        clsSql.Ejecutar strSql
        Set frmRecepcionListaEmbarque.VSFG.DataSource = clsSql.adorec_Def.DataSource
        frmRecepcionListaEmbarque.TxtTotal.Text = frmRecepcionListaEmbarque.VSFG.Rows - 1
    Unload Me
End Sub

Private Sub Command1_Click()
'    Dim RepStk As New frmReporte
'    RepStk.VSPrint.Device = ImpresoraEtiqueta
'    RepStk.VSPrint.PaperWidth = 7669.292
'    RepStk.VSPrint.PaperHeight = 3885.039
'    RepStk.strNumero = 1
'    RepStk.strReporte = "rptSTKListaEmbarque"
'    RepStk.Show
'    RepStk.Form_Activate
'    RepStk.VSPrint.PrintDoc
'    Unload RepStk
'    Dim RepEmpaque As New frmReporte
'    RepEmpaque.strNumero = 1
'    RepEmpaque.strReporte = "rptListaEmbarque"
'    RepEmpaque.Show
'    RepEmpaque.Form_Activate
'    RepEmpaque.VSPrint.PrintDoc
'    Unload RepEmpaque
'Dim clsAux As New clsGuiaUrbano
'clsAux.Generar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtFContenedor" _
       And Screen.ActiveControl.Name <> "txtFPedido" _
       And Screen.ActiveControl.Name <> "txtFFactura" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    clsSql.Inicializar AdoConn, AdoConnMaster
'    strSql = " SELECT par_texto FROM parametro WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='IED'"
'    clsSql.Ejecutar strSql
'    ImpresoraEtiqueta = clsSql.adorec_Def("par_texto")
        
    dtpFecha.Value = HoyDia
End Sub

Private Sub txtFContenedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarContenedor UCase(txtFContenedor.Text), "", ""
        txtFContenedor.Text = ""
        txtFPedido.Text = ""
        txtFFactura.Text = ""
    End If
End Sub

Private Sub txtFPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarContenedor "", UCase(txtFPedido.Text), ""
        txtFContenedor.Text = ""
        txtFPedido.Text = ""
        txtFFactura.Text = ""
    End If
End Sub

Private Sub txtFFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarContenedor "", "", UCase(txtFFactura.Text)
        txtFContenedor.Text = ""
        txtFPedido.Text = ""
        txtFFactura.Text = ""
    End If
End Sub

Private Sub BuscarContenedor(strContenedor As String, strPedido As String, strFactura As String)
    If Trim(strContenedor) <> "" Then
        strContenedor = strContenedor
    ElseIf Trim(strPedido) <> "" Then
        strSql = " SELECT contenedor.con_codigo " & _
                 " FROM pedido INNER JOIN det_contenedor " & _
                 " ON pedido.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND pedido.ped_egr_codigo=det_contenedor.egr_codigo AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo " & _
                 " INNER JOIN contenedor " & _
                 " ON det_contenedor.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_contenedor.con_codigo=contenedor.con_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo='" & strPedido & "'" & _
                 "  "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strContenedor = clsSql.adorec_Def("con_codigo")
        Else
            strContenedor = ""
        End If
    ElseIf Trim(strFactura) <> "" Then
        strSql = " SELECT contenedor.con_codigo " & _
                 " FROM det_contenedor INNER JOIN contenedor " & _
                 " ON det_contenedor.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_contenedor.con_codigo=contenedor.con_codigo " & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_contenedor.egr_codigo='" & strFactura & "'"
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strContenedor = clsSql.adorec_Def("con_codigo")
        Else
            strContenedor = ""
        End If
    End If
    strSql = " SELECT contenedor.per_codigo,con_fecha, CONCAT(per_apellido,' ',per_nombre) as cli,per_email, " & _
             " contenedor.cou_codigo,COALESCE(cou_nombre,'') as cou_nombre, con_guia,con_peso,con_observacion " & _
             " FROM contenedor INNER JOIN persona " & _
             " ON contenedor.emp_codigo=persona.emp_codigo " & _
             " AND contenedor.per_codigo=persona.per_codigo " & _
             " AND persona.cat_p_tipo='C' " & _
             " LEFT JOIN courier " & _
             " ON contenedor.emp_codigo=courier.emp_codigo " & _
             " AND contenedor.cou_codigo=courier.cou_codigo " & _
             " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
             " AND contenedor.con_codigo='" & FormatoD0(strContenedor) & "'"
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        txtContenedor.Text = strContenedor
        dtpFecha.Value = clsSql.adorec_Def("con_fecha")
        txtCliente.Text = clsSql.adorec_Def("cli")
        txtCliente.Tag = clsSql.adorec_Def("per_email")
        txtOperador.Text = clsSql.adorec_Def("cou_nombre")
        txtOperador.Tag = clsSql.adorec_Def("cou_codigo")
        txtGuia.Text = clsSql.adorec_Def("con_guia")
        txtPeso.Text = clsSql.adorec_Def("con_peso")
        TxtObserv.Text = clsSql.adorec_Def("con_observacion")
        TxtObserv.Tag = clsSql.adorec_Def("per_codigo")
        
        strSql = " SELECT pedido.ped_codigo,egreso.egr_codigo,egr_fecha," & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,pd.per_direccion2,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet, " & _
                 " est_descripcion,det_contenedor.det_con_fecha " & _
                 " FROM contenedor INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo " & _
                 " INNER JOIN est_contenedor ON det_contenedor.det_con_estado=est_contenedor.est_codigo " & _
                 " INNER JOIN egreso ON det_contenedor.emp_codigo=egreso.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_contenedor.tip_egr_codigo AND egreso.egr_anulado=0 " & _
                 " INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona pd ON egreso.emp_codigo=pd.emp_codigo " & _
                 " AND egreso.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo " & _
                 " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strContenedor & "' "
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        
        strSql = " SELECT pd.per_codigo," & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,det_contenedor_per.det_con_per_detalle, " & _
                 " pd.per_direccion2,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet " & _
                 " FROM contenedor INNER JOIN det_contenedor_per ON contenedor.emp_codigo=det_contenedor_per.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_per.con_codigo " & _
                 " INNER JOIN persona pd ON det_contenedor_per.emp_codigo=pd.emp_codigo " & _
                 " AND det_contenedor_per.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo " & _
                 " WHERE contenedor.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strContenedor & "' "
        clsSql.Ejecutar strSql
        Set VSFG2.DataSource = clsSql.adorec_Def.DataSource
        TxtTotal.Text = VSFG.Rows - 1 + VSFG2.Rows - 1
    Else
        txtContenedor.Text = ""
        dtpFecha.Value = HoyDia
        txtCliente.Text = ""
        txtOperador.Text = ""
        txtGuia.Text = ""
        TxtObserv.Text = ""
        VSFG.Clear 1
        VSFG.Rows = 1
        VSFG2.Clear 1
        VSFG2.Rows = 1
        MsgBox "No se encuentra registro", vbInformation, "Listado de Embarque"
    End If
End Sub
