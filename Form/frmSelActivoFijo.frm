VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelActivoFijo 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activos Fijos"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelActivoFijo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   6252
      TabIndex        =   16
      Top             =   4215
      Width           =   1700
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   450
      Left            =   4492
      TabIndex        =   15
      Top             =   4215
      Width           =   1700
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   450
      Left            =   2732
      TabIndex        =   14
      Top             =   4215
      Width           =   1700
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   972
      TabIndex        =   13
      Top             =   4200
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Activos Fijos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtFechaD 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2880
         Width           =   1665
      End
      Begin VB.TextBox txtFechaA 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2520
         Width           =   1665
      End
      Begin VB.TextBox txtFechaB 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox chkBaja 
         BackColor       =   &H00DDDDDD&
         Enabled         =   0   'False
         Height          =   240
         Left            =   1200
         TabIndex        =   4
         Top             =   2167
         Width           =   300
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1365
         Width           =   1695
      End
      Begin VB.TextBox txtMarca 
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   2745
      End
      Begin VB.TextBox txtTipo 
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2745
      End
      Begin VB.TextBox txtCustodio 
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1365
         Width           =   2745
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   675
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2745
      End
      Begin VB.TextBox txtVidautil 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3600
         Width           =   1065
      End
      Begin VB.TextBox txtUbicacion 
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1755
         Width           =   2745
      End
      Begin VB.TextBox txtValor_dep 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1755
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFDep 
         Height          =   1455
         Left            =   3600
         TabIndex        =   11
         Top             =   2160
         Width           =   4335
         _cx             =   7646
         _cy             =   2566
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelActivoFijo.frx":030A
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
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDDDDD&
         Caption         =   "Fecha de:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Adquisición:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   5760
         TabIndex        =   30
         Top             =   3705
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Años"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2280
         TabIndex        =   29
         Top             =   3705
         Width           =   390
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "De Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   26
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   25
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Vida Útil:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   3705
         Width           =   630
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Custodio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   22
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   21
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   405
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSelActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion del Activo_Fijo y poder modificar,                  #
'#  crear o eliminar Activo_Fijos                                               #
'#  frmSelActivoFijo V1.0                                                       #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los Activo_Fijos que al momento estan                #
'#  ingresados en el sistema. Desde esta ventana se puede crear un nuevo        #
'#  Activo_Fijo, modificar o eliminar los Activo_Fijos ya creados.              #
'#  Desde esta ventana se llama a la ventana frmActivoFijo en la que crea       #
'#  y modifica los Activo_Fijos                                                 #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    tipo_ativo_Fijo: En esta tabla se almacenan y se sacan los tipos de       #
'#               activos Fijos que se pueden asignar a los activos fijos        #
'#               con su respectivo codigo.                                      #
'#    marca_activo_Fijo: En esta tabla se almacenan y se sacan las marcas de    #
'#               de los activos fijos con sus nombres y codigos.                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private clsDet As New clsConsulta
Private clsDelete As New clsConsulta
Private clsDep_mod As New clsConsulta
Private clsDep As New clsConsulta
Private clsSql As New clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsDet = Nothing
    Set clsDelete = Nothing
    Set clsDep_mod = Nothing
    Set clsDep = Nothing
    Set clsSql = Nothing
End Sub


Private Sub cmdEliminar_Click()
    Dim strSql As String
    Mensaje = "Desea eliminar el Activo Fijo?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi"   ' Define el título.
    Respuesta = MsgBox(Mensaje, Estilo, Título)
    'Recorro el FlexGrid para almacenar los detalles del ingreso
    If Respuesta = vbYes Then
            ' Consulta para conocer si hay activo fijos en Detalle adquisicion
            strSql = " SELECT count(*) As Num " & _
                     " FROM det_adquisicion_af " & _
                     " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsDet.Ejecutar (strSql)
            
            ' Si existen activos fijos no se eliminan
            If clsDet.adorec_Def("Num") > 0 Then
                MsgBox "No Puede eliminar este activos fijos Tiene detalles de adquisicion", vbInformation, "Eliminación"
            ' Si no existen activos fijos se elimina
        '        Else
        '        ' Consulta para conocer si hay activos fijos en depreciacion de Activos Fijos
        '        strSql = " SELECT count(*) As Ing " & _
        '                 " FROM depreciacion_activo " & _
        '                 " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
        '                 " AND emp_codigo='" & strEmpresa & "'"
        '        clsCon_Def.Ejecutar (strSql)
        '        ' Si existen Activos Fijos no se elimina
        '        If clsCon_Def.adorec_Def("Ing") > 0 Then
        '            MsgBox "No Puede eliminar este Activo Fijo, esta depreciado en alguna area", vbInformation, "Eliminación"
        '                ' Si no existen Activo Fijo se elimina
                         Else
                         clsSql.Inicializar AdoConn, AdoConnMaster
                         clsDelete.Inicializar AdoConn, AdoConnMaster
                         
                         'Se elimina  Activos Fijos
                             strSql = " DELETE " & _
                                           " FROM depreciacion_activo " & _
                                           " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
                                           " AND emp_codigo='" & strEmpresa & "'"
                            clsDelete.Ejecutar (strSql), "M"
                                  
                                  strSql = " DELETE " & _
                                           " FROM activo_fijo " & _
                                           " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
                                           " AND emp_codigo='" & strEmpresa & "'"
                        clsSql.Ejecutar (strSql), "M"
                              
                        
                        MsgBox "Activo Fijo eliminado", vbInformation, "Eliminación"
        '        End If
            End If
    
            'Consulta los Activos Fijos que estan disponibles
            strSql = " SELECT activo_fijo.act_fij_codigo, activo_fijo.act_fij_nombre, marca_activo_fijo.mar_act_fij_nombre,  tipo_activo.tip_act_nombre, activo_fijo.act_fij_descripcion, activo_fijo.act_fij_vida_util,  activo_fijo.act_fij_fecha_adq,act_fij_fecha_baja, activo_fijo.act_fij_fecha_dep, activo_fijo.act_fij_fecha_baja,activo_fijo.act_fij_valor,  activo_fijo.act_fij_valor,act_fij_custodio, activo_fijo.act_fij_depreciado, activo_fijo.act_fij_ubicacion,activo_fijo.act_fij_baja" & _
                    " From activo_fijo" & _
                    " Inner join tipo_activo on activo_fijo.tip_act_codigo=tipo_activo.tip_act_codigo" & _
                    " and activo_fijo.emp_codigo=tipo_activo.emp_codigo" & _
                    " Inner join marca_activo_fijo on activo_fijo.mar_act_fij_codigo=marca_activo_fijo.mar_act_fij_codigo" & _
                    " WHERE activo_fijo.emp_codigo = '" & strEmpresa & "'" & _
                    " ORDER BY activo_fijo.act_fij_codigo"
            clsCon_Def.Ejecutar (strSql)
            If Not clsCon_Def.adorec_Def.EOF Then
            Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
            dcmbCodigo.ListField = "act_fij_codigo"
            dcmbCodigo.Text = clsCon_Def.adorec_Def("act_fij_codigo")
            Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
            dcmbNombre.ListField = "act_fij_nombre"
            dcmbNombre.BoundColumn = "act_fij_codigo"
            Else
            Set dcmbCodigo.RowSource = Nothing
            End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de un Activo_fijo, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del activo_fijo que se modificará
    Dim i As Integer
    Dim intPos As Integer
    Dim strCodAux As String
    frmActivoFijo.Show
    frmActivoFijo.txtCodigo.Text = Me.dcmbCodigo.Text
    frmActivoFijo.txtNombre.Text = Me.dcmbNombre.Text
    frmActivoFijo.txtDescripcion.Text = Me.txtDescripcion.Text
    frmActivoFijo.dcmbTipo.Text = Me.txtTipo.Text
    frmActivoFijo.dcmbMarca.Text = Me.txtMarca.Text
    frmActivoFijo.chkBaja.value = Me.chkBaja.value
    frmActivoFijo.txtValor.Text = Me.txtValor.Text
    frmActivoFijo.txtValor_dep.Text = Me.txtValor_dep.Text
    frmActivoFijo.cmbAñoA.Text = Year(txtFechaA)
    frmActivoFijo.cmbMesA.Text = Month(txtFechaA)
    frmActivoFijo.cmbDiaA.Text = Day(txtFechaA)
    frmActivoFijo.cmbAñoD.Text = Year(txtFechaD)
    frmActivoFijo.cmbMesD.Text = Month(txtFechaD)
    frmActivoFijo.cmbDiaD.Text = Day(txtFechaD)
    frmActivoFijo.cmbAñoB.Text = Year(txtFechaB)
    frmActivoFijo.cmbMesB.Text = Month(txtFechaB)
    frmActivoFijo.cmbDiaB.Text = Day(txtFechaB)
    frmActivoFijo.dcmbTipo.Text = Me.txtTipo.Text
    frmActivoFijo.dcmbMarca.Text = Me.txtMarca.Text
    frmActivoFijo.txtCustodio.Text = Me.txtCustodio.Text
    frmActivoFijo.txtUbicacion.Text = Me.txtUbicacion.Text
    frmActivoFijo.txtVidautil.Text = Me.txtVidautil.Text
    frmActivoFijo.Tag = "M"
        
End Sub

Private Sub cmdnuevo_Click()
' Crea un nuevo producto, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nuevo producto
    frmActivoFijo.Show
    frmActivoFijo.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
'Chequea el activo fijo seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
        If dcmbCodigo.Text = "" Then
            Call Borrar_Datos
            Call limpiarFxGD
            Exit Sub
        End If
        clsCon_Def.Actualizar
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = " act_fij_codigo = '" & dcmbCodigo.Text & "' "
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("act_fij_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            txtDescripcion.Text = clsCon_Def.adorec_Def("act_fij_descripcion")
            txtValor.Text = clsCon_Def.adorec_Def("act_fij_valor")
            txtValor_dep.Text = clsCon_Def.adorec_Def("act_fij_depreciado")
            chkBaja.value = clsCon_Def.adorec_Def("act_fij_baja")
            txtFechaA.Text = clsCon_Def.adorec_Def("act_fij_fecha_adq")
            txtFechaB.Text = clsCon_Def.adorec_Def("act_fij_fecha_baja")
            txtFechaD.Text = clsCon_Def.adorec_Def("act_fij_fecha_dep")
            txtTipo.Text = clsCon_Def.adorec_Def("tip_act_nombre")
            txtMarca.Text = clsCon_Def.adorec_Def("mar_act_fij_nombre")
            txtCustodio.Text = clsCon_Def.adorec_Def("act_fij_custodio")
            txtUbicacion.Text = clsCon_Def.adorec_Def("act_fij_ubicacion")
            txtVidautil.Text = clsCon_Def.adorec_Def("act_fij_vida_util")
            
            cmdNuevo.Enabled = True
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            Call Borrar_Datos
            Call limpiarFxGD
        End If
        dcmbCodigo.Tag = ""
        
        'llenar flexgrid
        strSql = " SELECT area.are_codigo,area.are_nombre,depreciacion_activo.dep_act_porcentaje  " & _
                 " FROM area INNER JOIN depreciacion_activo ON area.emp_codigo =depreciacion_activo.emp_codigo AND area.are_codigo =depreciacion_activo.are_codigo " & _
                 " WHERE area.emp_codigo = '" & strEmpresa & "' AND act_fij_codigo =  '" & dcmbCodigo.Text & "' " & _
                 " ORDER BY area.are_nombre"
        clsDep_mod.Ejecutar (strSql)
        
        If (clsDep_mod.adorec_Def.RecordCount > 0) Then
            txtTotal.Text = 0
            Set VSFDep.DataSource = clsDep_mod.adorec_Def.DataSource
            Call CalcuTotal
        Else
           Call limpiarFxGD
           txtTotal.Text = " "
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

Private Sub dcmbNombre_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbNombre.Text = "" Then
        Call Borrar_Datos
    End If
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub Form_Activate()
'Actualiza la lista de activos fijos al volver al formulario
    clsCon_Def.Actualizar
    dcmbCodigo.ListField = "act_fij_codigo"
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "act_fij_nombre"
    dcmbNombre.BoundColumn = "act_fij_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource

End Sub

Private Sub Form_Load()
'Centra esta forma dentro de la forma MDI
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        clsDet.Inicializar AdoConn, AdoConnMaster
        clsDep.Inicializar AdoConn, AdoConnMaster
        clsDep_mod.Inicializar AdoConn, AdoConnMaster
        
    strSql = " SELECT activo_fijo.act_fij_codigo, activo_fijo.act_fij_nombre, marca_activo_fijo.mar_act_fij_nombre,  tipo_activo.tip_act_nombre, activo_fijo.act_fij_descripcion, activo_fijo.act_fij_vida_util,  activo_fijo.act_fij_fecha_adq,act_fij_fecha_baja, activo_fijo.act_fij_fecha_dep, activo_fijo.act_fij_fecha_baja,activo_fijo.act_fij_valor,  activo_fijo.act_fij_valor,act_fij_custodio, activo_fijo.act_fij_depreciado, activo_fijo.act_fij_ubicacion,activo_fijo.act_fij_baja" & _
        " From activo_fijo" & _
        " Inner join tipo_activo on activo_fijo.tip_act_codigo=tipo_activo.tip_act_codigo" & _
        " and activo_fijo.emp_codigo=tipo_activo.emp_codigo" & _
        " Inner join marca_activo_fijo on activo_fijo.mar_act_fij_codigo=marca_activo_fijo.mar_act_fij_codigo" & _
        " WHERE activo_fijo.emp_codigo = '" & strEmpresa & "'" & _
        " ORDER BY activo_fijo.act_fij_codigo"
                     
                     
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.EOF = False Then
            Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbCodigo.ListField = "act_fij_codigo"
            Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbNombre.ListField = "act_fij_nombre"
                dcmbNombre.BoundColumn = "act_fij_codigo"
        Else
            Set dcmbCodigo.RowSource = Nothing
            Set dcmbNombre.RowSource = Nothing
        End If
        
        cmdNuevo.Enabled = True
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
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
        SendKeys "{TAB}"
    End If
End Sub
Public Sub Borrar_Datos()
        
            dcmbNombre.Text = ""
            txtDescripcion.Text = ""
            txtValor.Text = ""
            txtValor_dep.Text = ""
            chkBaja.value = 0
            txtFechaA.Text = ""
            txtFechaB.Text = ""
            txtFechaD.Text = ""
            txtTipo.Text = ""
            txtMarca.Text = ""
            txtCustodio.Text = ""
            txtUbicacion.Text = ""
            txtVidautil.Text = ""
        
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
End Sub



Private Sub txtValor_Change()
If (txtValor.Text <> "") Then
    txtValor.Text = FormatoD2(txtValor.Text)
End If
End Sub

Private Sub txtValor_dep_Change()
If (txtValor_dep.Text <> "") Then
    txtValor_dep.Text = FormatoD2(txtValor_dep.Text)
End If
End Sub
Private Sub CalcuTotal()
   'Calcula total
    Dim Subtotal As Double
    Total = 0
    For i = 1 To (VSFDep.Rows - 1)
        Total = Total + Val(VSFDep.TextMatrix(i, 3))
    Next i
    txtTotal.Text = Val(Total)
End Sub
Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim X, Y  As Integer
    VSFDep.Tag = "N"
    VSFDep.Clear 1
    VSFDep.Rows = 2
    VSFDep.Tag = "T"
    
End Sub
