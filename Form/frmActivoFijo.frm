VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmActivoFijo 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Activos Fijos"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActivoFijo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4557
      TabIndex        =   23
      Top             =   4440
      Width           =   1700
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2757
      TabIndex        =   22
      Top             =   4440
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
      Height          =   4215
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   3105
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   2745
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkBaja 
         BackColor       =   &H00DDDDDD&
         Height          =   240
         Left            =   1200
         TabIndex        =   4
         Top             =   2190
         Width           =   300
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtValor_dep 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cmbDiaD 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":030A
         Left            =   2880
         List            =   "frmActivoFijo.frx":036B
         TabIndex        =   19
         Text            =   "cmbDiaD"
         Top             =   3000
         Width           =   795
      End
      Begin VB.ComboBox cmbMesD 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":03E2
         Left            =   2040
         List            =   "frmActivoFijo.frx":040D
         TabIndex        =   18
         Text            =   "cmbMesD"
         Top             =   3000
         Width           =   795
      End
      Begin VB.ComboBox cmbAñoD 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":044D
         Left            =   1200
         List            =   "frmActivoFijo.frx":04AE
         TabIndex        =   17
         Text            =   "cmbAñoD"
         Top             =   3000
         Width           =   795
      End
      Begin VB.ComboBox cmbDiaB 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":056C
         Left            =   2880
         List            =   "frmActivoFijo.frx":05CD
         TabIndex        =   10
         Text            =   "cmbDiaB"
         Top             =   3360
         Width           =   795
      End
      Begin VB.ComboBox cmbMesB 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":0644
         Left            =   2040
         List            =   "frmActivoFijo.frx":066F
         TabIndex        =   9
         Text            =   "cmbMesB"
         Top             =   3360
         Width           =   795
      End
      Begin VB.ComboBox cmbAñoB 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":06AF
         Left            =   1200
         List            =   "frmActivoFijo.frx":0710
         TabIndex        =   8
         Text            =   "cmbAñoB"
         Top             =   3360
         Width           =   795
      End
      Begin VB.ComboBox cmbDiaA 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":07CE
         Left            =   2880
         List            =   "frmActivoFijo.frx":082F
         TabIndex        =   7
         Text            =   "cmbDiaA"
         Top             =   2640
         Width           =   795
      End
      Begin VB.ComboBox cmbMesA 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":08A6
         Left            =   2040
         List            =   "frmActivoFijo.frx":08D1
         TabIndex        =   6
         Text            =   "cmbMesA"
         Top             =   2640
         Width           =   795
      End
      Begin VB.ComboBox cmbAñoA 
         Height          =   330
         ItemData        =   "frmActivoFijo.frx":0911
         Left            =   1200
         List            =   "frmActivoFijo.frx":0972
         TabIndex        =   5
         Text            =   "cmbAñoA"
         Top             =   2640
         Width           =   795
      End
      Begin VB.TextBox txtUbicacion 
         Height          =   315
         Left            =   5280
         TabIndex        =   15
         Top             =   1800
         Width           =   2745
      End
      Begin VB.TextBox txtVidaUtil 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   3720
         Width           =   1065
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   675
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2745
      End
      Begin VB.TextBox txtCustodio 
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   1440
         Width           =   2745
      End
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   330
         Left            =   5280
         TabIndex        =   12
         Top             =   720
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFDep 
         Height          =   1455
         Left            =   4080
         TabIndex        =   20
         Top             =   2160
         Width           =   4335
         _cx             =   7646
         _cy             =   2566
         _ConvInfo       =   1
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmActivoFijo.frx":0A30
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
      End
      Begin MSDataListLib.DataCombo dcmbMarca 
         Height          =   330
         Left            =   5280
         TabIndex        =   13
         Top             =   1080
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   4680
         Picture         =   "frmActivoFijo.frx":0AC2
         Top             =   3720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   4320
         Picture         =   "frmActivoFijo.frx":0BEE
         Top             =   3720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   6120
         TabIndex        =   40
         Top             =   3645
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Años"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2400
         TabIndex        =   39
         Top             =   3765
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   38
         Top             =   1770
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   3045
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   3405
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Adquisición:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   2745
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   34
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Custodio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   33
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Vida Útil:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   3750
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   30
         Top             =   765
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   28
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   26
         Top             =   330
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "De Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   2190
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Activo Fijos con los que trabajará la    #
'#  empresa.                                                                    #
'#  frmActivoFijo V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de los Activo Fijos.                #
'#  Permitirá almacenar en la base de datos nuevos Activo Fijos y modificar     #
'#  sus datos, esto dependiendo de la propiedad Tag, la cual se cambiará en la  #
'#  ventana frmSelActivo Fijo y desde esta se llamará a esta ventana.           #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     Activo_Fijo: En esta tabla se almacenan los nuevos Activo Fijos y se     #
'#               modifican los datos de estos.                                  #
'#     Marca_activo_fijo : En esta tabla se sacan las marcas a las que se       #
'#               puede asignar a los diferentes Activo Fijos.                   #
'#     Tipo_activo : En esta tabla se sacan los tipo que se pueden asignar a    #
'#               los diferentes Activo Fijos.                                   #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#               clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private clsCon_Update As New clsConsulta
Private clsCon_Area As New clsConsulta
Private clsCon_Area_mod As New clsConsulta
Private clsDep_mod As New clsConsulta
Private clsCon_Delete As New clsConsulta
Private clsCon_Insert As New clsConsulta
Dim T As Integer
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Update = Nothing
    Set clsCon_Area = Nothing
    Set clsCon_Area_mod = Nothing
    Set clsDep_mod = Nothing
    Set clsCon_Delete = Nothing
    Set clsCon_Insert = Nothing
End Sub

Private Sub cmbAceptar_Click()
    Dim strSql As String
    Dim fa As Variant
    Dim fd As Variant
    Dim fb As Variant
    ' Si todos los campos estan llenos
    If txtCodigo.Text <> "" _
        And txtDescripcion.Text <> "" _
        And txtValor.Text <> "" _
        And txtValor_dep.Text <> "" _
        And txtNombre.Text <> "" _
        And dcmbTipo.Text <> "" _
        And dcmbMarca.Text <> "" _
        And txtCustodio.Text <> "" _
        And txtUbicacion.Text <> "" _
        And txtVidaUtil.Text <> "" Then
            fa = cmbAñoA.Text + "-" + cmbMesA + "-" + cmbDiaA
            fd = cmbAñoD.Text + "-" + cmbMesD + "-" + cmbDiaD
            fb = cmbAñoB.Text + "-" + cmbMesB + "-" + cmbDiaB
        If Not IsDate(fa) Then
            MsgBox "Fecha de Adquisición incorrecta" & vbCrLf & "Verífiquela, por favor", vbExclamation, "Sis-Admin"
            cmbAñoA.SetFocus
            Exit Sub
        End If
        If Not IsDate(fd) Then
             MsgBox "Fecha de depreciación incorrecta" & vbCrLf & "Verífiquela, por favor", vbExclamation, "Sis-Admin"
            cmbAñoD.SetFocus
            Exit Sub
        End If
        If Not IsDate(fb) Then
            MsgBox "Fecha de dada de Baja incorrecta" & vbCrLf & "Verífiquela, por favor", vbExclamation, "Sis-Admin"
            cmbAñoB.SetFocus
            Exit Sub
        End If
        If Not TxtTotal.Text = 100 Then
            MsgBox "La suma de la depreciación no es el 100 %" & vbCrLf & "Verífiquela, por favor", vbExclamation, "Sis-Admin - Activo Fijo"
            VSFDep.SetFocus
            Exit Sub
        End If
         'Verifica que existan datos en el FlexGrid
        For i = 1 To VSFDep.Rows - 1
         If (VSFDep.TextMatrix(i, 1) = "" And VSFDep.TextMatrix(i, 2) = "" And VSFDep.TextMatrix(i, 3) = "") Then
                If i = 1 Then
                    MsgBox "El area no tiene depreciación", vbExclamation, "Sis-Admi - Activo Fijo"
                End If
                i = i - 1
                Exit For
         Else
         For j = 1 To VSFDep.Cols - 1
                If (VSFDep.TextMatrix(i, j) = "") Then
                    MsgBox "Dato incorrecto o vacio en la fila: " & i, vbExclamation, "SisAdmi - Adquisición"
                    Exit Sub
                End If
            Next j
         End If
        Next i
        'valido que no haga filas vacias
        band = 0
        For i = 1 To VSFDep.Rows - 1
            For j = 1 To VSFDep.Cols - 1
                If VSFDep.TextMatrix(i, j) = "" Then band = band + 1
            Next j
            If band = VSFDep.Cols - 1 Then VSFDep.RemoveItem (i)
        Next i
        'Verifica que la depreciacion sea mayor que cero en el grid
        For h = 1 To VSFDep.Rows - 1
            If (VSFDep.TextMatrix(h, 3) = 0) Then
                MsgBox "No puede ser cero el % de depreciación en la fila: " & h, vbExclamation, "Sis-Admi - Activo Fijo"
                Exit Sub
            End If
        Next h

        ' Si se esta ingresando un nuevo Activo Fijo
        If Me.Tag = "N" Then
        'verifico que no se repita el codigo del Activo Fijo
            strSql = " SELECT act_fij_codigo " & _
                     " FROM activo_fijo " & _
                     " WHERE act_fij_codigo='" & txtCodigo.Text & "' "
            On Error GoTo errhandler
                clsCon_Def.Ejecutar (strSql)
            If clsCon_Def.adorec_Def.RecordCount <= 0 Then
            
            If (VSFDep.Rows - 1 <> 0) Then ' Si existen depreciaciones, almaceno.
                Mensaje = "Existen " & VSFDep.Rows - 1 & " Area(s) en la Activo Fijo, desea guardar?" ' Define el mensaje.
                Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
                Título = "SisAdmi "   ' Define el título.
                Respuesta = MsgBox(Mensaje, Estilo, Título)

                'Recorro el FlexGrid para almacenar las Area(s)
                If Respuesta = vbYes Then
            
            
                ' Almacenamiento de los datos del nuevo Activo Fijo
                strSql = " INSERT INTO activo_fijo(act_fij_codigo,mar_act_fij_codigo, " & _
                                            "emp_codigo,tip_act_codigo," & _
                                            "act_fij_nombre,act_fij_descripcion," & _
                                            "act_fij_valor,act_fij_depreciado," & _
                                            "act_fij_vida_util,act_fij_custodio," & _
                                            "act_fij_fecha_adq,act_fij_fecha_dep," & _
                                            "act_fij_fecha_baja,act_fij_ubicacion," & _
                                            "act_fij_baja,act_fij_fechamod," & _
                                            "act_fij_usumod)" & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & dcmbMarca.BoundText & "'," & _
                 "         '" & strEmpresa & "','" & dcmbTipo.BoundText & "', " & _
                 "         '" & UCase(txtNombre.Text) & "','" & UCase(txtDescripcion.Text) & "' ," & _
                 "          " & txtValor.Text & "," & txtValor_dep.Text & ", " & _
                 "          " & txtVidaUtil.Text & ",'" & UCase(txtCustodio.Text) & "', " & _
                 "          '" & Format(fa, "yyyy-mm-dd") & "','" & Format(fd, "yyyy-mm-dd") & "'," & _
                 "          '" & Format(fb, "yyyy-mm-dd") & "','" & UCase(txtUbicacion.Text) & "'," & _
                 "          " & chkBaja.Value & ",CURRENT_TIMESTAMP, '" & strUsuario & "')"
                 
                On Error GoTo errhandler
                clsCon_Def.Ejecutar (strSql)
            
                Dim aux As Integer
                           
                    For aux = 1 To i - 1
                        strSql = " INSERT INTO depreciacion_activo(emp_codigo,act_fij_codigo,are_codigo,dep_act_porcentaje," & _
                                "                   dep_act_fechamod,dep_act_usumod)" & _
                                "           VALUES ('" & strEmpresa & "' ,'" & UCase(txtCodigo) & "', " & _
                                "              '" & VSFDep.TextMatrix(aux, 1) & "','" & VSFDep.TextMatrix(aux, 3) & "', " & _
                                "                  CURRENT_TIMESTAMP, '" & strUsuario & "')"
                        
                        On Error GoTo errhandler
                        clsCon_Insert.Ejecutar (strSql)
                    Next
                    Else
                    Exit Sub
                End If
                    
              Else
                MsgBox "El activo Fijo que ingresó, ya existe." & vbCrLf & "Por favor cambie el código", vbExclamation, "Error Activo Fijo"
                txtCodigo.SetFocus
                txtCodigo.SelStart = 0
                txtCodigo.SelLength = Len(txtCodigo)
                Exit Sub
            End If
                
                
                
            End If
        '*******************
        ' Si se esta modificando al Activo Fijo
        ElseIf Me.Tag = "M" Then
        If (VSFDep.Rows - 1 <> 0) Then ' Si existen depreciaciones, almaceno.

                Mensaje = "Existen " & VSFDep.Rows - 1 & " Area(s) en la Activo Fijo, desea guardar?" ' Define el mensaje.
                Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
                Título = "SisAdmi "   ' Define el título.
                Respuesta = MsgBox(Mensaje, Estilo, Título)
                'Recorro el FlexGrid para almacenar las Area(s)
                    If Respuesta = vbYes Then
                    
            'Almacenamiento de los cambios realizados al Activo Fijo
            strSql = " UPDATE activo_fijo " & _
                     " SET act_fij_nombre='" & UCase(txtNombre.Text) & "', " & _
                     " act_fij_descripcion='" & UCase(txtDescripcion.Text) & "', " & _
                     " act_fij_custodio='" & UCase(txtCustodio.Text) & "', " & _
                     " act_fij_valor= '" & txtValor.Text & "', " & _
                     " mar_act_fij_codigo= '" & dcmbMarca.BoundText & "', " & _
                     " tip_act_codigo= '" & dcmbTipo.BoundText & "', " & _
                     " act_fij_vida_util= '" & txtVidaUtil.Text & "', " & _
                     " act_fij_depreciado='" & txtValor_dep.Text & "', " & _
                     " act_fij_ubicacion='" & UCase(txtUbicacion.Text) & "'," & _
                     " act_fij_fecha_adq='" & Format(fa, "yyyy-mm-dd") & "', act_fij_fecha_dep='" & Format(fd, "yyyy-mm-dd") & "', " & _
                     " act_fij_fecha_baja='" & Format(fb, "yyyy-mm-dd") & "', act_fij_baja='" & chkBaja.Value & "',act_fij_fechamod=CURRENT_TIMESTAMP,act_fij_usumod='" & strUsuario & "' " & _
                     " WHERE act_fij_codigo='" & txtCodigo.Text & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            On Error GoTo errhandler
            clsCon_Update.Ejecutar (strSql)
            '*******************
            'valido que no haga filas vacias
            band = 0
            For i = 1 To VSFDep.Rows - 1
                For j = 1 To VSFDep.Cols - 1
                    If VSFDep.TextMatrix(i, j) = "" Then band = band + 1
                Next j
                If band = VSFDep.Cols - 1 Then VSFDep.RemoveItem (i)
            Next i
            'Verifica que existan datos en el FlexGrid
            For i = 1 To VSFDep.Rows - 1
                If (VSFDep.TextMatrix(i, 1) = "" And VSFDep.TextMatrix(i, 2) = "" And VSFDep.TextMatrix(i, 3) = "") Then
                    If i = 1 Then
                    MsgBox "El Activo Fijo No tiene Depreciación", vbExclamation, "SisAdmi "
                   End If
                    i = i - 1
                    Exit For
                Else
                    For j = 1 To VSFDep.Cols - 1
                        If (VSFDep.TextMatrix(i, j) = "") Then
                            MsgBox "Dato incorrecto en: " & VSFDep.TextMatrix(i, j) & " ,fila: " & i, vbExclamation, "SisAdmi "
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
                   
                    strSql = " DELETE " & _
                             " FROM depreciacion_activo " & _
                             " WHERE emp_codigo='" & strEmpresa & "'AND act_fij_codigo='" & txtCodigo & "'"
                    On Error GoTo errhandler
                    clsCon_Delete.Ejecutar (strSql)
                                       
                    For aux = 1 To i - 1
                    strSql = " INSERT INTO depreciacion_activo(emp_codigo,act_fij_codigo,are_codigo,dep_act_porcentaje," & _
                             "                               dep_act_fechamod,dep_act_usumod)" & _
                             "                       VALUES ('" & strEmpresa & "' ,'" & txtCodigo & "', " & _
                             "                          '" & VSFDep.TextMatrix(aux, 1) & "','" & VSFDep.TextMatrix(aux, 3) & "', " & _
                             "                              CURRENT_TIMESTAMP, '" & strUsuario & "')"
            
                    On Error GoTo errhandler
                    clsCon_Insert.Ejecutar (strSql)
                    Next
                    End If
                End If
                
             Else
             Exit Sub
          End If
        '*******************
        Unload Me
        Exit Sub
    Else ' Si no estan llenos todos los campos
        MsgBox "Alguno de los campos esta vacío", vbExclamation, "ERROR"
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
Private Sub CmdCancelar_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    Dim strSql As String
    clsCon_Def.Inicializar AdoConn
    clsDep_mod.Inicializar AdoConn
    
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar los datos de los Activos Fijos"
        txtCodigo.Enabled = False
            'llenar flexgrid
            strSql = " SELECT area.are_codigo,area.are_nombre,depreciacion_activo.dep_act_porcentaje  " & _
                     " FROM area INNER JOIN depreciacion_activo ON area.emp_codigo =depreciacion_activo.emp_codigo AND area.are_codigo =depreciacion_activo.are_codigo " & _
                     " WHERE area.emp_codigo = '" & strEmpresa & "' AND act_fij_codigo =  '" & txtCodigo.Text & "' "
            clsDep_mod.Ejecutar (strSql)
    
            If (clsDep_mod.adorec_Def.RecordCount > 0) Then
            T = 1
                TxtTotal.Text = 0
                Set frmActivoFijo.VSFDep.DataSource = clsDep_mod.adorec_Def.DataSource
                CalcuTotal
                T = 0
            Else
               TxtTotal.Text = " "
            End If
            
         'Consulta las Area  de activos Fijos que estan disponibles
           strSql = " SELECT are_codigo,are_nombre " & _
                    " FROM area " & _
                    " ORDER BY are_codigo "
            clsCon_Area.Ejecutar (strSql)
        
        If Not clsCon_Area.adorec_Def.EOF Then
            VSFDep.ColComboList(1) = VSFDep.BuildComboList(clsCon_Area.adorec_Def, "*are_codigo,are_nombre", "are_codigo")
            VSFDep.ColComboList(2) = VSFDep.BuildComboList(clsCon_Area.adorec_Def, "are_codigo,*are_nombre", "are_nombre")
        Else
            VSFDep.Clear 1
            VSFDep.Rows = 2
            MsgBox "No existen áreas ingresadas en el sistema!", vbInformation, "NEED"
        End If
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nuevo Activo Fijo"
        'Consulta las Area  de activos Fijos que estan disponibles
           strSql = " SELECT are_codigo,are_nombre " & _
                    " FROM area " & _
                    " ORDER BY are_codigo "
            clsCon_Area.Ejecutar (strSql)
        
        If Not clsCon_Area.adorec_Def.EOF Then
            VSFDep.ColComboList(1) = VSFDep.BuildComboList(clsCon_Area.adorec_Def, "*are_codigo,are_nombre", "are_codigo")
            VSFDep.ColComboList(2) = VSFDep.BuildComboList(clsCon_Area.adorec_Def, "are_codigo,*are_nombre", "are_nombre")
        Else
            VSFDep.Clear 1
            VSFDep.Rows = 2
            MsgBox "No existen áreas ingresadas en el sistema!", vbInformation, "NEED"
        End If
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

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    d = CStr(Day(Date))
    m = Month(Date)
    Y = CStr(Year(Date))
    
    cmbDiaA.Text = d
    cmbAñoA.Text = Y
    cmbDiaB.Text = d
    cmbAñoB.Text = Y
    cmbDiaD.Text = d
    cmbAñoD.Text = Y
    'selecciona el mes actual
    For var = 0 To 11
        If (cmbMesA.ItemData(var) = m) Then
            cmbMesA.Text = cmbMesA.List(var)
            cmbMesB.Text = cmbMesB.List(var)
            cmbMesD.Text = cmbMesD.List(var)
            Exit For
        End If
    Next var
    If txtCodigo.Text = "" Or txtNombre.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
    On Error GoTo errhandler
        clsCon_Def.Inicializar AdoConn
        clsCon_Update.Inicializar AdoConn
        clsCon_Area_mod.Inicializar AdoConn
        clsCon_Area.Inicializar AdoConn
        clsCon_Insert.Inicializar AdoConn
        clsCon_Delete.Inicializar AdoConn
         
        'Consulta las Marcas de activos Fijos que estan disponibles
           strSql = " SELECT mar_act_fij_codigo,mar_act_fij_nombre " & _
                    " FROM marca_activo_fijo " & _
                   " ORDER BY mar_act_fij_nombre "
           clsCon_Def.Ejecutar (strSql)
               Set dcmbMarca.RowSource = clsCon_Def.adorec_Def.DataSource
               dcmbMarca.ListField = "mar_act_fij_nombre"
               dcmbMarca.BoundColumn = "mar_act_fij_codigo"
       
       'Consulta los Tipos de Activo Fijo que estan disponibles
           strSql = " SELECT tip_act_codigo,tip_act_nombre " & _
                 " FROM tipo_activo " & _
                 " ORDER BY tip_act_nombre "
           clsCon_Def.Ejecutar (strSql)
           Set dcmbTipo.RowSource = clsCon_Def.adorec_Def.DataSource
           dcmbTipo.ListField = "tip_act_nombre"
           dcmbTipo.BoundColumn = "tip_act_codigo"
                
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
Private Sub txtCodigo_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub TxtCustodio_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtNombre_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtTotal_Change()
    If txtCodigo.Text = "" And txtNombre.Text = "" And TxtTotal.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub
Private Sub txtUbicacion_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtValor_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> Asc(".")) Then
                KeyAscii = 0
    End If
End Sub
Private Sub txtValor_Validate(Cancel As Boolean)
'Pone los decimales en el txt de valor
    If txtValor.Text <> "" Then
        txtValor.Text = Format(CDbl(txtValor.Text), "##0.00")
    End If
End Sub
Private Sub txtValor_dep_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtValor_dep_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) And (KeyAscii <> Asc(".")) Then
                KeyAscii = 0
    End If
End Sub
Private Sub txtValor_dep_Validate(Cancel As Boolean)
'Pone los decimales en el txt de valor depreciacion
    If txtValor_dep.Text <> "" Then
        txtValor_dep.Text = Format(CDbl(txtValor_dep.Text), "##0.00")
    End If
End Sub
Private Sub TxtVidaUtil_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub TxtVidaUtil_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
                KeyAscii = 0
    End If
End Sub
'Private Sub VSFDep_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Col = 3 Then
'    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End If
'End Sub
Private Sub VSFDep_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If VSFDep.TextMatrix(Row, 1) = "" Or VSFDep.TextMatrix(Row, 2) = "" And Row > 0 Then
    VSFDep.TextMatrix(Row, 3) = ""
End If
'Verifica que solo se ingresen números en el % de depreciacion
        If Not IsNumeric(VSFDep.TextMatrix(Row, 3)) And VSFDep.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el % de depreciación.", vbInformation, "SisAdmi - Activos Fijos"
            VSFDep.TextMatrix(Row, 3) = ""
        End If
End Sub
Private Sub VSFDep_CellChanged(ByVal Row As Long, ByVal Col As Long)
If T = 0 Then
With VSFDep
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                 clsCon_Area.Filtrar ("are_codigo = '" & .TextMatrix(Row, 1) & "'")
                     .TextMatrix(Row, 2) = clsCon_Area.adorec_Def("are_nombre")
                 clsCon_Area.QuitarFiltro
             End If
             If Col = 2 Then
                 clsCon_Area.Filtrar ("are_nombre = '" & .TextMatrix(Row, 2) & "'")
                     .TextMatrix(Row, 1) = clsCon_Area.adorec_Def("are_codigo")
                 clsCon_Area.QuitarFiltro
             End If
         End If
    End With
    'Verifica que no se ingresen dos Areas iguales en el grid
    If Row > 1 And Col = 1 Or Col = 2 Then
        With VSFDep
            For i = 1 To .Rows - 1
                For j = i + 1 To .Rows - 1
                    If .TextMatrix(i, 1) = .TextMatrix(j, 1) Then
                        MsgBox "El Area ya ha sido ingresada, ingrese un Area diferente", vbExclamation, "SisAdmi - Activos Fijos"
                        .TextMatrix(Row, 1) = ""
                        .TextMatrix(Row, 2) = ""
                    End If
                    If j >= .Rows - 1 Then
                        Exit For
                    End If
                Next j
            Next i
        End With
    End If
    CalcuTotal
End If
End Sub
Private Sub VSFDep_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    ' get cell that was clicked
    Dim r&, c&
    r = VSFDep.MouseRow
    c = VSFDep.MouseCol
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    If (c <> 0 Or r = (VSFDep.Rows) Or VSFDep.Rows = 2) Then Exit Sub
    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFDep.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFDep.Cell(flexcpLeft, r, c) + VSFDep.Cell(flexcpWidth, r, c) - X
        If d > imgBtnDn.Width Then Exit Sub
        If r > 0 Then
        ' click was on a button: do the work
        VSFDep.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Proyecto de Ventas"   ' Define el título.
        Respuesta = MsgBox(Mensaje, Estilo, Título)
        'Recorro el FlexGrid para poner números a las filas
        If Respuesta = vbYes Then
            Dim i As Integer
            VSFDep.RemoveItem (r)
            PonerBotones
            CalcuTotal
            Else
            VSFDep.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    Cancel = True
End Sub
Private Sub VSFDep_KeyDown(KeyCode As Integer, Shift As Integer)
' Hace que cuando llegue al final del grid,
' al Presionar las teclas: enter, tab, izquierda y abajo ,
' se cree otra fila y ponga los botones correspondientes
    If VSFDep.Row = VSFDep.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If (VSFDep.TextMatrix(VSFDep.Row, 1) <> "" And VSFDep.TextMatrix(VSFDep.Row, 2) <> "" And (VSFDep.TextMatrix(VSFDep.Row, 3) <> "")) Then
            VSFDep.AddItem ""
            VSFDep.TextMatrix(VSFDep.Rows - 1, 0) = VSFDep.Rows - 1
            VSFDep.Cell(flexcpPicture, (VSFDep.Rows - 1), 0) = imgBtnUp
            VSFDep.Cell(flexcpPictureAlignment, (VSFDep.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub
Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFDep.Rows - 1)
        VSFDep.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFDep.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFDep.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub
Private Sub CalcuTotal()
   'Calcula total
    Dim SubTotal As Double
    Total = 0
    For i = 1 To (VSFDep.Rows - 1)
        Total = Total + Val(VSFDep.TextMatrix(i, 3))
    Next i
    TxtTotal.Text = Val(Total)
End Sub
