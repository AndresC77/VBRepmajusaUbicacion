VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerGrupos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Grupos de Productos"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerGrupos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Grupos Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   113
      TabIndex        =   8
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtNivel 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   13
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   5175
      End
      Begin MSDataGridLib.DataGrid dGrdSubGrps 
         Height          =   2175
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Listado de SubGrupos"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcmbGrupo 
         Height          =   330
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel a Revisar"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3720
         TabIndex        =   12
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   540
      End
      Begin VB.Label LblGrupo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   1125
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   2085
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3405
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   780
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4725
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmVerGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma ver grupos (frmVerGrupos.frm)                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Esta ventana muestra todas los grupos de productos que maneja una empresa.  #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  grupo:                                                                      #
'#      Tabla que contiene toda la información necesaria de todas los grupos    #
'#      de productos posibles que maneja una determinad empresa.                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#      1.  Al cargarse la forma se consultan todoslos grupos de maneja la      #
'#          empresa.                                                            #
'#      2.  Una vez que el usuario selecciona un código de grupo o su descrip_  #
'#          ción se procede a mostrar en un grid todos los subgrupos relacio_   #
'#          nados al mismo.                                                     #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#      1.  Maneja el mantenimiento de grupos de productos.                     #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#  clsConsu:                                                                   #
'#      Objeto para consultar a la base de datos todos los posiles grupos de    #
'#      productos de una empresa y desplegarlas en un comboox.                  #
'#  clsConsuGrid:                                                               #
'#      Objeto para consultar todas los subgrupos relacionados con un grupo     #
'#      seleccionado en el combobox y mostrarlos en un data grid.               #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsConsuGrid As New clsConsulta
Private strSql As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsConsuGrid = Nothing
End Sub

Private Sub cmdEliminar_Click()
Dim strSql As String
    ' Consulta para saber si hay productos dentro del grupo que se elimina
    strSql = " SELECT count(prd_codigo) as Num " & _
             " FROM producto " & _
             " WHERE gru_codigo = '" & dcmbCodigo.Text & "' " & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
          
    ' Si existen productos dentro del grupo no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este grupo ", vbInformation, "Eliminación"
    Else ' Si no hay productos en el grupo se elimina
        strSql = " DELETE " & _
                 " FROM grupo " & _
                 " WHERE gru_codigo= '" & dcmbCodigo.Text & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        MsgBox "Grupo eliminado", vbInformation, "Eliminación"
    End If
    CargaGrupo
End Sub
Private Sub CargaGrupo()
    'Consulta para actualizar los combos
    strSql = " Select gru_codigo,gru_nombre,gru_nivel,gru_descripcion " & _
             " From grupo " & _
             " Where emp_codigo = '" & strEmpresa & "' " & _
             " AND gru_nivel='" & txtNivel.Text & "' " & _
             " Order by gru_codigo"
    clsConsu.Ejecutar (strSql)
    'Muestra los códigos de los grupos en el combobox
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "gru_codigo"
    Set dcmbGrupo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbGrupo.ListField = "gru_nombre"
    dcmbGrupo.BoundColumn = "gru_codigo"
    'Muestra el primer grupo de la consulta
    If clsConsu.adorec_Def.RecordCount > 0 Then
        dcmbCodigo = clsConsu.adorec_Def(0)
    Else
        dcmbCodigo.Text = ""
        dcmbGrupo.Text = ""
    End If
End Sub
Private Sub cmdModificar_Click()
    If dcmbCodigo = "" Then
        MsgBox "Seleccione pimero un grupo.", vbInformation, "Grupo"
        Exit Sub
    End If
    frmModGrupo.MskCodigo = dcmbCodigo
    frmModGrupo.TxtGrupo = dcmbGrupo
    frmModGrupo.txtDescripcion = txtDescripcion
    frmModGrupo.Show
End Sub

Private Sub cmdNuevo_Click()
    frmNuevoGrupo.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
    'Muestra el nombre relacionado con el código de la cuenta en el momento de seleccionar otra cuenta del combobox
    If clsConsu.adorec_Def.RecordCount > 0 Then
        clsConsu.adorec_Def.MoveFirst
        clsConsu.adorec_Def.Find "gru_codigo = '" & dcmbCodigo & "'", , adSearchForward
        dcmbCodigo.Tag = "A"
        If clsConsu.adorec_Def.EOF = False Then
'            dcmbGrupo = ""
'            dcmbGrupo.BoundText = ""
'        Else
            'Muestra los subgrupos de un grupo seleccionado en el combo
            dcmbGrupo = clsConsu.adorec_Def("gru_nombre")
            dcmbGrupo.BoundText = dcmbCodigo.Text
            intNivCta = clsConsu.adorec_Def("gru_nivel")
            txtDescripcion = clsConsu.adorec_Def("gru_descripcion")
        End If
            strSql = " Select gru_codigo,gru_nombre,gru_nivel,gru_descripcion " & _
                     " From grupo " & _
                     " Where gru_codigo Like '" & dcmbCodigo & ".%' AND emp_codigo = '" & strEmpresa & "'"
            clsConsuGrid.Ejecutar strSql
            'Muestra la consulta anterior en el datagrid
            Set dGrdSubGrps.DataSource = clsConsuGrid.adorec_Def.DataSource
            dGrdSubGrps.Columns(0).Caption = "Código"
            dGrdSubGrps.Columns(1).Caption = "SubGrupo"
            dGrdSubGrps.Columns(2).Caption = "Nivel"
            dGrdSubGrps.Columns(3).Caption = "Descripción"
    Else
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbGrupo_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbCodigo.Tag <> "A" Then
        If dcmbGrupo.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbGrupo.BoundText
        End If
    End If
End Sub

Private Sub dcmbGrupo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbGrupo.BoundText
End Sub

Private Sub dcmbGrupo_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbGrupo.BoundText
    End If
End Sub

Private Sub Form_Activate()
    'Muestra la lista de cuentas actualizada
    clsConsu.Actualizar
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    Set dcmbGrupo.RowSource = clsConsu.adorec_Def.DataSource
    'Actualiza y muestra los datos de la consulta de grupos
    If Me.Tag <> "" Then
        dcmbCodigo = ""
        dcmbCodigo = Me.Tag
    ElseIf Not clsConsu.adorec_Def.EOF Then
        dcmbCodigo_Change
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsConsuGrid.Inicializar AdoConn, AdoConnMaster
    CargaGrupo
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

Private Sub txtNivel_Change()
    CargaGrupo
End Sub
