VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPlanCuenta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de cuentas"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlanCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Plan de Cuentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6255
      Begin MSDataGridLib.DataGrid dGrdPlan 
         Height          =   2055
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3625
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
         Caption         =   "Listado de SubCuentas"
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
         Left            =   840
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
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   330
         Left            =   3960
         TabIndex        =   1
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
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   540
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3240
         TabIndex        =   8
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3225
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1665
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdNueva 
      Caption         =   "&Nueva"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmPlanCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de plan de cuentas frmPlanCuenta                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana muestra todas las cuentas contables que maneja una determinada      #
'#  empresa.                                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  ctaconta   - tabla que contiene toda la información necesaria de todas las  #
'#               posibles cuentas que maneja una determinad empresa.            #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#  Objetos de la forma:                                                        #
'#  clsConsu  - Objeto para consultar a la base de datos todas las posiles      #
'#              cuentas que maneja una empresa y desplegarlas en un comboox     #
'#  clsConsuGrid  - Objeto para consultar todas las subcuentas relacionadas     #
'#                  con una cuenta seleccionada en el combobox y mostrarlas en  #
'#                  un data grid.                                               #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsConsuGrid As New clsConsulta

Private Sub cmdImprimir_Click()
    frmReporte.strReporte = "rptPlanCuenta"
    frmReporte.Show
End Sub

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
    ' Consulta para conocer si existen detalles con el codigo de cuenta a eliminar
    strSql = " SELECT count(*) As Det" & _
             " FROM det_asiento " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cta_codigo='" & dcmbCodigo.Text & "' AND (det_asi_debe!=0 OR det_asi_haber!=0)"
    clsConsu.Ejecutar (strSql)
    ' Si existen detalles con el código de la cuenta no se elimina
    If clsConsu.adorec_Def("Det") > 0 Then
        MsgBox "No Puede eliminar esta cuenta", vbInformation, "Eliminación"
    
    ' Si no existen detalles con el código de la cuenta nu tiene sucuentas, se elimina con la condición de que el identificador de subcuenta sea 0
    Else
        'Consulta para saber si la cuenta tiene subcuentas
        strSql = " SELECT count(*) As Num" & _
                 " FROM ctaconta " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cta_codigo LIKE '" & dcmbCodigo.Text & ".%' "
        clsConsu.Ejecutar (strSql)
        
        ' Si existen detalles con el código de la cuenta no se elimina
        If clsConsu.adorec_Def("Num") > 0 Then
            MsgBox "No Puede eliminar esta cuenta", vbInformation, "Eliminación"
        
        ' Si no existen detalles con el código de la cuenta nu tiene sucuentas, se elimina con la condición de que el identificador de subcuenta sea 0
        Else
            strSql = " DELETE " & _
                     " FROM ctaconta " & _
                     " WHERE cta_codigo='" & dcmbCodigo.Text & "'" & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsConsu.Ejecutar (strSql), "M"
            strSql = " DELETE " & _
                     " FROM det_asiento " & _
                     " WHERE cta_codigo='" & dcmbCodigo.Text & "'" & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsConsu.Ejecutar (strSql), "M"
            'Consulta para saber si la cuenta tiene subcuentas
            strSql = " SELECT count(*) As Num" & _
                     " FROM ctaconta " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cta_codigo like '" & Left(dcmbCodigo.Text, InStrRev(dcmbCodigo.Text, ".")) & "%' "
            clsConsu.Ejecutar (strSql)
            If clsConsu.adorec_Def("Num") = 0 Then
                strSql = " UPDATE  ctaconta" & _
                         " SET cta_subcta = 0 " & _
                         " WHERE cta_subcta = 1 " & _
                         " AND cta_codigo like '" & Left(dcmbCodigo.Text, InStrRev(dcmbCodigo.Text, ".") - 1) & "%'" & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsConsu.Ejecutar (strSql), "M"
            End If
            
        End If
    End If
    ' Consulta para actualizar los combos
    strSql = " SELECT cta_codigo,cta_nombre,cta_interviene,cta_nivel " & _
             " FROM ctaconta " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_codigo"
    clsConsu.Ejecutar (strSql)
'Muestra el resultado del SQl en un Data Grid
    Set dGrdPlan.DataSource = clsConsu.adorec_Def.DataSource
'Muestra los datos de una columna del resultado del SQL en un data combo
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "cta_codigo"
    Set dcmbCuenta.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCuenta.ListField = "cta_nombre"
    dcmbCuenta.BoundColumn = "cta_codigo"
    dcmbCodigo.Text = ""

End Sub

Private Sub cmdModificar_Click()
    If dcmbCodigo = "" Then
        MsgBox "Seleccione pimero una cuenta.", vbInformation, "Cuenta"
        Exit Sub
    End If
    strCodCuenta = dcmbCodigo
    strDescCuenta = dcmbCuenta
    frmModCuenta.Show
End Sub

Private Sub cmdNueva_Click()
    frmNuevaCuenta.Show
End Sub

Private Sub dcmbCodigo_Change()
    'Muestra el nombre relacionado con el código de la cuenta en el momento de seleccionar otra cuenta del combobox
    clsConsu.adorec_Def.MoveFirst
    clsConsu.adorec_Def.Find "cta_codigo = '" & dcmbCodigo & "'", , adSearchForward
    dcmbCodigo.Tag = "A"
    If clsConsu.adorec_Def.EOF = True Then
        dcmbCuenta = ""
        dcmbCuenta.BoundText = ""
    Else
        'If dcmbCuenta <> clsConsu.adorec_Def("cta_nombre") Then
        dcmbCuenta = clsConsu.adorec_Def("cta_nombre")
        dcmbCuenta.BoundText = dcmbCodigo.Text
        'End If
        clsConsuGrid.Ejecutar "select cta_codigo,cta_nombre,cta_interviene,cta_nivel from ctaconta where cta_codigo like '" & dcmbCodigo & ".%' AND emp_codigo = '" & strEmpresa & "';"
        Set dGrdPlan.DataSource = clsConsuGrid.adorec_Def.DataSource
        intNivCta = clsConsu.adorec_Def("cta_nivel")
        strEstadoPYG = clsConsu.adorec_Def("cta_interviene")
        dGrdPlan.Columns(0).Caption = "Código"
        dGrdPlan.Columns(1).Caption = "SubCuenta"
        dGrdPlan.Columns(2).Caption = "Interviene"
        dGrdPlan.Columns(3).Caption = "Nivel"
    End If
    dcmbCodigo.Tag = ""
End Sub

Private Sub dcmbCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "CN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCodigo
    End If
End Sub

Private Sub dcmbCuenta_Change()
    'Muestra el código relacionado con el nombre de la cuenta en el momento de seleccionar otra cuenta del combobox
     If dcmbCodigo.Tag <> "A" Then
        If dcmbCuenta.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbCuenta.BoundText
        End If
    End If
End Sub
Private Sub dcmbCuenta_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbCuenta.BoundText
End Sub

Private Sub dcmbCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbCuenta.BoundText
    End If
End Sub

Private Sub Form_Activate()
    Dim aux As String
    'Muestra la lista de cuentas actualizada
    clsConsu.Actualizar
    Set dGrdPlan.DataSource = clsConsu.adorec_Def.DataSource
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    Set dcmbCuenta.RowSource = clsConsu.adorec_Def.DataSource
    aux = dcmbCodigo
    dcmbCodigo = ""
    dcmbCodigo = aux
End Sub

Private Sub Form_Load()
    Dim strSql As String
    On Error GoTo errhandler
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn, AdoConnMaster
    clsConsuGrid.Inicializar AdoConn, AdoConnMaster
    'Ejecuta un SQL contra la base de datos
    strSql = " SELECT cta_codigo,cta_nombre,cta_interviene,cta_nivel " & _
             " FROM ctaconta " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_codigo"
    clsConsu.Ejecutar (strSql)
    'Muestra el resultado del SQl en un Data Grid
    Set dGrdPlan.DataSource = clsConsu.adorec_Def.DataSource
    'Muestra los datos de una columna del resultado del SQL en un data combo
    Set dcmbCodigo.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCodigo.ListField = "cta_codigo"
    Set dcmbCuenta.RowSource = clsConsu.adorec_Def.DataSource
    dcmbCuenta.ListField = "cta_nombre"
    dcmbCuenta.BoundColumn = "cta_codigo"
    'Muestra la primera cuenta de la consulta
    If Not clsConsu.adorec_Def.EOF Then
        dcmbCodigo = clsConsu.adorec_Def(0)
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

