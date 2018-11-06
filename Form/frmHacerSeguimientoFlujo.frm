VERSION 5.00
Begin VB.Form frmHacerSeguimientoFlujo 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de flujo"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHacerSeguimientoFlujo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   13230
   Begin VB.CommandButton cmdGuardarCerrar 
      Caption         =   "&Guardar y Cerrar"
      Height          =   375
      Left            =   5888
      TabIndex        =   28
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4208
      TabIndex        =   27
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtObservaciones 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   1125
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   2280
      Width           =   11775
   End
   Begin VB.TextBox txtFechaFinTareaReal 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFechaInicioTareaReal 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8130
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFechaFinTareaProgramada 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFechaInicioTareaProgramada 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtNumeroDias 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtDescripcionTarea 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7050
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox txtNombreTarea 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtFechaFinFlujo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtFechaInicioFlujo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtDescripcionFlujo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7050
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txtNombreFlujo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtTituloFlujo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   3735
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7568
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   2355
      Width           =   1155
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Fin Real:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   10440
      TabIndex        =   24
      Top             =   1635
      Width           =   795
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Inicio Real:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   7080
      TabIndex        =   22
      Top             =   1635
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Fin Prog:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   3360
      TabIndex        =   20
      Top             =   1635
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Inicio Prog:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   1635
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Días:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   1995
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrip. Tarea:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   5760
      TabIndex        =   14
      Top             =   1275
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarea:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1275
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Fin:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   3360
      TabIndex        =   10
      Top             =   795
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   795
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción del Flujo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   5400
      TabIndex        =   6
      Top             =   435
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Flujo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   435
      Width           =   1230
   End
   Begin VB.Label lblDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Flujo:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmHacerSeguimientoFlujo"
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

Private Sub cmdGuardar_Click()
    strSql = " UPDATE det_historia_flujo " & _
             " SET det_his_flu_observacion='" & txtObservaciones.Text & "', " & _
             " det_his_flu_fechamod=CURRENT_TIMESTAMP," & _
             " det_his_flu_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND flu_codigo='" & txtTituloFlujo.Tag & "'" & _
             " AND his_flu_codigo='" & txtNombreFlujo.Tag & "'" & _
             " AND det_flu_codigo='" & txtNombreTarea.Tag & "'"
    clsSql.Ejecutar strSql, "M"
    MsgBox "Datos almacenados", vbInformation, "Flujos"
    EnviarComunicacion "GUARDAR"
    Unload Me
End Sub

Private Sub EnviarComunicacion(strTipo As String)
Dim clsAuxCorreo As New clsConsulta
    Dim strAsunto As String
    Dim strCuerpo As String
    Dim strCopia As String
    strCopia = "acevallos@enlacedigital.com.ec"
    clsAuxCorreo.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT DISTINCT res_flu_email,res_flu_nombre " & _
             " FROM det_historia_flujo INNER JOIN det_flujo " & _
             " ON det_historia_flujo.emp_codigo=det_flujo.emp_codigo " & _
             " AND det_historia_flujo.flu_codigo=det_flujo.flu_codigo " & _
             " AND det_historia_flujo.det_flu_codigo=det_flujo.det_flu_codigo " & _
             " INNER JOIN responsable_flujo ON det_flujo.emp_codigo=responsable_flujo.emp_codigo" & _
             " AND det_flujo.res_flu_codigo=responsable_flujo.res_flu_codigo" & _
             " WHERE det_historia_flujo.emp_codigo = '" & strEmpresa & "'" & _
             " AND det_historia_flujo.flu_codigo = '" & txtTituloFlujo.Tag & "'" & _
             " AND det_historia_flujo.his_flu_codigo = '" & txtNombreFlujo.Tag & "'" & _
             " AND det_historia_flujo.det_flu_codigo >= '" & txtNombreTarea.Tag & "'" & _
             " ORDER BY det_flu_codigo "
    clsAuxCorreo.Ejecutar strSql
    If strTipo = "GUARDAR" Then
        strAsunto = "Observación al Flujo N." & txtTituloFlujo.Tag & " - " & txtTituloFlujo.Text
        While Not clsAuxCorreo.adorec_Def.EOF
            strCuerpo = "Estomad@." & vbNewLine & vbNewLine & _
                        "Se ha ingresado la siguiente observación a la tarea descrita a continuación:" & vbNewLine & _
                        "Tarea: " & txtNombreTarea.Text & vbNewLine & _
                        "Descripción: " & txtDescripcionTarea.Text & vbNewLine & _
                        "Fecha estimada de inicio de la tarea: " & txtFechaInicioTareaProgramada.Text & vbNewLine & _
                        "Fecha estimada de finalización de la tarea: " & txtFechaFinTareaProgramada.Text & vbNewLine & _
                        "Fecha real de inicio de la tarea: " & txtFechaInicioTareaReal.Text & vbNewLine & vbNewLine & _
                        "OBSERVACIONES: " & txtObservaciones.Text
                        
            EnviarMail "Flujos", "jefedeproductojsn@rbimportadores.com", clsAuxCorreo.adorec_Def("res_flu_nombre"), clsAuxCorreo.adorec_Def("res_flu_email"), strCopia, strAsunto, strCuerpo
            clsAuxCorreo.adorec_Def.MoveNext
        Wend
    Else
        strAsunto = "Finalización de tarea del Flujo N." & txtTituloFlujo.Tag & " - " & txtTituloFlujo.Text
        While Not clsAuxCorreo.adorec_Def.EOF
            strCuerpo = "Estomad@." & vbNewLine & vbNewLine & _
                        "Se ha finalizado la Tarea descrita a continuación con la siguiente observación:" & vbNewLine & _
                        "Tarea: " & txtNombreTarea.Text & vbNewLine & _
                        "Descripción: " & txtDescripcionTarea.Text & vbNewLine & _
                        "Fecha estimada de inicio de la tarea: " & txtFechaInicioTareaProgramada.Text & vbNewLine & _
                        "Fecha estimada de finalización de la tarea: " & txtFechaFinTareaProgramada.Text & vbNewLine & _
                        "Fecha real de inicio de la tarea: " & txtFechaInicioTareaReal.Text & vbNewLine & vbNewLine & _
                        "Fecha real de finalización de la tarea: " & HoyDia & vbNewLine & vbNewLine & _
                        "OBSERVACIONES: " & txtObservaciones.Text
            
            EnviarMail "Flujos", "jefedeproductojsn@rbimportadores.com", clsAuxCorreo.adorec_Def("res_flu_nombre"), clsAuxCorreo.adorec_Def("res_flu_email"), strCopia, strAsunto, strCuerpo
            clsAuxCorreo.adorec_Def.MoveNext
        Wend
    
    End If

End Sub

Private Sub cmdGuardarCerrar_Click()
    
    Dim i As Long
    Dim clsAux As New clsConsulta
    Dim FechaCalc As String
    
    strSql = " UPDATE det_historia_flujo " & _
             " SET det_his_flu_observacion='" & txtObservaciones.Text & "', " & _
             " det_his_flu_fecha_real_inicio=IF(det_his_flu_fecha_real_inicio>'" & HoyDia & "','" & HoyDia & "',det_his_flu_fecha_real_inicio)," & _
             " det_his_flu_fecha_real_fin='" & HoyDia & "'," & _
             " det_his_flu_fechamod=CURRENT_TIMESTAMP," & _
             " det_his_flu_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND flu_codigo='" & txtTituloFlujo.Tag & "'" & _
             " AND his_flu_codigo='" & txtNombreFlujo.Tag & "'" & _
             " AND det_flu_codigo='" & txtNombreTarea.Tag & "'"
    clsSql.Ejecutar strSql, "M"
    
    strSql = " SELECT det_flujo.det_flu_codigo,det_flu_nombre,det_flu_descripcion," & _
             " det_flu_tiempo,res_flu_nombre,LEFT(det_his_flu_fecha_prevista_inicio,10)," & _
             " LEFT(det_his_flu_fecha_prevista_fin,10),if(det_his_flu_fecha_real_inicio='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_inicio,10))," & _
             "  if(det_his_flu_fecha_real_fin='0000-00-00 00:00:00','',LEFT(det_his_flu_fecha_real_fin,10)),det_his_flu_observacion," & _
             " responsable_flujo.usu_codigo,det_his_flu_fechamod, det_his_flu_usumod, '0' as modi " & _
             " FROM det_historia_flujo INNER JOIN det_flujo " & _
             " ON det_historia_flujo.emp_codigo=det_flujo.emp_codigo " & _
             " AND det_historia_flujo.flu_codigo=det_flujo.flu_codigo " & _
             " AND det_historia_flujo.det_flu_codigo=det_flujo.det_flu_codigo " & _
             " INNER JOIN responsable_flujo ON det_flujo.emp_codigo=responsable_flujo.emp_codigo" & _
             " AND det_flujo.res_flu_codigo=responsable_flujo.res_flu_codigo" & _
             " WHERE det_historia_flujo.emp_codigo = '" & strEmpresa & "'" & _
             " AND det_historia_flujo.flu_codigo = '" & txtTituloFlujo.Tag & "'" & _
             " AND det_historia_flujo.his_flu_codigo = '" & txtNombreFlujo.Tag & "'" & _
             " AND det_flujo.det_flu_codigo > '" & txtNombreTarea.Tag & "'" & _
             " ORDER BY det_flujo.det_flu_codigo LIMIT 1 "
    clsSql.Ejecutar strSql
    clsAux.Inicializar AdoConn, AdoConnMaster
    FechaCalc = SumaDiasHabiles(HoyDia, 2)
    If clsSql.adorec_Def.RecordCount > 0 Then
        i = 0
        While Not clsSql.adorec_Def.EOF
            strSql = " UPDATE det_historia_flujo " & _
                     " SET det_his_flu_fecha_real_inicio='" & FechaCalc & "'"
            strSql = strSql & " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND his_flu_codigo='" & txtNombreFlujo.Tag & "'" & _
                     " AND flu_codigo='" & txtTituloFlujo.Tag & "'" & _
                     " AND det_flu_codigo='" & clsSql.adorec_Def("det_flu_codigo") & "'"
            clsAux.Ejecutar strSql, "M"
            FechaCalc = SumaDiasHabiles(FechaCalc, clsSql.adorec_Def("det_flu_tiempo") + 1)
            clsSql.adorec_Def.MoveNext
            i = i + 1
        Wend
    End If
    MsgBox "Datos almacenados", vbInformation, "Flujos"
    EnviarComunicacion "CERRARGUARDAR"
    Unload Me
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


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtObservaciones" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
End Sub

