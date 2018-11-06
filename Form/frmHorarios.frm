VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHorarios 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Horarios"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmHorarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6630
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Información general"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtNombre 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   4800
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   585
         Width           =   600
      End
   End
   Begin VB.Frame fraBotones 
      BackColor       =   &H00DDDDDD&
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   7200
      Width           =   6375
      Begin VB.CheckBox chkHabilitar 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Deshabilitar"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   3397
         TabIndex        =   31
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton btnAgregar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   1717
         TabIndex        =   30
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame fraHorario 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Horario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6375
      Begin VB.CheckBox chkFalta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Se considera falta al pasar el tiempo  de:"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   5280
         Width           =   3255
      End
      Begin VB.CheckBox chkSalida 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Tolerancia de Salida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Height          =   1215
         Left            =   3480
         TabIndex        =   40
         Top             =   3840
         Width           =   2535
         Begin MSComCtl2.DTPicker dtpSalidaMin 
            Height          =   255
            Left            =   1080
            TabIndex        =   28
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39615
         End
         Begin MSComCtl2.DTPicker dtpSalidaMax 
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previo:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   405
            TabIndex        =   42
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posterior:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   780
            Width           =   660
         End
      End
      Begin VB.CheckBox chkEntrada 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Tolerancia de Entrada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Frame fraEntrada 
         BackColor       =   &H00DDDDDD&
         Height          =   1215
         Left            =   480
         TabIndex        =   37
         Top             =   3840
         Width           =   2535
         Begin MSComCtl2.DTPicker dtpEntradaMin 
            Height          =   255
            Left            =   1125
            TabIndex        =   25
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39615
         End
         Begin MSComCtl2.DTPicker dtpEntradaMax 
            Height          =   255
            Left            =   1125
            TabIndex        =   26
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posterior:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previo:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   405
            TabIndex        =   38
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Domingo"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   6
         Left            =   720
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Sabado"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   5
         Left            =   720
         TabIndex        =   18
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Viernes"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   4
         Left            =   720
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Jueves"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Miercoles"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Martes"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame fraDia 
         BackColor       =   &H00DDDDDD&
         Height          =   3255
         Left            =   2280
         TabIndex        =   33
         Top             =   240
         Width           =   3255
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   13
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   22
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   8
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   11
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   14
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   17
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   20
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   23
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   59310081
            CurrentDate     =   39418
         End
         Begin VB.Label lblEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   35
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblSalida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   34
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Lunes"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFalta 
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   5280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59310081
         CurrentDate     =   39615
      End
   End
End
Attribute VB_Name = "frmHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public horariocodigo As String
Public Actualizar As Boolean
Private strSQL As String
Private clsSql As New clsConsulta
Private respuesta As Integer
Private i As Integer


Private Sub btnAgregar_Click()
    If btnAgregar.Caption = "&Aceptar" Then
        If ComprobarDatos = True Then
            AgregarDatos
            MsgBox "Se ha ingresado el Horario exitosamente!", vbInformation, "Definición de Horarios"
            Actualizar = False
            Limpiar
            frmVerHorarios.Carga
            Unload Me
        End If
    Else
        If ComprobarDatos = True Then
            Actualizar = True
            If ActualizarDatos = True Then
                MsgBox "Se ha actualizado el Horario exitosamente!", vbInformation, "Definición de Horarios"
                Actualizar = False
                Limpiar
                frmVerHorarios.Carga
                'cargar de nuevo el form de atras
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    Actualizar = False
    Limpiar
    Unload Me
End Sub

Private Sub chk_Click(Index As Integer)
    Check chk(Index), dtpEntrada(Index), dtpSalida(Index)
End Sub

Private Sub chkEntrada_Click()
    If chkEntrada.Value = 1 Then
        dtpEntradaMin.Enabled = True
        dtpEntradaMax.Enabled = True
    Else
        dtpEntradaMin.Enabled = False
        dtpEntradaMax.Enabled = False
    End If
    CrearHora dtpEntradaMin
    CrearHora dtpEntradaMax
End Sub

Private Sub chkFalta_Click()
    If chkFalta.Value = 1 Then
        dtpFalta.Enabled = True
    Else
        dtpFalta.Enabled = False
    End If
    CrearHora dtpFalta
End Sub

Private Sub chkSalida_Click()
    If chkSalida.Value = 1 Then
        dtpSalidaMin.Enabled = True
        dtpSalidaMax.Enabled = True
    Else
        dtpSalidaMin.Enabled = False
        dtpSalidaMax.Enabled = False
    End If
    CrearHora dtpSalidaMin
    CrearHora dtpSalidaMax
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    If Actualizar = True Then
        btnAgregar.Caption = "&Modificar"
        chkHabilitar.Visible = True
        Limpiar
        CargarDatosExistentes
    Else
        btnAgregar.Caption = "&Aceptar"
        chkHabilitar.Visible = False
        Limpiar
    End If
    
End Sub



Private Sub Check(chk As CheckBox, dtpI As DTPicker, dtpF As DTPicker)
    If CBool(chk.Value) = True Then
        dtpI.Enabled = True
        dtpF.Enabled = True
    Else
        dtpI.Enabled = False
        dtpF.Enabled = False
    End If
End Sub

Private Sub IniciarComponentes()
    For i = 0 To 6
        CrearHora dtpEntrada(i)
        CrearHora dtpSalida(i)
    Next i
End Sub

Private Sub CrearHora(objeto As DTPicker)
    objeto.Format = dtpCustom
    objeto.CustomFormat = "HH:mm"
    objeto.UpDown = True
    objeto.Hour = "00"
    objeto.Minute = "00"
    objeto.Second = "00"
    If Actualizar = False Then
        objeto.Value = "00:00:00"
    End If
End Sub

Private Function ComprobarDatos() As Boolean
    Dim contador As Integer
    contador = 0
    Actualizar = False
    If txtNombre.Text = "" Then
        MsgBox "Ingrese un Nombre para el Horario", vbCritical, "Definición de Horarios"
        ComprobarDatos = False
    Else
        For i = 0 To 6
            If CBool(chk(i).Value) = True Then
                If dtpEntrada(i).Value > dtpSalida(i).Value Or Format(dtpEntrada(i).Value, "hh:mm:ss") = "00:00:00" Or Format(dtpSalida(i).Value, "hh:mm:ss") = "00:00:00" Then
                    MsgBox "Ingrese una hora válida para el día " & MostrarDia(i) & " ", vbCritical, "Horarios"
                    ComprobarDatos = False
                    Exit Function
                End If
            End If
        Next i
        
        For i = 0 To 6
            If CBool(chk(i).Value) = True Then
                contador = contador + 1
            End If
        Next i
        If contador = 0 Then
            MsgBox "Ingrese por lo menos un día válido", vbCritical, "Horarios"
            ComprobarDatos = False
            Exit Function
        Else
            ComprobarDatos = True
        End If
        
    End If
End Function


Private Sub Limpiar()
    txtNombre.Text = ""
    chkEntrada.Value = 0
    chkEntrada_Click
    chkSalida.Value = 0
    chkSalida_Click
    IniciarComponentes
 
    chkFalta.Value = 0
    chkFalta_Click
    
    For i = 0 To 6
        chk(i).Value = 0
        chk_Click i
    Next i
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    horariocodigo = ""
    Actualizar = False
End Sub

Private Sub AgregarDatos()
    strSQL = " SELECT COALESCE(MAX(hor_codigo),0)+1 as codigo " & _
             " FROM horario " & _
             " WHERE emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSQL
    If clsSql.adorec_Def.RecordCount = 0 Then
        horariocodigo = "1"
    Else
        horariocodigo = clsSql.adorec_Def("codigo")
    End If
    
    strSQL = " SELECT hor_codigo " & _
             " FROM horario " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND hor_codigo='" & horariocodigo & "' "
    clsSql.Ejecutar strSQL
    
    If clsSql.adorec_Def.RecordCount = 0 Then
        Dim disponible As Integer
        disponible = IIf(CBool(chkHabilitar.Value) = True, 0, 1)
        strSQL = " INSERT INTO horario(emp_codigo,hor_codigo,hor_descripcion," & _
                 " hor_fechamod,hor_usumod,hor_disponible) " & _
                 " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & txtNombre.Text & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "','" & disponible & "')"
        clsSql.Ejecutar strSQL, "M"
        
        
        If chkEntrada.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tol_entrada_min='" & Format(dtpEntradaMin.Value, "HH:mm:SS") & "'," & _
                     " hor_tol_entrada_max='" & Format(dtpEntradaMax.Value, "HH:mm:SS") & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
                
        If chkSalida.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tol_salida_min='" & Format(dtpSalidaMin.Value, "HH:mm:SS") & "'," & _
                     " hor_tol_salida_max='" & Format(dtpSalidaMax.Value, "HH:mm:SS") & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
        
        If chkFalta.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tolerancia='" & Format(dtpFalta.Value, "HH:mm:SS") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
                
        For i = 0 To 6
            If CBool(chk(i)) = True Then
                If chkFalta.Value = 1 Then
                    strSQL = " INSERT INTO det_horario(emp_codigo,hor_codigo,det_hor_dia,det_hor_entrada,det_hor_salida," & _
                             " det_hor_fechamod,det_hor_usumod,det_hor_tolerancia) " & _
                             " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & i & "','" & dtpEntrada(i).Value & "','" & dtpSalida(i).Value & "'," & _
                             " CURRENT_TIMESTAMP,'" & strUsuario & "','" & Format(dtpEntrada(i).Value, "HH:mm:SS") + Format(dtpFalta.Value, "HH:mm:SS") & "')"
                    clsSql.Ejecutar strSQL, "M"
                Else
                    strSQL = " INSERT INTO det_horario(emp_codigo,hor_codigo,det_hor_dia,det_hor_entrada,det_hor_salida," & _
                             " det_hor_fechamod,det_hor_usumod) " & _
                             " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & i & "','" & dtpEntrada(i).Value & "','" & dtpSalida(i).Value & "'," & _
                             " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    clsSql.Ejecutar strSQL, "M"
                End If
                
            End If
        Next i
    Else
        MsgBox "El Código de Horario ya existe, intente nuevamente", vbCritical, "Definición de Horarios"
    End If
End Sub

Private Sub CargarDatosExistentes()
    Dim horcod As String
    strSQL = " SELECT horario.hor_codigo,hor_descripcion,TIME_FORMAT(hor_tol_entrada_min,'%H:%i:%s') as entmin,TIME_FORMAT(hor_tol_entrada_max,'%H:%i:%s') as entmax,TIME_FORMAT(hor_tol_salida_min,'%H:%i:%s') as salmin,TIME_FORMAT(hor_tol_salida_max,'%H:%i:%s') as salmax,TIME_FORMAT(hor_tolerancia,'%H:%i:%s') as tolerancia,hor_disponible " & _
             " FROM horario " & _
             " INNER JOIN det_horario ON det_horario.emp_codigo=horario.emp_codigo " & _
             " AND det_horario.hor_codigo=horario.hor_codigo " & _
             " WHERE horario.emp_codigo='" & strEmpresa & "' " & _
             " AND horario.hor_codigo='" & horariocodigo & "' "
    clsSql.Ejecutar strSQL
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        horariocodigo = clsSql.adorec_Def("hor_codigo")
        
        txtNombre.Text = clsSql.adorec_Def("hor_descripcion")
        
        If Not IsNull(clsSql.adorec_Def("entmin")) Or Not IsNull(clsSql.adorec_Def("entmax")) Then
            chkEntrada.Value = 1
            dtpEntradaMin.Value = clsSql.adorec_Def("entmin")
            dtpEntradaMax.Value = clsSql.adorec_Def("entmax")
        End If
        
        
        If Not IsNull(clsSql.adorec_Def("salmin")) Or Not IsNull(clsSql.adorec_Def("salmax")) Then
            chkSalida.Value = 1
            dtpSalidaMin.Value = clsSql.adorec_Def("salmin")
            dtpSalidaMax.Value = clsSql.adorec_Def("salmax")
        End If
            
        If Not IsNull(clsSql.adorec_Def("tolerancia")) Then
            chkFalta.Value = 1
            dtpFalta.Value = clsSql.adorec_Def("tolerancia")
        End If
            
        If FormatoD0(clsSql.adorec_Def("hor_disponible")) = 0 Then
            chkHabilitar.Value = 1
        Else
            chkHabilitar.Value = 0
        End If
            
        strSQL = " SELECT det_hor_dia,TIME_FORMAT(det_hor_entrada,'%H:%i:%s') as entrada ,TIME_FORMAT(det_hor_salida,'%H:%i:%s') as salida " & _
                 " FROM det_horario " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND hor_codigo='" & horariocodigo & "' " & _
                 " ORDER BY det_hor_dia "
        clsSql.Ejecutar strSQL
        
        While Not clsSql.adorec_Def.EOF
            i = clsSql.adorec_Def("det_hor_dia")
            chk(i).Value = 1
            chk_Click 1
            dtpEntrada(i).Value = clsSql.adorec_Def("entrada")
            dtpSalida(i).Value = clsSql.adorec_Def("salida")
            clsSql.adorec_Def.MoveNext
        Wend
        btnAgregar.Caption = "&Modificar"
    Else
        btnAgregar.Caption = "&Aceptar"
    End If
End Sub

Private Function ActualizarDatos() As Boolean
    strSQL = " SELECT COUNT(hor_codigo) as contar " & _
             " FROM horario_empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND hor_codigo='" & horariocodigo & "' "
    clsSql.Ejecutar strSQL
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        If FormatoD0(clsSql.adorec_Def(0)) > 0 Then
            If MsgBox("Existen empleados asignados al Horario, desea continuar?", vbQuestion + vbYesNo, "Definición de Horarios") = vbNo Then
                ActualizarDatos = False
                Exit Function
            End If
        End If
    End If

    
    strSQL = " DELETE FROM horario " & _
             " WHERE hor_codigo = '" & horariocodigo & "' " & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSQL, "M"
    
    strSQL = " DELETE FROM det_horario " & _
             " WHERE hor_codigo = '" & horariocodigo & "' " & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSQL, "M"
        
        Dim disponible As Integer
        disponible = IIf(CBool(chkHabilitar.Value) = True, 0, 1)
        
        strSQL = " INSERT INTO horario(emp_codigo,hor_codigo,hor_descripcion," & _
                 " hor_fechamod,hor_usumod,hor_disponible) " & _
                 " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & txtNombre.Text & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "','" & disponible & "')"
        clsSql.Ejecutar strSQL, "M"
        
        If chkEntrada.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tol_entrada_min='" & Format(dtpEntradaMin.Value, "HH:mm:SS") & "'," & _
                     " hor_tol_entrada_max='" & Format(dtpEntradaMax.Value, "HH:mm:SS") & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
                
        If chkSalida.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tol_salida_min='" & Format(dtpSalidaMin.Value, "HH:mm:SS") & "'," & _
                     " hor_tol_salida_max='" & Format(dtpSalidaMax.Value, "HH:mm:SS") & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
        
        If chkFalta.Value = 1 Then
            strSQL = " UPDATE horario SET " & _
                     " hor_tolerancia='" & Format(dtpFalta.Value, "HH:mm:SS") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND hor_codigo='" & horariocodigo & "' "
            clsSql.Ejecutar strSQL, "M"
        End If
                
        For i = 0 To 6
            If CBool(chk(i)) = True Then
                If chkFalta.Value = 1 Then
                    strSQL = " INSERT INTO det_horario(emp_codigo,hor_codigo,det_hor_dia,det_hor_entrada,det_hor_salida," & _
                             " det_hor_fechamod,det_hor_usumod,det_hor_tolerancia) " & _
                             " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & i & "','" & Format(dtpEntrada(i).Value, "HH:mm:SS") & "','" & Format(dtpSalida(i).Value, "HH:mm:SS") & "'," & _
                             " CURRENT_TIMESTAMP,'" & strUsuario & "','" & Format(dtpEntrada(i).Value + dtpFalta.Value, "HH:mm:SS") & "') "
                    clsSql.Ejecutar strSQL, "M"
                Else
                    strSQL = " INSERT INTO det_horario(emp_codigo,hor_codigo,det_hor_dia,det_hor_entrada,det_hor_salida," & _
                             " det_hor_fechamod,det_hor_usumod) " & _
                             " VALUES('" & strEmpresa & "','" & horariocodigo & "','" & i & "','" & Format(dtpEntrada(i).Value, "HH:mm:SS") & "','" & Format(dtpSalida(i).Value, "HH:mm:SS") & "'," & _
                             " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                    clsSql.Ejecutar strSQL, "M"
                End If
            End If
        Next i
    ActualizarDatos = True
End Function


