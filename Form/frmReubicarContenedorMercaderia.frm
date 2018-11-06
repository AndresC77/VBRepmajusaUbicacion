VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReubicarContenedorMercaderia 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reubicar Contenedores de Mercaderia"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmReubicarContenedorMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7680
   Begin VB.TextBox txtObservacion 
      Height          =   285
      Left            =   1410
      TabIndex        =   17
      Top             =   2280
      Width           =   6000
   End
   Begin VB.CommandButton cmdReubicarContenedor 
      Caption         =   "&Reubicar Contenedor"
      Height          =   360
      Left            =   1910
      TabIndex        =   2
      Top             =   2640
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4070
      TabIndex        =   1
      Top             =   2640
      Width           =   1700
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
      Begin VB.TextBox txtUbicacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtBodega 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   248
         Width           =   1815
      End
      Begin VB.TextBox TxtObser 
         Height          =   645
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   930
         Width           =   6000
      End
      Begin MSComCtl2.DTPicker dtpFecha 
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
         Left            =   4755
         TabIndex        =   8
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   69009411
         CurrentDate     =   37463
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contenedor:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Creación:"
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
         Left            =   3495
         TabIndex        =   9
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   3930
         TabIndex        =   6
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega:"
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
         Left            =   645
         TabIndex        =   5
         Top             =   645
         Width           =   600
      End
      Begin VB.Label LblObser 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
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
         Left            =   270
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cmbBodega 
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbUbicacion 
      Height          =   315
      Left            =   4875
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observación:"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega:"
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
      Left            =   840
      TabIndex        =   14
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación:"
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
      Left            =   4125
      TabIndex        =   13
      Top             =   1920
      Width           =   750
   End
End
Attribute VB_Name = "frmReubicarContenedorMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String

Private Sub cmbBodega_Validate(Cancel As Boolean)
    CargaUbica
End Sub

Private Sub CargaUbica()
    strSql = " SELECT ubi_bod_codigo " & _
             " FROM ubicacion_bodega " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND dep_codigo='" & CmbBodega.BoundText & "'" & _
             " ORDER BY ubi_bod_codigo "
    clsCon_Def.Ejecutar strSql
    Set cmbUbicacion.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbUbicacion.ListField = "ubi_bod_codigo"
    cmbUbicacion.BoundColumn = "ubi_bod_codigo"
End Sub

Private Sub cmdReubicarContenedor_Click()
    Dim clsConte As New clsContenedor
    If CmbBodega.MatchedWithList = True And cmbUbicacion.MatchedWithList = True Then
        clsConte.Inicializar AdoConn, AdoConnMaster
        clsConte.SetContenedor txtCodigo.Text
        clsConte.CambiarUbicacionContenedor CmbBodega.BoundText, cmbUbicacion.BoundText, txtObservacion.Text
        Set clsConte = Nothing
        Unload Me
    Else
        MsgBox "No ha seleccionado una Bodega y/o Ubicacion válida", vbCritical, "Ubicacion"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT dep_codigo, dep_nombre " & _
             " FROM deposito " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSql
    Set CmbBodega.RowSource = clsCon_Def.adorec_Def.DataSource
    CmbBodega.ListField = "dep_nombre"
    CmbBodega.BoundColumn = "dep_codigo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
