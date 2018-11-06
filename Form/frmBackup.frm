VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4095
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos para el Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "&Explorar"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   2640
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grabar en: "
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAceptar_Click()
    Dim sDir As String
    MyAppID = Shell(txtArchivo.Tag & "\respaldarneed.bat " & strUsuario & " " & strClave & " " & strBDD & " " & cdArchivo.FileName, 1)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    sDir = CurDir
    cdArchivo.ShowSave
    txtArchivo = cdArchivo.FileName
    ChDir sDir
End Sub

