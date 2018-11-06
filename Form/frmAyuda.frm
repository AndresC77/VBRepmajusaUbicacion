VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAyuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda"
   ClientHeight    =   4980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10665
   Icon            =   "frmAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser Br 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
    Unload Me
End Sub
