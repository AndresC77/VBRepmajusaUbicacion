VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBR 
   Caption         =   "Browser"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBR.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   5220
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   105
      Top             =   3405
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   3720
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5040
      ExtentX         =   8890
      ExtentY         =   6562
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
End
Attribute VB_Name = "frmBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  forma de Browser                                                            #
'#  frmBR V1.0                                                                  #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para abrir el browser para presentar las partes del sistema que se  #
'#  desarrollaron en PHP. Esta ventana se cargará cada ves que el ususario      #
'#  cambie de empresa.                                                          #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#  Objetos de la forma:                                                        #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'
Private Sub Form_Load()
    ' Se establese el tamaño de la ventana de acuerdo a la ventana principal
    Me.Top = 0
    Me.Left = 0
    Me.Height = mdiPrincipal.Height - 1245
    Me.Width = mdiPrincipal.Width - 180
'    webBrowser.Height = Me.Height - 660
'    webBrowser.Width = Me.Width - 300
    ' Se envia a cargar la página principal al browser
    webBrowser.Navigate "http://" & strServidorWeb & "/default_w.html", 2 + 4 + 8
End Sub

Private Sub Form_Resize()
    If Me.Height > 610 Then
        webBrowser.Height = Me.Height - 610
    End If
    If Me.Width > 300 Then
        webBrowser.Width = Me.Width - 300
    End If
End Sub

Private Sub Timer_Timer()
    ' Si acabó de cargar la páina default_w.html
    If webBrowser.ReadyState = READYSTATE_COMPLETE Then
    ' Se envia a cargar la página work_w.php enviando el usuario, clave y empresa al
    ' Frame al que debe ir esta página y asi presentar el menú de esta parte
        webBrowser.Navigate "http://" & strServidorWeb & "/work_w.php?" & "user=" & strUsuario & "&pass=" & strClave & "&suc=" & strSucursal & "&emp=" & strEmpresa & "&EmpresaN=0", 2 + 4 + 8
        Timer.Enabled = False
    End If
End Sub

