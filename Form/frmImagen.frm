VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImagen 
   BackColor       =   &H00DDDDDD&
   Caption         =   "Imagen"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   Icon            =   "frmImagen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4305
      Left            =   3840
      Picture         =   "frmImagen.frx":030A
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "&Explorar..."
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6480
      TabIndex        =   1
      Top             =   5400
      Width           =   1700
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      Top             =   5400
      Width           =   1700
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   2040
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgPic 
      BorderStyle     =   1  'Fixed Single
      Height          =   5000
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "frmImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fila As Long
Public Columna As Long
Private Sub cmdCargar_Click()
    SavePicture imgPic.Picture, Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg"
    'frmContenedorMercaderia.vsfgCaracteristica.TextMatrix(frmContenedorMercaderia.vsfgCaracteristica.Row, frmContenedorMercaderia.vsfgCaracteristica.Col) = Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg"
    frmContenedorMercaderia.vsfgCaracteristica.CellPicture = imgPic.Picture
    frmContenedorMercaderia.vsfgCaracteristica.Cell(flexcpPicture, frmContenedorMercaderia.vsfgCaracteristica.Row, frmContenedorMercaderia.vsfgCaracteristica.Col) = LoadPicture(Trim(App.Path) & "\" & frmContenedorMercaderia.vsfgCaracteristica.Row & ".jpeg")
    'frmContenedorMercaderia.vsfgCaracteristica.Cell(flexcpHeight, Fila, Columna) = 255
    
    'frmContenedorMercaderia.vsfgCaracteristica.CellHeight = 5000
    'frmContenedorMercaderia.vsfgCaracteristica.Refresh
    
    
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExplorar_Click()
    Dim factor As Double
    Dim anchoPic As Long
    Dim altoPic As Long
    Dim anchoImg As Long
    Dim altoImg As Long
    cdArchivo.ShowOpen
    If cdArchivo.FileName <> "" Then
        pic.Picture = LoadPicture(cdArchivo.FileName)
        anchoPic = pic.Width
        altoPic = pic.Height
        anchoImg = 8000
        altoImg = 5000
        If anchoImg / anchoPic > altoImg / altoPic Then
            factor = altoImg / altoPic
        Else
            factor = anchoImg / anchoPic
        End If
        pic.PaintPicture pic.Picture, 0, 0, FormatoD0(anchoPic / factor), FormatoD0(altoPic / factor)
        imgPic.Width = FormatoD0(anchoPic * factor)
        imgPic.Height = FormatoD0(altoPic * factor)
        imgPic.Picture = pic.Picture
    End If
End Sub
