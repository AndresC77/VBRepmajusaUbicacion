VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImprimirEnBloque 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Pedidos Enviados a Bodega"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmImprimirEnBloque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12225
   Begin VB.CommandButton cmdResumen 
      Caption         =   "Imprimir Resumen"
      Height          =   375
      Left            =   5400
      TabIndex        =   51
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSticker 
      Caption         =   "Imprimir Sticker"
      Height          =   375
      Left            =   3720
      TabIndex        =   47
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid VSFGFormaPago 
         Height          =   1935
         Left            =   8520
         TabIndex        =   49
         Top             =   2040
         Width           =   3300
         _cx             =   5821
         _cy             =   3413
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImprimirEnBloque.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "CONSULTAR"
         Height          =   375
         Left            =   8760
         TabIndex        =   46
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CheckBox chkCancelado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "SOLO PEDIDOS CANCELADOS"
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
         Height          =   375
         Left            =   7080
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Estado de Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   3960
         Width           =   6855
         Begin VB.OptionButton optPedAnulado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Anulados"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5760
            TabIndex        =   45
            Top             =   240
            Width           =   1020
         End
         Begin VB.OptionButton optPedFacturado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Facturados"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   4560
            TabIndex        =   44
            Top             =   240
            Width           =   1500
         End
         Begin VB.OptionButton optPedConfirmado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Confirmados"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   3360
            TabIndex        =   43
            Top             =   240
            Width           =   1500
         End
         Begin VB.OptionButton optPedAsignado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Procesado"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2280
            TabIndex        =   42
            Top             =   240
            Width           =   1500
         End
         Begin VB.OptionButton optPedGuardado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Borrador"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1080
            TabIndex        =   41
            Top             =   240
            Width           =   1380
         End
         Begin VB.OptionButton optPedTodos 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Todos"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   1815
         Left            =   8880
         TabIndex        =   34
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
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
            Height          =   255
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   225
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker Fecha1 
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
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   66453507
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
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
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   66453507
            CurrentDate     =   37463
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.OptionButton optN1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   15
         Top             =   765
         Width           =   255
      End
      Begin VB.OptionButton optN2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   14
         Top             =   1125
         Width           =   255
      End
      Begin VB.OptionButton optN3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   13
         Top             =   1485
         Width           =   255
      End
      Begin VB.OptionButton optN4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   12
         Top             =   1845
         Width           =   255
      End
      Begin VB.OptionButton optN5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   11
         Top             =   2205
         Width           =   255
      End
      Begin VB.OptionButton optN6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   10
         Top             =   2565
         Width           =   255
      End
      Begin VB.OptionButton optN7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   9
         Top             =   2925
         Width           =   255
      End
      Begin VB.OptionButton optN8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   8
         Top             =   3285
         Width           =   255
      End
      Begin VB.OptionButton optN9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1170
         TabIndex        =   7
         Top             =   3645
         Width           =   255
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   255
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   330
         Left            =   1530
         TabIndex        =   16
         Top             =   720
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   330
         Left            =   1530
         TabIndex        =   17
         Top             =   1080
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   330
         Left            =   1530
         TabIndex        =   18
         Top             =   1440
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   330
         Left            =   1530
         TabIndex        =   19
         Top             =   1800
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   330
         Left            =   1530
         TabIndex        =   20
         Top             =   2160
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   330
         Left            =   1530
         TabIndex        =   21
         Top             =   2520
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   330
         Left            =   1530
         TabIndex        =   22
         Top             =   2880
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   330
         Left            =   1530
         TabIndex        =   23
         Top             =   3240
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   330
         Left            =   1530
         TabIndex        =   24
         Top             =   3600
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eje.E. N4:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   270
         TabIndex        =   33
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empren N3:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dir N2:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   31
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.Zona N1:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   705
         TabIndex        =   29
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   705
         TabIndex        =   28
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   705
         TabIndex        =   27
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   705
         TabIndex        =   26
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   705
         TabIndex        =   25
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
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
         Left            =   600
         TabIndex        =   6
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Pedido"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Pedidos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   11700
         _cx             =   20637
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImprimirEnBloque.frx":0376
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin NEED2.uctrVSFG uctrVSFG1 
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   4815
         _extentx        =   8493
         _extenty        =   661
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
End
Attribute VB_Name = "frmImprimirEnBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para confirmar el stock en bodega de un pedido ya realizado con ante_ #
'#  rioridad.                                                                   #
'#  frmV_VerPedBod V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Opciones que permite:                                                       #
'#  *   En una lista se despliegan los pedidos y sus detalles para que la       #
'#      persona encargada pueda ver en mayor detalle el mismo y así poder       #
'#      confirmar la cantidad de los productos que se está pidiendo.            #
'#                                                                              #
'#  Procesos internos que maneja:                                               #
'#  *   La lista que muestra los distintos pedidos se refresca automáticamente  #
'#      cada 20 segundos para buscar un nuevo pedido generado.                  #
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los   #
'#      detalles del mismo en un segundo grid.                                  #
'#  *   Se controla que el usuario pida como máximo la cantidad de productos    #
'#      solicitada a la bodega.                                                 #
'#  *   Una vez que el pedido ha sido confirmado su estado pasa a revisado.     #
'#  *   Se pueden ver solo los pedidos que aún no están revisados o los que ya  #
'#      se han revisado el día de hoy.                                          #
'#                                                                              #
'#  Tablas que maneja:                                                          #
'#                                                                              #
'#  persona:                                                                    #
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el  #
'#      pedido que se está confirmando.                                         #
'#  *   También se extrae el nombre del vendedor asignado al pedido.            #
'#  pedido:                                                                     #
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.               #
'#  det_pedido:                                                                 #
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.      #
'#                                                                              #
'################################################################################

Private clsCon_Def As New clsConsulta
Private intDato As Variant

Private Sub CargarDistribuidores()
    Dim strSQL As String
    strSQL = " SELECT '%' as codigo, '  - Todos los N1 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_gz=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbGerente.RowSource = clsCon_Def.adorec_Def
    cmbGerente.BoundColumn = "codigo"
    cmbGerente.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N2 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_di=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbDirector.RowSource = clsCon_Def.adorec_Def
    cmbDirector.BoundColumn = "codigo"
    cmbDirector.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N3 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_em=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbEmprendedor.RowSource = clsCon_Def.adorec_Def
    cmbEmprendedor.BoundColumn = "codigo"
    cmbEmprendedor.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N4 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_ee=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbEjecutivo.RowSource = clsCon_Def.adorec_Def
    cmbEjecutivo.BoundColumn = "codigo"
    cmbEjecutivo.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N5 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n5=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbN5.RowSource = clsCon_Def.adorec_Def
    cmbN5.BoundColumn = "codigo"
    cmbN5.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N6 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n6=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbN6.RowSource = clsCon_Def.adorec_Def
    cmbN6.BoundColumn = "codigo"
    cmbN6.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N7 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n7=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbN7.RowSource = clsCon_Def.adorec_Def
    cmbN7.BoundColumn = "codigo"
    cmbN7.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N8 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n8=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbN8.RowSource = clsCon_Def.adorec_Def
    cmbN8.BoundColumn = "codigo"
    cmbN8.ListField = "nombre"
    
    strSQL = " SELECT '%' as codigo, '  - Todos los N9 -' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
             " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
            " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
            " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
            " AND p1.cat_p_tipo='C'" & _
            " AND p1.per_es_n9=1 AND p1.tip_ped_codigo like '" & cmbNegocio.BoundText & "'" & _
            " ORDER BY nombre "
    clsCon_Def.Ejecutar strSQL
    Set cmbN9.RowSource = clsCon_Def.adorec_Def
    cmbN9.BoundColumn = "codigo"
    cmbN9.ListField = "nombre"
    
End Sub

Private Sub chkFechas_Click()
    If chkFechas.Value = 1 Then
        Fecha1.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha1.Enabled = False
        Fecha2.Enabled = False
    End If
End Sub

Private Sub cmbNegocio_Change()
    Dim strCli As String
    If cmbNegocio.MatchedWithList = True Then
        CargarDistribuidores
    Else
        Exit Sub
    End If
End Sub


Private Sub cmdActualizar_Click()
    Dim strEstado As String
    Dim strRed As String
    Dim strFecha As String
    Dim strFormaPago As String
    Dim strSoloCancelado As String
    Dim i As Long
    strFormaPago = ""
    For i = 1 To VSFGFormaPago.Rows - 1
        If Abs(VSFGFormaPago.TextMatrix(i, 0)) = 1 Then
            strFormaPago = strFormaPago & "'" & VSFGFormaPago.TextMatrix(i, 1) & "',"
        End If
    Next i
    If strFormaPago <> "" Then
        strFormaPago = " AND persona.for_pag_codigo IN (" & Left(strFormaPago, Len(strFormaPago) - 1) & ") "
    End If
    strSoloCancelado = ""
    If chkCancelado.Value = 1 Then
        strSoloCancelado = " INNER JOIN doc_pago ON pedido.emp_codigo=doc_pago.emp_codigo AND pedido.ped_codigo=doc_pago.ped_codigo "
    End If
    
    If optPedTodos.Value = True Then
        strEstado = " "
    ElseIf optPedGuardado.Value = True Then
        strEstado = " AND ped_estado = -1 "
    ElseIf optPedAsignado.Value = True Then
        strEstado = " AND ped_estado = 0 "
    ElseIf optPedConfirmado.Value = True Then
        strEstado = " AND ped_estado = 1 "
    ElseIf optPedFacturado.Value = True Then
        strEstado = " AND ped_estado = 2 "
    ElseIf optPedAnulado.Value = True Then
        strEstado = " AND ped_estado = 3 "
    End If
    strRed = " "
    If optN1.Value = True Then
        strRed = " AND persona.per_codigo_ref LIKE '" & Me.cmbGerente.BoundText & "' "
    ElseIf optN2.Value = True Then
        strRed = " AND persona.per_codigo_ref2 LIKE '" & Me.cmbDirector.BoundText & "' "
    ElseIf optN3.Value = True Then
        strRed = " AND persona.per_codigo_ref3 LIKE '" & Me.cmbEmprendedor.BoundText & "' "
    ElseIf optN4.Value = True Then
        strRed = " AND persona.per_codigo_ref4 LIKE '" & Me.cmbEjecutivo.BoundText & "' "
    ElseIf optN5.Value = True Then
        strRed = " AND persona.per_codigo_ref5 LIKE '" & Me.cmbN5.BoundText & "' "
    ElseIf optN6.Value = True Then
        strRed = " AND persona.per_codigo_ref6 LIKE '" & Me.cmbN6.BoundText & "' "
    ElseIf optN7.Value = True Then
        strRed = " AND persona.per_codigo_ref7 LIKE '" & Me.cmbN7.BoundText & "' "
    ElseIf optN8.Value = True Then
        strRed = " AND persona.per_codigo_ref8 LIKE '" & Me.cmbN8.BoundText & "' "
    ElseIf optN9.Value = True Then
        strRed = " AND persona.per_codigo_ref9 LIKE '" & Me.cmbN9.BoundText & "' "
    End If
    If chkFechas.Value = 1 Then
        strFecha = " AND pedido.ped_fecha BETWEEN '" & Format(Fecha1.Value, "yyyy-mm-dd hh:mm") & ":00' AND '" & Format(Fecha2.Value, "yyyy-mm-dd hh:mm") & ":59' "
    Else
        strFecha = " "
    End If
    strSQL = " SELECT pedido.ped_codigo, LEFT(CURRENT_TIMESTAMP,10) as hoy,for_pag_nombre,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per," & _
             " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1,persona.per_direccion2,ciu_nombre, " & _
             " IF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,''))," & _
             " IF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')),''))))))))) as papa,pedido.ped_usumod,IF(persona.for_pag_codigo='CONT',0,1) as orden  " & _
             " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
             " INNER JOIN persona ON pedido.emp_codigo = persona.emp_codigo AND pedido.per_codigo = persona.per_codigo INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
             strSoloCancelado
    strSQL = strSQL & " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo  " & _
             " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo  AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1" & _
             " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo  AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1" & _
             " LEFT JOIN persona N3 ON persona.emp_codigo = N3.emp_codigo  AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1" & _
             " LEFT JOIN persona N4 ON persona.emp_codigo = N4.emp_codigo  AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1" & _
             " LEFT JOIN persona N5 ON persona.emp_codigo = N5.emp_codigo  AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1" & _
             " LEFT JOIN persona N6 ON persona.emp_codigo = N6.emp_codigo  AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1" & _
             " LEFT JOIN persona N7 ON persona.emp_codigo = N7.emp_codigo  AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1" & _
             " LEFT JOIN persona N8 ON persona.emp_codigo = N8.emp_codigo  AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1" & _
             " LEFT JOIN persona N9 ON persona.emp_codigo = N9.emp_codigo  AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1" & _
             " Where pedido.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C'" & strEstado & " " & strRed & _
             " " & strFecha & strFormaPago & " AND persona.tip_ped_codigo like '" & Me.cmbNegocio.BoundText & "' " & _
             " ORDER BY orden,nn1,papa,pedido.ped_codigo "
    clsCon_Def.Ejecutar (strSQL)
    Set VSFGPeds.DataSource = clsCon_Def.adorec_Def.DataSource
    VSFGPeds.SubtotalPosition = flexSTBelow
    VSFGPeds.Subtotal flexSTCount, -1, 1, , , , True

End Sub

Private Sub cmdImprimir_Click()
    Dim strListaPed As String
    Dim i As Long
    Dim RepPed As New frmReporte
    For i = 1 To Me.VSFGPeds.Rows - 2
        strListaPed = strListaPed & VSFGPeds.TextMatrix(i, 0) & ","
    Next i
    strListaPed = Left(strListaPed, Len(strListaPed) - 1)
    RepPed.strTipo = Me.VSFGPeds.Rows - 1
    RepPed.strNumero = strListaPed
    RepPed.strReporte = "rptPedido"
    RepPed.Show
End Sub

Private Sub cmdResumen_Click()
    Dim strListaPed As String
    Dim i As Long
    Dim RepStk As New frmReporte
    For i = 1 To Me.VSFGPeds.Rows - 2
        strListaPed = strListaPed & VSFGPeds.TextMatrix(i, 0) & " ,"
    Next i
    strListaPed = Left(strListaPed, Len(strListaPed) - 1)
    RepStk.strNumero = strListaPed
    RepStk.strTipo = VSFGPeds.Rows - 1
    RepStk.strReporte = "rptResumenDespacho"
    RepStk.Show

End Sub

Private Sub cmdSticker_Click()
    Dim strListaPed As String
    Dim i As Long
    Dim RepStk As New frmReporte
    For i = 1 To Me.VSFGPeds.Rows - 2
        strListaPed = strListaPed & VSFGPeds.TextMatrix(i, 0) & " ,"
    Next i
    strListaPed = Left(strListaPed, Len(strListaPed) - 1)
    RepStk.VSPrint.PrintDialog pdPrint
    RepStk.VSPrint.PaperWidth = 7669.292
    RepStk.VSPrint.PaperHeight = 3885.039
    RepStk.strNumero = strListaPed
    RepStk.strTipo = 5
    RepStk.strReporte = "rptSTKDespacho"
    RepStk.Show
    RepStk.Form_Activate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidosV = Nothing
    Set clsPed = Nothing
    Set clsExiPrd = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub cargarTipoPedido()
    strSQL = " SELECT 0 as sel,for_pag_codigo, for_pag_nombre " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSQL
    Set VSFGFormaPago.DataSource = clsCon_Def.adorec_Def.DataSource
    
    
    strSQL = " SELECT '%' as tip_ped_codigo, '- Todos los Negocios -' as tip_ped_nombre " & _
             " UNION " & _
             " SELECT tip_ped_codigo, tip_ped_nombre " & _
             " FROM tipo_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsCon_Def.Ejecutar strSQL
    Set cmbNegocio.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = "%"
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa los objetos de conexión con la base de datos
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    Fecha1.Value = HoyDia
    Fecha2.Value = HoyDia
    Set uctrVSFG1.VSFGControl = VSFGPeds
    uctrVSFG1.Inicializar False, False, False
    
    cargarTipoPedido
End Sub


Private Sub optN1_Click()
    ActivarCombo
End Sub

Private Sub optN2_Click()
    ActivarCombo
End Sub

Private Sub optN3_Click()
    ActivarCombo
End Sub

Private Sub optN4_Click()
    ActivarCombo
End Sub

Private Sub optN5_Click()
    ActivarCombo
End Sub

Private Sub optN6_Click()
    ActivarCombo
End Sub

Private Sub optN7_Click()
    ActivarCombo
End Sub

Private Sub optN8_Click()
    ActivarCombo
End Sub

Private Sub optN9_Click()
    ActivarCombo
End Sub

Private Sub ActivarCombo()
    cmbGerente.Locked = True
    cmbDirector.Locked = True
    cmbEmprendedor.Locked = True
    cmbEjecutivo.Locked = True
    cmbN5.Locked = True
    cmbN6.Locked = True
    cmbN7.Locked = True
    cmbN8.Locked = True
    cmbN9.Locked = True
    If optN1.Value = True Then
        cmbGerente.Locked = False
    ElseIf optN2.Value = True Then
        cmbDirector.Locked = False
    ElseIf optN3.Value = True Then
        cmbEmprendedor.Locked = False
    ElseIf optN4.Value = True Then
        cmbEjecutivo.Locked = False
    ElseIf optN5.Value = True Then
        cmbN5.Locked = False
    ElseIf optN6.Value = True Then
        cmbN6.Locked = False
    ElseIf optN7.Value = True Then
        cmbN7.Locked = False
    ElseIf optN8.Value = True Then
        cmbN8.Locked = False
    ElseIf optN9.Value = True Then
        cmbN9.Locked = False
    End If
End Sub


Private Sub cmbN9_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN9.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref8")
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN8_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN8.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN7_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN7.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN6_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN6.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN5_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN5.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEjecutivo_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEjecutivo.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEmprendedor_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEmprendedor.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbDirector_Validate(Cancel As Boolean)
    strSQL = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbDirector.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbGerente_Validate(Cancel As Boolean)
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbDirector.BoundText = ""
End Sub

Private Sub VSFGFormaPago_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

