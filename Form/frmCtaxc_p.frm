VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCtaxc_p 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas por Cobrar / Pagar"
   ClientHeight    =   8910
   ClientLeft      =   7185
   ClientTop       =   1395
   ClientWidth     =   9465
   Icon            =   "frmCtaxc_p.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9465
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   105
      TabIndex        =   25
      Top             =   120
      Width           =   9255
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente / Proveedor"
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
         Height          =   1215
         Left            =   1065
         TabIndex        =   28
         Top             =   600
         Width           =   7335
         Begin VB.OptionButton Optproveedores 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Proveedores"
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
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton OptCliente 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Clientes"
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
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DCmbPersona 
            Height          =   315
            Left            =   2880
            TabIndex        =   52
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbNegocio 
            Height          =   315
            Left            =   2880
            TabIndex        =   55
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
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
         Begin VB.Label lblNegocio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Negocio:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   1920
            TabIndex        =   56
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
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
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox TxtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   7800
         Width           =   1275
      End
      Begin VB.TextBox TxtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   7800
         Width           =   1275
      End
      Begin VB.TextBox TxtNumCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos de la cuenta"
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
         Height          =   3735
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   9015
         Begin NEED2.dtpFecha dtpFecha2 
            Height          =   285
            Left            =   7440
            TabIndex        =   54
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            Value           =   42115.4018634259
         End
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   285
            Left            =   7440
            TabIndex        =   53
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            Value           =   42115.4018634259
         End
         Begin VB.TextBox txtSTServ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtSTProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   10
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtICE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtBaseICE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtSTcero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            TabIndex        =   16
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            TabIndex        =   12
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAServ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            TabIndex        =   13
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtdocumento 
            Height          =   285
            Left            =   2640
            TabIndex        =   4
            Top             =   570
            Width           =   1455
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            Top             =   570
            Width           =   1095
         End
         Begin VB.TextBox txtAutorizacion 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TxtObservacion 
            Height          =   735
            Left            =   1560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   2880
            Width           =   5655
         End
         Begin MSDataListLib.DataCombo dcmbTipoDoc 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpCaduca 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   6
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/yyyy"
            Format          =   65929219
            CurrentDate     =   37463
         End
         Begin MSDataListLib.DataCombo dcmbICE 
            Height          =   315
            Left            =   4680
            TabIndex        =   8
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbIVA 
            Height          =   315
            Left            =   7440
            TabIndex        =   14
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbSustento 
            Height          =   315
            Left            =   6030
            TabIndex        =   48
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal Serv:"
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
            Left            =   405
            TabIndex        =   51
            Top             =   2070
            Width           =   1065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal Prod:"
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
            Left            =   420
            TabIndex        =   50
            Top             =   1710
            Width           =   1050
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Sustento:"
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
            Left            =   5040
            TabIndex        =   49
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "% IVA:"
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
            Left            =   6840
            TabIndex        =   47
            Top             =   1710
            Width           =   510
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Total ICE:"
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
            Left            =   6690
            TabIndex        =   46
            Top             =   1350
            Width           =   660
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICE:"
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
            Left            =   795
            TabIndex        =   45
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "% ICE:"
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
            Left            =   4140
            TabIndex        =   44
            Top             =   1350
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Caducidad del Doc:"
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
            Left            =   3195
            TabIndex        =   43
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   6795
            TabIndex        =   42
            Top             =   2430
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal 0%:"
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
            Left            =   3690
            TabIndex        =   41
            Top             =   2430
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Prod:"
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
            Left            =   3240
            TabIndex        =   40
            Top             =   1710
            Width           =   1365
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Serv:"
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
            Left            =   3225
            TabIndex        =   39
            Top             =   2070
            Width           =   1380
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Total IVA:"
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
            Left            =   6645
            TabIndex        =   38
            Top             =   2070
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de pago:"
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
            Left            =   6225
            TabIndex        =   37
            Top             =   945
            Width           =   1125
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de documento:"
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
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de emisión:"
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
            Left            =   6045
            TabIndex        =   35
            Top             =   570
            Width           =   1305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Autorizacion:"
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
            Left            =   0
            TabIndex        =   34
            Top             =   960
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Doc."
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
            Left            =   570
            TabIndex        =   33
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label8 
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
            Left            =   495
            TabIndex        =   27
            Top             =   2880
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGAsientos 
         Height          =   2055
         Left            =   120
         TabIndex        =   19
         Top             =   5760
         Width           =   9000
         _cx             =   28720643
         _cy             =   28708393
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
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCtaxc_p.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
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
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         Left            =   3600
         TabIndex        =   31
         Top             =   7815
         Width           =   735
      End
      Begin VB.Label LblNumCuenta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta por Cobrar No.:"
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
         Height          =   255
         Left            =   1065
         TabIndex        =   30
         Top             =   255
         Width           =   2055
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   360
         Picture         =   "frmCtaxc_p.frx":03E0
         ToolTipText     =   "Elimina una Fila"
         Top             =   7920
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   120
         Picture         =   "frmCtaxc_p.frx":0516
         Top             =   7920
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4785
      TabIndex        =   21
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3105
      TabIndex        =   20
      Top             =   8400
      Width           =   1575
   End
End
Attribute VB_Name = "frmCtaxc_p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsPersona As New clsConsulta
Private clsSql As New clsConsulta
Private clsAsi As New clsConsulta
Private clsParametro As New clsConsulta
Dim strSql As String
Private Var_NumCuenta As Long

Private Sub cmbNegocio_Validate(Cancel As Boolean)
    Llena_Cliente
End Sub

Private Sub dcmbICE_Change()
    TotalFac
End Sub

Private Sub dcmbIVA_Change()
    TotalFac
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPersona = Nothing
    Set clsSql = Nothing
    Set clsAsi = Nothing
    Set clsParametro = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la seginda columna del grid de todas las filas
    For i = 3 To (VSFGAsientos.Rows - 1)
        VSFGAsientos.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFGAsientos.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFGAsientos.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub

Private Sub Llena_Cliente()
        
        'Llena Combo de clientes
        dcmbPersona.Text = ""
        'DCmbNomPersona.Text = ""
        
        clsPersona.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (', tip_ped_nombre ,')',' (', per_ruc ,')') as nomb " & _
                 " FROM persona INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                 " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' AND cat_p_tipo='C'" & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " ORDER BY per_apellido,per_nombre"
        clsPersona.Ejecutar (strSql)
        If clsPersona.adorec_Def.EOF = False Then
            Set dcmbPersona.RowSource = clsPersona.adorec_Def.DataSource
            dcmbPersona.ListField = "nomb"
            dcmbPersona.BoundColumn = "per_codigo"
        Else
            Set dcmbPersona.RowSource = Nothing
        End If
End Sub

Private Sub Llena_Proveedor()
        
        'Llena Combo de proveedores
        dcmbPersona.Text = ""
        'DCmbNomPersona.Text = ""
        
        clsPersona.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (', per_ruc ,')') as nomb " & _
             " From persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='P'" & _
             " ORDER BY per_apellido,per_nombre"
        clsPersona.Ejecutar (strSql)
        If clsPersona.adorec_Def.EOF = False Then
            Set dcmbPersona.RowSource = clsPersona.adorec_Def.DataSource
            dcmbPersona.ListField = "nomb"
            dcmbPersona.BoundColumn = "per_codigo"

'            VSFGAsientos.TextMatrix(1, 1) = clsPersona.adorec_Def("per_ctaconta")
'            VSFGAsientos.TextMatrix(1, 2) = clsPersona.adorec_Def("cta_nombre")
        Else
            Set dcmbPersona.RowSource = Nothing
        End If
End Sub

Private Sub cmdAceptar_Click()
    Dim fechaEmision As String
    Dim FechaPago As String
    Dim FechaHoy As String
    Dim TotalCuenta As Double
    Dim ban As Boolean
    Dim clsAsiento As New clsContable
    
    TotalCuenta = 0
    'Control de fechas válidas
    fechaEmision = dtpFecha.Value
    FechaPago = dtpFecha2.Value
    FechaHoy = Format(HoyDia, "yyyy-mm-dd")
    If Not IsDate(fechaEmision) Then
        MsgBox "Fecha de Emisión NO válida", vbCritical
        Exit Sub
    End If
    
    If Not IsDate(FechaPago) Then
        MsgBox "Fecha de Pago NO válida", vbCritical
        Exit Sub
    End If
    
    If Trim(txtSerie.Text) = "" Or Trim(txtDocumento.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
        MsgBox "Debe llenar todos los datos del Docuento", vbInformation, "Pagos"
        Exit Sub
    End If
    
    TotalCuenta = FormatoD2(txtValor.Text)
    Var_Tipo = Me.Tag
    
    'Verificar que todos los datos se han llenado para ingresar en la base de datos
    If dcmbPersona.Text = "" Or txtDocumento = "" Or txtTotalDebe = "0,00" Or txtTotalHaber = "0,00" Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "Ingreso"
            dcmbPersona.SetFocus
            
    ElseIf Not Val(txtTotalDebe) = Val(txtTotalHaber) Then
            MsgBox "No cuadran los valores de las cuentas", vbInformation, "Ingreso"
            Exit Sub
    Else
            ban = True
'            If MsgBox("Desea generar el asiento?", vbYesNo, "Contabilidad") = vbYes Then
'                ban = True
'            End If
                If ban = True Then
                'Busca el código máximo de la tabla asiento
                    clsAsiento.Inicializar AdoConn, AdoConnMaster
                    clsAsiento.NuevoAsiento "D", fechaEmision, 0, 0, FormatoD2(txtTotalDebe.Text), _
                    "Persona: " & dcmbPersona.Text & vbNewLine & _
                    dcmbTipoDoc.Text & ": " & txtSerie.Text & " " & txtDocumento.Text & " Aut: " & txtAutorizacion.Text & vbNewLine & _
                    txtObservacion.Text
                
                    strMaximo = clsAsiento.NumAsiento
                End If
                
            'Ingreso de datos en cuenta_p_c
                Set clsIngCuentas = New clsConsulta
                clsIngCuentas.Inicializar AdoConn, AdoConnMaster
                Dim clsCxX As New clsCtaXx
                clsCxX.Inicializar AdoConn, AdoConnMaster
                clsCxX.NuevaCta Me.Tag, dcmbTipoDoc.BoundText, dcmbSustento.BoundText, fechaEmision, FechaPago, dcmbPersona.BoundText, txtObservacion.Text, txtSerie.Text, txtDocumento.Text, txtAutorizacion.Text, Format(dtpCaduca.Value, "mm/yyyy"), txtSTProd.Text, txtSTServ.Text, txtSTIVAProd.Text, txtSTIVAServ.Text, dcmbIVA.BoundText, TxtIva.Text, txtSTcero.Text, txtBaseICE.Text, dcmbICE.BoundText, txtICE.Text, txtValor.Text, clsAsiento.NumAsiento
                Me.TxtNumCuenta.Text = clsCxX.strNoCta
            'Ingreso de Detalle de asientos
            If ban = True Then
                With VSFGAsientos
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                            clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                        End If
                    Next i
                End With
            End If
            MsgBox " Los datos han sido ingresados", vbInformation, "Ingresos"
            Var_Tipo_Cuenta = Me.Tag
            
            '*******************
            If Me.Tag = "P" Then
                frmPagoRetencion.Tag = Me.Tag
                frmPagoRetencion.txtBeneficiario.Text = Me.dcmbPersona.Text
                frmPagoRetencion.txtCX.Text = Me.TxtNumCuenta.Text
                frmPagoRetencion.txtDocumento.Text = Me.dcmbTipoDoc.Text & " (" & Me.txtSerie.Text & Format(Me.txtDocumento.Text, "0000000") & ")"
                frmPagoRetencion.txtSTcero.Text = Me.txtSTcero.Text
                frmPagoRetencion.txtSTIVAProd.Text = Me.txtSTIVAProd.Text
                frmPagoRetencion.txtSTIVAServ.Text = Me.txtSTIVAServ.Text
                frmPagoRetencion.TxtIva.Text = Me.TxtIva.Text
                frmPagoRetencion.txtValor.Text = Me.txtValor.Text
                frmPagoRetencion.txtDocumento.Tag = VSFGAsientos.TextMatrix(1, 1)
                If ban = True Then
                    frmPagoRetencion.VSFG.Tag = strMaximo
                Else
                    frmPagoRetencion.VSFG.Tag = "NO"
                End If
                frmPagoRetencion.dtpFecha.Value = fechaEmision
                frmPagoRetencion.Show
            Else
            
                Dim Asien As New frmReporte
                Asien.strAsiento = clsAsiento.NumAsiento
                Asien.strReporte = "rptAsiento"
                Asien.Show
            End If
            '*******************
            Unload Me
        End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub






Private Sub DCmbPersona_Change()
 Dim strSql As String
 Dim Fact As String
    clsSql.Inicializar AdoConn, AdoConnMaster
    If clsPersona.adorec_Def.RecordCount > 0 Then
        clsPersona.adorec_Def.MoveFirst
    End If
    clsPersona.adorec_Def.Find "per_codigo = '" & dcmbPersona.BoundText & "'", , adSearchForward
    dcmbPersona.Tag = "A"
    If clsPersona.adorec_Def.EOF = True Then

    Else
       If Me.Tag = "C" Then
            
            strSql = " SELECT tip_ped_ptofac " & _
                     " FROM tipo_pedido " & _
                     " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                Fact = clsSql.adorec_Def(0)
            End If
            txtSerie.Text = strSucursal & Fact
            dtpCaduca.Value = Format(HoyDia, "MM/yyyy")
            txtAutorizacion.Text = "0"
        Else
            strSql = " SELECT cue_p_c_autorizacion,cue_p_c_caduca,cue_p_c_serie " & _
                     " FROM cuenta_p_c  " & _
                     " WHERE emp_codigo = '" & strEmpresa & "' AND cue_p_c_tipo='" & Me.Tag & "' AND per_codigo='" & dcmbPersona.BoundText & "' ORDER BY cue_p_c_fechaemision DESC"
                     'cue_p_c_numero DESC LIMIT 0,1
            
            clsSql.Ejecutar strSql
            If Not clsSql.adorec_Def.EOF Then
                txtSerie.Text = clsSql.adorec_Def("cue_p_c_serie")
                If clsSql.adorec_Def("cue_p_c_caduca") <> "00/0000" And clsSql.adorec_Def("cue_p_c_caduca") <> "" Then
                    dtpCaduca.Value = clsSql.adorec_Def("cue_p_c_caduca")
                End If
                txtAutorizacion.Text = clsSql.adorec_Def("cue_p_c_autorizacion")
            Else
                txtSerie.Text = ""
                dtpCaduca.Value = Format(HoyDia, "MM/yyyy")
                txtAutorizacion.Text = ""
            End If
        End If
    End If
    strSql = " SELECT cat_p_ctaconta, cta_nombre " & _
             " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
             " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
             " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
             " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
             " WHERE persona.per_codigo= '" & dcmbPersona.BoundText & "' AND persona.emp_codigo = '" & strEmpresa & "' "
    clsSql.Ejecutar strSql
    If Not clsSql.adorec_Def.EOF Then
        VSFGAsientos.TextMatrix(1, 1) = clsSql.adorec_Def("cat_p_ctaconta")
        'VSFGAsientos.TextMatrix(1, 2) = clsSql.adorec_Def("cta_nombre")
        VSFGAsientos.TextMatrix(1, 2) = clsSql.adorec_Def("cat_p_ctaconta")
    Else
        VSFGAsientos.TextMatrix(1, 1) = ""
        VSFGAsientos.TextMatrix(1, 2) = ""
    End If
    dcmbPersona.Tag = ""

End Sub

Private Sub Form_Activate()
    
    If Me.Tag = "C" Then
        LblTitulo.Caption = "Cuentas por Cobrar"
        LblNumCuenta.Caption = "Cuenta por Cobrar No.:"
        Me.Caption = "Cuentas por Cobrar"
        strSql = " SELECT cat_p_ctaconta, cta_nombre " & _
                 " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                 " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                 " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                 " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                 " WHERE persona.per_codigo= '" & dcmbPersona.BoundText & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsSql.Ejecutar strSql
        If Not clsSql.adorec_Def.EOF Then
            VSFGAsientos.TextMatrix(1, 1) = clsSql.adorec_Def("cat_p_ctaconta")
            VSFGAsientos.TextMatrix(1, 2) = clsSql.adorec_Def("cta_nombre")
        Else
            VSFGAsientos.Clear 1
            VSFGAsientos.Rows = 3
        End If
    
        'Consulta para saber el iva de ventas
        strSql = " SELECT par_numero,par_texto,cta_nombre " & _
                 " FROM parametro INNER JOIN ctaconta ON parametro.emp_codigo=ctaconta.emp_codigo AND parametro.par_texto=ctaconta.cta_codigo " & _
                 " WHERE parametro.emp_codigo='" & strEmpresa & "' " & _
                 " AND par_codigo='IVAV'"
        clsSql.Ejecutar strSql
        If Not clsSql.adorec_Def.EOF Then
            VSFGAsientos.TextMatrix(2, 1) = clsSql.adorec_Def("par_texto")
            VSFGAsientos.TextMatrix(2, 2) = clsSql.adorec_Def("cta_nombre")
        Else
            VSFGAsientos.Clear 1
            VSFGAsientos.Rows = 3
        End If
        TxtIva.Tag = FormatoD2(clsSql.adorec_Def("par_numero"))
    ElseIf Me.Tag = "P" Then
        LblTitulo.Caption = "Cuentas por Pagar"
        LblNumCuenta.Caption = "Cuenta por Pagar No.:"
        Me.Caption = "Cuentas por Pagar"
        
        strSql = " SELECT cat_p_ctaconta, cta_nombre " & _
                 " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                 " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                 " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                 " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                 " WHERE persona.per_codigo= '" & dcmbPersona.BoundText & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsSql.Ejecutar strSql
        If Not clsSql.adorec_Def.EOF Then
            VSFGAsientos.TextMatrix(1, 1) = clsSql.adorec_Def("cat_p_ctaconta")
            VSFGAsientos.TextMatrix(1, 2) = clsSql.adorec_Def("cta_nombre")
        Else
            VSFGAsientos.Clear 1
            VSFGAsientos.Rows = 3
        End If
        'Consulta para saber el iva de compras
        strSql = " SELECT COALESCE(par_numero,0) as par_numero,COALESCE(par_texto,'') as par_texto,COALESCE(cta_nombre,'') as cta_nombre " & _
                 " FROM parametro INNER JOIN ctaconta ON parametro.emp_codigo=ctaconta.emp_codigo AND parametro.par_texto=ctaconta.cta_codigo " & _
                 " WHERE parametro.emp_codigo='" & strEmpresa & "' " & _
                 " AND par_codigo='IVAC'"
        clsSql.Ejecutar strSql
        If Not clsSql.adorec_Def.EOF Then
            VSFGAsientos.TextMatrix(2, 1) = clsSql.adorec_Def("par_texto")
            'VSFGAsientos.TextMatrix(2, 2) = clsSql.adorec_Def("cta_nombre")
            VSFGAsientos.TextMatrix(2, 2) = clsSql.adorec_Def("par_texto")
        Else
            VSFGAsientos.Clear 1
            VSFGAsientos.Rows = 3
        End If
        If clsSql.adorec_Def.RecordCount > 0 Then
            TxtIva.Tag = FormatoD2(clsSql.adorec_Def("par_numero"))
        Else
            TxtIva.Tag = FormatoD2(0)
        End If
    End If
    Llena_Impuestos
     'LLena los datos de la primera fila
    Tipo_Cuenta
     'LLena los combolist en las columnas 1 y 2 despues de la primera fila
    LLena_CombosGrid
    AutoNumero_Cuenta
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub

Private Sub Form_Load()
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsParametro.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    

    cargarTipoPedido
     
    Llena_Cliente

    'Llena combos de fecha
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFecha2.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpCaduca = Format(HoyDia, "mm/yyyy")
    'Consulta para saber los tipos de documentos
    strSql = " SELECT cod_sus_com_codigo, cod_sus_com_nombre " & _
             " FROM codigo_sustento_comprobante " & _
             " ORDER BY cod_sus_com_codigo "
    clsAsi.Ejecutar strSql

    Set dcmbSustento.RowSource = clsAsi.adorec_Def.DataSource
    dcmbSustento.ListField = "cod_sus_com_nombre"
    dcmbSustento.BoundColumn = "cod_sus_com_codigo"
    'Consulta para saber los tipos de documentos
    strSql = " SELECT tip_doc_cue_codigo, tip_doc_cue_descripcion " & _
             " FROM tipo_doc_cuenta "
    clsAsi.Ejecutar strSql

    Set dcmbTipoDoc.RowSource = clsAsi.adorec_Def.DataSource
    dcmbTipoDoc.ListField = "tip_doc_cue_descripcion"
    dcmbTipoDoc.BoundColumn = "tip_doc_cue_codigo"
    
End Sub

Private Sub AutoNumero_Cuenta()
    'Pone el número de cuenta por cobrar / pagar siguiente
    Set clsNumCuentas = New clsConsulta
    clsNumCuentas.Inicializar AdoConn, AdoConnMaster
    If Me.Tag = "C" Then
        strSql = " SELECT COALESCE(max(cue_p_c_codigo),0) as num_cuenta" & _
                 " FROM cuenta_p_c " & _
                 " WHERE cue_p_c_tipo = 'C' AND emp_codigo = '" & strEmpresa & "'" & _
                 " GROUP BY emp_codigo"
    ElseIf Me.Tag = "P" Then
        strSql = " SELECT COALESCE(max(cue_p_c_codigo),0) as num_cuenta" & _
                 " FROM cuenta_p_c " & _
                 " WHERE cue_p_c_tipo = 'P' AND emp_codigo = '" & strEmpresa & "'" & _
                 " GROUP BY emp_codigo"
    End If
    clsNumCuentas.Ejecutar (strSql)
    If Not clsNumCuentas.adorec_Def.EOF Then
        Var_NumCuenta = Val(clsNumCuentas.adorec_Def("num_cuenta")) + 1
    Else
        Var_NumCuenta = 1
    End If
    TxtNumCuenta.Text = Var_NumCuenta
End Sub

Private Sub Llena_Impuestos()
    'Pone los combolist en las columnas 1 y 2 despues de la primera fila
    Dim au As Integer
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " WHERE cod_iva_enuso=1 "
     clsSql.Ejecutar (strSql)
     au = clsSql.adorec_Def("cod_iva_codigo")
     
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " ORDER BY cod_iva_porcentaje"
     clsSql.Ejecutar (strSql)
     Set dcmbIVA.RowSource = clsSql.adorec_Def.DataSource
     dcmbIVA.ListField = "cod_iva_porcentaje"
     dcmbIVA.BoundColumn = "cod_iva_codigo"
     dcmbIVA.BoundText = au
     strSql = " SELECT cod_ice_codigo, cod_ice_porcentaje" & _
                 " FROM codigo_ice " & _
                 " WHERE cod_ice_enuso=1 " & _
                 " ORDER BY cod_ice_porcentaje"
     clsSql.Ejecutar (strSql)
     Set dcmbICE.RowSource = clsSql.adorec_Def.DataSource
     dcmbICE.ListField = "cod_ice_porcentaje"
     dcmbICE.BoundColumn = "cod_ice_codigo"
     dcmbICE.BoundText = clsSql.adorec_Def("cod_ice_codigo")
End Sub

Private Sub LLena_CombosGrid()
    'Pone los combolist en las columnas 1 y 2 despues de la primera fila
    Set clsCuentas = New clsConsulta
    clsCuentas.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsCuentas.Ejecutar (strSql)
     VSFGAsientos.ColComboList(5) = VSFGAsientos.BuildComboList(clsCuentas.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
     
    strSql = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCuentas.Ejecutar (strSql)
              
     VSFGAsientos.ColComboList(1) = VSFGAsientos.BuildComboList(clsCuentas.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFGAsientos.ColComboList(2) = VSFGAsientos.BuildComboList(clsCuentas.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")
End Sub


Private Sub OptCliente_Click()
    cmbNegocio.Visible = True
    lblNegocio.Visible = True
    Llena_Cliente
End Sub

Private Sub Optproveedores_Click()
    cmbNegocio.Visible = False
    lblNegocio.Visible = False
    Llena_Proveedor
End Sub

Private Sub Tipo_Cuenta()

   If Me.Tag = "C" Then
        'Consulta parametros cuenta por cobrar
        strSql = " SELECT parametro.par_codigo, parametro.par_texto, ctaconta.cta_codigo, ctaconta.cta_nombre, ctaconta.emp_codigo" & _
                 " FROM parametro, ctaconta " & _
                 " WHERE parametro.par_codigo = 'CXC' AND ctaconta.emp_codigo=parametro.emp_codigo AND parametro.emp_codigo = '" & strEmpresa & "'" & _
                 " AND parametro.par_texto=ctaconta.cta_codigo "
    ElseIf Me.Tag = "P" Then
        'Consulta parametros cuenta por pagar
        strSql = " SELECT parametro.par_codigo, parametro.par_texto, ctaconta.cta_codigo, ctaconta.cta_nombre, ctaconta.emp_codigo" & _
                 " FROM parametro, ctaconta " & _
                 " WHERE parametro.par_codigo = 'CXP' AND ctaconta.emp_codigo=parametro.emp_codigo AND parametro.emp_codigo = '" & strEmpresa & "'" & _
                 " AND parametro.par_texto=ctaconta.cta_codigo "
    End If
    clsParametro.Ejecutar (strSql)
    VSFGAsientos.TextMatrix(1, 0) = "1"
    VSFGAsientos.TextMatrix(2, 0) = "2"
    If clsParametro.adorec_Def.EOF = False Then
        VSFGAsientos.TextMatrix(1, 1) = clsParametro.adorec_Def("par_texto")
        'VSFGAsientos.TextMatrix(1, 2) = clsparametro.adorec_Def("cta_nombre")
        VSFGAsientos.TextMatrix(1, 2) = clsParametro.adorec_Def("par_texto")
    End If
    Set clsCta = New clsConsulta
    clsCta.Inicializar AdoConn, AdoConnMaster
    strSql1 = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar (strSql1)
End Sub

Private Sub txtdocumento_Validate(Cancel As Boolean)
    If RevisaFac = True Then txtDocumento.Text = ""
End Sub

Private Function RevisaFac() As Boolean
    Dim clsConFac As New clsConsulta
    clsConFac.Inicializar AdoConn, AdoConnMaster
    strSql = "SELECT COUNT(*) as n FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND cue_p_c_tipo='" & Me.Tag & "'" & _
             " AND cue_p_c_numero='" & FormatoD0(txtDocumento.Text) & "'" & _
             " AND cue_p_c_serie='" & txtSerie.Text & "'" & _
             " AND cue_p_c_valor!=0 AND tip_doc_cue_codigo='" & Me.dcmbTipoDoc.BoundText & "'" & _
             " AND per_codigo='" & dcmbPersona.BoundText & "'"
    clsConFac.Ejecutar strSql
    RevisaFac = False
    If clsConFac.adorec_Def("n") > 0 Then
        RevisaFac = True
    End If
End Function

Private Sub txtSerie_Validate(Cancel As Boolean)
    If RevisaFac = True Then txtSerie.Text = ""
End Sub

Private Sub txtSTcero_Validate(Cancel As Boolean)
    TotalFac
End Sub

Private Sub txtSTIVAProd_Validate(Cancel As Boolean)
    TotalFac
End Sub

Private Sub txtSTIVAServ_Validate(Cancel As Boolean)
    TotalFac
End Sub

Private Sub txtSTProd_Validate(Cancel As Boolean)
    txtSTIVAProd.Text = FormatoD2(txtSTProd.Text)
    TotalFac
End Sub

Private Sub txtSTServ_Validate(Cancel As Boolean)
    txtSTIVAServ.Text = FormatoD2(txtSTServ.Text)
    TotalFac
End Sub

Private Sub txtBaseICE_Validate(Cancel As Boolean)
    TotalFac
End Sub

Private Sub txtPorICE_Validate(Cancel As Boolean)
    TotalFac
End Sub
Private Sub TotalFac()
    txtSTServ = FormatoD2(txtSTServ.Text)
    txtSTProd = FormatoD2(txtSTProd.Text)
    'txtSTcero.Text = FormatoD2(txtSTProd.Text) - FormatoD2(txtSTIVAProd.Text) + FormatoD2(txtSTServ.Text) - FormatoD2(txtSTIVAServ.Text)
    txtSTcero.Text = FormatoD2(txtSTcero.Text)
    txtSTIVAProd.Text = FormatoD2(txtSTIVAProd.Text)
    txtSTIVAServ.Text = FormatoD2(txtSTIVAServ.Text)
    TxtIva.Text = FormatoD2(FormatoD2(FormatoD2(txtSTIVAProd.Text) + FormatoD2(txtSTIVAServ.Text)) * FormatoD2(dcmbIVA.Text) / 100#)
    txtBaseICE.Text = FormatoD2(txtBaseICE.Text)
    txtICE.Text = FormatoD2(FormatoD2(txtBaseICE.Text) * FormatoD2(dcmbICE.Text) / 100#)
    txtValor.Text = FormatoD2(txtSTcero.Text) + FormatoD2(txtSTIVAProd.Text) + FormatoD2(txtSTIVAServ.Text) + FormatoD2(TxtIva.Text) + FormatoD2(txtICE.Text)
    If Me.Tag = "P" Then
        VSFGAsientos.TextMatrix(1, 4) = txtValor.Text
        VSFGAsientos.TextMatrix(2, 3) = TxtIva.Text
    Else
        VSFGAsientos.TextMatrix(1, 3) = txtValor.Text
        VSFGAsientos.TextMatrix(2, 4) = TxtIva.Text
    End If
End Sub

Private Sub VSFGAsientos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Not IsNumeric(VSFGAsientos.TextMatrix(Row, 3)) And VSFGAsientos.TextMatrix(Row, 3) <> "" Then
        MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
        VSFGAsientos.TextMatrix(Row, 3) = ""
    End If
    If Not IsNumeric(VSFGAsientos.TextMatrix(Row, 4)) And VSFGAsientos.TextMatrix(Row, 4) <> "" Then
        MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
        VSFGAsientos.TextMatrix(Row, 4) = ""
    End If
    Calcula_Total
End Sub

Private Sub Calcula_Total()
        'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    
    'Calcula total debe
    
    For i = 1 To VSFGAsientos.Rows - 1
        SumaDebe = SumaDebe + Val(VSFGAsientos.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    
    'Calcula total haber
    
    For i = 1 To VSFGAsientos.Rows - 1
        SumaHaber = SumaHaber + Val(VSFGAsientos.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")
End Sub

Private Sub VSFGAsientos_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <> OldRow Then
        If Year(dtpFecha.Value) >= 2018 And VSFGAsientos.Rows > 1 Then
            If (Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "4" Or Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "5" Or Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "6") And VSFGAsientos.TextMatrix(VSFGAsientos.Row, 5) = "" Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub VSFGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del grid, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes
    
    If VSFGAsientos.Row = VSFGAsientos.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1) <> "" And (VSFGAsientos.TextMatrix(VSFGAsientos.Row, 3) <> "" Or VSFGAsientos.TextMatrix(VSFGAsientos.Row, 4) <> "") Then
            If Year(dtpFecha.Value) >= 2018 And VSFGAsientos.Rows > 1 Then
                If Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) <> "4" And Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) <> "5" And Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) <> "6" Then
                    VSFGAsientos.AddItem ""
                ElseIf VSFGAsientos.TextMatrix(VSFGAsientos.Row, 5) <> "" Then
                    VSFGAsientos.AddItem ""
                End If
            Else
                VSFGAsientos.AddItem ""
            End If

            
            VSFGAsientos.TextMatrix(VSFGAsientos.Rows - 1, 0) = VSFGAsientos.Rows - 1
            VSFGAsientos.Cell(flexcpPictureAlignment, (VSFGAsientos.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub

Private Sub VSFGAsientos_CellChanged(ByVal Row As Long, ByVal Col As Long)
' Cambia el nombre y codigo de cuenta para los combos del grid escogidos
If Row > 0 Then
        Set clsCuentas = New clsConsulta
        clsCuentas.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
        clsCuentas.Ejecutar (strSql)
        
    With VSFGAsientos
        If .TextMatrix(Row, Col) <> "" Then
             If Col = 1 Then
                .TextMatrix(Row, 2) = .TextMatrix(Row, 1)
             End If
             
             If Col = 2 Then
                
                .TextMatrix(Row, 1) = .TextMatrix(Row, 2)
                
             End If
         End If
    End With
End If
End Sub

Private Sub VSFGAsientos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 1 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGAsientos_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            frmSelecCtaConta.Show
            Set frmSelecCtaConta.objEscribir = VSFGAsientos
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFGasientos_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

    ' only interesetd in left button
    If Button <> 1 Then Exit Sub

    ' get cell that was clicked
    Dim r&, c&
    r = VSFGAsientos.MouseRow
    c = VSFGAsientos.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub

    If (c <> 0 Or r = 1 Or r = 2) Then Exit Sub

    ' make sure the click was on a cell with a button
    If r > 0 Then
    If c > 1 Then
    If VSFGAsientos.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    End If
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = VSFGAsientos.Cell(flexcpLeft, r, c) + VSFGAsientos.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
        If r > 1 Then
        ' click was on a button: do the work
        VSFGAsientos.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Eliminar"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)

    'Recorro el FlexGrid para poner números a las filas

        If respuesta = vbYes Then
            Dim i As Integer
            VSFGAsientos.RemoveItem (r)
            PonerBotones
            Calcula_Total
        Else
            VSFGAsientos.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFGAsientos_Validate(Cancel As Boolean)
    If Year(dtpFecha.Value) >= 2018 And VSFGAsientos.Rows > 1 Then
        If (Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "4" Or Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "5" Or Left(VSFGAsientos.TextMatrix(VSFGAsientos.Row, 1), 1) = "6") And VSFGAsientos.TextMatrix(VSFGAsientos.Row, 5) = "" Then
            Cancel = True
        End If
    End If
End Sub
