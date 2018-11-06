VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmail 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de Mail"
   ClientHeight    =   7590
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7875
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7875
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos de Copia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   43
      Top             =   2640
      Width           =   6015
      Begin VB.TextBox txtCc 
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4200
      End
      Begin VB.TextBox txtCcName 
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label lblCC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   45
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblCcName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   44
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos de quien Recibe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   40
      Top             =   1680
      Width           =   6015
      Begin VB.TextBox txtToName 
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   4200
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4200
      End
      Begin VB.Label lblToName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   42
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   41
         Top             =   660
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos de quien Envia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   35
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txtFromName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   4200
      End
      Begin VB.TextBox txtFrom 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   4200
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   39
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   38
         Top             =   660
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   6420
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox txtPopServer 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   420
      Width           =   4200
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   720
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBcc 
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3660
      Width           =   4200
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   6420
      TabIndex        =   27
      Top             =   1620
      Width           =   1335
      Begin VB.CheckBox ckPopLogin 
         BackColor       =   &H00DDDDDD&
         Caption         =   "POP Login"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   2100
         Width           =   1095
      End
      Begin VB.CheckBox ckReceipt 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Receipt"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Request a Return Receipt"
         Top             =   1510
         Width           =   1035
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         ItemData        =   "frmEmail.frx":030A
         Left            =   120
         List            =   "frmEmail.frx":030C
         TabIndex        =   14
         Text            =   "cboPriority"
         ToolTipText     =   "Sets the Prioirty of the Mail Message"
         Top             =   840
         Width           =   1055
      End
      Begin VB.CheckBox ckHtml 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Html"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Mail Body is HTML / Plain Text"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   29
         Top             =   3180
         Width           =   1055
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2640
         Width           =   1055
      End
      Begin VB.CheckBox ckLogin 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Login"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   1800
         Width           =   915
      End
      Begin VB.OptionButton optEncodeType 
         BackColor       =   &H00DDDDDD&
         Caption         =   "MIME"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   110
         TabIndex        =   12
         ToolTipText     =   "Use MIME encoding for Mail & Attachments."
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optEncodeType 
         BackColor       =   &H00DDDDDD&
         Caption         =   "UUEncode"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   110
         TabIndex        =   13
         ToolTipText     =   "Use UU Encoding for Attachments."
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Password:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Username:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2460
         Width           =   975
      End
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   840
      Left            =   1980
      TabIndex        =   25
      Top             =   6120
      Width           =   4200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   6420
      TabIndex        =   8
      Top             =   5700
      Width           =   1275
   End
   Begin VB.TextBox txtAttach 
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5700
      Width           =   4200
   End
   Begin VB.TextBox txtMsg 
      Height          =   1260
      Left            =   1980
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4380
      Width           =   4200
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4020
      Width           =   4200
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   75
      Width           =   4200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   6420
      TabIndex        =   10
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   6420
      TabIndex        =   9
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblPopServer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor POP3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   33
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblBcc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bcc: Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   32
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   26
      Top             =   6180
      Width           =   465
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3780
      TabIndex        =   24
      Top             =   7080
      Width           =   870
   End
   Begin VB.Label lblAttach 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Adjunto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   23
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   22
      Top             =   4380
      Width           =   600
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   21
      Top             =   4020
      Width           =   510
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor SMTP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   795
      TabIndex        =   20
      Top             =   105
      Width           =   1035
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean
Private SiEnvio As Boolean
'Private poSendMail As clsSendMail

Public Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdSend_Click()

    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************
    Dim strSMTP As String
    Dim strSMTPPORT As String
    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_NONE     'VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
        
        If InStr(1, txtServer.Text, ":") > 0 Then
            strSMTP = Mid(txtServer.Text, 1, InStr(1, txtServer.Text, ":") - 1)
            strSMTPPORT = Mid(txtServer.Text, InStr(1, txtServer.Text, ":") + 1)
        Else
            strSMTP = txtServer.Text
            strSMTPPORT = 25
        End If
        
        
        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = strSMTP                  ' Required the fist time, optional thereafter
        .SMTPPort = strSMTPPORT                            ' Optional, default = 25
        .From = txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtFromName.Text         ' Optional, saved after first use
        .Recipient = txtTo.Text                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = txtToName.Text      ' Optional, separate multiple entries with delimiter character
        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = txtSubject.Text                  ' Optional
        .Message = txtMsg.Text                      ' Optional
        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .UserName = txtUserName                     ' Optional, default = Null String
        .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
         .ConnectTimeout = 20                      ' Optional, default = 10
         .ConnectRetry = 10                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
EnviarOtraVez:
        SiEnvio = True
        .send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        If SiEnvio = False Then
            If MsgBox("NO SE ENVIO EL CORREO, DESEA VOLVER A INTENTAR EL ENVIO?", vbQuestion + vbYesNo, "Envio de Correo") = vbYes Then
                GoTo EnviarOtraVez
            
            End If
        End If
        txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    cmdExit_Click
End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    SiEnvio = False
    cmdSend.Enabled = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    ' MsgBox "Send Successful!"
    lblProgress = ""

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()

    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************
    Set poSendMail = New clsSendMail
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width - 200)
    Me.Top = 0


    lblProgress = ""

    cboPriority.AddItem "Normal"
    cboPriority.AddItem "High"
    cboPriority.AddItem "Low"
    cboPriority.ListIndex = 1


    lblPopServer.Visible = False
    txtPopServer.Visible = False

    Me.Show

    RetrieveSavedValues
    If txtUserName.Text <> "" Then
        ckLogin = vbChecked
    Else
        ckLogin = vbUnchecked
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set poSendMail = Nothing

End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
    poSendMail.PersistentSettings = True
    optEncodeType(poSendMail.EncodeType).Value = True
    If poSendMail.UseAuthentication Then ckLogin = vbChecked Else ckLogin = vbUnchecked

End Sub

Private Sub optEncodeType_Click(Index As Integer)

    If optEncodeType(0).Value = True Then
        MyEncodeType = MIME_ENCODE
        cboPriority.Enabled = True
        ckHtml.Enabled = True
        ckReceipt.Enabled = True
        ckLogin.Enabled = True
    Else
        MyEncodeType = UU_ENCODE
        ckHtml.Value = vbUnchecked
        ckReceipt.Value = vbUnchecked
        ckLogin.Value = vbUnchecked
        cboPriority.Enabled = False
        ckHtml.Enabled = False
        ckReceipt.Enabled = False
        ckLogin.Enabled = False
    End If

End Sub

Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ckHtml_Click()

    If ckHtml.Value = vbChecked Then bHtml = True Else bHtml = False

End Sub

Private Sub ckLogin_Click()

    If ckLogin.Value = vbChecked Then
        bAuthLogin = True
        fraOptions.Height = 3555
    Else
        bAuthLogin = False
        If ckPopLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckPopLogin_Click()

    If ckPopLogin.Value = vbChecked Then
        bPopLogin = True
        lblPopServer.Visible = True
        txtPopServer.Visible = True
        fraOptions.Height = 3555
    Else
        bPopLogin = False
        lblPopServer.Visible = False
        txtPopServer.Visible = False
        If ckLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckReceipt_Click()

    If ckReceipt.Value = vbChecked Then bReceipt = True Else bReceipt = False

End Sub

Private Sub cmdBrowse_Click()

    Dim sFilenames()    As String
    Dim i               As Integer
    
    On Local Error GoTo Err_Cancel
  
    With cmDialog
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select File Attachment(s)"
        .MaxFileSize = &H7FFF
        .Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
        sFilenames = Split(.FileName, vbNullChar)
    End With
    
    If UBound(sFilenames) = 0 Then
        If txtAttach.Text = "" Then
            txtAttach.Text = sFilenames(0)
        Else
            txtAttach.Text = txtAttach.Text & ";" & sFilenames(0)
        End If
    ElseIf UBound(sFilenames) > 0 Then
        If Right$(sFilenames(0), 1) <> "\" Then sFilenames(0) = sFilenames(0) & "\"
        For i = 1 To UBound(sFilenames)
            If txtAttach.Text = "" Then
                txtAttach.Text = sFilenames(0) & sFilenames(i)
            Else
                txtAttach.Text = txtAttach.Text & ";" & sFilenames(0) & sFilenames(i)
            End If
        Next
    Else
        Exit Sub
    End If
    
Err_Cancel:
'
'End Sub
'
'For Each frm In Forms
'    Unload frm
'    Set frm = Nothing
'Next
'
'End

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    RetrieveSavedValues

End Sub

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com
    On Error Resume Next

    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Left = base.Left
        If StandardizeWidth Then cnts(i).Width = base.Width
    Next

End Sub

Private Sub CenterControlsVertical(space As Single, AlignLeft As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseLeft As Single

    Dim sngParentHeight As Single

    sngParentHeight = Me.ScaleHeight

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Height
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))
    cnts(0).Top = (sngParentHeight - sngTotalSpace) / 2

    sngBaseLeft = cnts(0).Left

    For i = 1 To UBound(cnts)
        cnts(i).Top = cnts(i - 1).Top + cnts(i - 1).Height + space
        If AlignLeft Then cnts(i).Left = sngBaseLeft
    Next

End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        If AlignTop Then cnts(i).Top = sngBaseTop
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Top = base.Top
        If StandardizeHeight Then cnts(i).Height = base.Height
    Next

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub SetHorizontalDistance(distance As Single, StandardizeWidth As Boolean, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(cnts)
        If StandardizeWidth Then cnts(i).Width = cnts(i - 1).Width
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + distance
        If AlignTop Then cnts(i).Top = cnts(i - 1).Top
    Next

End Sub

Public Sub CenterControlsRelativeHorizontal(RelativeTo As Object, space As Single, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim sngTotalWidth As Single
    Dim i As Integer
    For i = 0 To UBound(cnts)
        sngTotalWidth = sngTotalWidth + cnts(i).Width
        If i < UBound(cnts) Then sngTotalWidth = sngTotalWidth + space
    Next

    cnts(0).Left = RelativeTo.Left + ((RelativeTo.Width - sngTotalWidth) / 2)

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        cnts(i).Top = cnts(0).Top
    Next

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next

End Sub

