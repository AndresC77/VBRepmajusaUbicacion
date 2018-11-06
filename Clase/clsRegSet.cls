VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Option Explicit
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_STHOUSAND = &HF
Private Const LOCALE_SMONDECIMAL = &H16
Private Const LOCALE_SMONTHOUSAND = &H17
Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Property Get DateSeparator() As String
    DateSeparator = GetRegionalSetting("sDate")
End Property
Public Property Get DecimalSymbol() As String
    DecimalSymbol = GetRegionalSetting("sDecimal")
End Property
Public Property Get MonDecimalSymbol() As String
    MonDecimalSymbol = GetRegionalSetting("sMonDecimalSep")
End Property
Public Property Get ShortDate() As String
    ShortDate = GetRegionalSetting("sShortDate")
End Property
Public Property Get LongDate() As String
    LongDate = GetRegionalSetting("sLongDate")
End Property
Public Property Get CurrencyCode() As String
    CurrencyCode = GetRegionalSetting("sCurrency")
End Property
Public Property Get Country() As String
    Country = GetRegionalSetting("sCountry")
End Property
Public Property Get ThousandSeparator() As String
    ThousandSeparator = GetRegionalSetting("sThousand")
End Property
Public Property Get MonThousandSeparator() As String
    MonThousandSeparator = GetRegionalSetting("sMonThousandSep")
End Property
Public Property Get TimeSeparator() As String
    TimeSeparator = GetRegionalSetting("sTime")
End Property
Public Property Get ListSeparator() As String
   ListSeparator = GetRegionalSetting("sList")
End Property
Private Function GetRegionalSetting(strKeyName As String) As String
    On Error GoTo Err_Tag
    Dim strSecName As String
    Dim strRetString As String * 256
    Dim lngSuccess As Long
    strSecName = "Intl"
    GetRegionalSetting = ""
    lngSuccess = GetProfileString(strSecName, strKeyName, "", strRetString, Len(strRetString))
    If lngSuccess <> 0 Then
        GetRegionalSetting = Left$(strRetString, InStr(strRetString, Chr$(0)) - 1)
    End If
End_Tag:
    On Error Resume Next
    Exit Function
Err_Tag:
    MsgBox Err.Number & " " & Err.Description
    Resume End_Tag
End Function

Public Sub SetRegionalSetting(strDecimal As String, strThousand As String, strMDecimal As String, strMThousand As String, strDate As String)
    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()
    
    If SetLocaleInfo(dwLCID, LOCALE_SDECIMAL, strDecimal) = False Then
       MsgBox "Error Set Decimal Separator"
       Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_SMONDECIMAL, strMDecimal) = False Then
       MsgBox "Error Set Decimal Separator"
       Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_STHOUSAND, strThousand) = False Then
       MsgBox "Error Set Thousand Separator"
       Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_SMONTHOUSAND, strMThousand) = False Then
       MsgBox "Error Set Thousand Separator"
       Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, strDate) = False Then
       MsgBox "Error Set DateShort"
       Exit Sub
    End If
    
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub



