Attribute VB_Name = "ModSendKeys"
' Constantes para las teclas y otros
  
Const KEYEVENTF_KEYUP = &H2
Const KEYEVENTF_EXTENDEDKEY = &H1
  
  
'Declaración del Api keybd_event para la presión de tecla
  
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)
  
  
  
Sub SendKeys(Tecla As Long)
  
    Call keybd_event(Tecla, 0, 0, 0)
  
    Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
  
End Sub

