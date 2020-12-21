Attribute VB_Name = "Declarations"
Option Explicit

'Public Const WM_RBUTTONDOWN = &H204

Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function SetPixelV& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&, ByVal crColor&)

Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetinputState Lib "user32" () As Long

Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Boolean) As Boolean

Public Declare Function GetKeyState Lib "user32" ( _
  ByVal nVirtKey As Long) As Integer

Public Declare Sub keybd_event Lib "user32" ( _
  ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)
 
Public Const VK_NUMLOCK = &H90
Public Const KEYEVENTF_KEYUP = &H2

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long


Public Sub MouseNoRightClick(InvisiblePicBox As PictureBox)
    Const WM_RBUTTONDOWN = &H204
    ' Rechtsklick abfangen und an unsichtbare PictureBox weiterleiten
    ' KEIN PopUp-Menü anzeigen!
    InvisiblePicBox.Tag = ""
    SendMessage InvisiblePicBox.hwnd, WM_RBUTTONDOWN, 0&, 0&
End Sub

Public Sub DropDown(Combo As ComboBox, _
    Optional ByVal ShowHide As Boolean = True)
    
    Const CB_SHOWDROPDOWN = &H14F
    SendMessage Combo.hwnd, CB_SHOWDROPDOWN, ShowHide, 0

End Sub
Public Function IsDropped(Combo As ComboBox) As Boolean
    
    Const CB_GETDROPPEDSTATE = &H157
    IsDropped = CBool(SendMessage(Combo.hwnd, _
        CB_GETDROPPEDSTATE, 0, 0))
        
End Function
  
Public Function FileExists(strFullPath As String) As Boolean
    Dim objFile As New Scripting.FileSystemObject
    FileExists = objFile.FileExists(strFullPath)
End Function

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function
