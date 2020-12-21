Attribute VB_Name = "FormSetup"
' zunächst die benötigten Deklarationen
Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long

 
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
 
' Setzen der Minimieren bzw. Maximieren-Buttons
Public Sub FormAddMinMaxButtons(hwnd As Long, _
  Optional ByVal MinBtn As Boolean = True, _
  Optional ByVal MaxBtn As Boolean = True)
 
  Dim lStyle As Long
  lStyle = GetWindowLong(hwnd, GWL_STYLE)
 
  ' Minimieren-Button hinzufügen
  If (MinBtn) Then _
    lStyle = (lStyle Or WS_MINIMIZEBOX)
 
  ' Maximieren-Button hinzufügen
  If (MaxBtn) Then _
    lStyle = (lStyle Or WS_MAXIMIZEBOX)
 
  Call SetWindowLong(hwnd, GWL_STYLE, lStyle)
End Sub



