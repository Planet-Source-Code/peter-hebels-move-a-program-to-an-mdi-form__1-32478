Attribute VB_Name = "Module1"
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex&) As Long
Public Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hWnd&, ByVal lpString$, ByVal cb&)
Public Declare Function WindowFromPoint Lib "user32" (ByVal ptY As Long, ByVal ptX As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex&) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle&, ByVal nWidth&, ByVal crColor&) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject&) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc&, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&) As Long
Public Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long)
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance&, ByVal lpCursor&) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public mlngHwndCaptured As Long
Public Const IDC_UPARROW = 32516&
Public WinVal As Long
