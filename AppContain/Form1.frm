VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "App Container"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7350
   Begin VB.CommandButton Command1 
      Caption         =   "<-- Put Back"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'App Container project by Peter Hebels, Website "www.phsoft.cjb.net"                     *
'Iam not responsible for any damages may caused by this project or program               *
'*****************************************************************************************

'Warning wrong use of this program can wreck your Windows shell, if this happens
'you have to restart Windows as soon as possible!!
'Also do not select separate controls in programs for example the Textbox in Notepad, this also
'causes problems because you will rip the textbox from the program.

'If you are going to channge the code, BE CAREFULL what you change!!

'Do not forget to save any unsaved data before starting this app.

Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Boolean

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long

End Type

Private Type POINT

    X As Long
    Y As Long

End Type

Dim WinCount As Integer

Private Sub Command1_Click()
SetParent WinVal, GetDesktopWindow
WinCount = 0
End Sub

Private Sub Form_MouseDown(Button%, Shift%, X As Single, Y As Single)
    If SetCapture(hWnd) Then MousePointer = vbUpArrow
  End Sub

Private Sub Form_MouseMove(Button%, Shift%, X As Single, Y As Single)
    Dim pt As POINT
    Static hWndLast As Long
      
    If GetCapture() Then
        
        pt.X = CLng(X)
        pt.Y = CLng(Y)
        ClientToScreen Me.hWnd, pt
        
        mlngHwndCaptured = WindowFromPoint(pt.X, pt.Y)
        
        If hWndLast <> mlngHwndCaptured Then
            If hWndLast Then InvertTracker hWndLast
            InvertTracker mlngHwndCaptured
            hWndLast = mlngHwndCaptured
        End If
    End If
  End Sub

Private Sub Form_MouseUp(Button%, Shift%, X As Single, Y As Single)
    Dim strCaption$
    
    If mlngHwndCaptured Then
        
        strCaption = Space(1000)
               
        If Left(strCaption, GetWindowText(mlngHwndCaptured, strCaption, Len(strCaption))) = Me.Caption Then
        MsgBox "Can't select myself!", vbInformation, "Can't Do"
        InvalidateRect 0, 0, True
        Exit Sub
        End If
        
        If Left(strCaption, GetWindowText(mlngHwndCaptured, strCaption, Len(strCaption))) = Command1.Caption Then
        MsgBox "Can't select myself!", vbInformation, "Can't Do"
        InvalidateRect 0, 0, True
        Exit Sub
        End If
        
        If WinCount = 1 Then
        MsgBox "Already Filled", vbInformation, "Can't Do"
        InvalidateRect 0, 0, True
        Exit Sub
        End If
        
        Caption = Left(strCaption, GetWindowText(mlngHwndCaptured, strCaption, Len(strCaption))) & " Hwnd: " & mlngHwndCaptured
        WinCount = 1
        InvalidateRect 0, 0, True
          SetParent mlngHwndCaptured, Me.hWnd
          WinVal = mlngHwndCaptured
          mlngHwndCaptured = False
          MousePointer = vbNormal
      
      End If
  End Sub

Private Sub InvertTracker(hwndDest As Long)
    Dim hdcDest&, hPen&, hOldPen&, hOldBrush&
    Dim cxBorder&, cxFrame&, cyFrame&, cxScreen&, cyScreen&
    Dim rc As RECT, cr As Long
    Const NULL_BRUSH = 5
    Const R2_NOT = 6
    Const PS_INSIDEFRAME = 6
    
      cxScreen = GetSystemMetrics(0)
      cyScreen = GetSystemMetrics(1)
      cxBorder = GetSystemMetrics(5)
      cxFrame = GetSystemMetrics(32)
      cyFrame = GetSystemMetrics(33)
    

    GetWindowRect hwndDest, rc
   
    hdcDest = GetWindowDC(hwndDest)
    
    SetROP2 hdcDest, R2_NOT
    cr = RGB(0, 0, 0)
    hPen = CreatePen(PS_INSIDEFRAME, 3 * cxBorder, cr)
    
    hOldPen = SelectObject(hdcDest, hPen)
    hOldBrush = SelectObject(hdcDest, GetStockObject(NULL_BRUSH))
    Rectangle hdcDest, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top
    SelectObject hdcDest, hOldBrush
    SelectObject hdcDest, hOldPen
    
    ReleaseDC hwndDest, hdcDest
    DeleteObject hPen
End Sub

Private Sub Form_Load()
   
    ScaleMode = vbPixels
    AutoRedraw = True
    
End Sub


