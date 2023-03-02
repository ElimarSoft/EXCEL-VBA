Attribute VB_Name = "CanvasApiMod5"
Option Explicit

Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal Width As Long, ByVal height As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function apiSetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiSetActiveWindow Lib "user32" Alias "SetActiveWindow" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As winRect) As Long
    
Private Type winRect     'used by apiMoveWindow
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const SRCCOPY      As Long = &HCC0020      '/* dest = source                   */
Const SRCPAINT     As Long = &HEE0086      '/* dest = source OR dest           */
Const SRCAND       As Long = &H8800C6      '/* dest = source AND dest          */
Const SRCINVERT    As Long = &H660046      '/* dest = source XOR dest          */
Const SRCERASE     As Long = &H440328      '/* dest = source AND (NOT dest )   */
Const NOTSRCCOPY   As Long = &H330008      '/* dest = (NOT source)             */
Const NOTSRCERASE  As Long = &H1100A6      '/* dest = (NOT src) AND (NOT dest) */
Const MERGECOPY    As Long = &HC000CA      '/* dest = (source AND pattern)     */
Const MERGEPAINT   As Long = &HBB0226      '/* dest = (NOT source) OR dest     */
Const PATCOPY      As Long = &HF00021      '/* dest = pattern                  */
Const PATPAINT     As Long = &HFB0A09      '/* dest = DPSnoo                   */
Const PATINVERT    As Long = &H5A0049      '/* dest = pattern XOR dest         */
Const DSTINVERT    As Long = &H550009      '/* dest = (NOT dest)               */
Const BLACKNESS    As Long = &H42          '/* dest = BLACK                    */
Const WHITENESS    As Long = &HFF0062      '/* dest = WHITE

Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As LongPtr
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32.dll" ( _
ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const TMainCaneco5 As String = "TMainCaneco5"
Const TFDisjoncteur As String = "TFDisjoncteur"
Const TCirForm As String = "TCirForm"


Private Function WaitForWindow(WClass As String) As Long
    
    Dim Timeout As Integer: Timeout = 8
    Dim h1 As Long
    Do
        Delay (1)
        h1 = apiFindWindow(WClass, vbNullString)
        Timeout = Timeout - 1
        If Timeout <= 0 Then End
    Loop Until h1 <> 0

    WaitForWindow = h1

End Function

Public Sub Test08()

    Dim h1 As Long: h1 = WaitForWindow(TCirForm)
    apiSetFocus h1
     
    Dim hdcMemDC As Long: hdcMemDC = CreateBitmap(h1)
    
    Dim ws1 As Worksheet
    Set ws1 = Worksheets("Canvas")
    
    Dim sX As Integer, sY As Integer, w As Integer, h As Integer
    sX = 610
    sY = 575
    w = 76
    h = 35
    
    sX = 508
    sY = 577
    w = 80
    h = 25
    
    Dim x As Integer, y As Integer
    
        For x = sX To sX + w
        For y = sY To sY + h
            ws1.Cells(y - sY + 1, x - sX + 1).Interior.Color = GetPixel(hdcMemDC, x, y)
        Next
        Next
    
    
End Sub

Private Function CreateBitmap(h1 As Long) As Long
     
     Dim Rect1 As winRect: Call apiGetWindowRect(h1, Rect1)
     Dim w As Long: w = Rect1.Right - Rect1.Left
     Dim h As Long: h = Rect1.Bottom - Rect1.Top
     Dim Wdc As Variant: Wdc = GetWindowDC(h1)
     
     Dim hdcWindow As Long: hdcWindow = GetDC(h1)
     Dim hbmScreen As Variant: hbmScreen = CreateCompatibleBitmap(hdcWindow, w, h)
     Dim hdcMemDC As Variant: hdcMemDC = CreateCompatibleDC(hdcWindow)
     Call SelectObject(hdcMemDC, hbmScreen)
     Dim res As Boolean
     res = BitBlt(hdcMemDC, 0, 0, w, h, Wdc, 0, 0, SRCCOPY)
     'res = BitBlt(hdcMemDC, 0, 0, w, h, hdcWindow, 0, 0, SRCCOPY)

     CreateBitmap = hdcMemDC

End Function
