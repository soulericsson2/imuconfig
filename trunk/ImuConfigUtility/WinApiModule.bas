Attribute VB_Name = "WinApiModule"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetPixelV Lib "GDI32.dll" ( _
    ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal crColor As Long) As Long
    
    
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
         ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
         ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Const SRCCOPY = &HCC0020


Public Declare Function GetTickCount Lib "kernel32" () As Long
