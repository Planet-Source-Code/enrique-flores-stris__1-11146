Attribute VB_Name = "Stris"
Option Explicit

Type STRISLEVEL
   px2 As Long          ' position of x2 (x1 always null) = width for this level
   mTime As Single      ' speed or time to start with
   R(21) As String      ' level data
End Type

Public OK As Boolean
Public DialogTitle As String
Public DialogText As String

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
