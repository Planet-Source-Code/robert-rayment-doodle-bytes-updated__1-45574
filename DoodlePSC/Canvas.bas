Attribute VB_Name = "Module1"
'Module1.bas

Option Base 1  ' Arrays base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings

'--------------------------------------------------------------------------
' Shaping APIs
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal _
 X2 As Long, ByVal Y2 As Long, _
 ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
(ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

' --------------------------------------------------------------
' Windows API - For timing

Public Declare Function timeGetTime& Lib "winmm.dll" ()

' --------------------------------------------------------------
Public Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const VERTRES = 10
'------------------------------------------------------------------------------

Public Declare Function GetSystemMetrics Lib "user32" _
(ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0  ' Screen Width
Public Const SM_CYSCREEN = 1  ' Screen Height
Public Const SM_CYCAPTION = 4 ' Height of window caption
Public Const SM_CYMENU = 15   ' Height of menu
Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)

' -----------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

' -----------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in memory to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'------------------------------------------------------------------------------
' -----------------------------------------------------------

Public ExtraBorder, ExtraHeight
Public ScanLine      ' Num bytes in a scan line

Public FWidth, FHeight        ' Form pixel width & height
Public FMemW, FMemH
Public bFMem()

Public Const pi# = 3.14159265


Public Sub GetExtras(BStyle As Byte)

' Public ExtraBorder, ExtraHeight

' BStyle 1 to 5 (not 0)
' BStyle = Form1.BorderStyle

Border = GetSystemMetrics(SM_CXDLGFRAME)
If BStyle = 2 Or BStyle = 5 Then Border = Border + 1 ' Sizable
If BStyle > 3 Then
   CapHeight = GetSystemMetrics(SM_CYSMCAPTION) ' Small cap - ToolWindow
Else
   CapHeight = GetSystemMetrics(SM_CYCAPTION)   ' Standard cap
End If
ExtraBorder = 2 * Border
ExtraHeight = CapHeight + ExtraBorder

'MenuHeight = GetSystemMetrics(SM_CYMENU)
'ExtraHeight = CapHeight + MenuHeight + ExtraBorder

' Win98  ExtraBorder=6 or 8, ExtraHeight= 41 - 46
' WinXP  ExtraBorder=6 or 8, ExtraHeight= 44 - 54

End Sub

Public Sub FillBMPStruc()

' Public FMemW, FMemH   ' Form pixel width & height

With bm.bmiH
  .biSize = 40&
  .biwidth = FMemW
  .biheight = FMemH
  .biPlanes = 1
  .biBitCount = 8
  .biCompression = 0&
  ' Ensure expansion to 4B boundary
  ScanLine = (Int((FMemW + 3) \ 4)) * 4
   
   If ScanLine <> FMemW Then
      MsgBox ("Byte array width (FMemW) must be a multiple of 4")
   End If
   
  .biSizeImage = ScanLine * Abs(.biheight)
  .biXPelsPerMeter = 0&
  .biYPelsPerMeter = 0&
  .biClrUsed = 0&
  .biClrImportant = 0&
End With
End Sub

Public Sub ReadPAL(PalSpec$)
' Read JASC-PAL palette file
' Any error shown by PalSpec$ = ""
' Else RGB into Colors(i) Long

'Public red As Byte, green As Byte, blue As Byte
On Error GoTo palerror

'ReDim Colors(0 To 255)

Open PalSpec$ For Input As #1
Line Input #1, A$
p = InStr(1, A$, "JASC")
If p = 0 Then PalSpec$ = "": Exit Sub
   
   'JASC-PAL
   '0100
   '256
   Line Input #1, A$
   Line Input #1, A$

   For n = 0 To 255
      If EOF(1) Then Exit For
      With bm.Colors(n)
         Input #1, .rgbRed, .rgbGreen, .rgbBlue
      End With
      bm.Colors(n).rgbReserved = 0
   Next n
   Close #1
Exit Sub
'===========
palerror:
PalSpec$ = vbNullString
   
   MsgBox "Palette file error or not there"
   On Error GoTo 0

End Sub

