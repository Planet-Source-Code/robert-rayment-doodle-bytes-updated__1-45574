VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6600
   DrawWidth       =   2
   Icon            =   "Doodle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Flame  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   8
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2925
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Pulse  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   7
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2550
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shapes  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   6
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2190
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Signal  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   5
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Falls  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   4
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Jet  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   3
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Smoke "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   2
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shear "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   1
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Bunsen "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   0
      Left            =   0
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1 (Doodle.frm)

' Doodle bytes  by Robert Rayment  May 2003

' Animation using byte arrays & 256 palettes.

' Palettes designed using Palette Designer PSC CodeId=28440

' Code could be compacted with more Subs but
' left stretched out.

Option Base 1

' NB DO NOT USE Option Explicit with Def system
' Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings


Dim aDone   ' Boolean to allow exitting Loop
Dim FMH     ' For detecting Screen res change on the fly


Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

aDone = True

For i = 0 To 8
   If i <> Index Then Check1(i).Value = 0 'Unchecked
Next i
   
Select Case Index
Case 0: Bunsen
Case 1: Shear
Case 2: Smoke
Case 3: Jet
Case 4: Falls
Case 5: Signal
Case 6: Shapes
Case 7: Pulse
Case 8: Flame
End Select

End Sub

Private Sub Form_Load()

If App.PrevInstance Then End

Caption = " Doodle Bytes  by Robert Rayment"

PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height

' NB NB FMemW MUST ALWAYS BE A MULTIPLE OF 4 !!
FMemW = 256 '512
FMemH = 480 '400

' Back Canvas
ReDim bFMem(FMemW, FMemH)
FillBMPStruc

' Set Form size to accomodate required display rectangle
' Public ExtraBorder, ExtraHeight
GetExtras Form1.BorderStyle

Form1.Width = (FMemW + ExtraBorder) * Screen.TwipsPerPixelX
Form1.Height = (FMemH + ExtraHeight) * Screen.TwipsPerPixelY

' Shape Check boxes
For i = 0 To 8
   RoundCheck = CreateRoundRectRgn(10, 10, 70, 30, 15, 15)
   SetWindowRgn Check1(i).hWnd, RoundCheck, True
   DeleteObject RoundCheck
Next i

Form1.Show

'' Test four corners
'ReadPAL PathSpec$ & "YellBlue.pal"
'bFMem(1, FMemH) = 127      ' TL
'bFMem(FMemW, FMemH) = 255  ' TR
'bFMem(1, 1) = 200          ' BL
'bFMem(FMemW, 1) = 100      ' BR
'
'ShowWholePicture

End Sub

Private Sub ShowWholePicture()

If StretchDIBits(Form1.hdc, _
   0, 0, FMemW, FMemH, _
   0, 0, FMemW, FMemH, _
   bFMem(1, 1), bm, _
   DIB_RGB_COLORS, vbSrcCopy) = 0 Then
      MsgBox "Blit Error", vbCritical, "Doodle Whole"
      Form_Unload 0
End If
Form1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

aDone = True
Unload Me
End

End Sub

Private Sub INITDoodle(PalSpec$)

' Back Canvas
ReDim bFMem(FMemW, FMemH)
FillBMPStruc

' Set Form size to accomodate required display rectangle
' Public ExtraBorder, ExtraHeight
GetExtras Form1.BorderStyle

Form1.Width = (FMemW + ExtraBorder) * Screen.TwipsPerPixelX
Form1.Height = (FMemH + ExtraHeight) * Screen.TwipsPerPixelY

FMH = GetDeviceCaps(Me.hdc, VERTRES)

Form1.Show

ReadPAL (PalSpec$)

End Sub

Private Sub Bunsen()

Me.BackColor = RGB(0, 0, 0)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 256 '512
FMemH = 480 '400

INITDoodle PathSpec$ & "YellBlue.pal"

h = 120  ' Height of start of flame
mmid = FMemW \ 2

' Draw bunsen stem
For j = 1 To h - 1
   k = 0
   For i = mmid - 26 To mmid + 25
       bFMem(i, j) = 80 + k
       k = k + 1
   Next i
Next j

aDone = False

Do
   
   culindex = 7
   For k = -32 To -1
      If k > -10 And k < -4 Then
         ' Introduce randomness into part of flame
         culplus = 40 * (Rnd - 0.5)
         bFMem(mmid + k, h) = culindex + culplus
         bFMem(mmid + (-k - 1), h) = culindex + culplus
      Else
         bFMem(mmid + k, h) = culindex
         bFMem(mmid + (-k - 1), h) = culindex
      End If
      culindex = culindex + 8
   Next k
   
   For j = h To FMemH - 1
      For k = -32 To 31
         If k > -3 And k < 4 Then
            ' Reduce height by decreasing palette index more rapidly
            ' than 1 near center of flame.
            ' Changing .35 varies height.
            subindex = 1 + CInt(Rnd - 0.35)
            If bFMem(mmid + k, j) > 2 Then bFMem(mmid + k, j + 1) = bFMem(mmid + k, j) - subindex
         Else
            ' Reduce pal index by 1 -> towards black at low indexes
            If bFMem(mmid + k, j) > 1 Then bFMem(mmid + k, j + 1) = bFMem(mmid + k, j) - 1
         End If
      Next k
   Next j
   
'   Show central 64 bytes of picture

   If StretchDIBits(Form1.hdc, _
      FMemW \ 2 - 32, 0, 64, (FMemH), _
      mmid - 33, 0, 64, FMemH, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Bunsen"
         Form_Unload 0
   End If
   
   Form1.Refresh
   
   DoEvents

Loop Until aDone

End Sub

Private Sub Shear()

Me.BackColor = RGB(0, 100, 255)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 480 '512
FMemH = 256 '400

INITDoodle PathSpec$ & "Shear.pal"

h = 1 ' Initial right height

aDone = False

TLIM = 3 '6 ' Set speed, smaller values faster

Do
   
   If GetDeviceCaps(Me.hdc, VERTRES) <> FMH Then
'      'Screen res changed
      FMH = GetDeviceCaps(Me.hdc, VERTRES)
      DoEvents
      Form1.Show
   End If
   
   ' Zero right section
   For j = 1 To FMemH \ 2 + 2
      bFMem(FMemW, j) = 0
      bFMem(FMemW - 1, j) = 0
      bFMem(FMemW - 2, j) = 0
      bFMem(FMemW - 3, j) = 0
      bFMem(FMemW - 4, j) = 0
      bFMem(FMemW - 5, j) = 0
      bFMem(FMemW - 6, j) = 0
      bFMem(FMemW - 7, j) = 0
      bFMem(FMemW - 8, j) = 0
   Next j
   
   ' Set random right height
   zrn = 16 * Rnd * (Rnd - 0.5)
   IY = h + zrn
   If IY < 1 Then IY = 1
   If IY > FMemH \ 2 Then IY = FMemH \ 2
   
   ' Color right section
   For j = 1 To IY
      cul = 200 + 100 * (Rnd - 0.5)
      jj = j '+ Abs(10 * Sin(j))
      bFMem(FMemW - 8, jj) = cul
      bFMem(FMemW - 7, jj) = cul
      bFMem(FMemW - 6, jj) = cul
      bFMem(FMemW - 5, jj) = cul
      bFMem(FMemW - 4, jj) = cul
      bFMem(FMemW - 3, jj) = cul
      bFMem(FMemW - 2, jj) = cul
      bFMem(FMemW - 1, jj) = cul
      bFMem(FMemW, jj) = cul
   Next j
   
   h = IY
   
   ' Scroll sections to left
   ' CopyMemory Dest,Source
   For j = 1 To FMemH \ 2
      
      Select Case j
      Case Is < FMemH \ 16:      CopyMemory bFMem(1, j), bFMem(2, j), (FMemW - 1)
      Case Is < FMemH \ 8:       CopyMemory bFMem(1, j), bFMem(3, j), (FMemW - 2)
      Case Is < 3 * FMemH \ 16:  CopyMemory bFMem(1, j), bFMem(4, j), (FMemW - 3)
      Case Is < FMemH \ 4:       CopyMemory bFMem(1, j), bFMem(5, j), (FMemW - 4)
      Case Is < 5 * FMemH \ 16:  CopyMemory bFMem(1, j), bFMem(6, j), (FMemW - 5)
      Case Is < 3 * FMemH \ 8:   CopyMemory bFMem(1, j), bFMem(7, j), (FMemW - 6)
      Case Is < 7 * FMemH \ 16:  CopyMemory bFMem(1, j), bFMem(8, j), (FMemW - 7)
      Case Else:                 CopyMemory bFMem(1, j), bFMem(9, j), (FMemW - 8)
      End Select
      
   Next j
   
   ' Show Lower Half Picture
   
   If StretchDIBits(Form1.hdc, _
      0, FMemH \ 2, FMemW, FMemH \ 2, _
      0, 0, FMemW, FMemH \ 2, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Lower"
         Form_Unload 0
   End If
   
   Form1.Refresh
   
   ' Timer
   T = timeGetTime
   Do: DoEvents
   Loop Until (timeGetTime() - T >= TLIM Or (timeGetTime() - T) < 0 Or aDone = True)

Loop Until aDone

End Sub

Private Sub Smoke()

Me.BackColor = RGB(0, 0, 0)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 256 '512
FMemH = 480 '400

INITDoodle PathSpec$ & "Smoke.pal"

mmid = FMemW \ 2

zk = 0  ' Start Period varier
aDone = False
k = FMemH

TLIM = 1

Do
   
   Period = 1.195 * FMemH
   
   ' Clear stick
   j = 28  ' Height of start of smoke
   IX = mmid + (16 + 24 * j / FMemH) * Sin(2 * pi# * j / Period - zk)
   For k = 1 To 26
      bFMem(IX - 3, k) = 0
      bFMem(IX - 2, k) = 0
      bFMem(IX - 1, k) = 0
      bFMem(IX, k) = 0
      bFMem(IX + 1, k) = 0
      bFMem(IX + 2, k) = 0
   Next k
   
   ' Clear smoke
   For j = 28 To FMemH
      IX = mmid + (16 + 24 * j / FMemH) * Sin(2 * pi# * j / Period - zk)
      bFMem(IX - 3, j) = 0
      bFMem(IX - 2, j) = 0
      bFMem(IX - 1, j) = 0
      bFMem(IX, j) = 0
      bFMem(IX + 1, j) = 0
      Period = Period - 1
   Next j
   
   zk = zk + Rnd / 16
   Period = 1.2 * FMemH
   
   ' Draw stick
   j = 28
   IX = mmid + (16 + 24 * j / FMemH) * Sin(2 * pi# * j / Period - zk)
   For k = 1 To 26
      bFMem(IX - 3, k) = 100
      bFMem(IX - 2, k) = 100
      bFMem(IX - 1, k) = 100
      bFMem(IX, k) = 100
      bFMem(IX + 1, k) = 100
      bFMem(IX + 2, k) = 100
   Next k
   
   ' Draw smoke
   For j = 28 To FMemH
      IX = mmid + (16 + 24 * j / FMemH) * Sin(2 * pi# * j / Period - zk)
      bFMem(IX - 3, j) = 120
      bFMem(IX - 2, j) = 180
      'bFMem(IX - 1, j) = 200
      'bFMem(IX, j) = 255
      bFMem(IX + 1, j) = 255
      Period = Period - 1
   Next j

   ' Show Central 80 bytes of Picture
   
   If StretchDIBits(Form1.hdc, _
      (FMemW) \ 2 - 40, 0, 80, (FMemH), _
      mmid - 41, 0, 80, FMemH, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Bunsen"
         Form_Unload 0
   End If
   
   Form1.Refresh
   DoEvents
Loop Until aDone

End Sub

Private Sub Jet()

Me.BackColor = RGB(0, 0, 0)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 256 '512
FMemH = 480 '400

INITDoodle PathSpec$ & "Jet.pal"

mmid = FMemW \ 2

zk = 1      ' Start Period varier
Period = 4  'Intial Period

aDone = False

Do
   
   zk = zk + Rnd / 16
   
   ' Draw jet
   For j = 1 To FMemH
      IX = mmid + (16 + 48 * j / FMemH) * Sin(2 * pi# * j / Period - zk)
      If j = 1 Then
         bFMem(IX, j) = 0
      Else
         bFMem(IX, j) = 255 * Rnd
      End If
   Next j
   ' Add random stars
   If Rnd < 0.5 Then
      IX = mmid + 40 * (Rnd - 0.5)
      IY = FMemH \ 2 + 180 * (Rnd - 0.5)
      bFMem(IX, IY) = 255
      bFMem(IX, IY - 1) = 255
      bFMem(IX, IY + 1) = 255
      bFMem(IX, IY - 2) = 255
      bFMem(IX, IY + 2) = 255
      bFMem(IX - 1, IY) = 255
      bFMem(IX + 1, IY) = 255
      bFMem(IX + 2, IY) = 255
      bFMem(IX - 2, IY) = 255
      bFMem(IX - 1, IY - 1) = 255
      bFMem(IX - 1, IY + 1) = 255
      bFMem(IX + 1, IY - 1) = 255
      bFMem(IX + 1, IY + 1) = 255
   End If
   
   ' Scroll up
   For j = FMemH - 1 To 2 Step -1
         CopyMemory bFMem(mmid - 56, j), bFMem(mmid - 56, j - 1), 112
   Next j
   
   ' Show Central 112 bytes of Picture
   
   If StretchDIBits(Form1.hdc, _
      (FMemW) \ 2 - 56, 0, 112, (FMemH), _
      mmid - 57, 0, 112, FMemH, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Bunsen"
         Form_Unload 0
   End If
   
   Form1.Refresh
   
   DoEvents
Loop Until aDone

End Sub

Private Sub Falls()

Me.BackColor = RGB(100, 180, 220)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 320 '512
FMemH = 320 '400

INITDoodle PathSpec$ & "Water.pal"

h = 240 ' Top height

aDone = False

TLIM = 1 ' Set speed, smaller values faster

Do
   
   If GetDeviceCaps(Me.hdc, VERTRES) <> FMH Then
'      'Screen res changed
      FMH = GetDeviceCaps(Me.hdc, VERTRES)
      DoEvents
      Form1.Show
   End If
   
   ' Fill row h with random colors
   For i = 1 To FMemW
      bFMem(i, h) = Rnd * 255
   Next i
   
   For j = 2 To h
      If Rnd < 0.9 Then
         ii = 2 - CInt(Rnd)
         CopyMemory bFMem(1, j - 1), bFMem(ii, j), FMemW
      End If
   Next j
   
   ' Show lower 3/4 of picture
   
   If StretchDIBits(Form1.hdc, _
      0, (FMemH) \ 4, (FMemW), 3 * (FMemH) \ 4, _
      0, 0, FMemW, 3 * FMemH \ 4, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Lower"
         Form_Unload 0
   End If
   
   Form1.Refresh
   
   ' Timer
   T = timeGetTime
   Do: DoEvents
   Loop Until (timeGetTime() - T >= TLIM Or (timeGetTime() - T) < 0 Or aDone = True)

Loop Until aDone

End Sub

Private Sub Signal()

Me.BackColor = 0 'RGB(0, 100, 255)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 480 '512
FMemH = 256 '400

INITDoodle PathSpec$ & "YellBlue.pal"

h = 1 ' Initial right height

mmid = FMemH \ 2

aDone = False

TLIM = 1 '3 '6 ' Set speed, smaller values faster
zang = 0
zt = 100
s = 1
Do
   
   If GetDeviceCaps(Me.hdc, VERTRES) <> FMH Then
'      'Screen res changed
      FMH = GetDeviceCaps(Me.hdc, VERTRES)
      DoEvents
      Form1.Show
   End If
   
   
   ' Zero right section
   For j = mmid - 50 To mmid + 50
      bFMem(FMemW, j) = 0
   Next j
   
   ' Set right height
   iy1 = mmid + 50 * Rnd * Cos(zang / 0.2)
   bFMem(FMemW, iy1) = 200
   iy2 = mmid - 50 * Rnd * Cos(zang / 0.2 + 2 * Sin(zang) + Rnd)
   bFMem(FMemW, iy2) = 200
   If iy2 <> iy1 Then
      For IY = iy1 To iy2 Step Sgn(iy2 - iy1)
         bFMem(FMemW, IY) = Rnd * 200
      Next IY
   End If
   
   zang = zang + 0.01
   If zang > 2 * pi# Then zang = 0
   
   ' Scroll section to left
   ' CopyMemory Dest,Source
   For j = mmid - 50 To mmid + 50
      CopyMemory bFMem(1, j), bFMem(2, j), (FMemW - 1)
   Next j
   
   ' Show middle part of Picture
   
   If StretchDIBits(Form1.hdc, _
      0, FMemH \ 4, FMemW, FMemH \ 2, _
      0, FMemH \ 4, FMemW, FMemH \ 2, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Lower"
         Form_Unload 0
   End If

   Form1.Refresh

   ' Timer  NB OK on Win98 but slow & flickering on WinXP
   'T = timeGetTime
   'Do: DoEvents
   'Loop Until (timeGetTime() - T >= TLIM Or (timeGetTime() - T) < 0 Or aDone = True)
   DoEvents
Loop Until aDone

End Sub

Private Sub Shapes()

Me.BackColor = 0 'RGB(0, 100, 255)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 320 '512
FMemH = 320 '400

INITDoodle PathSpec$ & "YellBlue.pal"

aDone = False

Do
   
   If GetDeviceCaps(Me.hdc, VERTRES) <> FMH Then
'      'Screen res changed
      FMH = GetDeviceCaps(Me.hdc, VERTRES)
      DoEvents
      Form1.Show
   End If
   
   SPIRAL
      ScrollUp
   SPRING
      ScrollUp
   HYPOCYCLOID
      ScrollUp
   ROSE
      ScrollUp
   ROSE2
      ScrollUp
   LIMACON
      ScrollUp
   LEMNISCATE
      ScrollUp
   
   ReDim bFMem(FMemW, FMemH)
   ShowWholePicture
   DoEvents
Loop Until aDone


End Sub


Private Sub ScrollUp()
   
   For i = 1 To FMemW
      bFMem(i, 1) = 0
   Next i
   For k = 1 To FMemH
      For j = FMemH To 2 Step -1
         CopyMemory bFMem(1, j), bFMem(1, j - 1), FMemW
      Next j
      ShowWholePicture
   Next k

End Sub

Private Sub SPIRAL()
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   zAmp = 1.6
   For zang = 0 To 20 * pi# Step 0.03
      IX = xcen + zAmp * zang * Cos(zang)
      IY = ycen + zAmp * zang * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
' Clears shape
'   zange = zang - 0.02
'   For zang = zange To 0 Step -0.02
'      IX = xcen + zAmp * zang * Cos(zang)
'      IY = ycen + zAmp * zang * Sin(zang)
'      bFMem(IX, IY) = 0
'      bFMem(IX - 1, IY) = 0
'      bFMem(IX, IY + 1) = 0
'      bFMem(IX - 1, IY + 1) = 0
'      ShowWholePicture
'      DoEvents
'      If aDone = True Then Exit For
'   Next zang
End Sub

Private Sub SPRING()
   ycen = FMemH \ 4
   xcen = FMemW \ 2
   zAmp = 0.5
   For zang = 0 To 30 * pi# Step 0.03
      IX = xcen + zAmp * zang * Cos(zang)
      IY = ycen + zAmp * zang * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      ycen = ycen + 0.05
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub HYPOCYCLOID()
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   zAmp = 100
   For zang = 0 To 2 * pi# Step 0.01
      IX = xcen + zAmp * Cos(zang) ^ 3
      IY = ycen + zAmp * Sin(zang) ^ 3
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub ROSE()
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   zAmp = 100
   For zang = 0 To pi# Step 0.005
      zR = zAmp * Cos(3 * zang)
      IX = xcen + zR * Cos(zang)
      IY = ycen + zR * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub ROSE2()
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   zAmp = 100
   For zang = 0 To 2 * pi# Step 0.005
      zR = zAmp * Cos(4 * zang)
      IX = xcen + zR * Cos(zang)
      IY = ycen + zR * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub LIMACON()
   ycen = FMemH \ 2
   xcen = FMemW \ 3
   zAmp = 100
   zB = 80
   For zang = 0 To 2 * pi# Step 0.005
      zR = zB + zAmp * Cos(zang)
      IX = xcen + zR * Cos(zang)
      IY = ycen + zR * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub LEMNISCATE()   ' Sort of
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   zAmp = 60
   For zang = 0 To pi# Step 0.005
      zR = Sgn(Cos(2 * zang)) * zAmp * Sqr(Abs(Cos(2 * zang)))
      IX = xcen + zR * Cos(zang)
      IY = ycen + zR * Sin(zang)
      bFMem(IX, IY) = 100
      bFMem(IX - 1, IY) = 100
      bFMem(IX, IY + 1) = 100
      bFMem(IX - 1, IY + 1) = 100
      ShowWholePicture
      DoEvents
      If aDone = True Then Exit For
   Next zang
End Sub

Private Sub Pulse()

Me.BackColor = 0 'RGB(0, 100, 255)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 320 '512
FMemH = 320 '400

INITDoodle PathSpec$ & "YellBlue.pal"

aDone = False
TLIM = 700
Do
   ycen = FMemH \ 2
   xcen = FMemW \ 2
   
   cul = 255
   For zR = 11 To 100
      For zang = 0 To 2 * pi# Step 0.02
         IX = xcen + zR * Cos(zang)
         IY = ycen - zR * Sin(zang)
         bFMem(IX, IY) = cul
      Next zang
      If aDone Then Exit For
      cul = cul - 1
      If cul < 0 Then cul = 0
      ShowWholePicture
      DoEvents
      If aDone Then Exit For
   Next zR
   
   cul = 0
   For zR = 100 To 11 Step -1
      For zang = 0 To 2 * pi# Step 0.02
         IX = xcen + zR * Cos(zang)
         IY = ycen - zR * Sin(zang)
         bFMem(IX, IY) = cul
      Next zang
      If aDone Then Exit For
      cul = cul + 1
      If cul > 255 Then cul = 255
      ShowWholePicture
      DoEvents
      If aDone Then Exit For
   Next zR
   
   ' Timer
   T = timeGetTime
   Do: DoEvents
   Loop Until (timeGetTime() - T >= TLIM Or (timeGetTime() - T) < 0 Or aDone = True)

Loop Until aDone

End Sub

Private Sub Flame()
'aDone = True
'Me.Cls

Me.BackColor = RGB(0, 0, 0)
Me.Cls

' Set display size wanted
' Public FMemW, FMemH   ' bFMem() width & height
FMemW = 256 '512
FMemH = 480 '400

INITDoodle PathSpec$ & "Flame.pal"

h = 10  ' Height of start of flame

mmid = FMemW \ 2


' Draw bunsen stem
For j = 1 To h - 1
   k = 0
   For i = mmid - 64 To mmid + 63
       bFMem(i, j) = 80
       k = k + 1
   Next i
Next j

aDone = False

Do
   
   For k = -64 To 63
      bFMem(mmid + k, h) = 204 + 100 * (Rnd - 0.5)
      bFMem(mmid + k, h - 1) = 180 + 100 * (Rnd - 0.5)
   Next k
   For j = 10 To FMemH - 1
      For k = -64 To 63
         ' Reduce pal index by 1 -> towards black at low indexes
         If bFMem(mmid + k, j) > 2 Then
            bFMem(mmid + k, j + 1) = bFMem(mmid + k, j) - 2
            
'            ' Ooze
'            bFMem(mmid + k, j + 1) = _
'            (1& * bFMem(mmid + k - 1, j - 1) + bFMem(mmid + k - 1, j + 1) _
'            + bFMem(mmid + k + 1, j - 1) + bFMem(mmid + k + 1, j + 1) _
'            ) \ 4

            ' Flame
            If Rnd < 0.5 Then
               bFMem(mmid + k, j + 1) = _
               (1& * bFMem(mmid + k - 1, j + 1) + bFMem(mmid + k + 1, j - 1) _
               ) \ 2
            Else
               bFMem(mmid + k, j + 1) = _
               (1& * bFMem(mmid + k - 1, j - 1) + bFMem(mmid + k + 1, j - 1) _
               ) / 2.02
            End If
            
         End If
      Next k
   Next j
   
'   Show central 128 bytes of picture

   If StretchDIBits(Form1.hdc, _
      FMemW \ 2 - 64, 0, 128, FMemH, _
      mmid - 63, 0, 128, FMemH, _
      bFMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Doodle Bunsen"
         Form_Unload 0
   End If
   
   Form1.Refresh
   
   DoEvents

Loop Until aDone

End Sub

