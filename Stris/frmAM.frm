VERSION 5.00
Begin VB.Form frmAM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AstroMover"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmAM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTZ 
      Height          =   285
      Left            =   1260
      TabIndex        =   23
      Text            =   "1"
      ToolTipText     =   "Time Zone"
      Top             =   4515
      Width           =   315
   End
   Begin VB.PictureBox picRadix 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   4560
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   330
      TabIndex        =   22
      Top             =   0
      Width           =   4950
   End
   Begin VB.PictureBox picWorld 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   60
      MousePointer    =   2  'Cross
      Picture         =   "frmAM.frx":030A
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   21
      ToolTipText     =   "Positie"
      Top             =   3060
      Width           =   2730
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   495
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   20
      Top             =   6045
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picSymb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   225
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   19
      Top             =   5775
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   2865
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   18
      Top             =   0
      Width           =   1650
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1920
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   17
      Top             =   5865
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton cmdTimeDown 
      Height          =   345
      Index           =   1
      Left            =   705
      Picture         =   "frmAM.frx":0E38
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeUp 
      Height          =   345
      Index           =   1
      Left            =   705
      Picture         =   "frmAM.frx":136A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1725
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeDown 
      Height          =   345
      Index           =   3
      Left            =   1755
      Picture         =   "frmAM.frx":189C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeUp 
      Height          =   345
      Index           =   3
      Left            =   1755
      Picture         =   "frmAM.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1725
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeDown 
      Height          =   345
      Index           =   4
      Left            =   2160
      Picture         =   "frmAM.frx":2300
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeDown 
      Height          =   345
      Index           =   2
      Left            =   1110
      Picture         =   "frmAM.frx":2832
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2550
      Width           =   615
   End
   Begin VB.CommandButton cmdTimeDown 
      Height          =   345
      Index           =   0
      Left            =   300
      Picture         =   "frmAM.frx":2D64
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeUp 
      Height          =   345
      Index           =   4
      Left            =   2160
      Picture         =   "frmAM.frx":3296
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1725
      Width           =   375
   End
   Begin VB.CommandButton cmdTimeUp 
      Height          =   345
      Index           =   2
      Left            =   1110
      Picture         =   "frmAM.frx":37C8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1725
      Width           =   615
   End
   Begin VB.CommandButton cmdTimeUp 
      Height          =   345
      Index           =   0
      Left            =   300
      Picture         =   "frmAM.frx":3CFA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1725
      Width           =   375
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   285
      MaxLength       =   16
      TabIndex        =   6
      Text            =   "21/03/1990 12.00"
      ToolTipText     =   "Datum/Tijd"
      Top             =   2100
      Width           =   2250
   End
   Begin VB.CommandButton cmdPlace 
      Height          =   285
      Index           =   3
      Left            =   2505
      Picture         =   "frmAM.frx":422C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4515
      Width           =   285
   End
   Begin VB.CommandButton cmdPlace 
      Height          =   285
      Index           =   2
      Left            =   1620
      Picture         =   "frmAM.frx":475E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4515
      Width           =   285
   End
   Begin VB.CommandButton cmdPlace 
      Height          =   285
      Index           =   1
      Left            =   930
      Picture         =   "frmAM.frx":4C90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4515
      Width           =   285
   End
   Begin VB.CommandButton cmdPlace 
      Height          =   285
      Index           =   0
      Left            =   45
      Picture         =   "frmAM.frx":51C2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4515
      Width           =   285
   End
   Begin VB.TextBox txtNB 
      Height          =   285
      Left            =   1935
      TabIndex        =   1
      Text            =   "51.05"
      ToolTipText     =   "Breedte"
      Top             =   4515
      Width           =   525
   End
   Begin VB.TextBox txtOL 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "02.40"
      ToolTipText     =   "Lengte"
      Top             =   4515
      Width           =   525
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1170
      MouseIcon       =   "frmAM.frx":56F4
      MousePointer    =   99  'Custom
      Picture         =   "frmAM.frx":59FE
      ToolTipText     =   "Help"
      Top             =   990
      Width           =   480
   End
   Begin VB.Menu mnuRadix 
      Caption         =   "ASTRO"
      Visible         =   0   'False
      Begin VB.Menu mnuAscMc 
         Caption         =   "Ascendant/&MC"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHouses 
         Caption         =   "&Houses"
      End
      Begin VB.Menu mnuAspects 
         Caption         =   "&Aspects"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFixedZodiak 
         Caption         =   "&Fixed Zodiak"
      End
      Begin VB.Menu mnuVariantRadii 
         Caption         =   "&Variant radii"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MseDwn As Boolean
Dim MaxPause As Long
Dim MinPause As Long
Private Radix As RADIXPARAM

' calibrates the speed of the commandbuttons
Private Sub GetMinMaxPause()
   Dim I As Long
   Dim sTime As Variant
   
   sTime = Timer
   While Timer - sTime < 0.375: I = I + 1: DoEvents: Wend
   MaxPause = I
   MinPause = I / 10
End Sub

' generates 40 symbols with a certain square-size
' only used ones in this program, but offers more possibilities
Private Sub SymbolsNewSize(ByVal NewSize As Long)
   SymbSize = NewSize
   picSymb.Move picSymb.Left, _
                picSymb.Top, _
                (SymbSize + 1) * 10, _
                (SymbSize + 1) * 4
   picSymb.Cls
   picMask.Move picMask.Left, _
                picMask.Top, _
                (SymbSize + 1) * 10, _
                (SymbSize + 1) * 4
   picMask.Cls
   SymbolsRedraw
End Sub

' in other programs (not here), the backcolor could be changed
' to white for printing purposes
' here SymbolsNewSize & SymbolsRedraw could be fused
Private Sub SymbolsRedraw()
   Dim Nr As Long
   Dim X As Long, Y As Long
   SetFont picSymb, ZodiakLgFnt, "WingDings", SymbSize * 22 / 32, False, False
 
   For Nr = 0 To 39
      X = (Nr Mod 10) * (SymbSize + 1)
      Y = (Nr \ 10) * (SymbSize + 1)
      If Nr < 12 Then
         SymbolsDrawZodiac picSymb.hdc, Nr, X + SymbSize / 2, Y + SymbSize / 2, SymbSize
      ElseIf Nr >= 12 And Nr < 34 Then
         SymbolsDrawObject picSymb.hdc, X, Y, SymbSize, SymbSize, DOB(Nr - 12)
      End If
   Next Nr
   picSymb.Refresh
   BWMask = True
   For Nr = 0 To 39
      X = (Nr Mod 10) * (SymbSize + 1)
      Y = (Nr \ 10) * (SymbSize + 1)
      If Nr < 12 Then
         SymbolsDrawZodiac picMask.hdc, Nr, X + SymbSize / 2, Y + SymbSize / 2, SymbSize
      ElseIf Nr >= 12 And Nr < 34 Then
         SymbolsDrawObject picMask.hdc, X, Y, SymbSize, SymbSize, DOB(Nr - 12)
      End If
   Next Nr
   picMask.Refresh
   BWMask = False
End Sub

' draws 1 of 12 zodiac symbols using WingDings font
Private Sub SymbolsDrawZodiac(ByVal hdc As Long, _
                  ByVal Nr As Long, _
                  ByVal pX As Long, ByVal pY As Long, _
                  ByVal Size As Long)
                  
   Dim Centre As Long
   Dim B As Long
   Dim X As Long, Y As Long
   Dim pt As POINTAPI, sze As Size
   Dim hPen As Long, hOldPen As Long
   Dim hBrush As Long, hOldBrush As Long
   Dim LgFnt As LOGFONT
   Dim hFnt As Long, hOldFnt As Long
   Dim hMemDC As Long
   Dim hOldBitmap As Long
   
   Centre = Size / 2
   hMemDC = CreateCompatibleDC(hdc)
   hOldBitmap = SelectObject(hMemDC, CreateCompatibleBitmap(hdc, Size, Size))
   'cls
   hPen = CreatePen(0, 0, IIf(BWMask, QBColor(15), SymbColor(skAchtergrond)))
   hOldPen = SelectObject(hMemDC, hPen)
   hBrush = CreateSolidBrush(IIf(BWMask, QBColor(15), SymbColor(skAchtergrond)))
   hOldBrush = SelectObject(hMemDC, hBrush)
   Rectangle hMemDC, 0, 0, Size, Size
   hOldBrush = SelectObject(hMemDC, hOldBrush)
   DeleteObject hBrush
   hOldPen = SelectObject(hMemDC, hOldPen)
   DeleteObject hPen
   
   SetBkMode hMemDC, 1
   
   hFnt = CreateFontIndirect(ZodiakLgFnt)
   hOldFnt = SelectObject(hMemDC, hFnt)
   GetTextExtentPoint32 hMemDC, Chr(94 + Nr), 1, sze
   SetTextColor hMemDC, 0
   If BWMask = False Then
      If SymbColor(skAchtergrond) = QBColor(7) Then
         SetTextColor hMemDC, Choose((Nr Mod 4) + 1, QBColor(14), QBColor(15), QBColor(0), QBColor(11))
         Else
         SetTextColor hMemDC, Choose((Nr Mod 4) + 1, QBColor(14), QBColor(7), QBColor(0), QBColor(11))
         End If
      End If
   X = Centre - sze.cx / 2 + Choose(Nr + 1, 1, 1, 0, 1, 1, 1, 0, 1, 1, 1, 0, 1)
   Y = Centre - sze.cy / 2
   TextOut hMemDC, X, Y, Chr(94 + Nr), 1
   If BWMask = False Then
      SetTextColor hMemDC, Choose((Nr Mod 4) + 1, Red, Black, Yellow, Blue)
      End If
   TextOut hMemDC, X - 1, Y - 1, Chr(94 + Nr), 1
   SelectObject hMemDC, hOldFnt
   DeleteObject hFnt
   
   BitBlt hdc, pX - Centre, pY - Centre, Size, Size, hMemDC, 0, 0, SRCCOPY
   DeleteObject (SelectObject(hMemDC, hOldBitmap))
   DeleteDC hMemDC
End Sub

' the user changed latitude and longitude
Private Sub PlaceShowPosition()
   Dim X As Single
   Dim Y As Single

   X = Val(txtOL) / 2 + 90
   Y = Val(txtNB) / 2 + 45
   picWorld.Cls
   picWorld.Line (0, picWorld.ScaleHeight - Y)-(picWorld.ScaleWidth, picWorld.ScaleHeight - Y), QBColor(12)
   picWorld.Line (X, 0)-(X, picWorld.ScaleHeight), QBColor(12)
   
   CalcI.pWidth = Val(txtNB)
   CalcI.pLength = Val(txtOL)
   Call NewRadix
End Sub

Private Sub PlaceShowText(ByVal X As Single, ByVal Y As Single)
   txtOL.Text = ((X - 90) Mod 90) * 2
   txtNB.Text = ((45 - Y) Mod 45) * 2
End Sub

Private Sub NewRadix()
   CalcRadix CalcI, CalcO
   picRadix_Paint
   picPos_Paint
End Sub

' only at the start of the program
Private Sub DrawBackground()
   Dim W As Long, H As Long, Wa As Long, Ha As Long
   Dim X As Long, Y As Long, dX As Long, dY As Long
   Dim I As Long
   Static RndDone As Boolean
   
   'Random background
   If RndDone = False Then
      Randomize
      dX = picBG.ScaleWidth - 1
      dY = picBG.ScaleHeight - 1
      picBG.Cls
      For I = 0 To 150
         X = Int(dX * Rnd + 1)
         Y = Int(dY * Rnd + 1)
         picBG.PSet (X, Y), QBColor(8)
         picBG.PSet (X + 1, Y + 1), QBColor(15)
      Next I
      RndDone = True
      End If
   
   'Tile
   W = 128: H = 128
   Wa = (Me.ScaleWidth \ W) + 1
   Ha = (Me.ScaleHeight \ H) + 1
   For Y = Ha To 0 Step -1
      For X = 0 To Wa
         BitBlt Me.hdc, X * W, Y * H, W, H, picBG.hdc, 0, 0, SRCCOPY
      Next X
   Next Y
   
   'Title reliëf (3D)
   Fontname = "Times New roman": Fontsize = 26
   X = 10: Y = 10
   ForeColor = QBColor(0)
   CurrentX = X - 1: CurrentY = Y - 1: Print "Astro Mover"
   ForeColor = QBColor(15)
   CurrentX = X + 1: CurrentY = Y + 1: Print "Astro Mover"
   ForeColor = QBColor(7)
   CurrentX = X: CurrentY = Y: Print "Astro Mover"
  
End Sub

' continu-push-button-system
Private Sub cmdPlace_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Pze As Long, Pause As Long
   Dim pX As Single
   Dim pY As Single

   pX = Val(txtOL) / 2 + 90   ' get current position
   pY = 45 - Val(txtNB) / 2
   
   Pause = MaxPause           ' max calibrated at start
   MseDwn = True
   While MseDwn = True
      Select Case Index
         Case 0: pX = pX - 1
         Case 1: pX = pX + 1
         Case 2: pY = pY - 1
         Case 3: pY = pY + 1
      End Select
      If pX > 0 And pX < 180 And pY > 0 And pY < 90 Then
         PlaceShowText pX, pY
         PlaceShowPosition
         Else
         MseDwn = False
         End If
      For Pze = 0 To Pause: DoEvents: Next Pze     ' delay
      If Pause > MinPause Then Pause = Pause - 100 ' acceleration
   Wend

End Sub

' terminate loop in cmdPlace_MouseDown
Private Sub cmdPlace_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   MseDwn = False
End Sub

Private Sub cmdTimeDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim datum As Variant
   Dim ndatum As Variant
   Dim Pze As Long, Pause As Long
   
   Pause = MaxPause
   datum = txtDate.Text
   MseDwn = True
   While MseDwn = True
      Select Case Index
      Case 0: ndatum = DateAdd("d", -1, datum)     ' day
      Case 1: ndatum = DateAdd("m", -1, datum)     ' month
      Case 2: ndatum = DateAdd("yyyy", -1, datum)  ' year
      Case 3: ndatum = DateAdd("h", -1, datum)     ' hour
      Case 4: ndatum = DateAdd("n", -1, datum)     ' minute
      End Select
      txtDate.Text = Format(ndatum, "dd/mm/yyyy hh.nn")
      CalcI.Day = Day(txtDate.Text)
      CalcI.Month = Month(txtDate.Text)
      CalcI.Year = Year(txtDate.Text)
      CalcI.Hour = Hour(txtDate.Text) + Minute(txtDate.Text) / 100
      Call NewRadix
      datum = ndatum
      For Pze = 0 To Pause: DoEvents: Next Pze     ' delay
      If Pause > MinPause Then Pause = Pause - 50  ' acceleration
   Wend
End Sub

Private Sub cmdTimeDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   MseDwn = False
End Sub

' TimeUp and Down could be reassembled in one subroutine
' only the DateAdd direction is different
Private Sub cmdTimeUp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim datum As Variant
   Dim ndatum As Variant
   Dim Pze As Long, Pause As Long
   
   Pause = MaxPause
   datum = txtDate.Text
   MseDwn = True
   While MseDwn = True
      Select Case Index
      Case 0: ndatum = DateAdd("d", 1, datum)
      Case 1: ndatum = DateAdd("m", 1, datum)
      Case 2: ndatum = DateAdd("yyyy", 1, datum)
      Case 3: ndatum = DateAdd("h", 1, datum)
      Case 4: ndatum = DateAdd("n", 1, datum)
      End Select
      txtDate.Text = Format(ndatum, "dd/mm/yyyy hh.nn")
      CalcI.Day = Day(txtDate.Text)
      CalcI.Month = Month(txtDate.Text)
      CalcI.Year = Year(txtDate.Text)
      CalcI.Hour = Hour(txtDate.Text) + Minute(txtDate.Text) / 100
      Call NewRadix
      datum = ndatum
      For Pze = 0 To Pause: DoEvents: Next Pze
      If Pause > MinPause Then Pause = Pause - 50
   Wend
End Sub

Private Sub cmdTimeUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   MseDwn = False
End Sub

Private Sub Form_Load()
   GetConstantData
   SymbSize = 24
   SymbolsNewSize SymbSize
  
   CalcI.Day = Day(Now)
   CalcI.Month = Month(Now)
   CalcI.Year = Year(Now)
   CalcI.Hour = Hour(Now)
   CalcI.Hour = Hour(Now) + Minute(Now) / 100
   CalcI.TimeZone = 1
   CalcI.SummerTime = 0
   CalcI.pLength = 2.4
   CalcI.pWidth = 51.05
   
   CalcRadix CalcI, CalcO

   SetFont picRadix, Radix.HouseNoLgFnt, "Arial", 6, False, False
   Radix.ShowAspects = True
   Radix.ShowAscMc = True
   
   txtNB.Text = Str(CalcI.pWidth) ' latitude
   txtOL.Text = Str(CalcI.pLength)  ' longitude
   txtDate.Text = Format(Now, "dd/mm/yyyy hh.nn")
   
   PosWidth = picPos.ScaleWidth
   PosHeight = picPos.ScaleHeight
   SetFont picPos, PosLgFnt, "Courier New", 10, True, False

   DrawBackground
   Radix.Width = picRadix.ScaleWidth
   Radix.Height = picRadix.ScaleHeight

   GetMinMaxPause
End Sub

Private Sub imgHelp_Click()
   On Error Resume Next
   AppActivate "help.wri - WordPad", False
   If Err = 0 Then Exit Sub
   Err = 0
   Shell "write.exe " & Path & "\help.wri", vbNormalFocus
   If Err <> 0 Then MsgBox Err.Description
   On Error GoTo 0
End Sub

Private Sub mnuFixedZodiak_Click()
   If mnuFixedZodiak.Checked = True Then
      mnuFixedZodiak.Checked = False
      Radix.TurnWhat = 0
      Else
      mnuFixedZodiak.Checked = True
      Radix.TurnWhat = 1
      End If
   DrawRadix picRadix, Radix, CalcO
End Sub

Private Sub mnuVariantRadii_Click()
   If mnuVariantRadii.Checked = True Then
      mnuVariantRadii.Checked = False
      Radix.RadiiType = 2
      Else
      mnuVariantRadii.Checked = True
      Radix.RadiiType = 0
      End If
   DrawRadix picRadix, Radix, CalcO
End Sub

' fix ToolTipText with position info
Private Sub picPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim I As Long, txt As String
   I = Y \ 25 + 1
   If I > 13 Then Exit Sub
   txt = " " & hemell(I) & " is on degree "
   txt = txt & CalcO.Ps(I).Degree & " of " & Trim(Zd(CalcO.Ps(I).Zod).nm) & " "
   
   If CalcO.Ps(I).Retro <> " " Then
      txt = txt & "and moves "
      If CalcO.Ps(I).Retro = "R" Then txt = txt & "retrograde (backward) "
      If CalcO.Ps(I).Retro = "D" Then txt = txt & "direct (forward)"
      End If
   
   picPos.ToolTipText = txt
End Sub

Private Sub picPos_Paint()
   frmAMhDC = frmAM.hdc
   SymbolhDC = picSymb.hdc
   SymbMaskhDC = picMask.hdc
   DrawPositionList picPos, PosWidth, PosHeight, CalcO
End Sub

Private Sub picRadix_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PopupMenu mnuRadix, 4
End Sub

Private Sub picWorld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PlaceShowText X, Y
   PlaceShowPosition
End Sub

Private Sub picWorld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   If X < picWorld.ScaleWidth And _
      Y < picWorld.ScaleHeight And _
      X > 0 And _
      Y > 0 Then
      PlaceShowText X, Y
      PlaceShowPosition
      End If
   End If
End Sub

Private Sub picWorld_Paint()
   PlaceShowPosition
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      CalcI.Day = Day(txtDate.Text)
      CalcI.Month = Month(txtDate.Text)
      CalcI.Year = Year(txtDate.Text)
      CalcI.Hour = Hour(txtDate.Text) + Minute(txtDate.Text) / 100
      Call NewRadix
      End If
End Sub

Private Sub txtNB_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      PlaceShowPosition
      End If
End Sub

Private Sub txtOL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      PlaceShowPosition
      End If
End Sub

Private Sub mnuAscMc_Click()
   If mnuAscMc.Checked = True Then
      mnuAscMc.Checked = False
      Else
      mnuAscMc.Checked = True
      End If
   Radix.ShowAscMc = mnuAscMc.Checked
   DrawRadix picRadix, Radix, CalcO
End Sub

Private Sub mnuAspects_Click()
   If mnuAspects.Checked = True Then
      mnuAspects.Checked = False
      Else
      mnuAspects.Checked = True
      End If
   Radix.ShowAspects = mnuAspects.Checked
   DrawRadix picRadix, Radix, CalcO
End Sub

Private Sub mnuHouses_Click()
   If mnuHouses.Checked = True Then
      mnuHouses.Checked = False
      Else
      mnuHouses.Checked = True
      End If
   Radix.ShowHouses = mnuHouses.Checked
   CalcRadix CalcI, CalcO
   DrawRadix picRadix, Radix, CalcO
End Sub

Private Sub picRadix_Paint()
   frmAMhDC = frmAM.hdc
   SymbolhDC = picSymb.hdc
   SymbMaskhDC = picMask.hdc
   DrawRadix picRadix, Radix, CalcO
End Sub

Private Sub txtTZ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      CalcI.TimeZone = Val(txtTZ.Text)
      PlaceShowPosition
      End If
End Sub


