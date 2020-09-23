VERSION 5.00
Begin VB.Form frmStris 
   AutoRedraw      =   -1  'True
   Caption         =   "Stris by: Enrique A. Flores B."
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "frmStris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBalls 
      Caption         =   "&Extent with balls"
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Top             =   5295
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   315
      TabIndex        =   16
      Top             =   5175
      Width           =   1515
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   7365
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Timer tmrPlay 
      Left            =   7305
      Top             =   1005
   End
   Begin VB.Timer Tmr1 
      Left            =   7305
      Top             =   360
   End
   Begin VB.PictureBox picSqrs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   7335
      Picture         =   "frmStris.frx":030A
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   4860
      Left            =   2100
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   1
      Top             =   150
      Width           =   4995
   End
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      Height          =   1020
      Left            =   540
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   870
      Width           =   1020
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   810
      MouseIcon       =   "frmStris.frx":1048
      MousePointer    =   99  'Custom
      Picture         =   "frmStris.frx":1352
      ToolTipText     =   "Help"
      Top             =   4410
      Width           =   480
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   3
      Left            =   300
      TabIndex        =   11
      Top             =   3825
      Width           =   855
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   300
      TabIndex        =   10
      Top             =   3405
      Width           =   855
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   300
      TabIndex        =   9
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Index           =   0
      Left            =   585
      TabIndex        =   8
      Top             =   2025
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmStris.frx":165C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   0
      Left            =   1170
      TabIndex        =   5
      Top             =   3795
      Width           =   570
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   0
      Left            =   1170
      TabIndex        =   4
      Top             =   3390
      Width           =   570
   End
   Begin VB.Label lblLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   0
      Left            =   1170
      TabIndex        =   3
      Top             =   2985
      Width           =   570
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   2385
      Width           =   1515
   End
   Begin VB.Label lblLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   1200
      TabIndex        =   14
      Top             =   2985
      Width           =   570
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   3390
      Width           =   570
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   3795
      Width           =   570
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   1
      Left            =   300
      TabIndex        =   15
      Top             =   2400
      Width           =   1515
   End
End
Attribute VB_Name = "frmStris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*** playing Field ***
Dim mField(31, 21) As Long
'***  objects / 'pieces' build with '4 squares'  ***
Dim obj(8, 6) As Long            ' objectdata
Dim nxtobjnr As Long, objnr As Long
Dim ox As Long, oy As Long       ' place in mField
Dim hox(3) As Long               ' rel. place of squares now
Dim hoy(3) As Long
Dim nox(3) As Long               ' rel. place of squares next (preview)
Dim noy(3) As Long
Dim vox(3) As Long               ' 4! squares to remember (to erase previous)
Dim voy(3) As Long
Dim vobjfl As Long
Dim SLevel(99) As STRISLEVEL
'*** other ***
Dim Score As Long, Level As Long, Lines As Long, vLin As Long
Dim mTime  As Single             ' the speed of the falling pieces (increasing from start-speed)
Dim Stat As Long                 ' is a piece still moveable
Dim KeyPze As Boolean            ' =False when SPACE key is pushed = piece falls fast
Dim Playing As Boolean           ' game is on, no pauzing currently
Dim ContFl As Boolean            ' ok to continue flag
Dim StartLevel As Long, StartedAt As Long
Dim PieceMode As Long            ' determines the no. of pieces-types
                                 '     7='without balls'   8='with balls'
Dim Gr As Long                   ' size of one square in a piece = constant
Dim px1 As Long, px2 As Long     ' size of the playing field of the current level
Dim py1 As Long, py2 As Long

Private Sub GameOver()
   Dim i As Long
   ReDim names(1 To 7) As String, scores(1 To 7) As Long
   Dim txt As String
   
   ' turn everything off
   Playing = False
   KeyPreview = False
   KeyPze = True
   tmrPlay.Enabled = False
   Tmr1.Enabled = False
      
   'judge score
   names(1) = "Super       ": scores(1) = 1000000
   names(2) = "Expert      ": scores(2) = 800000
   names(3) = "Experienced ": scores(3) = 600000
   names(4) = "Good        ": scores(4) = 400000
   names(5) = "Pupil       ": scores(5) = 200000
   names(6) = "Beginner    ": scores(6) = 100000
   names(7) = "Poor        ": scores(7) = 50000
   For i = 1 To 7
      If Score > scores(i) Then Exit For
   Next i
   If i > 7 Then i = 7
   DialogTitle = "Stris - Game over"
   txt = "With a score of   " & Format(Score) & vbCrLf
   txt = txt & "your achievement is situated" & vbCrLf
   txt = txt & "on the niveau :" & vbCrLf & vbCrLf
   txt = txt & names(i) & vbCrLf
   DialogText = txt
   frmDial.Show 1
   
   ' next Start from which Level?
   StartLevel = (Level \ 3) * 3
   Level = StartLevel
   SetLevel
   
   ' turn everything off again to be on the save side
   tmrPlay.Enabled = False
   Tmr1.Enabled = False
   Playing = False
   chkBalls.Enabled = True
   cmdStartStop.Caption = "&Start"
End Sub
' Sets a certain level
Private Sub SetLevel()
   Static fout
   Dim X As Long, Y As Long
   Dim i As Long, vx2 As Long
   Dim T As String, R As String
   
   On Error GoTo SetLevelError:
   px2 = SLevel(Level).px2       ' get new width of playing field
   mTime = SLevel(Level).mTime   ' set mTime to start time (the higher, the slower)
   For Y = 0 To 21               ' get mField from Level-data
     R = SLevel(Level).R(Y)
     For X = 0 To px2
       mField(X, Y) = Val(Mid(R, X + 1, 1))
     Next X
   Next Y

SetLevelNext:
   On Error GoTo 0
   vx2 = (px2 - 1) * Gr          ' calc. size in pixels of curr. mField
   P1.Width = vx2 + 60           ' and set picturebox
   ShowField
   Width = P1.Left + P1.Width + 180
   lblLevel(0).Caption = Level: lblLevel(1).Caption = lblLevel(0).Caption
   lblSpeed(0).Caption = mTime: lblSpeed(1).Caption = lblSpeed(0).Caption
   Exit Sub
   
SetLevelError:
   ClsField                      ' generate a random mField in case of an error or all levels are done.
   Randomize
   For i = 0 To 9
     X = Int(Rnd(px2 - 1) * (px2 - 1)) + 1
     Y = 20 - Int(Rnd(10) * 10) + 1
     mField(X, Y) = 8
   Next i
   mTime = 950
   Resume SetLevelNext:

End Sub

' prepare next piece
Private Sub MakeNextObj()
   Dim i As Integer, X As Long, Y As Long

   Randomize Timer
   nxtobjnr = Int(Rnd(PieceMode) * PieceMode) + 1
   For i = 0 To 2
      nox(i + 1) = obj(nxtobjnr, i * 2)
      noy(i + 1) = obj(nxtobjnr, i * 2 + 1)
   Next i
   
   ' show in preview window(picturebox)
   P2.Cls
   For X = 1 To 4: For Y = 1 To 4
      DrawSquare P2.hDC, X * 16, Y * 16, 0
   Next Y: Next X
   For i = 0 To 3
      X = (nox(i) + 2) * 16
      Y = (noy(i) + 2) * 16
      DrawSquare P2.hDC, X, Y, nxtobjnr
   Next i
End Sub

' make ready for start
Private Sub MakeReady()
    MakeNextObj
    SetLevel
    lblScore(0).Caption = Score: lblScore(1).Caption = lblScore(0).Caption
    lblLines(0).Caption = 0: lblLines(1).Caption = lblLines(0).Caption
    TakeNextObj
End Sub

' next Level --> next piece at the top-middle
Private Sub NextLevel()
    Level = Level + 1: SetLevel: ox = px2 / 2: oy = 1
End Sub
' drop pieces (Obj) 1 position down
Private Sub ObjDown()
   oy = oy + 1          ' test next position
   CheckSituation
   If Stat = 1 Then
      If oy <= 2 Then GameOver: Exit Sub
      oy = oy - 1       ' couldn't go further --> keep previous pos.
      StoreField        ' and store it in mField (pieces remain there from now one)
      CheckLines          ' new full-linnes made?
      TakeNextObj       ' take already prepared next piece from preview
      Score = Score + 100
      lblScore(0).Caption = Score: lblScore(1).Caption = lblScore(0).Caption
      KeyPze = True
      vobjfl = 1
      DrawPiece
      vobjfl = 0
      tmrPlay.Interval = mTime
      Else
      DrawPiece
      vobjfl = 0
      End If
End Sub

' draw a piece (obj) in playing field
Private Sub DrawPiece()
   Dim i As Long
   
   If vobjfl = 0 Then ' clear previous position
      For i = 0 To 3: DrawSquare P1.hDC, vox(i) * 16, voy(i) * 16, 0: Next i
      End If
   For i = 0 To 3 ' draw new position
      DrawSquare P1.hDC, (ox + hox(i)) * 16, (oy + hoy(i)) * 16, objnr
      vox(i) = ox + hox(i): voy(i) = oy + hoy(i)
   Next i
   P1.Refresh
End Sub

' store a piece that got stuck and make it permanent
Private Sub StoreField()
   Dim i As Long
   For i = 0 To 3: mField(ox + hox(i), oy + hoy(i)) = objnr: Next i
End Sub

' take the prepared piece (in preview), make it the current
' prepare a new one
Private Sub TakeNextObj()
   Dim i As Long
   ox = px2 \ 2: oy = 1
   objnr = nxtobjnr
   For i = 1 To 3: hox(i) = nox(i): hoy(i) = noy(i): Next i
   MakeNextObj
End Sub

' draw the whole playing field
Private Sub ShowField()
   Dim X As Integer, Y As Integer
   
   P1.Cls
   For X = px1 To px2 - 1: For Y = py1 To py2 - 1
      DrawSquare P1.hDC, X * 16, Y * 16, mField(X, Y)
   Next Y: Next X
   P1.Refresh
End Sub

' turn a piece (obj)
Private Sub Turn(R As Integer)
   Dim z As Integer, i As Integer
   
   If objnr > 6 Then Exit Sub
   For i = 1 To 3
      If R = 1 Then z = hox(i): hox(i) = hoy(i): hoy(i) = -z
      If R = 2 Then z = hox(i): hox(i) = -hoy(i): hoy(i) = z
   Next i
End Sub
' generate pauze
Private Sub Pauze(ds As Integer)
   Dim sec As Single
   Dim td As Variant
   
   sec = ds / 1000
   td = Timer
   While Timer - td < sec: DoEvents: Wend
End Sub

' draw one single piece-square
Private Sub DrawSquare(ByVal phDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal tpe As Integer)
BitBlt phDC, X - 16, Y - 16, 16, 16, frmStris.picSqrs.hDC, 0, tpe * 16, SRCCOPY
End Sub

' at the start of this program...
Private Sub DrawBackground()
   Dim W As Long, H As Long, Wa As Long, Ha As Long
   Dim X As Long, Y As Long, dX As Long, dY As Long
   Dim i As Long
   Static RndDone As Boolean
   
   'Random background
   If RndDone = False Then
      Randomize
      dX = picBg.ScaleWidth - 1
      dY = picBg.ScaleHeight - 1
      picBg.Cls
      For i = 0 To 750
         X = Int(dX * Rnd + 1)
         Y = Int(dY * Rnd + 1)
         picBg.PSet (X, Y), QBColor(8)
         picBg.PSet (X + 1, Y + 1), QBColor(15)
      Next i
      RndDone = True
      End If
   
   'Tile
   W = 128: H = 128
   Wa = (Me.ScaleWidth \ W) + 1
   Ha = (Me.ScaleHeight \ H) + 1
   For Y = Ha To 0 Step -1
      For X = 0 To Wa
         BitBlt Me.hDC, X * W, Y * H, W, H, picBg.hDC, 0, 0, SRCCOPY
      Next X
   Next Y
   
   'Title
   FontName = "Times New roman": FontSize = 26
   X = 780: Y = 110
   ForeColor = QBColor(0)
   CurrentX = X - 15: CurrentY = Y - 1: Print "Stris"
   ForeColor = QBColor(15)
   CurrentX = X + 15: CurrentY = Y + 1: Print "Stris"
   ForeColor = QBColor(7)
   CurrentX = X: CurrentY = Y: Print "Stris"
End Sub

' check if new full lines are made
Private Sub CheckLines()
Dim fl As Integer, X As Integer, Y As Integer, xX As Integer, yY As Integer
    
    While fl = 0: fl = 1      ' repeat until there are no more full lines found
      For Y = 20 To 2 Step -1 ' from top to bottom
        X = 0
        Do: X = X + 1: Loop Until mField(X, Y) = 0 Or X = px2
        If X = px2 Then       ' no hole = full line
          Lines = Lines + 1
          lblLines(0).Caption = Lines: lblLines(1).Caption = lblLines(0).Caption
          For yY = Y To 1 Step -1: For X = 1 To px2 ' erase line
          mField(X, yY) = mField(X, yY - 1): Next X: Next yY
          For X = 1 To px2 - 1: mField(X, 1) = 0: Next X
          Score = Score + 1000
          lblScore(0).Caption = Score: lblScore(1).Caption = lblScore(0).Caption
          fl = 0
          ShowField
          End If
      Next Y
    Wend
    ' look for remaining balls(=8) in mField
    fl = 0
    For X = 1 To px2 - 1: For Y = 1 To 20
    If mField(X, Y) = 8 Then fl = 1: Exit For
    Next Y: Next X
    If fl = 0 Then ' no more balls
        NextLevel
        Score = Score + Abs(100000 \ (Lines - vLin)) ' the less lines, the more points
        If (Level - StartedAt) Mod 5 = 0 Then Score = Score + 50000 ' 6 levels since start?
        lblScore(0).Caption = Score: lblScore(1).Caption = lblScore(0).Caption
        vLin = Lines
        Pauze 2000
        Exit Sub
        End If

End Sub
' is a piece able to turn, drop further
Private Sub CheckSituation()
   Dim i As Long
   
   Stat = 0
   For i = 0 To 3
   If mField(ox + hox(i), oy + hoy(i)) <> 0 Then Stat = 1: Exit Sub
   Next i
End Sub

' clear playing field
Private Sub ClsField()
Dim X As Integer, i As Integer, Y As Integer
    For X = 0 To px2: For Y = 0 To 21: mField(X, Y) = 0: Next Y: Next X
    For i = 0 To 21: mField(0, i) = 1: mField(px2, i) = 1: Next i
    For i = 0 To px2: mField(i, 21) = 1: mField(i, 0) = 1: Next i
End Sub

' read all level data
Private Sub LoadLevels()
   Dim ch As Long
   Dim X As Long, Y As Long
   Dim i As Long
   Dim T As String, R As String

   ch = FreeFile
   Open App.Path & "\Levels.dat" For Input As ch
   For i = 0 To 99
      Line Input #ch, T
      Line Input #ch, T: SLevel(i).px2 = Val(T)
      Line Input #ch, T: SLevel(i).mTime = Val(T)
      For Y = 0 To 21
        Line Input #ch, SLevel(i).R(Y)
      Next Y
   Next i
   Close ch

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyS Then cmdStartStop_Click: KeyCode = 0: Exit Sub
   If KeyCode = 18 And Playing = True Then KeyCode = 0: Exit Sub
   If Playing = False Then Exit Sub
   KeyPze = True
   Select Case KeyCode
      Case vbKeyLeft: ox = ox - 1:  CheckSituation: If Stat = 1 Then ox = ox + 1 Else DrawPiece
      Case vbKeyRight: ox = ox + 1:  CheckSituation: If Stat = 1 Then ox = ox - 1 Else DrawPiece
      Case vbKeyUp: Turn 1: CheckSituation: If Stat = 1 Then Turn 2 Else DrawPiece
      Case vbKeyDown: ObjDown
      Case vbKeySpace
         KeyPze = False: tmrPlay.Enabled = False
         While KeyPze = False: ObjDown: Pauze 20: Wend
         If KeyPze = True And Playing = True Then tmrPlay.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   px1 = 1:  py1 = 1: py2 = 21: px2 = 20
   Gr = 16 * 15
   PieceMode = 7
   LoadLevels
   obj(1, 0) = -1: obj(1, 1) = 0: obj(1, 2) = 1: obj(1, 3) = 0: obj(1, 4) = 2: obj(1, 5) = 0
   obj(2, 0) = -1: obj(2, 1) = 0: obj(2, 2) = 1: obj(2, 3) = 0: obj(2, 4) = 1: obj(2, 5) = -1
   obj(3, 0) = -1: obj(3, 1) = 0: obj(3, 2) = 1: obj(3, 3) = 0: obj(3, 4) = 1: obj(3, 5) = 1
   obj(4, 0) = -1: obj(4, 1) = 0: obj(4, 2) = 1: obj(4, 3) = 0: obj(4, 4) = 0: obj(4, 5) = 1
   obj(5, 0) = -1: obj(5, 1) = 0: obj(5, 2) = 0: obj(5, 3) = 1: obj(5, 4) = 1: obj(5, 5) = 1
   obj(6, 0) = -1: obj(6, 1) = 0: obj(6, 2) = 0: obj(6, 3) = -1: obj(6, 4) = 1: obj(6, 5) = -1
   obj(7, 0) = -1: obj(7, 1) = 0: obj(7, 2) = -1: obj(7, 3) = 1: obj(7, 4) = 0: obj(7, 5) = 1
   obj(8, 0) = 0: obj(8, 1) = 0: obj(8, 2) = 0: obj(8, 3) = 0: obj(8, 4) = 0: obj(8, 5) = 0
   DrawBackground
   MakeReady
End Sub

Private Sub Form_Resize()
   DrawBackground
End Sub

Private Sub imgHelp_Click()
    frmAbout.Show
End Sub

Private Sub chkBalls_Click()
   If chkBalls.Value = 0 Then
      PieceMode = 7
      Else
      PieceMode = 8
      End If
End Sub

Private Sub cmdStartStop_Click()
   Dim txt As String
   
   Select Case cmdStartStop.Caption
   
   Case "&Start"
   
   DialogTitle = "Stris - Start game"
   txt = "After clicking the Start button," & vbCrLf
   txt = txt & "you will hear 3 beeps, followed by" & vbCrLf
   txt = txt & "the first piece starting to fall." & vbCrLf
   txt = txt & "Ready to start?"
   DialogText = txt
   OK = False: frmDial.Show 1
   If OK = False Then Exit Sub
         
   chkBalls.Enabled = False
   Level = StartLevel
   StartedAt = StartLevel
   Score = 0: Lines = 0: vLin = 0
   MakeReady
   Playing = True
   Beep
   Pauze 1000: Beep: Pauze 1000: Beep: Pauze 1000
   P1.SetFocus
   Tmr1.Enabled = True
   KeyPreview = True
   tmrPlay.Interval = mTime
   tmrPlay.Enabled = True
   cmdStartStop.Caption = "&Stop"
   
   Case "&Stop"
   
   KeyPze = True
   tmrPlay.Enabled = False
   Tmr1.Enabled = False
   Playing = False
   
   DialogTitle = "Stris - Pauze"
   txt = "Dear Stris Player" & vbCrLf & vbCrLf
   txt = txt & "You want to rest for a while," & vbCrLf
   txt = txt & "get some coffee, smoke a sigaret (please don't) ... ?" & vbCrLf
   txt = txt & "Or is it not your day, you have enough of it ?!" & vbCrLf
   txt = txt & "If I were you, I would continue!" & vbCrLf
   DialogText = txt
   frmDial.Show 1
   If OK = True Then
      P1.SetFocus
      Playing = True
      tmrPlay.Enabled = True
      Tmr1.Enabled = True
      tmrPlay.Interval = mTime
      Else
      chkBalls.Enabled = True
      Level = 0
      SetLevel
      Playing = False
      KeyPreview = False
      cmdStartStop.Caption = "&Start"
      End If

   End Select
End Sub

' preview
Private Sub P2_Click()
   If Playing = True Then Exit Sub
   MakeNextObj
End Sub

' tempo increase Timer
Private Sub Tmr1_Timer()
   If mTime > 150 Then mTime = mTime - 25
   lblSpeed(0).Caption = Format(mTime, "000")
   lblSpeed(1).Caption = lblSpeed(0).Caption
   tmrPlay.Interval = mTime
End Sub

' drop pieces (obj) Timer
Private Sub TmrPlay_Timer()
   ObjDown
End Sub

