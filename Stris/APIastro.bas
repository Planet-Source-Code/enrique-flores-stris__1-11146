Attribute VB_Name = "APIastro"
' The engine of this program, calculation of radixes, comes
' from old times (the beginning of the eighties) out of a
' magazine for the Commodore 64. Over years it changed a lot
' but there still some old variables and technics remaining.

' The routines are extracted from a more develloped and complex
' astrology program. Most of the routine still have
' the original potentials and can be used in more ways
' For example: radix and symbols can be sized

' For speed purposes, all drawings are done with API functions
' first in memory, then copied to the screen. This technic is
' often used to obtain flicker free moving images

Option Explicit
DefDbl A-Z
Public PI As Double, Deg As Double, Rad As Double
Public G0  As Double
Public Axx(12), H01(12), H11(12), H02(12), H12(12)
Public H03(12), H13(12), H04(12), H14(12), H05(12), H15(12)
Public asp(6), Aspc(6), asps(6) As String
Public hemell(13) As String   ' an object in the sky ???
Public Kra(10, 12) As String * 2
Public Red, Black, Yellow, Blue
Public AspectColor(6) As Long
Public PosWidth As Long
Public PosHeight As Long
Public PosLgFnt As LOGFONT
Public SymbolhDC As Long
Public SymbMaskhDC As Long
Public frmAMhDC As Long

Public Path As String

Type POSITION
    HelioZ As Double
    HelioX As Double
    HelioY As Double
    pLength As Double
    pWidth As Double
    Retro As String * 1
    Degree As String * 5
    Zod As Integer
End Type

Type HOUSEPOSITION
    pLength As Double
    Degree As String * 5
    Zod As Integer
End Type

Type CALC_INPUT
   Day As Long
   Month As Long
   Year As Long
   Hour As Double
   SummerTime As Long
   TimeZone As Long
   pLength As Single
   pWidth As Single
End Type
Public CalcI As CALC_INPUT

Type CALC_OUTPUT
   Ps(13) As POSITION
   House(12) As HOUSEPOSITION
End Type
Public CalcO As CALC_OUTPUT

Type RADIXPARAM
   Left As Long
   Top As Long
   Width As Long
   Height As Long
   ShowHouses As Boolean
   ShowAspects As Boolean
   ShowAscMc As Boolean
   TurnWhat As Long
   RadiiType As Long
   HouseNoLgFnt As LOGFONT
End Type

Type ZodiacSign
    nm As String * 12 ' name
    el As String * 2  ' element
    kw As String * 2  ' kwality
    hr As Integer     ' master
End Type
Public Zd(12) As ZodiacSign

'------------------------------- assembling symbols
Public Enum SYMBOLCOLORS
   skZon          ' sun
   skMaansikkel   ' moon sickle
   skAardekruis   ' earth cross
   skVuur         ' fire
   skAarde        ' earth
   skLucht        ' ear
   skWater        ' water
   skConjunctie   ' 0°
   skSextiel      ' 60°
   skVierkant     ' 90°
   skDriehoek     ' 120°
   skInconjunct   ' 150°
   skOppositie    ' 180°
   skAchtergrond  ' background
End Enum

Public Enum DRAWPIECETYPE
   DPMoveTo = 1
   DPLineTo = 2
   DPPolyLine = 3
   DPPolyBezierTo = 4
   DPPolygon = 5
   DPEllipse = 6
   DPRectangle = 7
End Enum

Public Const MAXPOINT = 50
Public Const MAXPIECE = 10
Public Const MAXOBJ = 30

Type DRAWPIECE
   Type As DRAWPIECETYPE
   Color As Long
   PTCount As Long
   pt(MAXPOINT) As POINTAPI
End Type

Type DRAWOBJECT
   Name As String * 25
   DPCount As Long
   DP(MAXPIECE) As DRAWPIECE
End Type
Public DOB(MAXOBJ) As DRAWOBJECT

Public ObjCount As Long
Public BWMask As Boolean
Public SymbColor(15) As Long
Public ZodiakLgFnt As LOGFONT
Public SymbSize As Long

Public Sub SetFont(pic As Control, LF As LOGFONT, ByVal Fontname As String, ByVal Fontsize As Single, ByVal FontBold As Boolean, ByVal FontItalic As Boolean)
   Dim TM As TEXTMETRIC
   Dim API As Long
   
   pic.Fontname = Fontname
   pic.Fontsize = Fontsize
   pic.FontBold = FontBold
   pic.FontItalic = FontItalic
   API = GetTextMetrics(pic.hdc, TM)
   LF.lfHeight = TM.tmHeight
   LF.lfWidth = TM.tmAveCharWidth
   LF.lfEscapement = 0
   LF.lfOrientation = 0
   LF.lfWeight = TM.tmWeight
   LF.lfItalic = TM.tmItalic
   LF.lfUnderline = TM.tmUnderlined
   LF.lfStrikeOut = TM.tmStruckOut
   LF.lfCharSet = TM.tmCharSet
   LF.lfOutPrecision = "0"
   LF.lfClipPrecision = "0"
   LF.lfQuality = "0"
   LF.lfPitchAndFamily = TM.tmPitchAndFamily
   API = GetTextFace(pic.hdc, 32, LF.lfFaceName)

End Sub


Public Sub SymbolsDrawObject(ByVal hdc As Long, _
                           ByVal pLeft As Long, ByVal pTop As Long, _
                           ByVal pWidth As Long, ByVal pHeight As Long, _
                           DOB As DRAWOBJECT)
   Dim I As Long
   Dim hMemDC As Long
   Dim hOldBitmap As Long
   Dim hPen As Long, hOldPen As Long
   Dim hBrush As Long, hOldBrush As Long
   Dim pt As POINTAPI, sze As Size
   Dim scaleX As Single, scaleY As Single

   scaleX = pWidth / 64
   scaleY = pHeight / 64
   hMemDC = CreateCompatibleDC(hdc)
   hOldBitmap = SelectObject(hMemDC, CreateCompatibleBitmap(hdc, pWidth, pHeight))

   'cls
   hPen = CreatePen(0, 0, IIf(BWMask, QBColor(15), SymbColor(skAchtergrond)))
   hOldPen = SelectObject(hMemDC, hPen)
   hBrush = CreateSolidBrush(IIf(BWMask, QBColor(15), SymbColor(skAchtergrond)))
   hOldBrush = SelectObject(hMemDC, hBrush)
   Rectangle hMemDC, 0, 0, pWidth, pHeight
   hOldBrush = SelectObject(hMemDC, hOldBrush)
   DeleteObject hBrush
   hOldPen = SelectObject(hMemDC, hOldPen)
   DeleteObject hPen

   For I = 0 To DOB.DPCount - 1
      SymbolsDrawPiece hMemDC, DOB.DP(I), scaleX, scaleY
   Next I

   BitBlt hdc, pLeft, pTop, pWidth, pHeight, hMemDC, 0, 0, SRCCOPY
   DeleteObject (SelectObject(hMemDC, hOldBitmap))
   DeleteDC hMemDC

End Sub

Public Sub SymbolsDrawPiece(ByVal hdc As Long, _
                           DP As DRAWPIECE, _
                           ByVal scaleX As Single, _
                           ByVal scaleY As Single)
   Dim I As Long
   Dim pt As POINTAPI
   Dim npt(MAXPOINT) As POINTAPI
   Dim hPen As Long, hOldPen As Long
   Dim hBrush As Long, hOldBrush As Long
   
   For I = 0 To DP.PTCount - 1
      npt(I).X = DP.pt(I).X * scaleX
      npt(I).Y = DP.pt(I).Y * scaleY
   Next I
   hPen = CreatePen(0, 0, IIf(BWMask, 0, SymbColor(DP.Color)))
   hOldPen = SelectObject(hdc, hPen)
   hBrush = CreateSolidBrush(IIf(BWMask, 0, SymbColor(DP.Color)))
   hOldBrush = SelectObject(hdc, hBrush)
   Select Case DP.Type
      Case DPMoveTo
         MoveToEx hdc, npt(0).X, npt(0).Y, pt
      Case DPLineTo
         LineTo hdc, npt(0).X, npt(0).Y
      Case DPPolyLine
         Polyline hdc, npt(0), DP.PTCount
      Case DPPolyBezierTo
         BeginPath hdc
         PolyBezierTo hdc, npt(0), DP.PTCount
         EndPath hdc
         StrokeAndFillPath hdc
      Case DPPolygon
         Polygon hdc, npt(0), DP.PTCount
      Case DPEllipse
         Ellipse hdc, npt(0).X, npt(0).Y, npt(1).X, npt(1).Y
      Case DPRectangle
         Rectangle hdc, npt(0).X, npt(0).Y, npt(1).X, npt(1).Y
  End Select
   hOldBrush = SelectObject(hdc, hOldBrush)
   DeleteObject hBrush
   hOldPen = SelectObject(hdc, hOldPen)
   DeleteObject hPen
   
End Sub
             
' hyperbool-cosinus ??? (dutch=boogcosinus)
Function BoogCosinus#(X#): BoogCosinus# = Atn(Sqr(1 - X# * X#) / X#) - PI * (X# < 0)
End Function

'boogsinus
Function BoogSinus#(X#)
On Error Resume Next
BoogSinus# = Atn(X# / Sqr(1 - X# * X#))
End Function



'minutes to degree
Function Minut2Degr#(X#)
Minut2Degr# = Int(X#) + Int((X# - Int(X#)) * 100 + 0.5) / 60
End Function


'sinus in degrees
Function SinDegr#(X#): SinDegr# = Sin(X# * PI / 180)
End Function

'mod 360
Function MinusMulti360#(X#): MinusMulti360# = X# - (Int(X# / 360) * 360)
End Function

'mod 24
Function MinusMulti24#(X#): MinusMulti24# = X# - (Int(X# / 24) * 24): End Function

Sub GetConstantData()
   Dim ch As Long, I As Long, Z As Long, P As Long
   Dim msg As String
   Dim ob As Long, DP As Long, pt As Long
   
   PI = 4 * Atn(1)
   Deg = PI / 180
   Rad = 180 / PI
   G0 = Sqr(0.0002959)
  
   'Kleuren
   Red = QBColor(12)
   Black = QBColor(0)
   Yellow = QBColor(14)
   Blue = QBColor(9)
   AspectColor(1) = RGB(0, 0, 0) 'conjunctie
   AspectColor(2) = RGB(0, 0, 196) 'sextiel
   AspectColor(3) = RGB(196, 0, 0) 'vierkant
   AspectColor(4) = RGB(0, 196, 0) 'driehoek
   AspectColor(5) = RGB(196, 196, 0) 'inconjunct
   AspectColor(6) = RGB(196, 0, 196) 'oppositie
   SymbColor(0) = QBColor(4)
   SymbColor(1) = QBColor(9)
   SymbColor(2) = QBColor(0)
   SymbColor(3) = Red
   SymbColor(4) = Black
   SymbColor(5) = Yellow
   SymbColor(6) = Blue
   SymbColor(7) = RGB(0, 0, 0)
   SymbColor(8) = RGB(0, 0, 128)
   SymbColor(9) = RGB(128, 0, 0)
   SymbColor(10) = RGB(0, 128, 0)
   SymbColor(11) = RGB(128, 128, 128)
   SymbColor(12) = RGB(128, 0, 128)
   SymbColor(13) = QBColor(7)

   Path = App.Path
   Path = IIf(Right(Path, 1) = "\", Left(Path, Len(Path) - 1), Path)
   On Error GoTo GetConstantDataError1:
   ch = FreeFile
   Open Path & "\astro.dat" For Input As ch
   ' signs of the zodiac
   For I = 1 To 12:
      Input #ch, Zd(I).nm, Zd(I).el, Zd(I).kw, Zd(I).hr
   Next I
   ' sky objects
   For I = 1 To 13:
      Input #ch, hemell(I): hemell(I) = Trim(hemell(I))
      If I < 11 Then
         Input #ch, Axx(I), H01(I), H11(I), H02(I), H12(I), H03(I), H13(I), H04(I), H14(I), H05(I), H15(I)
         End If
   Next I
   ' energy table of sky objects in signs - not used here
   For Z = 1 To 12: For P = 1 To 10: Input #ch, Kra(P, Z): Next P: Next Z
   ' aspects
   For I = 1 To 6
      Input #ch, asps(I), asp(I), Aspc(I)
   Next I
   Close ch
   
   On Error GoTo GetConstantDataError2:
   ch = FreeFile
   Open Path & "\symbols.dat" For Input As ch
   ob = 0
   Do While Not (EOF(ch))
      Input #ch, DOB(ob).Name, DOB(ob).DPCount
      For DP = 0 To DOB(ob).DPCount - 1
         Input #ch, DOB(ob).DP(DP).Type, _
                    DOB(ob).DP(DP).Color, _
                    DOB(ob).DP(DP).PTCount
         For pt = 0 To DOB(ob).DP(DP).PTCount - 1
            Input #ch, DOB(ob).DP(DP).pt(pt).X, DOB(ob).DP(DP).pt(pt).Y
         Next pt
      Next DP
      ob = ob + 1
   Loop
   Close ch
   ObjCount = ob
   On Error GoTo 0
   
   Exit Sub

GetConstantDataEinde:
   MsgBox msg
   End
GetConstantDataError1:
   msg = "Error:" & Str(Err) & " " & Error & vbCrLf
   msg = msg & App.Path & "\ASTRO.DAT is missing!" & vbCrLf & "AstroMover cannot start."
   Resume GetConstantDataEinde:
GetConstantDataError2:
   msg = "Error:" & Str(Err) & " " & Error & vbCrLf
   msg = msg & App.Path & "\SYMBOLS.DAT is missing!" & vbCrLf & "AstroMover cannot start."
   Resume GetConstantDataEinde:

End Sub

Sub DrawPositionList(pic As Control, ByVal Width As Long, ByVal Height As Long, RO As CALC_OUTPUT)
   Dim I As Long
   Dim X As Long, Y As Long
   Dim hMemDC As Long, hdc As Long       ' handles of memory and picturebox
   Dim hOldBitmap As Long                ' buffer
   Dim hBrush As Long, hOldBrush As Long ' handle-buffer brush
   Dim hFnt As Long, hOldFnt As Long     ' handle-buffer font

   hdc = pic.hdc
   hMemDC = CreateCompatibleDC(hdc)
   hOldBitmap = SelectObject(hMemDC, CreateCompatibleBitmap(hdc, Width, Height))
   
   SetBkMode hMemDC, 1
   BitBlt hMemDC, 0, 0, Width, Height, frmAMhDC, pic.Left, pic.Top, SRCCOPY
   hFnt = CreateFontIndirect(PosLgFnt)
   hOldFnt = SelectObject(hMemDC, hFnt)
   X = 15
   For I = 1 To 13
      Y = 2 + (I - 1) * 25
      If RO.Ps(I).Retro <> " " Then TextOut hMemDC, X - 10, Y + 5, RO.Ps(I).Retro, Len(RO.Ps(I).Retro)
      SymbolMasked hMemDC, SymbolhDC, SymbMaskhDC, 11 + I, X, Y
      TextOut hMemDC, X + 25, Y + 5, RO.Ps(I).Degree, Len(RO.Ps(I).Degree)
      SymbolMasked hMemDC, SymbolhDC, SymbMaskhDC, RO.Ps(I).Zod - 1, X + 67, Y
   Next I
   SelectObject hMemDC, hOldFnt
   DeleteObject hFnt

   BitBlt hdc, 0, 0, Width, Height, hMemDC, 0, 0, SRCCOPY
   DeleteObject (SelectObject(hMemDC, hOldBitmap))
   DeleteDC hMemDC

End Sub

Sub CalcRadix(Rinput As CALC_INPUT, Routput As CALC_OUTPUT)
   Dim M, Yy, Z, f, ZHour, L4, L5, B4, B5, Jd, r, t, u, v
   Dim I As Integer, K As Integer
   Dim S0, S1, S2
   Dim Ek, E1, E2
   Dim Mo, M1, Mc, mX
   Dim Ao, A1, A2, Ac, Ax
   Dim Dm, hh
   Dim Retro$
   Dim Exc, Inc, Kno, Per, Man, Arg, Ea
   Dim S8, S9, C8, C9, L8, L9
   Dim X, Y, X1, Y1, Y9, B9
   Dim Xo, Yo, Lo
   Dim RO, R8, R9
   Dim LM, Mm, Km, Ls, Ms
   Dim Lh, L1, Bh, B1, Bk
   Dim Rx, P1, P9, Vo, V1, V2, V3, po
   Dim Zz$, Z1s$, Z2s$, Z1 As Integer, Z2 As Integer, Z3 As Integer
   Dim Qq As Integer
   Dim Q$
   
   M = Rinput.Month
   Yy = Rinput.Year
   Z = Minut2Degr(Rinput.Hour) - Rinput.TimeZone - Rinput.SummerTime
   ZHour = Z
   f = Z / 24 'world time
   '
   L4 = Rinput.pLength: L5 = Sgn(L4) * Minut2Degr(Abs(L4))
   B4 = Rinput.pWidth: B5 = Deg * (Sgn(B4) * Minut2Degr(Abs(B4)))
   'calc Julian date
   '  jd = Day
   '  T = century counting from 01.01.1950
   Yy = Yy + (M <= 2): M = M - 12 * (M <= 2)
   Jd = Int(365.25 * Yy) + Int(30.6001 * (M + 1)) + Int(Yy / 400) - Int(Yy / 100)
   Jd = Jd - 712286 + Rinput.Day + f
   t = Jd / 36525
   '  S0 = greenwich-sideral time
   '  S1 = local sideral time in houres
   '  S2 = local sideral time in degrees resp. radials
   S0 = 6.67170278 + 0.0657098232 * Int(Jd) + 1.00273791 * ZHour
   S0 = MinusMulti24(S0)
   S1 = S0 + L5 / 15
   S2 = Deg * (MinusMulti360(S1 * 15))
   If S1 = 6 Then S1 = S1 + 0.001
   Q$ = Format(Int((S1 - Int(S1)) * 60 + 0.5), "00")
   Ek = Deg * (23.4458 - 0.13 * t)
   E1 = Cos(Ek)
   E2 = Sin(Ek)
   ' calculate ascendant & mc
   Mo = Tan(S2) / E1: M1 = Atn(Mo) ' labda MC
   If S1 <= 6 Then
      Mc = M1
   ElseIf S1 > 6 And S1 <= 18 Then
      Mc = M1 + PI
   Else
      Mc = M1 + 2 * PI
   End If
   mX = Mc
   Mc = MinusMulti360(Rad * Mc): Routput.Ps(13).pLength = Mc
   Ao = Sin(mX) * E2
   Dm = BoogSinus(Ao)
   A1 = Cos(S2) * E2 * Tan(B5 - Dm)
   A2 = Atn(A1)
   Ac = mX + PI / 2 + A2
   Ac = MinusMulti360(Rad * Ac): Routput.Ps(12).pLength = Ac
   Retro$ = " "
   Z = Ac: GoSub AbsCoordToZodiakCoord:
   Routput.Ps(12).Retro = Retro$: Routput.Ps(12).Degree = Zz$: Routput.Ps(12).Zod = Qq
   Z = Mc: GoSub AbsCoordToZodiakCoord
   Routput.Ps(13).Retro = Retro$: Routput.Ps(13).Degree = Zz$: Routput.Ps(13).Zod = Qq
   
   'planet positions
   For I = 1 To 10
      If I = 2 Then I = 3
      Ax = Axx(I) ' large axes of ellipsoide course
      hh = H01(I) + H11(I) * t: Exc = hh 'excentricity
      hh = H02(I) + H12(I) * t: Inc = Deg * (MinusMulti360(hh)) 'inclination
      hh = H03(I) + H13(I) * t: Kno = Deg * (MinusMulti360(hh)) 'Length of the raising lunar knot
      hh = H04(I) + H14(I) * t: Per = Deg * (MinusMulti360(hh)) 'perihelium-Length
      hh = H05(I) + H15(I) * t: Man = Deg * (MinusMulti360(hh)) 'average deviation
      Arg = Per - Kno 'argument of the knot
      Ea = Man
      For K = 1 To 5: Ea = Man + Exc * Sin(Ea): Next K
      'calc coordinates on the surface of the orbital flight
      u = Ax * (Cos(Ea) - Exc)
      v = Ax * Sqr(1 - Exc * Exc) * Sin(Ea)
      'calc heliocentr. cartesi. ecl. coord.
      S9 = Sin(Arg): C9 = Cos(Arg)
      S8 = Sin(Kno): C8 = Cos(Kno)
      X1 = C9 * u - S9 * v
      Y9 = S9 * u + C9 * v
      Y1 = Y9 * Cos(Inc)
      Z = Y9 * Sin(Inc)
      X = C8 * X1 - S8 * Y1
      Y = S8 * X1 + C8 * Y1
      Routput.Ps(I).HelioX = X: Routput.Ps(I).HelioY = Y: Routput.Ps(I).HelioZ = Z
      'calc helioc. cart. ---> helioc. poolcoord.
      GoSub CartCoordToPoolCoord
      If I = 1 Then
         Xo = X: Yo = Y
         Lo = L9: RO = r
         Z = MinusMulti360(Rad * (Lo + PI))
         Routput.Ps(1).pLength = Z: GoSub CheckRetrograde
         Else
         L8 = L9: R8 = r 'heliopolair planet
         'geocentric cart.
         X = X - Xo: Y = Y - Yo: GoSub CartCoordToPoolCoord
         R9 = r
         Routput.Ps(I).pLength = MinusMulti360(Rad * L9)
         Routput.Ps(I + 1).pWidth = B9
         Z = Routput.Ps(I).pLength: GoSub CheckRetrograde
         End If
      GoSub AbsCoordToZodiakCoord
      Routput.Ps(I).Retro = Retro$: Routput.Ps(I).Degree = Zz$: Routput.Ps(I).Zod = Qq
   Next I
   'moon
   LM = 64.3755 + 481267.882 * t - 0.0022 * t * t        ' avg. course moon
   Mm = 215.5315 + 477198.859 * t + 0.009199998 * t * t  ' avg. deviation
   Km = 12.1128 - 1934.1399 * t - 0.0021 * t * t         ' raising lunar knot
   Ls = 280.6967 + 36000.7692 * t + 0.0003 * t * t       ' avg. course sun
   Ms = 358.0007 + 35999.0496 * t - 0.0002 * t * t       ' avg. deviation sun
   Z = MinusMulti360(Km): Routput.Ps(11).pLength = Z: GoSub AbsCoordToZodiakCoord
   Routput.Ps(11).Retro = Retro$: Routput.Ps(11).Degree = Zz$: Routput.Ps(11).Zod = Qq
   'start calc lunar-length LM1
   Lh = 2 * (LM - Ls): Retro$ = " "
   L1 = LM + 6.2889 * SinDegr(Mm) + 1.275 * SinDegr(Lh - Mm)
   L1 = L1 + 0.6583 * SinDegr(Lh) + 0.2136 * SinDegr(2 * Mm) - 0.1144 * SinDegr(2 * (LM - Km))
   L1 = L1 - 0.1872 * SinDegr(Ms) + 0.0575 * SinDegr(Lh - Mm - Ms) + 0.0533 * SinDegr(Lh + Mm)
   L1 = L1 + 0.0461 * SinDegr(Lh - Ms) + 0.0411 * SinDegr(Mm - Ms) - 0.0339 * SinDegr(Lh / 2)
   L1 = L1 - 0.0303 * SinDegr(Mm + Ms) + 0.0589 * SinDegr(Lh - 2 * Mm)
   L1 = MinusMulti360(L1)
   Z = L1: GoSub AbsCoordToZodiakCoord
   Routput.Ps(2).Retro = Retro$: Routput.Ps(2).Degree = Zz$: Routput.Ps(2).Zod = Qq
   Routput.Ps(2).pLength = L1
   'start calc lunar-length BM1
   Bh = L1 + Km - 2 * Ls: Bk = L1 - Km
   B1 = 5.15 * SinDegr(Bk) + 0.1467 * SinDegr(Bh)
   B1 = B1 + 0.0072 * SinDegr(2 * Mm + Bk) + 0.0069 * SinDegr(Bk - Ms) + 0.0067 * SinDegr(Bk + Ms)
   B1 = B1 + 0.0061 * SinDegr(Bh - Ms) - 0.0044 * SinDegr(Bh - Mm)
   B1 = B1 - 0.0038 * SinDegr(Mm) - 0.0028 * SinDegr(Bh + Ms) + 0.0036 * SinDegr(2 * Bk)
   
   'houses
   Dim J As Long
   Dim Ad, Sa, Th, aa, La, Dl, Vx, Lb, Ld, Ha, hhb
   Routput.House(1).pLength = Mc
   Ax = Tan(B5) * Tan(Dm)
   Ad = BoogSinus(Ax)
   Sa = Ad + PI / 2
   Th = Sa / 3
   Ao = S2 - Sa
   For J = 1 To 5
      aa = Ao + J * Th: La = Atn(Tan(aa) / E1)
      If aa <= PI / 2 Then La = La: GoTo L2270
      Dl = MinusMulti360(Mc + 180 - Ac): Vx = Ac
      If aa > PI / 2 And aa <= 3 * PI / 2 Then La = La + PI: GoTo L2270
      La = La + 2 * PI
L2270:
      La = La - Int(La / (2 * PI)) * 2 * PI: Lb = Sin(La) * E2: Ld = BoogSinus(Lb)
      Ha = Tan(B5 - Ld) * Cos(aa) * E2
      hhb = Atn(Ha) + La + PI / 2
      Routput.House(J + 1).pLength = MinusMulti360(Rad * hhb)
   Next J
   For J = 7 To 12: Routput.House(J).pLength = MinusMulti360(Routput.House(J - 6).pLength + 180): Next J
   ReDim Hb(12): For J = 1 To 12: Hb(J) = Routput.House((J + 8) Mod 12 + 1).pLength: Next J
   For J = 1 To 12: Routput.House(J).pLength = Hb(J): Next J
   For J = 1 To 12
      Z = Routput.House(J).pLength + 180: If Z > 360 Then Z = Z - 360
      GoSub AbsCoordToZodiakCoord: Routput.House(J).Degree = Zz$: Routput.House(J).Zod = Qq
   Next J
   Exit Sub

' subroutines
' abs. coord. ---> zodiac coord.
AbsCoordToZodiakCoord:
   Z3 = Int(Z)
   Qq = Int(Z3 / 30) + 1
   Z1 = Int((Z3 / 30 - Int(Z3 / 30)) * 30)
   Z2 = Int(Int(((Z - Z3) * 60) * 10 + 0.5) / 10)
   Z1s$ = Str$(Z1): If Len(Z1s$) = 2 Then Z1s$ = " " + Z1s$
   Z2s$ = Str$(Z2): If Len(Z2s$) = 2 Then Z2s$ = " " + Z2s$
   Zz$ = Right$(Z1s$, 2) + "°" + Right$(Z2s$, 2)
   Return
' cartesis coord. ---> poolcoord.
CartCoordToPoolCoord:
   r = Sqr(X * X + Y * Y + Z * Z)
   Rx = Sqr(X * X + Y * Y): B9 = Atn(Z / Rx)
   If X = 0 And Y = 0 Then L9 = 0: Return
   P9 = 2 * Atn(Y / (Abs(X) + Rx))
   If X < 0 Then L9 = PI - P9: Return
   L9 = P9: If Y < 0 Then L9 = L9 + 2 * PI
   Return
' check retrograde
CheckRetrograde:
   If I = 1 Then
      Retro$ = " "
      Vo = G0 * Sqr(2 / RO - 1)
      po = BoogCosinus(G0 * Exc * Sin(Ea) / (Vo * RO))
      V3 = (Rad * Vo) / RO
      Return
      Else
      V1 = G0 * Sqr(2 / R8 - 1 / Ax)
      P1 = BoogCosinus(G0 * Sqr(Ax) * Exc * Sin(Ea) / R8 / V1)
      V2 = V1 * Sin(P1 + L8 - L9) - Vo * Sin(po + Lo - L9)
      V3 = (Rad * V2) / R9
      Retro$ = " "
      If Abs(V3) * R9 < 0.01 Then Retro$ = "D": Return
      If V3 < 0 Then Retro$ = "R"
      Return
      End If
End Sub


Public Sub DrawRadix(pic As Control, _
                      RI As RADIXPARAM, _
                      RO As CALC_OUTPUT)

   Dim A As Integer, I As Integer, J1 As Integer, J2 As Integer
   Dim X, Y, X1, X2, Y1, Y2
   Dim Xcentr, Ycentr, C1, C2, C3, C6
   Dim straal, CP, HLtot
   Dim hoek As Double
   Dim orb As Integer
   Dim hMemDC As Long, hdc As Long
   Dim hOldBitmap As Long
   Dim hPen As Long, hOldPen As Long
   Dim hBrush As Long, hOldBrush As Long
   Dim hFnt As Long, hOldFnt As Long
   Dim pt As POINTAPI, sze As Size
   Dim Ascendant
   
   hdc = pic.hdc
   If RI.TurnWhat = 0 Then
       Ascendant = 360 - RO.Ps(12).pLength
       Else
       Ascendant = 180
       End If
   Xcentr = RI.Width / 2
   Ycentr = RI.Height / 2
   C1 = IIf(Xcentr < Ycentr, Xcentr, Ycentr) - 5
   C2 = C1 * 0.78: C3 = C1 * 0.75: C6 = C1 * 0.25
   
   hMemDC = CreateCompatibleDC(hdc)
   hOldBitmap = SelectObject(hMemDC, CreateCompatibleBitmap(hdc, RI.Width, RI.Height))
   
   hFnt = CreateFontIndirect(RI.HouseNoLgFnt)
   hOldFnt = SelectObject(hMemDC, hFnt)
   SetBkMode hMemDC, 1
   SetTextColor hMemDC, QBColor(0)
   'Cls
   hBrush = CreateSolidBrush(QBColor(7))
   hOldBrush = SelectObject(hMemDC, hBrush)
   BitBlt hMemDC, 0, 0, RI.Width, RI.Height, frmAMhDC, pic.Left, pic.Top, SRCCOPY
   'cirkels
   Arc hMemDC, Xcentr - C1, Ycentr - C1, Xcentr + C1, Ycentr + C1, 0, 0, 0, 0
   Arc hMemDC, Xcentr - C2, Ycentr - C2, Xcentr + C2, Ycentr + C2, 0, 0, 0, 0
   Arc hMemDC, Xcentr - C3, Ycentr - C3, Xcentr + C3, Ycentr + C3, 0, 0, 0, 0
   Arc hMemDC, Xcentr - C6, Ycentr - C6, Xcentr + C6, Ycentr + C6, 0, 0, 0, 0
   'streepjes (NL)
   For A = 1 To 360 Step 30
      X1 = Xcentr + C2 * Cos(Deg * (A + Ascendant))
      X2 = Xcentr + C1 * Cos(Deg * (A + Ascendant))
      Y1 = Ycentr + C2 * Sin(Deg * (A + Ascendant - 180))
      Y2 = Ycentr + C1 * Sin(Deg * (A + Ascendant - 180))
      MoveToEx hMemDC, X1, Y1, pt
      LineTo hMemDC, X2, Y2
   Next A
   For A = 1 To 360 Step 10
      X1 = Xcentr + C3 * Cos(Deg * (A + Ascendant))
      X2 = Xcentr + C2 * Cos(Deg * (A + Ascendant))
      Y1 = Ycentr + C3 * Sin(Deg * (A + Ascendant - 180))
      Y2 = Ycentr + C2 * Sin(Deg * (A + Ascendant - 180))
      MoveToEx hMemDC, X1, Y1, pt
      LineTo hMemDC, X2, Y2
   Next A
   
   'signs zodiac
   For A = 0 To 11
      straal = C2 + (C1 - C2) / 2
      X = Xcentr + straal * Cos(Deg * (30 * A + 195 + Ascendant))
      Y = Ycentr + straal * Sin(Deg * (30 * A + 15 + Ascendant))
      SymbolMasked hMemDC, SymbolhDC, SymbMaskhDC, A, X - SymbSize / 2, Y - SymbSize / 2
   Next A
   
   'planets
   If RI.ShowAscMc = True Then HLtot = 13 Else HLtot = 11
   For A = 1 To HLtot
      hoek = Abs(RO.Ps(A).pLength) + 181 + Ascendant
      CP = C3 - SymbSize * 0.8
      Select Case RI.RadiiType
      Case 0
         For I = 1 To A - 1
            If Abs(RO.Ps(I).pLength - RO.Ps(A).pLength) < 10 Then CP = CP * 0.79
         Next I
      Case 1
         CP = CP - (C3 - C6) / 14 * A
      Case 2
      End Select
      If A > 11 Then CP = C3 + SymbSize / 2
      X = Xcentr - SymbSize / 2 + CP * Cos(Deg * (hoek))
      Y = Ycentr - SymbSize / 2 + CP * Sin(Deg * (hoek - 180))
      SymbolMasked hMemDC, SymbolhDC, SymbMaskhDC, 11 + A, X, Y
      If RO.Ps(A).Retro <> " " Then
          Y = Y + SymbSize / 3
          X = X - 5
          TextOut hMemDC, X, Y, RO.Ps(A).Retro, Len(RO.Ps(A).Retro)
          End If
      X1 = Xcentr + C3 * Cos(Deg * (hoek))
      X2 = Xcentr + (C3 - 6) * Cos(Deg * (hoek))
      Y1 = Ycentr + C3 * Sin(Deg * (hoek - 180))
      Y2 = Ycentr + (C3 - 6) * Sin(Deg * (hoek - 180))
      MoveToEx hMemDC, X1, Y1, pt
      LineTo hMemDC, X2, Y2
   Next A
   
  'houses
   Dim dw As Long, J As Long, Nr As String
   If RI.ShowHouses = True Then
      straal = C2 + 7
      If RI.ShowAscMc = False Then
         X = Xcentr - 2 + straal * Cos(Deg * (RO.House(1).pLength + Ascendant))
         Y = Ycentr - 5 + straal * Sin(Deg * (RO.House(1).pLength + Ascendant - 180))
         TextOut hMemDC, X, Y, "A", 1
         X = Xcentr - 5 + straal * Cos(Deg * (RO.House(4).pLength + 180 + Ascendant))
         Y = Ycentr - 5 + straal * Sin(Deg * (RO.House(4).pLength + Ascendant))
         TextOut hMemDC, X, Y, "M", 1
         X = Xcentr - 2 + straal * Cos(Deg * (RO.House(4).pLength + Ascendant))
         Y = Ycentr - 8 + straal * Sin(Deg * (RO.House(4).pLength + Ascendant - 180))
         TextOut hMemDC, X, Y, "I", 1
         X = Xcentr - 5 + straal * Cos(Deg * (RO.House(1).pLength + 180 + Ascendant))
         Y = Ycentr - 8 + straal * Sin(Deg * (RO.House(1).pLength + Ascendant))
         TextOut hMemDC, X, Y, "D", 1
         End If
      For J = 1 To 12
         If J = 1 Or J = 4 Or J = 7 Or J = 10 Then dw = 2 Else dw = 1
         hPen = CreatePen(0, dw, QBColor(0))
         hOldPen = SelectObject(hMemDC, hPen)
         hoek = (RO.House(J).pLength) + Ascendant + 1
         X1 = Xcentr + C6 * Cos(Deg * (hoek))
         X2 = Xcentr + C3 * Cos(Deg * (hoek))
         Y1 = Ycentr + C6 * Sin(Deg * (hoek - 180))
         Y2 = Ycentr + C3 * Sin(Deg * (hoek - 180))
         MoveToEx hMemDC, X1, Y1, pt
         LineTo hMemDC, X2, Y2
         hoek = hoek + 10
         X = Xcentr + (C6 + 7) * Cos(Deg * (hoek))
         Y = Ycentr + (C6 + 7) * Sin(Deg * (hoek - 180))
         Nr = Format(J)
         GetTextExtentPoint32 hMemDC, Nr, Len(Nr), sze
         TextOut hMemDC, X - sze.cx / 2, Y - sze.cy / 2, Nr, Len(Nr)
         hOldPen = SelectObject(hMemDC, hOldPen)
         DeleteObject hPen
      Next J
      End If
   
   'angle-houses (1-4-7-10)
   If RI.ShowAscMc = True Then
      hoek = (RO.Ps(12).pLength) + Ascendant + 1
      X1 = Xcentr + C3 * Cos(Deg * (hoek))
      X2 = Xcentr + C3 * Cos(Deg * (hoek - 180))
      Y1 = Ycentr + C3 * Sin(Deg * (hoek - 180))
      Y2 = Ycentr + C3 * Sin(Deg * (hoek))
      MoveToEx hMemDC, X1, Y1, pt
      LineTo hMemDC, X2, Y2
      hoek = (RO.Ps(13).pLength) + 181 + Ascendant
      X1 = Xcentr + C3 * Cos(Deg * (hoek))
      X2 = Xcentr + C3 * Cos(Deg * (hoek - 180))
      Y1 = Ycentr + C3 * Sin(Deg * (hoek - 180))
      Y2 = Ycentr + C3 * Sin(Deg * (hoek))
      MoveToEx hMemDC, X1, Y1, pt
      LineTo hMemDC, X2, Y2
      End If

'aspect lines
   If RI.ShowAspects = False Then GoTo TekenEinde:
   For J1 = 1 To 10: For J2 = J1 + 1 To 11
      X = MinusMulti360(Abs(RO.Ps(J2).pLength) - Abs(RO.Ps(J1).pLength))
      orb = 2
      If J1 = 1 Then orb = 5
      If J1 = 2 Or J2 = 2 Then orb = 4
      If J1 = 11 Then orb = 3
      For I = 1 To 6: A = X - asp(I)
         If Abs(A) <= orb Then GoSub L4200
      Next I
      For I = 1 To 6: A = X - Aspc(I)
         If Abs(A) <= orb Then GoSub L4200
      Next I
   Next J2: Next J1

TekenEinde:
   hOldBrush = SelectObject(hMemDC, hOldBrush)
   DeleteObject hBrush
   SelectObject hMemDC, hOldFnt
   DeleteObject hFnt
   
   BitBlt hdc, RI.Left, RI.Top, RI.Width, RI.Height, hMemDC, 0, 0, SRCCOPY
   DeleteObject (SelectObject(hMemDC, hOldBitmap))
   DeleteDC hMemDC

Exit Sub

L4200:
  hoek = Abs(RO.Ps(J1).pLength) + 181 + Ascendant
  X1 = Xcentr + C6 * Cos(Deg * (hoek))
  Y1 = Ycentr + C6 * Sin(Deg * (hoek - 180))
  hoek = Abs(RO.Ps(J2).pLength) + 181 + Ascendant
  X2 = Xcentr + C6 * Cos(Deg * (hoek))
  Y2 = Ycentr + C6 * Sin(Deg * (hoek - 180))
  hPen = CreatePen(0, 0, AspectColor(I))
  hOldPen = SelectObject(hMemDC, hPen)
  MoveToEx hMemDC, X1, Y1, pt
  LineTo hMemDC, X2, Y2
  hOldPen = SelectObject(hMemDC, hOldPen)
  DeleteObject hPen
  Return

End Sub

Sub SymbolMasked(ByVal Phdc As Long, Shdc As Long, Mhdc As Long, ByVal Nr As Integer, ByVal X As Integer, ByVal Y As Integer)
   Dim ox As Long, oy As Long, Size As Long
   ox = Nr Mod 10
   oy = Int(Nr / 10)
   Size = SymbSize
   BitBlt Phdc, X, Y, SymbSize, SymbSize, Shdc, ox * (SymbSize + 1), oy * (SymbSize + 1), SRCINVERT
   BitBlt Phdc, X, Y, SymbSize, SymbSize, Mhdc, ox * (SymbSize + 1), oy * (SymbSize + 1), SRCAND
   BitBlt Phdc, X, Y, SymbSize, SymbSize, Shdc, ox * (SymbSize + 1), oy * (SymbSize + 1), SRCINVERT
End Sub

