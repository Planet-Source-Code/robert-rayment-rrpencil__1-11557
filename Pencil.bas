Attribute VB_Name = "Module1"
'RRPencil  BAS MODULE

Option Base 1
DefInt A-Q  'a-q integers
DefSng R-Z  'rst, uvw, xyz real

'Copy one array to another of same number of bytes
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'(Destination As Any, Source As Any, ByVal Length As Long)
'------------------------------------------------------------------------------

'Used to extract small bitmap from a large one and show shrunken bitmap
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------------------------------

'To get dimensions of GIF's & JPG's
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'Dim bmp As BITMAP
'------------------------------------------------------------------------------

'To move & get cursor postion
Public Declare Sub SetCursorPos Lib "user32" (ByVal IX As Long, ByVal IY As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'------------------------------------------------------------------------------

'Windows API - much faster then VB's PSet and Point
Public Declare Function SetPixelV Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'------------------------------------------------------------------------------
'Draw polyline
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, _
lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Type POINTAPI: kX As Long: kY As Long: End Type
'------------------------------------------------------------------------------
'For shading & fill

Public Declare Function CreatePen Lib "gdi32" _
(ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

'------------------------------------------------------------------------------
'Selecting pens & clearing up API Objects & Device Contexts
Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'API's & Structure for Rotating Text
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
(lpLogFont As LOGFONT) As Long
'Logical Font
Global Const LF_FACESIZE = 32
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
'        lfFaceName(LF_FACESIZE - 1) As Byte
        lfFaceName As String * LF_FACESIZE
End Type
Global RotateFont As LOGFONT
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------

'For finding item in ListBox
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
       Integer, ByVal lParam As Any) As Long
Global Const LB_FINDSTRING = &H18F
'------------------------------------------------------------------------------

Global Instruction$()   'TOOL instructions
Global CurrentDirec$
Global LoadFileSpec$

Global TOOL, prevTOOL   'DRAWING TOOLS 0-20
Global LCount  'Click count, allows drawing with buttons up
Global zrad, zratio, zradx, zrady, zrad2, zratio2 'For cirlipses
Global xsto(), ysto(), PCount, zpspac  'For multi-segments & double line spacing
Global mx, my        'For AirBrush Bricks & Tiles spacing
Global xprev, yprev  'Saved X,Y

'ACTIVE RECTANGLE
Global arleft, artop, arright, arbottom, arwidth, arheight 'Active rect coords
Global karleft, kartop, karright, karbottom, karwidth, karheight 'Kept Active rect coords

'ZOOM PARAMETERS
Global xzoom, yzoom, GridStep, ZoomSize(), ZoomSpan
Global pzoomtool, zoomdrawing, zoomend
Global xs, ys, xp, yp   'Draw item & zoomRect coords
Global ZoomInstr$

'MCRBox dragging
Global zDragX, zDragY

'COLORS
Global LeftCul&, RightCul&    'Left Button, Right button colors
Global DrawCul&, RubberCul&   'Left Button, Right button colors
Global cn, rubcn              'Left/Right color numbers
Global RubRectCul&            'Rubber erase color
Global zoomRectCul&           'Zoom Active rectangle color
Global PenCul&      'Draw color for shapes colored in after completion

Global ScrollStep   'Amount to scroll picture

'SWITCHES
Global ActiveRectExists    'True or False
Global ShowPlusLines, ShowXLines  'True/False for Hair Lines
Global DrawingMode   'True if drawing in process
Global ZoomMode      'True if Zoom ON
Global TextSW     'When True picCanvas.MouseDown ignored
Global MCRSW      'When True picCanvas.MouseDown ignored
Global ResizeSW   'When True picCanvas.MouseDown ignored
Global PerspectiveSW    'When True picCanvas.MouseDown ignored
Global RotateSW   'When True picCanvas.MouseDown ignored
Global AddBMPSW   'When True picCanvas.MouseDown ignored
Global RedrawPolCoTu
Global prevIndex

'PERSPECTIVE HAIR LINE PARAMETERS
Global ShowPerspecLines, SetPerspecPts, YPr, XPr1, XPr2, XPr3, NumPPts
Global PerSpecInstr$
'21 TOOL NUMBERS
Global Const Brush = 0
Global Const AirBrush = 1
Global Const ALine = 2
Global Const APolyline = 3
Global Const Spline = 4
Global Const Rectangle = 5
Global Const Cirlipse = 6
Global Const Cone = 7
Global Const Tube = 8
Global Const Arch = 9
Global Const TPiece = 10
Global Const AText = 11
Global Const Fill = 12
Global Const Rubber = 13
Global Const ActiveRect = 14
Global Const Smudge = 15
Global Const MCR = 16
Global Const Resize = 17
Global Const PerspecShear = 18
Global Const Rotate = 19
Global Const Tile = 20

'9 SUB_TOOL TYPES
Global BrushType
Global AirBrushType
Global LineType
Global RectangleType
Global TPieceType
Global FillType
Global RubberType
Global SmudgeType
Global MCRType
Global ResizeType
Global RotateType

Global HelpExist

'FILL TYPES 7-15
Global BitArray() As Integer

Global Const pi# = 3.1415927
Global Const r2d# = 180 / pi#


Public Function ActionInProgress()

If Form1.picINFO.Visible Then
Form1.picINFO.Cls
a$ = " DMode=" + Str$(DrawingMode) + "  ZMode=" + Str$(ZoomMode)
a$ = a$ + "  TextSW=" + Str$(TextSW) + "  MCRSW=" + Str$(MCRSW)
a$ = a$ + "  ResSW=" + Str$(ResizeSW) + "  PerSpecSW=" + Str$(PerspectiveSW) + "  RotSW=" + Str$(RotateSW)
a$ = a$ + "  +BMPSW=" + Str$(AddBMPSW) + "  SetPP=" + Str$(SetPerspecPts)
Form1.picINFO.Print a$
End If

ActionInProgress = True
If DrawingMode Then Exit Function
If ZoomMode Then Exit Function
If TextSW Then Exit Function
If MCRSW Then Exit Function
If ResizeSW Then Exit Function
If PerspectiveSW Then Exit Function
If RotateSW Then Exit Function
If AddBMPSW Then Exit Function
If SetPerspecPts Then Exit Function
If Form1.picMCRBox.Visible = True Then Exit Function
ActionInProgress = False
End Function

Public Sub FixRectCoords(ByVal x1, ByVal x2, ByVal y1, ByVal y2)
If x2 < x1 Then
   TX = x1: x1 = x2: x2 = TX
   arleft = x1: arright = x2
   karleft = x1: karright = x2
End If
If y2 < y1 Then
   TY = y1: y1 = y2: y2 = TY
   artop = y1: arbottom = y2
   kartop = y1: karbottom = y2
End If
End Sub

Public Sub ShowXY(ByVal IX, ByVal IY)
Form1.picXY.Cls
Form1.picXY.Print " X=" & Str$(IX)
Form1.picXY.Print " Y=" & Str$(IY)
End Sub

Public Sub SetInstructions()
ReDim Instruction$(0 To 20)
Instruction$(0) = Space$(8) & "BRUSH: LB(Left Button, Left Color)/RB(Right Button, Right Color) - Click - Hold && Move"
Instruction$(1) = Space$(8) & "AIRBRUSH: LB/RB - Click or Hold && Move"
Instruction$(2) = Space$(8) & "LINE: LC(Left Click, Left Color)/RC(Right Click, Right Color) - Move - LC - Move to Locate - Click to Fix"
Instruction$(3) = Space$(8) & "POLYLINE: LC/RC - Move - LC - Move - etc - - RC to End then Move to Locate - Click to Fix"
Instruction$(4) = Space$(8) & "SPLINE: LC/RC - Move - LC - Move - etc - - RC to End then Move to Locate - Click to Fix"
Instruction$(5) = Space$(8) & "RECTANGLE: LC/RC - Move - LC - Move to Locate - Click to Fix"
Instruction$(6) = Space$(8) & "CIRCLIPSE: LC/RC - Move - LC - Move to Locate - Click to Fix"
Instruction$(7) = Space$(8) & "CONE: LC/RC - Move (Base) - LC - Move (Axis) - Click to End then Move to Locate - Click to Fix"
Instruction$(8) = Space$(8) & "TUBE: LC/RC - Move (Base) - LC - Move (Axis) - Click to End then Move to Locate - Click to Fix"
Instruction$(9) = Space$(8) & "ARCH: LC/RC - Move - LC - Move to Locate - Click to Fix"
Instruction$(10) = Space$(8) & "T-PIECE: LC/RC - Move - LC - Move to Locate - Click to Fix (This also takes the Line Type apart from shading)"
Instruction$(11) = Space$(8) & "TEXT: Input Text, LC(Left Color)/RC(Right Color) on Picture, LC - Hold - Move to Locate, RC to Fix"
Instruction$(12) = Space$(8) & "FILL: LC(Left Color)/RC(Right Color) inside closed shape on Picture"
Instruction$(13) = Space$(8) & "RUBBER (Takes Right Color): LC to draw rubber - Move to rub out - LC to End or Click INSIDE Active Rectangle"
Instruction$(14) = Space$(8) & "ACTIVE RECTANGLE: LC - Move - Click to End then Move to Locate - Click to Fix. LC to Clear, LC to Redo"
Instruction$(15) = Space$(8) & "SMUDGE: LC Repeatedly on Picture or INSIDE Active Rectangle"
Instruction$(16) = Space$(8) & "MOVE COPY REFLECT: LC INSIDE Active Rectangle, LC - Hold - Move - RC to Fix"
Instruction$(17) = Space$(8) & "RESIZE (+/-N % INSIDE Active Rectangle): LC INSIDE Active Rectangle, LC - Hold - Move - RC to Fix"
Instruction$(18) = Space$(8) & "PERSPECTIVE SHEAR: LC Repeatedly OUTSIDE Active Rectangle Until Satisfied - RC to Fix"
Instruction$(19) = Space$(8) & "ROTATE: LC INSIDE Active Rectangle to Rotate +/-N degrees, LC - Hold - Move - RC to Fix"
Instruction$(20) = Space$(8) & "TILE: LC INSIDE Active Rectangle to Tile across WHOLE Picture"

PerSpecInstr$ = Space$(8) & "SET HORIZON && PERSPECTIVE POINTS:  LC on Picture for each perspective point (max 3) - RC to Stop"
ZoomInstr$ = Space$(8) & "ZOOM:  LC ON PICTURE, ADJUST MAG WITH UPDOWN BUTTONS, THEN LC IN ZOOM BOX"

End Sub
Public Sub SetToolTips()
For I = 0 To 20
   Form1.chkTOOLS(I).Left = 92
   Form1.chkTOOLS(I).Top = 8 + 20 * I
   Form1.chkTOOLS(I).Width = 21
Next I
Form1.chkTOOLS(0).ToolTipText = "Brush"
Form1.chkTOOLS(1).ToolTipText = "AirBrush"
Form1.chkTOOLS(2).ToolTipText = "Line"
Form1.chkTOOLS(3).ToolTipText = "PolyLine"
Form1.chkTOOLS(4).ToolTipText = "Spline"
Form1.chkTOOLS(5).ToolTipText = "Rectangle"
Form1.chkTOOLS(6).ToolTipText = "Cirlipse"
Form1.chkTOOLS(7).ToolTipText = "Cone"
Form1.chkTOOLS(8).ToolTipText = "Tube"
Form1.chkTOOLS(9).ToolTipText = "Arch"
Form1.chkTOOLS(10).ToolTipText = "T-Piece"
Form1.chkTOOLS(11).ToolTipText = "Text"
Form1.chkTOOLS(12).ToolTipText = "Fill"
Form1.chkTOOLS(13).ToolTipText = "Rubber"
Form1.chkTOOLS(14).ToolTipText = "Active Rectangle"
Form1.chkTOOLS(15).ToolTipText = "Smudge"
Form1.chkTOOLS(16).ToolTipText = "MoveCopyReflect"
Form1.chkTOOLS(17).ToolTipText = "Resize, percent"
Form1.chkTOOLS(18).ToolTipText = "Perspective Shear"
Form1.chkTOOLS(19).ToolTipText = "Rotate, degrees"
Form1.chkTOOLS(20).ToolTipText = "Tile"
Form1.chkPerspec.ToolTipText = "Set Perspec Pts"
End Sub

Public Sub GetGreyCN(cul&, culb As Byte)
'Input: cul& long color
'Output: culb colour number 0-255
Dim redc As Byte, greenc As Byte, bluec As Byte
redc = cul& And &HFF&
greenc = (cul& And &HFF00&) / &H100&
bluec = (cul& And &HFF0000) / &H10000
culnum = (1 * redc + greenc + bluec) / 3
If culnum < 0 Then culnum = 0
If culnum > 255 Then culnum = 255
culb = culnum
End Sub

Public Function ExtractFileName$(FSpec$)
ExtractFileName$ = " "
'In:  FSpec$ = Full FileSpec
'Out: FileName =ExtractFileName$
If FSpec$ = "" Then Exit Function
'Find pbs on last backslash \
p = 0: pbs = 0
Do: p = InStr(p + 1, FSpec$, "\")
    If p <> 0 Then pbs = p Else Exit Do
Loop
If pbs > 0 Then
   ExtractFileName$ = Mid$(FSpec$, pbs + 1)
End If
End Function
Public Function ExtractPath$(FSpec$)
ExtractPath$ = ""
'In:  FSpec$ = Full FileSpec
'Out: Path =ExtractPath$
If FSpec$ = "" Then Exit Function
'Find pbs on last backslash \
p = 0: pbs = 0
Do: p = InStr(p + 1, FSpec$, "\")
    If p <> 0 Then pbs = p Else Exit Do
Loop
If pbs > 0 Then
   ExtractPath$ = Left$(FSpec$, pbs - 1) 'ie including last \
End If
End Function
Public Sub FixFileExtension(SaveFileSpec$, a$)
'eg a$="bmp" or "jpg"
B$ = "." + a$
pdot = InStr(1, SaveFileSpec$, ".")
If pdot = 0 Then
   SaveFileSpec$ = SaveFileSpec$ + B$
   Else
      Ext$ = LCase$(Mid$(SaveFileSpec$, pdot))
      If Ext$ <> B$ Then
         SaveFileSpec$ = Left$(SaveFileSpec$, pdot - 1) + B$
      End If
   End If
End Sub

Public Function zAtn2(Y, X)
'0° to right, 0 to -pi#(-180°) anticlockwise, 0 to +pi#(+180°) clockwise
If X = 0 Then
    If Abs(Y) > Abs(X) Then   'Must be an overflow
        If Y > 0 Then zAtn2 = pi# / 2 Else zAtn2 = -pi# / 2
    Else
        zAtn2 = 0   'Must be an underflow
    End If
Else
    zAtn2 = Atn(Y / X)
    If (X < 0) Then
        If (Y < 0) Then zAtn2 = zAtn2 - pi# Else zAtn2 = zAtn2 + pi#
    End If
End If
End Function

Public Sub Findxeye(zd, x1, y1, x2, y2, xdd, ydd, xa, ya, xb, yb)
'In:  Coords of line xy1->xy2, paraspacing zd
'Out: Coords of parallel line to right xya->xyb
'     Increment xdd=x1-x0, ydd=y0-y1, needed for Findxiyi when used
'Find angle to horizontal (downwards)
'and hence increments onto parallel line
xd12 = x2 - x1: yd12 = y2 - y1
If xd12 = 0 Then
   ydd = 0: xdd = Sgn(yd12) * zd
ElseIf yd12 = 0 Then
   xdd = 0: ydd = Sgn(xd12) * zd
Else
   zang = Atn(yd12 / xd12)
   xdd = Sgn(xd12) * zd * Sin(zang)
   ydd = Sgn(xd12) * zd * Cos(zang)
End If
xa = x1 - xdd: ya = y1 + ydd
xb = x2 - xdd: yb = y2 + ydd
End Sub

Public Sub Findxiyi(x1, y1, x2, y2, x3, y3, xdd, ydd, xa, ya, xb, yb, xi, yi)
'In:  Coords of 2 intersecting lines xya-(xyi)-xyb zd from xy1-xy2-xy3
'      From Findxeye:-
'      xa,ya start coords of the 1st parallel line zd from xy1-xy2
'      xb,yb end coords of the 2nd parallel line zd away from xy2-xy3
'      xdd,ydd incr to x2,y2 to give xi,yi IF the lines have the same slope
'Out: xi,yi intersection of the 2 parallel lines  xya-xyi-xyb
'Find slopes
xd12 = x2 - x1: yd12 = y2 - y1
xd23 = x3 - x2: yd23 = y3 - y2
xd12 = x2 - x1: yd12 = y2 - y1
xd13 = x3 - x1: yd13 = y3 - y1

'Find slopes
If xd12 <> 0 Then
   zm12 = yd12 / xd12
   zc1 = y1 - zm12 * x1
Else  'Vertical
   zm12 = Sgn(yd12) * 10000
   zc1 = Sgn(zm12) * 10000
End If

If xd23 <> 0 Then
   zm23 = yd23 / xd23
   zc2 = y2 - zm23 * x2
Else   'Vertical
   zm23 = Sgn(yd23) * 10000
   zc2 = Sgn(zm23) * 10000
End If
If xd13 <> 0 Then
   zm13 = yd13 / xd13
   zc3 = y3 - zm13 * x3
Else   'Vertical
   zm13 = Sgn(yd13) * 10000
   zc3 = Sgn(zm13) * 10000
End If

'Find intersection
If zm12 <> zm23 Then
      
      If Abs(zm12) > 9000 Then
         xi = xa
         yi = zm23 * xi - xb * zm23 + yb
      ElseIf Abs(zm23) > 9000 Then
         xi = xb
         yi = zm12 * xi - xa * zm12 + ya
      Else
         xi = xa * zm12 - ya - xb * zm23 + yb
         xi = xi / (zm12 - zm23)
         yi = zm23 * xi - xb * zm23 + yb
      End If

Else
   xi = x2 - xdd
   yi = y2 + ydd
End If
'If xi,yi lies outside triangle x1y1,x2y2,x3y3 then xiyi=(x1y1+x3y3)/2
'Restrict magnitude (NB could deal with shallow angles better)

'sin12 = Sgn(zd12 * xi + yi + zc1)
'sin23 = Sgn(zd23 * xi + yi + zc2)
'sin13 = Sgn(zd13 * xi + yi + zc3)
'If sin12 = sin23 And sin12 = sin13 Then
'Else
'   xi = (x1 + x3) / 2: yi = (y1 + y3) / 2
'End If

If xi > 32500 Then xi = 32500
If xi < -32500 Then xi = -32500
If yi > 32500 Then yi = 32500
If yi > 32500 Then yi = 32500
End Sub

Public Sub EvalTangents(xc, yc, zradius, xp, yp, x1, y1, x2, y2)
'IN: Circle centre xc,yc radius zradius, point outside xp,yp
'OUT: Tangents xp,yp to x1,y1 & x2,y2
zL = Sqr((xp - xc) * (xp - xc) + (yp - yc) * (yp - yc))
If zL > zradius Then
   zd = Sqr(zL * zL - zradius * zradius)
   x1 = xp - zd * ((xp - xc) * zd - (yp - yc) * zradius) / zL ^ 2
   y1 = yp - zd * ((yp - yc) * zd + (xp - xc) * zradius) / zL ^ 2
   x2 = xp - zd * ((xp - xc) * zd + (yp - yc) * zradius) / zL ^ 2
   y2 = yp + zd * (-(yp - yc) * zd + (xp - xc) * zradius) / zL ^ 2
Else  'xp,yp inside circle
   x1 = xp: y1 = yp
   x2 = xc: y2 = yc
End If
End Sub

Public Sub EvalDiameters(xs, ys, zrad, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4)
'IN: Circ1: xs,ys,zrad Circ2: xp,yp,zrad
'OUT: Circ1 diam cords: x1,y1-x2,y2  Circ2 diam coords x3,y3-x4,y4
ztheta = zAtn2(yp - ys, xp - xs)
x1 = xs + zrad * Sin(ztheta)
y1 = ys - zrad * Cos(ztheta)
x2 = xs - zrad * Sin(ztheta)
y2 = ys + zrad * Cos(ztheta)

x3 = xp + zrad * Sin(ztheta) + 1.5 * Cos(ztheta) 'more accurate diameter coords
y3 = yp - zrad * Cos(ztheta) + 1.5 * Sin(ztheta) 'for this point
x4 = xp - zrad * Sin(ztheta)
y4 = yp + zrad * Cos(ztheta)

End Sub
Public Sub EvalZradZratio(X, Y)
'Global zrad, zratio, zradx, zrady, zrad2, zratio2 'For cirlipses
'From Start & Draw Cirlipse
zradx = Abs(X - xs)
zrady = Abs(Y - ys)
If zradx = 0 Then
   zrad = zrady
   zratio = 10
ElseIf zradx >= zrady Then
   zrad = zradx
   zratio = zrady / zradx
Else  'zradx<zrady
   zrad = zrady
   zratio = zrady / zradx
End If

End Sub
Public Sub EvalZrad2Zratio2()
'Global zrad, zratio, zradx, zrady, zrad2, zratio2 'For cirlipses
'From Start & Draw Cirlipse
zrad2 = zrad - zpspac
If zrad2 < 0 Then zrad2 = 0
If zradx = zpspac Then
   zratio2 = 10   'Produces a narrow and wide ellipse
Else
   zratio2 = (zrady - zpspac) / (zradx - zpspac)
End If
If zratio2 < 0 Then zratio2 = 1

End Sub
      
Public Function InRectangle(ByVal xp, ByVal yp, ByVal x1, ByVal y1, ByVal x2, ByVal y2)
InRectangle = False
If xp <= x1 Or xp >= x2 Then Exit Function
If yp <= y1 Or yp >= y2 Then Exit Function
InRectangle = True
End Function


Public Sub ClearAllSubToolbars()
Form1.frmBrush.Visible = False
Form1.frmAirBrush.Visible = False
Form1.frmLine.Visible = False
Form1.frmRect.Visible = False
Form1.frmTPiece.Visible = False
Form1.frmFill.Visible = False
Form1.frmRubber.Visible = False
Form1.frmSmudge.Visible = False
Form1.frmMCR.Visible = False
Form1.frmResize.Visible = False
Form1.frmRotate.Visible = False
End Sub

Public Sub SetZOOMParams()
'=========================================
'ZOOM BOX PARAMETERS
'Global ZoomMode, xzoom, yzoom, GridStep, ZoomSize(), ZoomSpan
'GridStep = 240 / (2 * ZoomSize(ZoomSpan))
'ZoomSpan    GridStep  PixelBox
' (ZoomSize) Mag
'-    2      60        4 X
'0    3      40        6 *
'1    4      30        8 *
'2    5      24       10 *
'3    6      20       12 *
'4    8      15       16 *
'5    10     12       20 *
'6    12     10       24 *
'7    15      8       30 *
'8    20      6       40 *
'9    24      5       48 *
'10   30      4       60 *
'11   40      3       80
'12   60      2      120 *
ReDim ZoomSize(0 To 12)
ZoomSize(0) = 3
ZoomSize(1) = 4
ZoomSize(2) = 5
ZoomSize(3) = 6
ZoomSize(4) = 8
ZoomSize(5) = 10
ZoomSize(6) = 12
ZoomSize(7) = 15
ZoomSize(8) = 20
ZoomSize(9) = 24
ZoomSize(10) = 30
ZoomSize(11) = 40
ZoomSize(12) = 60
'Start ZoomSpan
ZoomSpan = 9
GridStep = 240 / (2 * ZoomSize(ZoomSpan))
Form1.picZOOMBox.Visible = False
'=========================================
End Sub

Public Sub DrawZoomGrid()
'GridStep = 240 / (2 * ZoomSize(ZoomSpan))
For kX = 0 To 240 Step GridStep
   Form1.picZOOMBox.Line (kX, 0)-(kX, 480), QBColor(1) 'QBColor(12)
Next kX
For kY = 0 To 480 Step GridStep
   Form1.picZOOMBox.Line (0, kY)-(240, kY), QBColor(1) 'QBColor(12)
Next kY
End Sub
Public Sub PositionZoomBox(X)
Form1.picZOOMBox.DrawMode = 13
If X > 300 Then ZoomLeft = 4 Else ZoomLeft = Form1.picCanvas.Width - 246 '348
Form1.picZOOMBox.Top = 35 - Form1.VScroll1.Value
Form1.picZOOMBox.Left = ZoomLeft
Form1.picZOOMBox.Width = 244
Form1.picZOOMBox.Height = 484
Form1.picZOOMBox.AutoRedraw = True
Form1.picZOOMBox.BackColor = QBColor(15)
Form1.picZOOMBox.Cls
End Sub

Public Sub SETUPSCREEN()
With Form1.picCanvas
   .ScaleMode = vbPixels
   .BackColor = RGB(255, 255, 255)
   .BorderStyle = vbBSNone
   .Top = 4
   .Left = 148
   .Width = 600 '604
   .Height = 520 '524
   .ZOrder 0
   .AutoRedraw = True
   .FontName = "Arial"
   .FontSize = 8
   .ForeColor = 0
   .FontBold = False
   .FontItalic = False
   .FontUnderline = False
End With

With Form1.picCanvasStore
   .ScaleMode = vbPixels
   .BackColor = RGB(255, 255, 255)
   .BorderStyle = vbBSNone
   .Top = 4
   .Left = 148
   .Width = Form1.picCanvas.Width
   .Height = Form1.picCanvas.Height
   .AutoRedraw = True
   .ZOrder 1
   .Visible = False
End With

With Form1.picMCRBox
   .ScaleMode = vbPixels
   .BackColor = RGB(255, 255, 255)
   .BorderStyle = vbFixedSingle
   .AutoRedraw = True
   .ZOrder 1
   .Visible = False
End With
With Form1.picZOOMBox
   .ScaleMode = vbPixels
   .BackColor = RGB(255, 255, 255)
   .BorderStyle = vbFixedSingle
   .AutoRedraw = True
   .ZOrder 1
   .Visible = False
End With
With Form1.picPAL
   .ScaleMode = vbPixels
   .BackColor = RGB(255, 255, 255)
   .AutoRedraw = True
   .BorderStyle = vbBSNone
   .Top = 4 + 15
   .Left = 756
   .Width = 36
   .Height = 256
End With
'Fill grey palette strip
For ny = 0 To 255
   cul& = RGB(255 - ny, 255 - ny, 255 - ny)
   Form1.picPAL.Line (0, ny)-Step(Form1.picPAL.Width, 0), cul&
Next ny

'Position Color Labels & Set Initial Colors
cn = 255
LeftCul& = RGB(0, 0, 0)
DrawCul& = LeftCul&
rubcn = 0
RightCul& = RGB(255, 255, 255)
RubberCul& = RightCul&
'Set colour for intermediate drawing
PenCul& = Form1.picCanvas.BackColor Xor QBColor(0)

'Position color boxes
L = 760: H = 21: W = 33
Form1.LabCN.Caption = "255"
Form1.LabCUL.Caption = ""
Form1.LabCUL.Left = L
Form1.LabCUL.Width = W: Form1.LabCUL.Height = H
Form1.LabCUL.BackColor = DrawCul&

Form1.LabLEFTCN.Caption = "255"
Form1.LabLEFTCUL.Caption = ""
Form1.LabLEFTCUL.Left = L
Form1.LabLEFTCUL.Width = W: Form1.LabLEFTCUL.Height = H
Form1.LabLEFTCUL.BackColor = DrawCul&

Form1.LabRIGHTCN.Caption = "0"
Form1.LabRIGHTCUL.Caption = ""
Form1.LabRIGHTCUL.Left = L
Form1.LabRIGHTCUL.Width = W: Form1.LabRIGHTCUL.Height = H
Form1.LabRIGHTCUL.BackColor = RightCul&

SetCursorPos 110, 60

'TOOLS
SetToolTips

'-------------------------------------------
'SET SUB-TOOL FRAMES & INITIALIZE START SUB-TOOLS
'Start TOOL = Brush
TOOL = 0
Form1.chkTOOLS(TOOL).Value = 1
With Form1.frmBrush
   .Top = 0
   .Left = 116
End With
For I = 0 To 14
   Form1.optBrush(I).Left = 50
Next I
Form1.optBrush(TOOL).Value = True   'Start Brush option = single dot
BrushType = B1Dot

With Form1.frmAirBrush
   .Top = 0
   .Left = 116
End With
For I = 0 To 10
   Form1.optAirBrush(I).Left = 50
Next I
Form1.optAirBrush(0).Value = True   'Start AirBrush option = small spread

With Form1.frmLine
   .Top = 0
   .Left = 116
End With
With Form1.frmRect
   .Top = 0
   .Left = 116
End With
For I = 0 To 12
   Form1.optLine(I).Left = 50
   If I < 10 Then Form1.optRect(I).Left = 50
Next I
Form1.optLine(0).Value = True 'Start Line option = 1 pix line
Form1.optRect(0).Value = True 'Start Rect option = 1 pix line

With Form1.frmTPiece
   .Top = 0
   .Left = 116
End With
For I = 0 To 1
   Form1.optTPiece(I).Left = 50
Next I
Form1.optTPiece(0).Value = True 'Start 3-Legged T-Piece

With Form1.frmFill
   .Top = 0
   .Left = 116
End With
For I = 0 To 15
   Form1.optFill(I).Left = 50
Next I
Form1.optFill(0).Value = True 'Start Solid fill

With Form1.frmRubber
   .Top = 0
   .Left = 116
End With
For I = 0 To 3
   Form1.optRubber(I).Left = 50
Next I
Form1.optRubber(0).Value = True 'Start Small rubber

With Form1.frmSmudge
   .Top = 0
   .Left = 116
End With
For I = 0 To 1
   Form1.optSmudge(I).Left = 50
Next I
Form1.optSmudge(0).Value = True 'Start brush smudge

With Form1.frmMCR
   .Top = 0
   .Left = 116
End With
For I = 0 To 3
   Form1.optMCR(I).Left = 50
Next I
Form1.optMCR(0).Value = True 'Start with Move

With Form1.frmResize
   .Top = 0
   .Left = 116
End With
For I = 0 To 1
   Form1.optResize(I).Left = 50
Next I
Form1.optResize(0).Value = True 'Start with Move

With Form1.frmRotate
   .Top = 0
   .Left = 116
End With
For I = 0 To 1
   Form1.optRotate(I).Left = 50
Next I
Form1.optRotate(0).Value = True 'Start with Move

'-------------------------------------------
'INIT RESIZE & ROTATE INFO
Form1.picRSR.Cls
Form1.picRSR.Print "+/- percent or"
Form1.picRSR.Print "+/- degrees"
Form1.txtRSR.Text = "10"

End Sub
Public Sub SetBitArray()
'For FillTypes 7-15
'Can only use 8x8 bit patterns

'NB 0 draws black, 1 draws white,
'opposite of what's expected!

   ReDim BitArray(1 To 72)
   
   'Light bricks  8
   BitArray(1) = Not &HFF
   BitArray(2) = Not &H20
   BitArray(3) = Not &H20
   BitArray(4) = Not &H20
   BitArray(5) = Not &HFF
   BitArray(6) = Not &H4
   BitArray(7) = Not &H4
   BitArray(8) = Not &H4
   'Dark bricks   9
   BitArray(9) = &HFF
   BitArray(10) = &H20
   BitArray(11) = &H20
   BitArray(12) = &H20
   BitArray(13) = &HFF
   BitArray(14) = &H4
   BitArray(15) = &H4
   BitArray(16) = &H4
   'Light squares 10
   BitArray(17) = Not &H0
   BitArray(18) = Not &H0
   BitArray(19) = Not &H3E
   BitArray(20) = Not &H22
   BitArray(21) = Not &H22
   BitArray(22) = Not &H3E
   BitArray(23) = Not &H0
   BitArray(24) = Not &H0
   'Heavy squares 11
   BitArray(25) = &H0
   BitArray(26) = &H0
   BitArray(27) = &H3E
   BitArray(28) = &H22
   BitArray(29) = &H22
   BitArray(30) = &H3E
   BitArray(31) = &H0
   BitArray(32) = &H0

   'Light dots  12
   BitArray(33) = Not &H80
   BitArray(34) = Not &H10
   BitArray(35) = Not &H0
   BitArray(36) = Not &H8
   BitArray(37) = Not &H0
   BitArray(38) = Not &H20
   BitArray(39) = Not &H0
   BitArray(40) = Not &H2
   'Heavy dots   13
   BitArray(41) = Not &H11
   BitArray(42) = Not &H44
   BitArray(43) = Not &H11
   BitArray(44) = Not &H44
   BitArray(45) = Not &H11
   BitArray(46) = Not &H44
   BitArray(47) = Not &H11
   BitArray(48) = Not &H44
   
   'Horz waves 14
   BitArray(49) = Not &H0
   BitArray(50) = Not &H0
   BitArray(51) = Not &H60
   BitArray(52) = Not &H90
   BitArray(53) = Not &H9
   BitArray(54) = Not &H6
   BitArray(55) = Not &H0
   BitArray(56) = Not &H0
   'Vert waves 15
   BitArray(57) = Not &H10
   BitArray(58) = Not &H20
   BitArray(59) = Not &H20
   BitArray(60) = Not &H10
   BitArray(61) = Not &H8
   BitArray(62) = Not &H4
   BitArray(63) = Not &H4
   BitArray(64) = Not &H8

   'Dark tiles 16
   BitArray(65) = &HFF
   BitArray(66) = &H80
   BitArray(67) = &H80
   BitArray(68) = &H80
   BitArray(69) = &H80
   BitArray(70) = &H80
   BitArray(71) = &H80
   BitArray(72) = &H80

End Sub

Public Sub OpenLoadDialog(Title$, Choice$, FileSpec$, InitDir$)
Form1.CommonDialog1.DialogTitle = Title$
'&H2 forces save to be same directory as open
'&H8 checks if file exists
Form1.CommonDialog1.Flags = &H8
Form1.CommonDialog1.CancelError = True
On Error GoTo cancelload
Form1.CommonDialog1.Filter = Choice$
Form1.CommonDialog1.InitDir = InitDir$
Form1.CommonDialog1.FileName = ""
Form1.CommonDialog1.ShowOpen
SetCursorPos 20, 120
FileSpec$ = Form1.CommonDialog1.FileName
Exit Sub
'============
cancelload:
Close
FileSpec$ = ""
SetCursorPos 20, 120
Exit Sub
Resume
End Sub
Public Sub OpenSaveDialog(Title$, Choice$, FileSpec$, InitDir$, SFile$)
Form1.CommonDialog1.DialogTitle = Title$
'&H8 forces save to be same directory as open
'&H2 checks if file exists & queries overwriting
Form1.CommonDialog1.Flags = &H2
Form1.CommonDialog1.CancelError = True
On Error GoTo cancelsave
Form1.CommonDialog1.Filter = Choice$
Form1.CommonDialog1.InitDir = InitDir$
Form1.CommonDialog1.FileName = SFile$
Form1.CommonDialog1.ShowSave

SetCursorPos 20, 120
FileSpec$ = Form1.CommonDialog1.FileName
Exit Sub
'============
cancelsave:
Close
FileSpec$ = ""
SetCursorPos 20, 120
Exit Sub
Resume
End Sub


