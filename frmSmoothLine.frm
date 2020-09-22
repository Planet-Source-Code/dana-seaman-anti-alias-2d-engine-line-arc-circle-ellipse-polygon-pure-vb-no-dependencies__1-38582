VERSION 5.00
Begin VB.Form frmSmoothLine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   -1965
   ClientWidth     =   9510
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Black Chancery"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Begin VB.OptionButton optFace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rolex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   7380
      TabIndex        =   16
      Top             =   6960
      Width           =   1155
   End
   Begin VB.OptionButton optFace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pooky"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   7380
      TabIndex        =   15
      Top             =   6600
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.PictureBox picV 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   5580
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   14
      Top             =   6600
      Width           =   1590
   End
   Begin VB.PictureBox picToxic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   -3000
      Picture         =   "frmSmoothLine.frx":0000
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   6060
      Width           =   4815
      Begin VB.CheckBox chk3D 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3D Clock Hands"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   100
         Left            =   2280
         Max             =   1000
         Min             =   25
         SmallChange     =   25
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1140
         Value           =   25
         Width           =   2295
      End
      Begin VB.CheckBox chkStop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkAA 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clock Anti-alias ON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   2070
      End
      Begin VB.CheckBox chkColor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Random Color ON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Get/Set Pixel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIBits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Radial/Star speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lblElapse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   105
      End
   End
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   0
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   9
      Top             =   3000
      Width           =   3000
   End
   Begin VB.PictureBox picClock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6225
      Left            =   3240
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   5
      Top             =   180
      Width           =   6225
      Begin VB.Image imgCenter 
         Height          =   180
         Left            =   3060
         Picture         =   "frmSmoothLine.frx":1EED
         Top             =   3060
         Width           =   180
      End
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   100
      Left            =   6420
      Top             =   120
   End
   Begin VB.PictureBox picSmooth 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   0
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   4
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmSmoothLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const sCap              As String = "2D Anti-Alias Engine (Line,Arc,Circle,Ellipse,Polygon). Pure Vb (NO dependencies)"
Const Pi                As Single = 3.141593 '3.14159265358979
Const Rads              As Single = Pi / 180
Const MAX_PATH          As Long = 260

Dim Buffer              As String * MAX_PATH
Dim sc(359, 3)          As Single  'look-up table for DrawRadial
Dim Radius              As Integer
Dim ClockRadius         As Integer
Dim m_OldHour           As Integer
Dim m_OldMin            As Integer
Dim m_OldSec            As Integer
Dim m_HandGradientEnd   As OLE_COLOR
Dim m_HandGradientStart As OLE_COLOR
Dim Star(4)             As Integer 'Star angles
Dim Start               As Long
Dim AA1                 As New LineGS 'DrawRadial
Dim AA2                 As New LineGS 'DrawStar
Dim AA3                 As New LineGS 'DrawPolygons(Clock hands)
'Sin/Cos look-up table
Dim Rotate(359, 1)      As Single
'Polygons point to 3 O'clock
Dim ptHour()            As POINTAPI
Dim ptMinute()          As POINTAPI
Dim ptSecond()          As POINTAPI
'Rotated polygons
Dim ptNewHour()         As POINTAPI
Dim ptNewMinute()       As POINTAPI
Dim ptNewSecond()       As POINTAPI
'for Gradients
Dim HandEndRGB          As RGB
Dim HandStartRGB        As RGB

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type TRIVERTEX
   X     As Long
   Y     As Long
   Red   As Integer
   Green As Integer
   Blue  As Integer
   Alpha As Integer
End Type

Private Type RGB
   Red   As Integer
   Green As Integer
   Blue  As Integer
End Type
  
Private Type GradientTRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type
   Private Const BF_LEFT = &H1
   Private Const BF_BOTTOM = &H8
   Private Const BF_RIGHT = &H4
   Private Const BF_TOP = &H2
   Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
   ' Private Const BDR_INNER = &HC
   ' Private Const BDR_OUTER = &H3
   ' Private Const BDR_RAISED = &H5
   Private Const BDR_RAISEDINNER = &H4
   Private Const BDR_RAISEDOUTER = &H1
   ' Private Const BDR_SUNKEN = &HA
   Private Const BDR_SUNKENINNER = &H8
   Private Const BDR_SUNKENOUTER = &H2
   ' Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
   ' Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
   Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
   Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Const GradientFILL_TRIANGLE As Long = &H2
'Private Declare Function GradientFillTri Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GradientTRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Used for Multilanguage Support
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'#####NB Use GetUserDefaultLCID in lieu of GetSystemDefaultLCID
'        to get correct current user LCID.
Const LOCALE_SSHORTDATE As Long = &H1F
Const LOCALE_IDATE      As Long = &H21    'short date format ordering
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Sub LoadClockFace()
   Dim sName As String
   
   sName = App.Path & "\" & IIf(optFace(0), "pooky.jpg", "rolex.jpg")
   picClock.AutoSize = True
   Set picClock = LoadPicture(sName)
   ClockRadius = picClock.ScaleHeight \ 2
   BuildPolygon
   'Position center image(anti-aliased in Corel)
   imgCenter.Move ClockRadius - imgCenter.Width \ 2, _
                  ClockRadius - imgCenter.Height \ 2

End Sub
Private Sub DrawRadialLines()
   Dim i       As Integer
   
   picSmooth.Cls

   Start = GetTickCount()

   If Option1(1) Then   'DIBits
      With picSmooth
         'Copy DIBits into  byte array
         AA1.DIB .hdc, .Image.Handle, .ScaleWidth, .ScaleHeight
      End With
   End If

   If Option1(1) Then
      For i = 0 To 359 Step 6
         AA1.LineDIB sc(i, 2), _
                    sc(i, 3), _
                    sc(i, 0), _
                    sc(i, 1), _
                    RGBRandom
      Next
      AA1.CircleDIB Radius, Radius, Radius * 0.75, Radius * 0.75, vbBlack
   ElseIf Option1(0) Then
      For i = 0 To 359 Step 6
         AA1.LineGP picSmooth.hdc, _
                   sc(i, 2), _
                   sc(i, 3), _
                   sc(i, 0), _
                   sc(i, 1), _
                   RGBRandom
      Next
      AA1.CircleGP picSmooth.hdc, _
                  Radius, _
                  Radius, _
                  Radius * 0.75, _
                  Radius * 0.75, _
                  vbBlack
   End If
   
   If Option1(1) Then
      'If using DIBits copy array back to hDC
      AA1.Array2Pic
   End If
   
   lblElapse = FormatNumber$(((GetTickCount() - Start)) / 60, 4) & _
               " ms per " & _
               Int(Radius * 0.5) & _
               " pixel line (Faster if compiled)"

End Sub
Private Sub DrawText(ByVal obj As Object, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal txt As String, _
   ByVal Effect As Long, _
   ByVal FirstColor As Long, _
   ByVal LastColor As Long, _
   ByVal MainColor As Long, _
   ByVal Depth As Long)
                     
   'Public Enum TextEffect
   '   [Normal] = 0
   '   [Engraved] = 1
   '   [Embossed] = 2
   '   [Shadowed] = 3
   'End Enum
                     
   Dim i As Integer
   Dim j As Integer
   obj.ScaleMode = vbPixels
   Select Case Effect
      Case 0
         obj.ForeColor = MainColor
         TextOut obj.hdc, X, Y, txt, Len(txt)
      Case Else
         '### First Color
         obj.ForeColor = FirstColor
         For i = 1 To Depth
            j = IIf(Effect = 1, i, -i)
            TextOut obj.hdc, X + j, Y + j, txt, Len(txt)
         Next
         '### Last Color
         If Effect <> 3 Then
            obj.ForeColor = LastColor
            For i = 1 To Depth
               j = IIf(Effect = 1, -i, i)
               TextOut obj.hdc, X + j, Y + j, txt, Len(txt)
            Next
         End If
         '### Main Color
         obj.ForeColor = MainColor
         TextOut obj.hdc, X, Y, txt, Len(txt)
   End Select
End Sub

Private Sub DrawStar()

   Dim L4 As Long
   Dim M4 As Long
   Dim N4 As Long
   
   For L4 = 0 To 4
      'Increment indexes by 6°
      Star(L4) = (Star(L4) + 6) Mod 360
   Next
   
   picStar.Cls

   If Option1(1) Then   'DIBits
      With picStar
         'Copy DIBits into  byte array
         AA2.DIB .hdc, .Image.Handle, .ScaleWidth, .ScaleHeight
      End With
   End If

   'Draw 5-point star
   For L4 = 0 To 4
      'Second point (+144°)
      M4 = (L4 + 2) Mod 5
      'Smoothline
      DrawStarLine sc(Star(L4), 0), _
                   sc(Star(L4), 1), _
                   sc(Star(M4), 0), _
                   sc(Star(M4), 1), _
                   vbBlack
      
      N4 = (L4 Mod 2) * 10
      If Option1(0) Then
         'Get/Set Pixel
         AA2.CircleGP picStar.hdc, _
                     sc(Star(M4), 0), _
                     sc(Star(M4), 1), _
                     N4 + 10, _
                     10, _
                     vbBlack
         AA2.ArcGP picStar.hdc, _
                  sc(Star(M4), 0), _
                  sc(Star(M4), 1), _
                  N4 + 15, _
                  15, _
                  90, _
                  180, _
                  vbBlack
         AA2.ArcGP picStar.hdc, _
                  sc(Star(M4), 0), _
                  sc(Star(M4), 1), _
                  N4 + 15, _
                  15, _
                  315, _
                  45, _
                  vbBlack
      Else
         'DIBits
         AA2.CircleDIB sc(Star(M4), 0), _
                      sc(Star(M4), 1), _
                      N4 + 10, _
                      10, _
                      vbBlack
         AA2.ArcDIB sc(Star(M4), 0), _
                   sc(Star(M4), 1), _
                   N4 + 15, _
                   15, _
                   90, _
                   180, _
                   vbBlack
         AA2.ArcDIB sc(Star(M4), 0), _
                   sc(Star(M4), 1), _
                   N4 + 15, _
                   15, _
                   315, _
                   45, _
                   vbBlack
      End If
   Next
  
   If Option1(1) Then
      'If using DIBits copy array back to hDC
      AA2.Array2Pic
   End If
   
End Sub

Private Sub DrawStarLine(ByVal X1 As Integer, _
   ByVal Y1 As Integer, _
   ByVal X2 As Integer, _
   ByVal Y2 As Integer, _
   ByVal Color As Long)

   If Option1(0) Then
      AA2.LineGP picStar.hdc, X1, Y1, X2, Y2, Color
   Else
      AA2.LineDIB X1, Y1, X2, Y2, Color
   End If

End Sub
Private Function RGBRandom() As Long
   If chkColor Then
      RGBRandom = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
   Else
      RGBRandom = vbBlack
   End If
End Function

Private Sub chk3D_Click()
   BuildPolygon
End Sub

Private Sub chkColor_Click()
   DrawRadialLines
End Sub

'Private Sub cmdVote_Click()
'  frmVote.Show 1
'End Sub

Private Sub Form_Load()
   Dim L4 As Long
   Dim SinRads As Single
   Dim CosRads As Single

   Me.Caption = sCap
   Radius = picSmooth.ScaleHeight \ 2
   LoadClockFace

   For L4 = 0 To 359
      '##### Sin/Cos look-up array(singles) for polygon
      '      rotation and Tick/Numeral positions
      SinRads = Sin(L4 * Rads)
      CosRads = Cos(L4 * Rads)
      Rotate(L4, 0) = SinRads
      Rotate(L4, 1) = CosRads
      'prebuild all x/y points for radial demo
      sc(L4, 0) = SinRads * Radius * 0.75 + Radius
      sc(L4, 1) = CosRads * Radius * 0.75 + Radius
      sc(L4, 2) = SinRads * Radius * 0.25 + Radius
      sc(L4, 3) = CosRads * Radius * 0.25 + Radius
   Next

   Randomize
                  
   picSmooth.BorderStyle = 0
   picStar.BorderStyle = 0

   'Init star points
   For L4 = 0 To 4
      Star(L4) = 36 + L4 * 72
   Next
   
   'Set Timers to HScroll1 value
   HScroll1_Change
   
   'for Demo...these are properties in real control
   m_HandGradientStart = vbWhite
   m_HandGradientEnd = vbBlack
   HandStartRGB = GetRGBColours(m_HandGradientStart)
   HandEndRGB = GetRGBColours(m_HandGradientEnd)
   
   'Paint clock hands
   Timer1_Timer 0
End Sub

Private Function GetRGBColours(lColour As Long) As RGB

   Dim HexColour As String
   TranslateColor lColour, 0, lColour
   HexColour = String(6 - Len(Hex$(lColour)), "0") & Hex$(lColour)
   GetRGBColours.Red = "&H" & Mid$(HexColour, 5, 2) & "00"
   GetRGBColours.Green = "&H" & Mid$(HexColour, 3, 2) & "00"
   GetRGBColours.Blue = "&H" & Mid$(HexColour, 1, 2) & "00"

End Function

Private Sub BuildPolygon()
   Dim n             As Long
   Dim m_HandRadius  As Long
   
   n = IIf(chk3D, 18, 4)

   'Points to 3Hr mark (90°)
   ReDim ptHour(n)
   ReDim ptMinute(n)
   ReDim ptSecond(n)
   'Rotated
   ReDim ptNewHour(n)
   ReDim ptNewMinute(n)
   ReDim ptNewSecond(n)
   
   'Init to bogus values so hands will update immediately
   m_OldHour = -1
   m_OldMin = -1
   m_OldSec = -1
   
   If chk3D Then
      
      m_HandRadius = ClockRadius * 0.7
      'Define upper half of second hand
      ptSecond(0).X = -m_HandRadius * 0.15
      'ptSecond(0).Y = 0
      ptSecond(1).X = -m_HandRadius * 0.1
      ptSecond(1).Y = m_HandRadius * 0.025
      ptSecond(2).X = -m_HandRadius * 0.05
      ptSecond(2).Y = m_HandRadius * 0.025
      ptSecond(3).X = -m_HandRadius * 0.05
      ptSecond(3).Y = m_HandRadius * 0.01
      ptSecond(4).X = m_HandRadius * 0.05
      ptSecond(4).Y = m_HandRadius * 0.01
      ptSecond(5).X = m_HandRadius * 0.05
      ptSecond(5).Y = m_HandRadius * 0.025
      ptSecond(6).X = m_HandRadius * 0.425
      ptSecond(6).Y = m_HandRadius * 0.05
      ptSecond(7).X = m_HandRadius * 0.8
      ptSecond(7).Y = m_HandRadius * 0.025
      ptSecond(8).X = m_HandRadius * 0.8
      ptSecond(8).Y = m_HandRadius * 0.07
      ptSecond(9).X = m_HandRadius * 0.95
      'ptSecond(9).Y = 0
   
      'Replicate the Second Hand
      CopySecToMin 9
   
      'Define upper half of hour hand
      ptHour(0).X = -m_HandRadius * 0.15
      'ptHour(0).Y = 0
      ptHour(1).X = -m_HandRadius * 0.1
      ptHour(1).Y = m_HandRadius * 0.025
      ptHour(2).X = -m_HandRadius * 0.05
      ptHour(2).Y = m_HandRadius * 0.025
      ptHour(3).X = -m_HandRadius * 0.05
      ptHour(3).Y = m_HandRadius * 0.01
      ptHour(4).X = m_HandRadius * 0.05
      ptHour(4).Y = m_HandRadius * 0.01
      ptHour(5).X = m_HandRadius * 0.05
      ptHour(5).Y = m_HandRadius * 0.025
      ptHour(6).X = m_HandRadius * 0.2
      ptHour(6).Y = m_HandRadius * 0.05
      ptHour(7).X = m_HandRadius * 0.45
      ptHour(7).Y = m_HandRadius * 0.025
      ptHour(8).X = m_HandRadius * 0.45
      ptHour(8).Y = m_HandRadius * 0.09
      ptHour(9).X = m_HandRadius * 0.6
      'ptHour(9).Y = 0

      'Replicate upper half to bottom for all 3 hands
      MirrorVerticals 10, 18, 18 'From, To, Index
   
   Else
      'Define upper half of second hand
      ptSecond(0).X = -ClockRadius * 0.2
      'ptSecond(0).y = 0
      'ptSecond(1).x = 0
      ptSecond(1).Y = ClockRadius * 0.05
      ptSecond(2).X = ClockRadius * 0.6   'Outermost Point
      'ptSecond(2).y = 0

      'Replicate the Second Hand
      CopySecToMin 2
           
      'Define upper half of hour hand
      ptHour(0).X = -ClockRadius * 0.2
      'ptHour(0).y = 0
      'ptHour(1).x = 0
      ptHour(1).Y = ClockRadius * 0.075
      ptHour(2).X = ClockRadius * 0.4     'Outermost Point
      'ptHour(2).y = 0
   
      'Replicate upper half to bottom for all 3 hands
      MirrorVerticals 3, 4, 4 'From, To, Index
   End If
End Sub
Private Sub MirrorVerticals(ByVal Start As Integer, ByVal Finish As Integer, ByVal Idx As Integer)
   Dim n As Integer
   For n = Start To Finish
      ptSecond(n).X = ptSecond(Idx - n).X
      ptSecond(n).Y = -ptSecond(Idx - n).Y
      ptMinute(n).X = ptMinute(Idx - n).X
      ptMinute(n).Y = -ptMinute(Idx - n).Y
      ptHour(n).X = ptHour(Idx - n).X
      ptHour(n).Y = -ptHour(Idx - n).Y
   Next
End Sub

Private Sub CopySecToMin(ByVal Finish As Integer)
   Dim n As Integer
   For n = 0 To Finish
      ptMinute(n).X = ptSecond(n).X
      ptMinute(n).Y = ptSecond(n).Y
   Next
End Sub

Private Sub RotatePoints(Points() As POINTAPI, NewPoints() As POINTAPI, ByVal Angle As Single)
   Dim i       As Integer
   Dim P       As Integer
   
   P = UBound(Points)
   'Use Sin/Cos lookup table Rotate() for speed
   For i = 0 To P
      NewPoints(i).X = Points(i).X * Rotate(Angle, 0) + _
                       Points(i).Y * Rotate(Angle, 1) + ClockRadius
      NewPoints(i).Y = -Points(i).X * Rotate(Angle, 1) + _
                       Points(i).Y * Rotate(Angle, 0) + ClockRadius
   Next
   
End Sub
Private Function StripNulls(ByVal sText As String) As String
   ' Remove nulls from string
   Dim nPosition&
   StripNulls = sText
   nPosition = InStr(sText, vbNullChar)
   If nPosition Then StripNulls = Left$(sText, nPosition - 1)
   If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = ""
End Function
Private Sub DrawPolygon(ptNew() As POINTAPI, OutlineColor As Long, FillColor As Long)
   Dim L4      As Integer
   Dim P       As Integer
   Dim hdc     As Long

   With picClock
           
      If chk3D Then
         .FillStyle = 1
         OutlineColor = vbBlack

         hdc = .hdc
         'Fill all the triangles with Gradients
         'Create 3D twisted effect
         DrawTriangleGradient hdc, ptNew(), 0, 1, 17, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 16, 1, 17, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 1, 2, 16, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 15, 3, 14, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 4, 3, 14, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 5, 6, 12, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 12, 5, 13, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 6, 7, 11, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 11, 6, 12, HandStartRGB, HandEndRGB
         DrawTriangleGradient hdc, ptNew(), 9, 8, 10, HandStartRGB, HandEndRGB
         If chkAA Then 'only if we are anti-aliasing
            'This line fixes glitch where the Gradients meet
            'It occurs only at some angles.
            AA3.LineGP picClock.hdc, _
                      ptNew(7).X, _
                      ptNew(7).Y, _
                      ptNew(11).X, _
                      ptNew(11).Y, _
                      m_HandGradientStart
            'These lines accentuate the twist effect as
            'well as hiding the jaggies where gradients meet .
            AA3.LineGP picClock.hdc, _
                      ptNew(1).X, _
                      ptNew(1).Y, _
                      ptNew(16).X, _
                      ptNew(16).Y, _
                      m_HandGradientStart
            AA3.LineGP picClock.hdc, _
                      ptNew(5).X, _
                      ptNew(5).Y, _
                      ptNew(12).X, _
                      ptNew(12).Y, _
                      m_HandGradientStart
            AA3.LineGP picClock.hdc, _
                      ptNew(6).X, _
                      ptNew(6).Y, _
                      ptNew(11).X, _
                      ptNew(11).Y, _
                      m_HandGradientStart
         End If
      Else

         '##### Draw API polygon
         .ForeColor = OutlineColor
         If FillColor <> -1 Then
            'Fill polygon
            .FillColor = FillColor
            .FillStyle = 0
            Polygon .hdc, ptNew(0), UBound(ptNew)
         ElseIf chkAA = False Then
            'Transparent
            .FillStyle = 1
            Polygon .hdc, ptNew(0), UBound(ptNew)
         End If
      End If

      If chkAA Then
          '##### Anti-alias outline
          '      Overwrites API polygon.
          P = UBound(ptNew) - 1
          For L4 = 0 To P
             AA3.LineGP picClock.hdc, _
                       ptNew(L4).X, _
                       ptNew(L4).Y, _
                       ptNew(L4 + 1).X, _
                       ptNew(L4 + 1).Y, _
                       OutlineColor
          Next
      End If
   End With
End Sub
Private Function DrawTriangleGradient(hdc As Long, _
   ptNew() As POINTAPI, _
   A As Integer, _
   B As Integer, _
   c As Integer, _
   lStart As RGB, _
   lEnd As RGB) As Long

   Dim V(2) As TRIVERTEX
   Dim Triangle As GradientTRIANGLE
    
   V(0).X = ptNew(A).X
   V(0).Y = ptNew(A).Y
   V(0).Red = lEnd.Red
   V(0).Green = lEnd.Green
   V(0).Blue = lEnd.Blue
   V(1).X = ptNew(B).X
   V(1).Y = ptNew(B).Y
   V(1).Red = lStart.Red
   V(1).Green = lStart.Green
   V(1).Blue = lStart.Blue
   V(2).X = ptNew(c).X
   V(2).Y = ptNew(c).Y
   V(2).Red = lStart.Red
   V(2).Green = lStart.Green
   V(2).Blue = lStart.Blue
   Triangle.Vertex1 = 0
   Triangle.Vertex2 = 1
   Triangle.Vertex3 = 2
   GradientFillRect hdc, _
      V(0), _
      3, _
      Triangle, _
      1, _
      GradientFILL_TRIANGLE _

End Function

Private Sub Form_Resize()
   Me.Move (Screen.Width - Me.Width) \ 2, 0
End Sub

Private Sub HScroll1_Change()
   Timer1(1).Interval = HScroll1.Value
End Sub

Private Sub optFace_Click(Index As Integer)
   LoadClockFace
End Sub

Private Sub Timer1_Timer(Index As Integer)
   Dim tim              As Date
   Dim m_Hour           As Integer
   Dim m_Min            As Integer
   Dim m_Sec            As Integer
   Dim bChange          As Boolean
   Static Position      As Long
   
   If chkStop Then Exit Sub

   Select Case Index
      Case 1
         DrawStar
         DrawRadialLines
      Case 2
         Position = (Position + 1) Mod 5
         BitBlt picV.hdc, 0, 0, 106, 26, _
            picToxic.hdc, 0, Position * 26, _
            vbSrcCopy
      Case 0 'Paint clock hands.
         tim = Time 'prevents calling Time multiple times
         'Get Time components and convert to degrees °
         m_Sec = Second(tim) * 6
         m_Min = Minute(tim) * 6
         'Adjust hour hand to include minute angle \12
         m_Hour = (Hour(tim) Mod 12) * 30 + m_Min \ 12
         'Rotate Hours only when change is detected
         If m_Hour <> m_OldHour Then
            m_OldHour = m_Hour
            RotatePoints ptHour, ptNewHour, m_Hour
            bChange = True
         End If
         'Rotate Minutes only when change is detected
         If m_Min <> m_OldMin Then
            m_OldMin = m_Min
            RotatePoints ptMinute, ptNewMinute, m_Min
            bChange = True
         End If
         'Rotate Seconds only when change is detected
         If m_Sec <> m_OldSec Then
            m_OldSec = m_Sec
            RotatePoints ptSecond, ptNewSecond, m_Sec
            bChange = True
         End If

         If bChange Then 'Update hands
            'Clock face is frozen in Picture property.
            picClock.Cls  'Clear prior paints.
            DrawDigitalTime tim
            Draw_WeekDay_Month
            'Draw hands (FillColor = -1 if transparent)
            DrawPolygon ptNewHour(), vbWhite, -1
            DrawPolygon ptNewMinute(), vbWhite, -1
            DrawPolygon ptNewSecond(), vbWhite, vbRed
            picClock.Refresh
         End If
   End Select

End Sub

Private Sub DrawDigitalTime(tim As Date)
   Dim temp       As String
   Dim X          As Long
   Dim Y          As Long
      
   temp = FormatDateTime(tim, vbLongTime)
   picClock.FontSize = 24
   Set Me.Font = picClock.Font
   X = (picClock.ScaleWidth - TextWidth(temp)) / 2
   Y = (ClockRadius * 1.35) - TextHeight(temp)
       
   DrawText picClock, _
      X, Y, _
      temp, 1, _
      vbWhite, &HC0C0C0, _
      vbBlack, 1
   
End Sub
Private Sub Draw_WeekDay_Month()
   Dim TR       As RECT
   Dim h_brush  As Long
   Dim dat      As Date
   Dim temp     As String
   Dim X        As Long
   Dim Y        As Long
   
   dat = Date
   picClock.FontSize = 16
   Set Me.Font = picClock.Font
   picClock.ForeColor = vbWhite
   temp = Format$(dat, "ddd")
   X = imgCenter.Left + 35
   Y = ClockRadius - TextHeight(temp) \ 2
   SetRect TR, X, Y, X + TextWidth(temp), Y + TextHeight(temp)
   InflateRect TR, 4, 2
   'Fill BackColor
   h_brush = CreateSolidBrush(vbBlack)
   FillRect picClock.hdc, TR, h_brush
   'Sunken Box
   DrawEdge picClock.hdc, TR, EDGE_SUNKEN, BF_RECT
   TextOut picClock.hdc, _
      X, _
      Y, _
      temp, Len(temp)

   'Get day_month order so this will work
   'internationally
   X = GetLocaleInfo(GetUserDefaultLCID(), LOCALE_IDATE, Buffer, MAX_PATH)
   If StripNulls(Buffer) = "1" Then
      temp = Format$(dat, "dd MMM") 'day preceeds month
   Else
      temp = Format$(dat, "MMM dd") 'month preceeds day
   End If

   X = imgCenter.Left - 10 - TextWidth(temp)
   Y = ClockRadius - TextHeight(temp) \ 2
   SetRect TR, X, Y, X + TextWidth(temp), Y + TextHeight(temp)
   InflateRect TR, 4, 2
   'Fill BackColor
   h_brush = CreateSolidBrush(vbBlack)
   FillRect picClock.hdc, TR, h_brush
   'Sunken Box
   DrawEdge picClock.hdc, TR, EDGE_SUNKEN, BF_RECT
   TextOut picClock.hdc, _
      X, _
      Y, _
      temp, Len(temp)

End Sub
