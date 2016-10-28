VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Hx19 3D interface"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   Icon            =   "hx19xyzDDEii.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "hx19"
   ScaleHeight     =   10230
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Text            =   "99999,10000 >"
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "< 0,0 (x,z)"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Text            =   "99000,99000 >"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   9
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Text            =   "< 0,0 (x,y)"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TX"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sync"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "DDE"
      ToolTipText     =   "Dynamic Data Exchange"
      Top             =   480
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Log"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "log incoming data \dataFiles\"
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Ready for data"
      ToolTipText     =   "tag, X,Y,Z, time(10mS),record,#receivers,(detecting receivers)"
      Top             =   480
      Width           =   4095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   8
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim linebuffer(200)
Private Sub positionTag(tag)
Static prox%

freeze = count100ms
scount = scount + 1
If count100ms > 8640000 Then count100ms = 0: scount = 0

If pt(tag) > 20 Then pt(tag) = 20
If pt(tag) < 1 Then GoTo skipover

'the following places receiver ID and receiver distance for the tag being positioned into arrays ordered with the closest receiver first
For j = 0 To pt(tag) - 1
 xlow = 100000000000#
 For i = 0 To pt(tag) - 1
  If xt(tag, i) <= xlow And xt(tag, i) <= 15000 And xt(tag, i) > 0 And rx(tag, i) > 0 Then xlow = xt(tag, i): mn = i
 Next
 xtSorted(j) = xlow: rxSorted(j) = rx(tag, mn): rx(tag, mn) = 0:
Next
'the following is for display only and not important
receivers = ""
For i = 0 To pt(tag) - 1
receivers = receivers + Format(rxSorted(i), " 0")
Next


If maxPoints = 1 Or pt(tag) < 2 Then    'if not enough receivers exist to compute 3d or 2d then place the point close to the mapped location of the detecting receiver
 xp = xx(rxSorted(0)) + 100 * Rnd()
 yp = yy(rxSorted(0)) + 100 * Rnd()
 zp = zz(rxSorted(0))
Text1 = Format(tag, " 0") + Format(xp, " 0") + Format(yp, " 0") + Format(zp, " 0") + Format(freeze, " 0") + Format(scount, " 0 ") + Format(pt(tag), " 0 ") + "(" + receivers + ")"
' If Check1.Value = 1 Then Print #1, tag, Round(xp, 0), Round(yp, 0), Round(zp, 0), freeze, scount, pt(tag); "("; receivers; ")"
' SYNTAX ERROR HERE
'  GoTo skipover
' End If

If pt(tag) < minPoints Then GoTo skipover
If maxPoints < pt(tag) Then pt(tag) = maxPoints
'for 3d or 2d the following is executed
'following task
ii = 0
pts = mm
For j = 0 To pt(tag) - 2
 For i = j + 1 To pt(tag) - 1
  di = (xtSorted(i) - xtSorted(j))
   maxInterval = Sqr((xx(rxSorted(i)) - xx(rxSorted(j))) ^ 2 + (yy(rxSorted(i)) - yy(rxSorted(j))) ^ 2)
   If maxInterval > di And di >= 0 Then
   diff(ii) = di                            'this array holds the difference in the time of flight relative to a pair of receivers and is used to estimate the x and y coordinates
   rxA(ii) = rxSorted(i): rxB(ii) = rxSorted(j)
   ii = ii + 1
  End If
 Next
Next
good = ii   'the good variable states how many pairs qualified the sorting proceedure

ymax = -10000: xmax = -10000
ymin = 100000: xmin = 100000

'the following looks for the position of the receivers in the map file, and determines the range by which the computational iteration is to take place.
For i = 0 To good - 1
If xx(rxA(i)) > xmax Then xmax = xx(rxA(i))
If xx(rxA(i)) < xmin Then xmin = xx(rxA(i))
If xx(rxB(i)) > xmax Then xmax = xx(rxB(i))
If xx(rxB(i)) < xmin Then xmin = xx(rxB(i))
If yy(rxA(i)) > ymax Then ymax = yy(rxA(i))
If yy(rxA(i)) < ymin Then ymin = yy(rxA(i))
If yy(rxB(i)) > ymax Then ymax = yy(rxB(i))
If yy(rxB(i)) < ymin Then ymin = yy(rxB(i))
Next

z = 0
dIteration 0    'first itereation is started for Z=0
getZ (pt(tag))  'Z is estimated and feed to next iteration and so forth.
dIteration zp
getZ (pt(tag))
dIteration zp
getZ (pt(tag))

If xp < 0 Or xp > 100000 Then xp = xf(tag)  'if error is unacceptable we us x forcasted from the double exponential filter.
If yp < 0 Or yp > 100000 Then yp = yf(tag)
If zp < 0 Or zp > 15000 Then zp = zf(tag)

'the followin is a double exponential filter as presented by wikipedia see (http://en.wikipedia.org/wiki/Exponential_smoothing)
'there are probably better ways of filtering, but this is fairly good
If alpha <> 0 Then exponentialFilter tag
Text1 = Format(tag, " 0") + Format(xp, " 0") + Format(yp, " 0") + Format(zp, " 0") + Format(freeze, " 0") + Format(scount, " 0 ") + Format(pt(tag), " 0 ") + "(" + receivers + ")"
'If Check1.Value = 1 Then Print #1, tag, Round(xp, 0), Round(yp, 0), Round(zp, 0), freeze, scount, pt(tag); "("; receivers; ")"
' SYNTAX ERROR HERE
'sixa% = tag
'the following just plots the position of the tag on the window set in the setup file
xyDotPlot sixa
xzDotPlot sixa
prox = (prox + 1) And 7: If prox = 0 Then refreshMap ': dogrid
skipover:
pt(tag) = 0
End Sub
Private Sub dIteration(zz1)

'this routine selects the steps at which the x, y and z are to be computationally tested against measured results
'at first a coarse minima is found, then the steps get finer and finer until a precise minima is found
'note that when a coarse minima is found, for next iteration there is a slight backstep and then the next steps are devided by 2
'so it gets finer and finer.

xstep = (xmax - xmin) / 4
xmin = xmin - 4 * xstep
xmax = xmax + 4 * xstep

ystep = (ymax - ymin) / 4
ymin = ymin - 4 * ystep
ymax = ymax + 4 * ystep
zzz = zz1
dscanXYZ xmin, xmax, ymin, ymax, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
donn:
End Sub

