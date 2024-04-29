VERSION 5.00
Begin VB.Form dance 
   BackColor       =   &H00000000&
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "dance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmddance 
      BackColor       =   &H80000009&
      Caption         =   "CLICK TO DANCE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "CLICK WILL MAKE DANCE A MOUSE POINTER"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "PRESS ESC TO STOP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "dance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Dim xx As Integer
Dim yy As Integer
Dim a As Integer
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim sh As Integer
Dim sw As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim s As String
Dim s1 As String
Dim l As Integer
Private Sub cmddance_Click()
Timer1.Enabled = True
xx = Rnd * 10 + 1
yy = Rnd * 10 + 1
End Sub



Private Sub cmddance_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyEscape Then
    Timer1.Enabled = False
    dance.Caption = s
End If
End Sub

Private Sub Form_Load()
s = "DANCING POINTER BY DHEEREN PATEL"
l = 1
End Sub

Private Sub Timer1_Timer()
Dim pp As POINTAPI
GetCursorPos pp
sh = (Screen.Height / 15) - 1
sw = (Screen.Width / 15) - 1
If pp.x <= 0 Then xx = -xx
If pp.x >= sw Then xx = -xx
If pp.y <= 0 Then yy = -yy
If pp.y >= sh Then yy = -yy
DoEvents
pp.x = pp.x + xx
pp.y = pp.y + yy
SetCursorPos pp.x, pp.y
lblinfo.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
s1 = Left(s, l)
dance.Caption = s1
l = l + 1
If l >= Len(s) + 3 Then
    l = 1
    s1 = ""
End If
End Sub
