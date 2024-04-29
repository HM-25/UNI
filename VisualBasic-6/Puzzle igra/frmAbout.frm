VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FAEDE2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5145
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2132.773
   ScaleMode       =   0  'User
   ScaleWidth      =   4831.421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4560
      Top             =   840
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   465
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3960
      Picture         =   "frmAbout.frx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   600
      Picture         =   "frmAbout.frx":0614
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":091E
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   600
      Picture         =   "frmAbout.frx":0C28
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "sarfraznawaz2005@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmAbout.frx":0F32
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.angelfire.com/ultra/sarfraz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   525
      MouseIcon       =   "frmAbout.frx":123C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "To check for an update of this program or to check out other programs, please visit this website:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4605
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   3765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   4
      Top             =   780
      Width           =   2925
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: SARFRAZ AHMED CHANDIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   135
      TabIndex        =   2
      Top             =   2505
      Width           =   3870
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1170
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " (10/10/2006)"
    lblTitle.Caption = App.Title
    Me.Icon = frmMain.Icon
End Sub

Private Sub Label1_Click()
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End Sub

Private Sub Label2_Click()
On Error Resume Next
Shell "start mailto:sarfrazahmed_pk@yahoo.com", vbHide
End Sub

'Animate balls
Private Sub Timer1_Timer()
    Image1.Picture = Image2.Picture
    Image2.Picture = Image3.Picture
    Image3.Picture = Image4.Picture
    Image4.Picture = Image5.Picture
    Image5.Picture = Image1.Picture
End Sub
