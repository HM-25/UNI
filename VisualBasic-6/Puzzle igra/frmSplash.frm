VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FAEDE2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classic Puzzle Game"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtRed 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Text            =   "Player Two"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtGreen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1140
      TabIndex        =   0
      Text            =   "Player One"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox Pic1 
      BackColor       =   &H00FAEDE2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   4335
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FCCDAB&
      BorderWidth     =   15
      X1              =   240
      X2              =   4920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   540
      Picture         =   "frmSplash.frx":030A
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   540
      Picture         =   "frmSplash.frx":0614
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3120
      Picture         =   "frmSplash.frx":091E
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   3120
      Picture         =   "frmSplash.frx":0C28
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FBBA8A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Recommended screen resolution 800x600)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -240
      TabIndex        =   6
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   540
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   540
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FCCDAB&
      BorderWidth     =   15
      FillColor       =   &H00FAEDE2&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   120
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnStart_Click()

On Error Resume Next
SaveSetting App.Title, "Settings", "User Green", txtGreen.Text
SaveSetting App.Title, "Settings", "User Red", txtRed.Text

GreenUser = txtGreen.Text
RedUser = txtRed.Text

Load frmMain
frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()

' Increasing the maximum value of i (here it is 5) ,
' increases the thickness of text & more deeper shadow to it

Dim i As Integer
Pic1.AutoRedraw = True
Pic1.ScaleMode = 2 'Point
For i = 1 To 7
    Pic1.CurrentX = i
    Pic1.CurrentY = i
    Pic1.Print "CLASSIC PUZZLE GAME"
    If i < 3 Then Pic1.ForeColor = vbYellow
    If i > 3 Then Pic1.ForeColor = vbRed
Next

'get the green/upper username
If Len(GetSetting(App.Title, "Settings", "User Green")) > 0 Then
    txtGreen.Text = GetSetting(App.Title, "Settings", "User Green")
End If

'get the red/lower username
If Len(GetSetting(App.Title, "Settings", "User Red")) > 0 Then
    txtRed.Text = GetSetting(App.Title, "Settings", "User Red")
End If

'get the ball colors
If GetSetting(App.Title, "Settings", "RedBlue") = "True" Then
    Image1.Picture = Image6.Picture
    Image2.Picture = Image7.Picture
    txtGreen.ForeColor = &HC0&          'Red
    txtRed.ForeColor = vbBlue
ElseIf GetSetting(App.Title, "Settings", "GreenBlue") = "True" Then
    Image1.Picture = Image5.Picture
    Image2.Picture = Image7.Picture
    txtRed.ForeColor = vbBlue
ElseIf GetSetting(App.Title, "Settings", "GreenYellow") = "True" Then
    Image1.Picture = Image5.Picture
    Image2.Picture = Image9.Picture
    txtRed.ForeColor = &H8080&          'Yellow
ElseIf GetSetting(App.Title, "Settings", "RedYellow") = "True" Then
    Image1.Picture = Image6.Picture
    Image2.Picture = Image9.Picture
    txtGreen.ForeColor = &HC0&          'Red
    txtRed.ForeColor = &H8080&          'Yellow
ElseIf GetSetting(App.Title, "Settings", "YellowBlue") = "True" Then
    Image1.Picture = Image9.Picture
    Image2.Picture = Image7.Picture
    txtGreen.ForeColor = &H8080&        'Yellow
    txtRed.ForeColor = vbBlue
Else
    Image1.Picture = Image5.Picture
    Image2.Picture = Image6.Picture
    FirstColor = &H8000&                'Green
    SecondColor = &HC0&                 'Red
End If

'set text colors as per ball color
FirstColor = txtGreen.ForeColor
SecondColor = txtRed.ForeColor

'Associate gam files to this program
Associate App.Path & "\" & App.EXEName & ".EXE", "gam", "Classic Puzzle Game File", App.Path & "\File.ico"

If Len(Command) > 0 Then
    If LCase(Right(Command, 3)) = "gam" Then
        txtGreen.Enabled = False
        txtRed.Enabled = False
        btnStart_Click
    End If
End If

End Sub

Private Sub txtGreen_Click()
If txtGreen.Text = "Player One" Then txtGreen.Text = ""
End Sub

Private Sub txtRed_Click()
If txtRed.Text = "Player Two" Then txtRed.Text = ""
End Sub

