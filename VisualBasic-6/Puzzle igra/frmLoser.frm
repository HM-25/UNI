VERSION 5.00
Begin VB.Form frmLoser 
   BackColor       =   &H00FAEDE2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surrender"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEDE2&
      Caption         =   "Select Loser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Surrender"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Surrender"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lower Player"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upper Player"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Which player surrenders?"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmLoser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
    Me.Hide
    'stop the music
    StopMIDI Song
    'play winning sound
    sndPlaySound App.Path & "\Win.wav", SND_ASYNC
    frmMain.Image8.Visible = False
    frmMain.BackColor = frmMain.Label2.ForeColor
    MsgBox RedUser & " wins the match !!", vbInformation
    'Default form color
    frmMain.BackColor = &HFAEDE2
    If frmMain.Picture Then
        frmMain.Image8.Visible = False
    Else
        frmMain.Image8.Visible = True
    End If
    'red/lower player has won one more game
    RedWins = RedWins + 1
    Label3.Caption = RedUser & " = " & RedWins
    frmMain.Timer1.Enabled = True
    'start the new game
    NewGame
    Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Me.Hide
    StopMIDI Song
    sndPlaySound App.Path & "\Win.wav", SND_ASYNC
    frmMain.Image8.Visible = False
    frmMain.BackColor = frmMain.Label1.ForeColor
    MsgBox GreenUser & " wins the match !!", vbInformation
    frmMain.BackColor = &HFAEDE2
    If frmMain.Picture Then
        frmMain.Image8.Visible = False
    Else
        frmMain.Image8.Visible = True
    End If
    GreenWins = GreenWins + 1
    frmMain.Label4.Caption = GreenUser & " = " & GreenWins
    frmMain.Timer1.Enabled = True
    NewGame
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Label2.ForeColor = FirstColor
    Label3.ForeColor = SecondColor
    Label2.Caption = GreenUser
    Label3.Caption = RedUser
End Sub
