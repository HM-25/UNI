VERSION 5.00
Begin VB.Form frmOpt 
   BackColor       =   &H00FAEDE2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
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
   ScaleHeight     =   6810
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDefaults 
      Cancel          =   -1  'True
      Caption         =   "&Default"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FAEDE2&
      Caption         =   "Balls"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   3015
      Begin VB.OptionButton optYellowBlue 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Yellow-Blue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton optRedYellow 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Red-Yellow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton optGreenYellow 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Green-Yellow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optGreenBlue 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Green-Blue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optRedBlue 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Red-Blue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optRedGreen 
         BackColor       =   &H00FAEDE2&
         Caption         =   "Green-Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":0000
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":030A
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":0614
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":091E
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":0C28
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":0F32
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":123C
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":1546
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":1850
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":1B5A
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1800
         Picture         =   "frmOpt.frx":1E64
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   2400
         Picture         =   "frmOpt.frx":216E
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEDE2&
      Caption         =   "Set Colors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         BackColor       =   &H00EAFFEA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Board color"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Border color"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEDE2&
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   3015
      Begin VB.CommandButton btnRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnApply 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnApply_Click()
Dim sFileName As String
    
sFileName = OpenDialog(Me, "Picture Files|*.jpg;*.bmp;*.gif;*.wmf", "Open", "")

If Len(sFileName) > 0 Then
    frmMain.Picture = LoadPicture(sFileName)
    SaveSetting App.Title, "Settings", "Picture", sFileName
    frmMain.Shape7.Visible = False
    frmMain.Shape9.Visible = False
    frmMain.Image8.Visible = False
End If
End Sub

Private Sub btnDefaults_Click()

If MsgBox("Are you sure you want to restore the default settings?", vbQuestion + vbYesNo) = vbYes Then
    On Error Resume Next
    DeleteSetting App.Title, "Settings"
    
    If Len(GetSetting(App.Title, "Settings", "User Green")) = 0 Then
        MsgBox "Default settings restored successfully !" & vbNewLine & "Restart the game for the change to take effect !!", vbInformation
        Unload Me
    End If
End If

End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnRemove_Click()
On Error Resume Next
frmMain.Picture = LoadPicture("")
frmMain.Shape7.Visible = True
frmMain.Image8.Visible = True
frmMain.Shape9.Visible = True

If Len(GetSetting(App.Title, "Settings", "Picture")) > 0 Then
    DeleteSetting App.Title, "Settings", "Picture"
End If

If Len(GetSetting(App.Title, "Settings", "Picture")) > 0 Then
    MsgBox "Unknown error while deleting the picture entry from the Registry!!", vbCritical
End If
End Sub

Private Sub Command1_Click()
Dim myColor As Long

On Error Resume Next
myColor = ShowColor(Me)
Command1.BackColor = myColor

SaveSetting App.Title, "Settings", "BoardColor", Command1.BackColor
frmMain.Shape7.FillColor = Command1.BackColor

End Sub

Private Sub Command2_Click()
Dim myColor As Long

On Error Resume Next
myColor = ShowColor(Me)
Command2.BackColor = myColor

SaveSetting App.Title, "Settings", "BorderColor", Command2.BackColor

frmMain.Shape1.BorderColor = Command2.BackColor
frmMain.Shape2.BorderColor = Command2.BackColor
frmMain.Shape3.BorderColor = Command2.BackColor
frmMain.Line1.BorderColor = Command2.BackColor
frmMain.Line2.BorderColor = Command2.BackColor
frmMain.Line3.BorderColor = Command2.BackColor
frmMain.Line4.BorderColor = Command2.BackColor

End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon

If Len(GetSetting(App.Title, "Settings", "BoardColor")) > 0 Then
    Command1.BackColor = GetSetting(App.Title, "Settings", "BoardColor")
End If

If Len(GetSetting(App.Title, "Settings", "BorderColor")) > 0 Then
    Command2.BackColor = GetSetting(App.Title, "Settings", "BorderColor")
End If

If GetSetting(App.Title, "Settings", "RedBlue") = "True" Then
    optRedBlue.Value = True
ElseIf GetSetting(App.Title, "Settings", "GreenBlue") = "True" Then
    optGreenBlue.Value = True
ElseIf GetSetting(App.Title, "Settings", "GreenYellow") = "True" Then
    optGreenYellow.Value = True
ElseIf GetSetting(App.Title, "Settings", "RedYellow") = "True" Then
    optRedYellow.Value = True
ElseIf GetSetting(App.Title, "Settings", "YellowBlue") = "True" Then
    optYellowBlue.Value = True
Else
    optRedGreen.Value = True
End If

End Sub

Private Sub optGreenBlue_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "GreenRed"
    DeleteSetting App.Title, "Settings", "RedBlue"
    DeleteSetting App.Title, "Settings", "GreenYellow"
    DeleteSetting App.Title, "Settings", "RedYellow"
    DeleteSetting App.Title, "Settings", "YellowBlue"


    SaveSetting App.Title, "Settings", "GreenBlue", optGreenBlue.Value

    frmMain.Label1.ForeColor = &H8000&        'Green
    frmMain.Label4.ForeColor = &H8000&        'Green
    frmMain.Label2.ForeColor = vbBlue
    frmMain.Label3.ForeColor = vbBlue
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
    Next
        
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
    Next

End Sub

Private Sub optGreenYellow_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "GreenRed"
    DeleteSetting App.Title, "Settings", "RedBlue"
    DeleteSetting App.Title, "Settings", "GreenBlue"
    DeleteSetting App.Title, "Settings", "RedYellow"
    DeleteSetting App.Title, "Settings", "YellowBlue"

    SaveSetting App.Title, "Settings", "GreenYellow", optGreenYellow.Value

    frmMain.Label1.ForeColor = &H8000&      'Green
    frmMain.Label4.ForeColor = &H8000&      'Green
    frmMain.Label2.ForeColor = &H8080&      'Yellow
    frmMain.Label3.ForeColor = &H8080&      'Yellow
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
    Next
        
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image9.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image9.Picture
    Next

End Sub

Private Sub optRedBlue_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "GreenRed"
    DeleteSetting App.Title, "Settings", "GreenBlue"
    DeleteSetting App.Title, "Settings", "GreenYellow"
    DeleteSetting App.Title, "Settings", "RedYellow"
    DeleteSetting App.Title, "Settings", "YellowBlue"

    SaveSetting App.Title, "Settings", "RedBlue", optRedBlue.Value
    
    frmMain.Label1.ForeColor = &HC0&     'Red
    frmMain.Label4.ForeColor = &HC0&     'Red
    frmMain.Label2.ForeColor = vbBlue
    frmMain.Label3.ForeColor = vbBlue
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image6.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image6.Picture
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
    Next

End Sub

Private Sub optRedGreen_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "RedBlue"
    DeleteSetting App.Title, "Settings", "GreenBlue"
    DeleteSetting App.Title, "Settings", "GreenYellow"
    DeleteSetting App.Title, "Settings", "RedYellow"
    DeleteSetting App.Title, "Settings", "YellowBlue"

    SaveSetting App.Title, "Settings", "GreenRed", optRedGreen.Value
    
    frmMain.Label1.ForeColor = &H8000&       'Green
    frmMain.Label4.ForeColor = &H8000&       'Green
    frmMain.Label2.ForeColor = &HC0&         'Red
    frmMain.Label3.ForeColor = &HC0&         'Red
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image6.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image6.Picture
    Next

End Sub

Private Sub optRedYellow_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "GreenRed"
    DeleteSetting App.Title, "Settings", "RedBlue"
    DeleteSetting App.Title, "Settings", "GreenBlue"
    DeleteSetting App.Title, "Settings", "GreenYellow"
    DeleteSetting App.Title, "Settings", "YellowBlue"

    SaveSetting App.Title, "Settings", "RedYellow", optRedYellow.Value

    frmMain.Label1.ForeColor = &HC0&        'Red
    frmMain.Label4.ForeColor = &HC0&        'Red
    frmMain.Label2.ForeColor = &H8080&      'Yellow
    frmMain.Label3.ForeColor = &H8080&      'Yellow
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image6.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image6.Picture
    Next
        
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image9.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image9.Picture
    Next


End Sub

Private Sub optYellowBlue_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    
    DeleteSetting App.Title, "Settings", "GreenRed"
    DeleteSetting App.Title, "Settings", "RedBlue"
    DeleteSetting App.Title, "Settings", "GreenBlue"
    DeleteSetting App.Title, "Settings", "GreenYellow"
    DeleteSetting App.Title, "Settings", "RedYellow"

    SaveSetting App.Title, "Settings", "YellowBlue", optYellowBlue.Value

    frmMain.Label1.ForeColor = &H8080&      'Yellow
    frmMain.Label4.ForeColor = &H8080&      'Yellow
    frmMain.Label2.ForeColor = vbBlue
    frmMain.Label3.ForeColor = vbBlue
    
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image9.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image9.Picture
    Next
        
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
    Next

End Sub
