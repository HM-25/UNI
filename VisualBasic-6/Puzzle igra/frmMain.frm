VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FAEDE2&
   BorderStyle     =   0  'None
   Caption         =   "Classic Puzzle Game by SARFRAZ AHMED CHANDIO"
   ClientHeight    =   9000
   ClientLeft      =   450
   ClientTop       =   0
   ClientWidth     =   10905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10905
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "&Minimize"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Play Music"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   8775
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit Game"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   5760
      Top             =   3000
   End
   Begin VB.CommandButton btnConfess 
      Caption         =   "S&urrender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Op&tions"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   0
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   5280
      Picture         =   "frmMain.frx":030A
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   465
      Left            =   2760
      Picture         =   "frmMain.frx":0614
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   4920
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   5280
      Picture         =   "frmMain.frx":3B66
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   1800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Balls Beaten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   1800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "( 9 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "( 9 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player Turn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "frmMain.frx":3E70
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game # "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   930
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   120
      Tag             =   "Red"
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   120
      Tag             =   "Green"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5280
      Picture         =   "frmMain.frx":417A
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   5280
      Picture         =   "frmMain.frx":4484
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   45
   End
   Begin VB.Image Image4 
      Height          =   675
      Left            =   6600
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "frmMain.frx":478E
      Stretch         =   -1  'True
      Tag             =   "Fire"
      ToolTipText     =   "Put beaten balls here for removal"
      Top             =   4080
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   45
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   8880
      X2              =   11280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   2760
      X2              =   5160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   6960
      X2              =   6960
      Y1              =   5760
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   6960
      X2              =   6960
      Y1              =   720
      Y2              =   2880
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   2895
      Left            =   5160
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   5055
      Left            =   3960
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   7215
      Left            =   2760
      Top             =   720
      Width           =   8535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      FillColor       =   &H00E9FEF4&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00EAFFEA&
      FillStyle       =   0  'Solid
      Height          =   7215
      Left            =   2760
      Top             =   720
      Width           =   8535
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      FillStyle       =   7  'Diagonal Cross
      Height          =   2775
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FCCDAB&
      FillStyle       =   0  'Solid
      Height          =   9200
      Left            =   0
      Top             =   -120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   CLASSIC PUZZLE GAME
'   MADE BY
'   SARFRAZ AHMED CHANIDO
'   (sarfraznawaz2005@yahoo.com)

'   Date:Wed, October 11,2006

Option Explicit

Private Sub btnConfess_Click()
    frmLoser.Show 1
End Sub

Private Sub btnNew_Click()
If MsgBox("Are you sure you want to play a New game?", vbQuestion + vbYesNo) = vbYes Then
'play the new game
    NewGame
End If
End Sub

'opens already saved games
Private Sub btnOpen_Click()
Dim sFileName As String

sFileName = OpenDialog(Me, "Classic Puzzle Game Files (*.gam)|*.gam", "Open", App.Path)

If Len(sFileName) > 0 Then
    NewGame
    Dim i As Integer
    Dim j As Integer
    Dim FSO As Object
    Dim File As Object
    Dim strLine As String
    Dim FileOneNumber As Integer
    Dim FileTwoNumber As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.OpenTextFile(sFileName)
    
    'hide all upper balls
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Visible = False
    Next

    'hide all lower balls
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Visible = False
    Next

    'read the file until end
    Do While File.AtEndOfStream = False
        'read each line and asign to strLine variable
        strLine = ToggleText(File.ReadLine)
        
        If strLine = "END" Then
        'total upper balls tobe placed at top-left corner
            FileOneNumber = CInt(File.Line) - 1 '-2 for total files visible
        ElseIf strLine = "ENDAGAIN" Then
        'total lower balls tobe placed at top-left corner
            FileTwoNumber = CInt(File.Line) - 1
        End If
       
On Error Resume Next
        
        'now move all available red balls to their original
        'location when game was saved by reading their
        'location from the game file
        For i = 1 To FileOneNumber
            Image3.Item(GetBallNumber(strLine)).Move GetLeft(strLine), GetTop(strLine)
            Image3.Item(GetBallNumber(strLine)).Visible = True
            'Red/lower ball moved to the board
            Image3.Item(GetBallNumber(strLine)).Tag = "RedMoved"
        Next

        For j = FileOneNumber To FileTwoNumber
            Image2.Item(GetBallNumber(strLine)).Move GetLeft(strLine), GetTop(strLine)
            Image2.Item(GetBallNumber(strLine)).Visible = True
            Image2.Item(GetBallNumber(strLine)).Tag = "GreenMoved"
        Next
            
        
        'after reading the game file, asign each line
        'properly to the form fields/labels
        If File.Line = 2 Then
            Label7.Caption = strLine
            CurrentGreenBall = CInt(Mid(strLine, 3, 1))
        ElseIf File.Line = 3 Then
            Label8.Caption = strLine
            CurrentRedBall = CInt(Mid(strLine, 3, 1))
        ElseIf File.Line = 4 Then
            Label1.Caption = strLine
            iCountGreen = Right(strLine, 1)
        ElseIf File.Line = 5 Then
            Label2.Caption = strLine
            iCountRed = Right(strLine, 1)
        ElseIf File.Line = 6 Then
            Label5.Caption = strLine
            GameNumber = Right(strLine, 1)
        ElseIf File.Line = 7 Then
            Label4.Caption = strLine
            GreenWins = Right(strLine, 1)
        ElseIf File.Line = 8 Then
            Label3.Caption = strLine
            RedWins = Right(strLine, 1)
        ElseIf File.Line = 9 Then
            GreenUser = strLine
        ElseIf File.Line = 10 Then
            RedUser = strLine
        End If
    
    Loop
    
    Me.Caption = "Classic Puzzle Game - " & UCase(GreenUser) & " <--Vs--> " & UCase(RedUser)
    Image4.Visible = True
    File.Close
    Set FSO = Nothing
    Set File = Nothing
    btnNew.Enabled = True
    btnConfess.Enabled = True
    
    If Check1.Value = 1 Then
        'start the game music
        PlayMIDI Song
    End If
    
End If

End Sub

'saves games in game's folder
Private Sub btnSave_Click()
Dim Value As String

Start:
Value = InputBox("Enter the file name (without extention)", "File Name")

If Len(Trim(Value)) > 0 Then
    Dim i As Integer
    Dim j As Integer
    Dim FSO As Object
    Dim File As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'file already exists?
    If FSO.FileExists(App.Path & "\" & Value & ".gam") Then
        MsgBox "A file with this name already exists!!" & vbNewLine & "Please enter a different name.", vbExclamation
        GoTo Start
    End If
    
    'create a new game file
    Set File = FSO.CreateTextFile(App.Path & "\" & Value & ".gam")
        
    'write these encrypted lines in the above file
    File.WriteLine ToggleText(Label7.Caption)
    File.WriteLine ToggleText(Label8.Caption)
    
    File.WriteLine ToggleText(Label1.Caption)
    File.WriteLine ToggleText(Label2.Caption)
    
    File.WriteLine ToggleText(Label5.Caption)
    File.WriteLine ToggleText(Label4.Caption)
    File.WriteLine ToggleText(Label3.Caption)
    
    File.WriteLine ToggleText(GreenUser)
    File.WriteLine ToggleText(RedUser)
        
    For i = 0 To Image2.Count - 1
        If Image2.Item(i).Visible = True Then
            'writes the location of each upper ball to the file
            File.WriteLine ToggleText(i & "=" & Image2.Item(i).Left & "." & Image2.Item(i).Top)
        End If
    Next
    'marks end of upper balls location
    File.WriteLine ToggleText("END")
    
    For j = 0 To Image3.Count - 1
        If Image3.Item(j).Visible = True Then
           'writes the location of each lower ball to the file
            File.WriteLine ToggleText(j & "=" & Image3.Item(j).Left & "." & Image3.Item(j).Top)
        End If
    Next
    'marks end of lower balls location
    File.WriteLine ToggleText("ENDAGAIN")
    File.Close
    Set File = Nothing
     
    'file saved
    If FSO.FileExists(App.Path & "\" & Value & ".gam") Then
        MsgBox "File saved successfully !!" & vbNewLine & App.Path & "\" & Value & ".gam", vbInformation
    End If
        
    'thanks FSO
    Set FSO = Nothing

End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    PlayMIDI Song
Else
    StopMIDI Song
End If
End Sub

Private Sub Command1_Click()
frmAbout.Show 1
End Sub

Private Sub Command2_Click()
If MsgBox("Are you sure you want to quit the game?", vbQuestion + vbYesNo) = vbYes Then
    'quit the game
    Unload Me
End If
End Sub

Private Sub btnHelp_Click()
frmHelp.Show 1
End Sub

Private Sub Command3_Click()
frmOpt.Show 1
End Sub

Private Sub Command4_Click()
On Error Resume Next
Me.WindowState = 1
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
If X < 2000 Or X > 11475 Or Y > 8000 Then Exit Sub

On Error Resume Next

 X = X - Source.Width / 2
 Y = Y - Source.Height / 2

If Check1.Value = 1 Then
    'play the game music
    PlayMIDI Song
End If

On Error Resume Next
'move the file to the dragged location
Source.Move X, Y
'play the moving sound
sndPlaySound App.Path + "\Move.wav", SND_ASYNC
btnConfess.Enabled = True
btnNew.Enabled = True
btnSave.Enabled = True
Image4.Visible = True

'in order to determine the player turn
If Source.Tag = "GreenMoved" Then
    Shape5.FillColor = &HC0C0FF
Else
    Shape5.FillColor = &HC0FFC0
End If


'a green ball has been moved from top-left to the game board
'so decrease the ball number by 1 from their total at top-left.
If CurrentGreenBall <> 0 Then
    If Source.Tag = "Green" Then
        CurrentGreenBall = CurrentGreenBall - 1
        Label7.Caption = "( " & CurrentGreenBall & " )"
        'green ball moved to board
        Source.Tag = "GreenMoved"
    End If
End If

'a red ball has been moved from top-left to the game board
'so decrease the ball number by 1 from their total at top-left.
If CurrentRedBall <> 0 Then
    If Source.Tag = "Red" Then
        CurrentRedBall = CurrentRedBall - 1
        Label8.Caption = "( " & CurrentRedBall & " )"
        'red ball moved to board
        Source.Tag = "RedMoved"
    End If
End If

End Sub

Private Sub Form_Load()

On Error Resume Next

'first game tobe started
GameNumber = 1
Label5.Caption = "Game # " & GameNumber

'fill these labels with values of these vars
Label1.Caption = GreenUser & " = " & iCountGreen
Label2.Caption = RedUser & " = " & iCountRed
Label4.Caption = GreenUser & " = " & GreenWins
Label3.Caption = RedUser & " = " & RedWins

Dim i As Integer
Dim j As Integer

'total ball needed for each of the two players 9
CurrentGreenBall = 9
CurrentRedBall = 9

'these registry keys determine the color the balls tobe used
If GetSetting(App.Title, "Settings", "RedBlue") = "True" Then
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image6.Picture
        Image2.Item(i).DragIcon = Image6.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image7.Picture
        Image3.Item(j).DragIcon = Image7.Picture
    Next
ElseIf GetSetting(App.Title, "Settings", "GreenBlue") = "True" Then
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image5.Picture
        Image2.Item(i).DragIcon = Image5.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image7.Picture
        Image3.Item(j).DragIcon = Image7.Picture
    Next
ElseIf GetSetting(App.Title, "Settings", "GreenYellow") = "True" Then
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image5.Picture
        Image2.Item(i).DragIcon = Image5.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image9.Picture
        Image3.Item(j).DragIcon = Image9.Picture
    Next
ElseIf GetSetting(App.Title, "Settings", "RedYellow") = "True" Then
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image6.Picture
        Image2.Item(i).DragIcon = Image6.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image9.Picture
        Image3.Item(j).DragIcon = Image9.Picture
    Next
ElseIf GetSetting(App.Title, "Settings", "YellowBlue") = "True" Then
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image9.Picture
        Image2.Item(i).DragIcon = Image9.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image7.Picture
        Image3.Item(j).DragIcon = Image7.Picture
    Next
Else
    'default ball colors
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Picture = Image5.Picture
        Image2.Item(i).DragIcon = Image5.Picture
    Next
    
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Picture = Image6.Picture
        Image3.Item(j).DragIcon = Image6.Picture
    Next
End If

'Set Text colors
frmMain.Label1.ForeColor = FirstColor
frmMain.Label4.ForeColor = FirstColor
frmMain.Label2.ForeColor = SecondColor
frmMain.Label3.ForeColor = SecondColor

'extract resource sound files to the game's folder
ExtractResourceFiles

'needed later to play the game music
Song = GetShortPath(App.Path & "\Music.mid")
'sets normal attribute of the music file
On Error Resume Next
SetAttr App.Path & "\Music.mid", vbNormal

'set the game title caption
Me.Caption = "Classic Puzzle Game - " & UCase(GreenUser) & " <--Vs--> " & UCase(RedUser)

'move the game name label to the bottom
Image8.Move 4600, 8500

If Image8.Left = 4600 And Image8.Top = 8500 Then
    Image8.Visible = True
End If

'put picture in the game board
If Len(GetSetting(App.Title, "Settings", "Picture")) > 0 Then
    Me.Picture = LoadPicture(GetSetting(App.Title, "Settings", "Picture"))
    Shape7.Visible = False
    Image8.Visible = False
    Shape9.Visible = False
End If

'board color
If Len(GetSetting(App.Title, "Settings", "BoardColor")) > 0 Then
    Shape7.FillColor = GetSetting(App.Title, "Settings", "BoardColor")
End If

'border color
If Len(GetSetting(App.Title, "Settings", "BorderColor")) > 0 Then
    Shape1.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Shape2.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Shape3.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Line1.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Line2.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Line3.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
    Line4.BorderColor = GetSetting(App.Title, "Settings", "BorderColor")
End If





'Opens gam type files if dragged or directly clicked
Dim sFileName As String

If Len(Command) > 0 Then
    If LCase(Right(Command, 3)) = "gam" Then
        sFileName = Command
    End If
End If

If Len(sFileName) > 0 Then
    NewGame
    Dim FSO As Object
    Dim File As Object
    Dim strLine As String
    Dim FileOneNumber As Integer
    Dim FileTwoNumber As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.OpenTextFile(sFileName)
    
    'hide all upper balls
    For i = 0 To Image2.Count - 1
        Image2.Item(i).Visible = False
    Next

    'hide all lower balls
    For j = 0 To Image3.Count - 1
        Image3.Item(j).Visible = False
    Next

    'read the file until end
    Do While File.AtEndOfStream = False
        'read each line and asign to strLine variable
        strLine = ToggleText(File.ReadLine)
        
        If strLine = "END" Then
        'total upper balls tobe placed at top-left corner
            FileOneNumber = CInt(File.Line) - 1 '-2 for total files visible
        ElseIf strLine = "ENDAGAIN" Then
        'total lower balls tobe placed at top-left corner
            FileTwoNumber = CInt(File.Line) - 1
        End If
       
On Error Resume Next
        
        'now move all available red balls to their original
        'location when game was saved by reading their
        'location from the game file
        For i = 1 To FileOneNumber
            Image3.Item(GetBallNumber(strLine)).Move GetLeft(strLine), GetTop(strLine)
            Image3.Item(GetBallNumber(strLine)).Visible = True
            'Red/lower ball moved to the board
            Image3.Item(GetBallNumber(strLine)).Tag = "RedMoved"
        Next

        For j = FileOneNumber To FileTwoNumber
            Image2.Item(GetBallNumber(strLine)).Move GetLeft(strLine), GetTop(strLine)
            Image2.Item(GetBallNumber(strLine)).Visible = True
            Image2.Item(GetBallNumber(strLine)).Tag = "GreenMoved"
        Next
            
        
        'after reading the game file, asign each line
        'properly to the form fields/labels
        If File.Line = 2 Then
            Label7.Caption = strLine
            CurrentGreenBall = CInt(Mid(strLine, 3, 1))
        ElseIf File.Line = 3 Then
            Label8.Caption = strLine
            CurrentRedBall = CInt(Mid(strLine, 3, 1))
        ElseIf File.Line = 4 Then
            Label1.Caption = strLine
            iCountGreen = Right(strLine, 1)
        ElseIf File.Line = 5 Then
            Label2.Caption = strLine
            iCountRed = Right(strLine, 1)
        ElseIf File.Line = 6 Then
            Label5.Caption = strLine
            GameNumber = Right(strLine, 1)
        ElseIf File.Line = 7 Then
            Label4.Caption = strLine
            GreenWins = Right(strLine, 1)
        ElseIf File.Line = 8 Then
            Label3.Caption = strLine
            RedWins = Right(strLine, 1)
        ElseIf File.Line = 9 Then
            GreenUser = strLine
        ElseIf File.Line = 10 Then
            RedUser = strLine
        End If
    
    Loop
    
    Me.Caption = "Classic Puzzle Game - " & UCase(GreenUser) & " <--Vs--> " & UCase(RedUser)
    Image4.Visible = True
    File.Close
    Set FSO = Nothing
    Set File = Nothing
    btnNew.Enabled = True
    btnConfess.Enabled = True
    
    If Check1.Value = 1 Then
        'start the game music
        PlayMIDI Song
    End If
    
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'determines which player wins more games upper or lower
If GreenWins <> 0 Or RedWins <> 0 Then
    If GreenWins > RedWins Then
        'stop game music
        StopMIDI Song
        'play winning sound
        sndPlaySound App.Path & "\Win.wav", SND_ASYNC
        MsgBox "Winner = " & GreenUser, vbInformation
    ElseIf GreenWins < RedWins Then
        StopMIDI Song
        sndPlaySound App.Path & "\Win.wav", SND_ASYNC
        MsgBox "Winner = " & RedUser, vbInformation
    ElseIf GreenWins = RedWins Then
        MsgBox "Both users tied at = " & RedWins, vbInformation
    End If
End If

'stop game music
StopMIDI Song
'quit the game
Unload Me
End Sub

Private Sub Image4_DragDrop(Source As Control, X As Single, Y As Single)
   'balls moved to the board and their calculation
   If Source.Tag = "RedMoved" Then
        iCountRed = iCountRed + 1
        Label2.Caption = RedUser & " = " & iCountRed
        Source.Visible = False
        GameInitiated = True
        On Error Resume Next
        sndPlaySound App.Path & "\Fire.wav", SND_ASYNC
        Timer2.Enabled = True
        Image1.Visible = True
    ElseIf Source.Tag = "GreenMoved" Then
        iCountGreen = iCountGreen + 1
        Label1.Caption = GreenUser & " = " & iCountGreen
        Source.Visible = False
        GameInitiated = True
        On Error Resume Next
        sndPlaySound App.Path & "\Fire.wav", SND_ASYNC
        Timer2.Enabled = True
        Image1.Visible = True
   End If
   
End Sub

'when two balls are visible of any player the game is
'won and lost.
Private Sub Timer1_Timer()
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim L As Integer

On Error Resume Next

If Me.Picture Then
    Image8.Visible = False
Else
    Image8.Visible = True
End If


For i = 0 To Image2.Count - 1
    If Image2.Item(i).Left = 120 And _
    Image2.Item(i).Top = 120 Then
        Image2.Item(i).Tag = "Green"
    End If
    
    'counts no. of upper visible balls
    If Image2.Item(i).Visible = True Then
        K = K + 1
'        Debug.Print K
    End If
    
Next

For j = 0 To Image3.Count - 1
    If Image3.Item(j).Left = 120 And _
    Image3.Item(j).Top = 720 Then
        Image3.Item(j).Tag = "Red"
    End If
    
    'counts no. of lower visible balls
    If Image3.Item(j).Visible = True Then
        L = L + 1
'        Debug.Print L
    End If
    
Next


'When seven balls become invisible, game is lost
If K = 2 Then
    If iCountGreen > iCountRed Then
        'again same code
        Image8.Visible = False
        Me.BackColor = Label2.ForeColor
        StopMIDI Song
        sndPlaySound App.Path & "\Win.wav", SND_ASYNC
        MsgBox RedUser & " wins the match !!", vbInformation
        Me.BackColor = &HFAEDE2
        Image8.Visible = True
        Timer1.Enabled = False
        RedWins = RedWins + 1
        Label3.Caption = RedUser & " = " & RedWins
        btnConfess.Enabled = False
        btnNew.Enabled = False
        GameInitiated = False
        Image4.Visible = False
        NewGame
        K = 0
    End If
End If

If L = 2 Then
    If iCountRed > iCountGreen Then
        Image8.Visible = False
        Me.BackColor = Label1.ForeColor
        StopMIDI Song
        sndPlaySound App.Path & "\Win.wav", SND_ASYNC
        MsgBox GreenUser & " wins the match !!", vbInformation
        Me.BackColor = &HFAEDE2
        Image8.Visible = True
        Timer1.Enabled = False
        GreenWins = GreenWins + 1
        Label4.Caption = GreenUser & " = " & GreenWins
        btnConfess.Enabled = False
        Image4.Visible = False
        btnNew.Enabled = False
        GameInitiated = False
        NewGame
        L = 0
    End If
End If

End Sub

'determines which player wins/not utilized in game though
Sub WhoWins()
    If iCountGreen > iCountRed Then
        MsgBox RedUser & " wins the match !!", vbInformation
    ElseIf iCountGreen = iCountRed Then
        MsgBox "Draw Match", vbExclamation
    Else
        MsgBox GreenUser & " wins the match !!", vbInformation
    End If
End Sub

Sub ExtractResourceFiles()
Dim FileWin() As Byte
Dim FileMove() As Byte
Dim FileKill() As Byte
Dim FileMusic() As Byte
Dim FF As Integer

FF = FreeFile

'read the resource file data
FileMove = LoadResData(102, "WAVE")

'write resouce file data
If Not FileExists(App.Path + "\Move.wav") Then
    Open App.Path + "\Move.wav" For Binary Access Write As #FF
        Put #FF, , FileMove
    Close #FF
End If

'read the resource file data
FileWin = LoadResData(101, "WAVE")

'write resouce file data
If Not FileExists(App.Path + "\Win.wav") Then
    Open App.Path + "\Win.wav" For Binary Access Write As #FF
        Put #FF, , FileWin
    Close #FF
End If

'read the resource file data
FileKill = LoadResData(103, "WAVE")

'write resouce file data
If Not FileExists(App.Path + "\Fire.wav") Then
    Open App.Path + "\Fire.wav" For Binary Access Write As #FF
        Put #FF, , FileKill
    Close #FF
End If

'read the resource file data
FileMusic = LoadResData(104, "CUSTOM")

'write resouce file data
If Not FileExists(App.Path + "\Music.mid") Then
    Open App.Path + "\Music.mid" For Binary Access Write As #FF
        Put #FF, , FileMusic
    Close #FF
End If

End Sub

' returns True when file exists, and False when not,
'   (returns True for directory also):
Public Function FileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
On Error GoTo NotFound
    
    i = GetAttr(sFileName)
    FileExists = True
    Exit Function

NotFound:
    FileExists = False
End Function

'moves/animates cloud picture on the form
Private Sub Timer2_Timer()
    If Image1.Top > 0 Then
        Image1.Move Image1.Left - 50, Image1.Top - 75
        If Image1.Top < -10 Then Me.Refresh
    Else
        Image1.Visible = False
        Image1.Top = 3360
        Image1.Left = 6720
        Timer2.Enabled = False
    End If
End Sub

'gets Left position of the ball
Function GetLeft(sLine As String) As String
If Len(sLine) > 0 Then
Dim Char As String
Dim i As Integer
    
    For i = 3 To Len(sLine)
        Char = Mid(sLine, i, 1)
        If Char = "." Then Exit For
        GetLeft = GetLeft & Char
    Next
End If
End Function

'gets Top position of the ball
Function GetTop(sLine As String) As String
If Len(sLine) > 0 Then
Dim Char As String
Dim i As Integer
 
    For i = 1 To Len(sLine)
        Char = Mid(sLine, i, 1)
        If Char = "." Then Exit For
        GetTop = Mid(sLine, i + 2, Len(sLine))
    Next
End If
End Function

'gets ball number
Function GetBallNumber(sLine As String) As String
If Len(sLine) > 0 Then
    GetBallNumber = Left(sLine, 1)
End If
End Function

'Toggles b/w encrypted and decrypted text
Function ToggleText(Text As String) As String
Dim n As Long
Dim CodedText As String

For n = 1 To Len(Text)
    CodedText = CodedText & Chr$(255 - Asc(Mid$(Text, n, 1)))
Next
ToggleText = CodedText
End Function

