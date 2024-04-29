Attribute VB_Name = "modPuzzle"
Option Explicit
Public GreenUser As String
Public RedUser As String
Public FirstColor As Long
Public SecondColor As Long
Public RedWins As Integer          'no. of wins of lower user
Public GreenWins As Integer        'no. of wins of upper user
Public iCountRed As Integer        'no. of red balls/lower balls
Public iCountGreen  As Integer     'no. of green balls/upper balls
Public GameNumber As Integer       'no. of games played
Public GameInitiated As Boolean    'has first ball been played?
Public CurrentGreenBall As Integer 'current upper ball no.
Public CurrentRedBall As Integer   'current lower ball no.


Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000

Type CHOOSECOLOR
lStructSize As Long
hwndOwner As Long
hInstance As Long
rgbResult As Long
lpCustColors As String
flags As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String

'Syntax:
'varFileName = OpenDialog(Form1, Filter, Title, InitDir)

'Text1 = varFileName
  
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.flags = OFN_HIDEREADONLY + OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST
  a = GetOpenFileName(ofn)

  If (a) Then
      OpenDialog = Trim$(ofn.lpstrFile)
  Else
      OpenDialog = ""
  End If
End Function

Function ShowColor(ByVal uObject As Object)  'Creates the color selection dialog.
Dim cc As CHOOSECOLOR
Dim Custcolor(16) As Long
Dim lReturn As Long
Dim CustomColors() As Byte

cc.lStructSize = Len(cc)
cc.hwndOwner = uObject.hWnd
cc.hInstance = App.hInstance
cc.lpCustColors = StrConv(CustomColors, vbUnicode)
cc.flags = 3
If CHOOSECOLOR(cc) <> 0 Then
ShowColor = cc.rgbResult
CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
Else
ShowColor = -1
End If
End Function

Public Sub ColorSelect(ByVal uObject As Object)  'Color selection dialog
Dim NewColor As Long
NewColor = ShowColor(uObject)
If NewColor <> -1 Then
uObject.BackColor = NewColor
Else
MsgBox "No color has been selected.", vbInformation
End If
End Sub

'start the new game
Public Sub NewGame()
On Error Resume Next

    'set the vars and labels
    iCountRed = 0
    iCountGreen = 0
    
    frmMain.Label3.Caption = RedUser & " = " & RedWins
    frmMain.Label4.Caption = GreenUser & " = " & GreenWins
    frmMain.Label1.Caption = GreenUser & " = " & iCountGreen
    frmMain.Label2.Caption = RedUser & " = " & iCountRed
    
    Dim i As Integer
    Dim j As Integer
    
'determins ball colors tobe placed for new game
If GetSetting(App.Title, "Settings", "RedBlue") = "True" Then
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image6.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image6.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
ElseIf GetSetting(App.Title, "Settings", "GreenBlue") = "True" Then
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
ElseIf GetSetting(App.Title, "Settings", "GreenYellow") = "True" Then
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image9.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image9.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
ElseIf GetSetting(App.Title, "Settings", "RedYellow") = "True" Then
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image6.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image6.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image9.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image9.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
ElseIf GetSetting(App.Title, "Settings", "YellowBlue") = "True" Then
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image9.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image9.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image7.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image7.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
Else
    'default ball color:Green for upper user, Red for lower user
    For i = 0 To frmMain.Image2.Count - 1
        frmMain.Image2.Item(i).Picture = frmMain.Image5.Picture
        frmMain.Image2.Item(i).DragIcon = frmMain.Image5.Picture
        frmMain.Image2.Item(i).Visible = True
        frmMain.Image2.Item(i).Top = 120
        frmMain.Image2.Item(i).Left = 120
        frmMain.Image2.Item(i).Tag = "Green"
    Next
    
    For j = 0 To frmMain.Image3.Count - 1
        frmMain.Image3.Item(j).Picture = frmMain.Image6.Picture
        frmMain.Image3.Item(j).DragIcon = frmMain.Image6.Picture
        frmMain.Image3.Item(j).Visible = True
        frmMain.Image3.Item(j).Top = 720
        frmMain.Image3.Item(j).Left = 120
        frmMain.Image3.Item(j).Tag = "Red"
    Next
End If
    
    
    frmMain.Timer1.Enabled = True
    'one more game started
    GameNumber = GameNumber + 1
    frmMain.Label5.Caption = "Game # " & GameNumber
    GameInitiated = True
    frmMain.btnConfess.Enabled = False
    frmMain.btnSave.Enabled = False
    frmMain.btnNew.Enabled = False
    frmMain.Image4.Visible = False
    frmMain.Shape5.FillColor = &HFFC0C0
    
    frmMain.Label7.Caption = "( 9 )"
    frmMain.Label8.Caption = "( 9 )"
    CurrentGreenBall = 9
    CurrentRedBall = 9
    
    StopMIDI Song
    
    frmMain.Refresh
    
End Sub


