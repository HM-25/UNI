VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color changer"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmCol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1260
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   300
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim Color As Long
Dim r1 As Long
Dim g1 As Long
Dim b1 As Long
Dim c As Long
frmMain.picMain.Cls
If Sel = True Then
    Sel = False

    For i = FX + 1 To LX - 1
        For j = FY + 1 To LY - 1
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            r = Abs(r + Slider1.value)
            g = Abs(g + Slider2.value)
            b = Abs(b + Slider3.value)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
         Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
    frmMain.picMain.Refresh
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            r = Abs(r + Slider1.value)
            g = Abs(g + Slider2.value)
            b = Abs(b + Slider3.value)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
         Next
    Bar.Max = frmMain.picMain.ScaleWidth
    Bar.value = Bar.value + 1
    Next
    frmMain.picMain.Refresh

End If
Bar.value = 0
Call Setpicture
Save = MDIForm1.picSave.UBound
Undo = Undo + 1
If IsUndo = True Then
    IsUndo = False
    For s = Undo To Save
    Unload MDIForm1.picSave(s)
    Next
Save = MDIForm1.picSave.UBound
End If

MDIForm1.Toolbar2.Buttons(4).Enabled = True
Load MDIForm1.picSave(Save + 1)
MDIForm1.picSave(Save + 1).Picture = frmMain.picMain.Picture
MDIForm1.Toolbar2.Buttons(5).Enabled = False
MDIForm1.mnuBackward.Enabled = True
MDIForm1.mnuForward.Enabled = False

End Sub

Private Sub Command2_Click()

On Error Resume Next
Dim Color As Long
Dim r1 As Long
Dim g1 As Long
Dim b1 As Long
Dim c As Long
frmMain.picMain.Cls
If Sel = True Then
    Sel = False

    For i = FX + 1 To LX - 1
        For j = FY + 1 To LY - 1
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            r = Abs(r + Slider1.value)
            g = Abs(g + Slider2.value)
            b = Abs(b + Slider3.value)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
         Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
    frmMain.picMain.Refresh
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            r = Abs(r + Slider1.value)
            g = Abs(g + Slider2.value)
            b = Abs(b + Slider3.value)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
         Next
    Bar.Max = frmMain.picMain.ScaleWidth
    Bar.value = Bar.value + 1
    Next
    frmMain.picMain.Refresh

End If
Bar.value = 0
Call Setpicture
Unload Me
Save = MDIForm1.picSave.UBound
Undo = Undo + 1
If IsUndo = True Then
    IsUndo = False
    For s = Undo To Save
    Unload MDIForm1.picSave(s)
    Next
Save = MDIForm1.picSave.UBound
End If

MDIForm1.Toolbar2.Buttons(4).Enabled = True
Load MDIForm1.picSave(Save + 1)
MDIForm1.picSave(Save + 1).Picture = frmMain.picMain.Picture
MDIForm1.Toolbar2.Buttons(5).Enabled = False
MDIForm1.mnuBackward.Enabled = True
MDIForm1.mnuForward.Enabled = False


End Sub
