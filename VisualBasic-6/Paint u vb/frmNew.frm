VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Page color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         Caption         =   "C&ustom color"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1030
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Background color"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&White"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2775
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "500"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "500"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Pixels"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   880
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bground As Byte

Private Sub Command1_Click()

On Error Resume Next
Set frmMain = New frmMain
Load frmMain
frmMain.Show

If Len(Trim(Text1.Text)) >= 1 Then
    frmMain.Caption = Trim(Text1.Text)
Else
    frmMain.Caption = "Untitled"
End If
frmMain.picMain.Height = Val(Text2.Text) * 15.065913370998
frmMain.picMain.Width = Val(Text3.Text) * 15.065913370998
frmMain.picMain.Left = 0
frmMain.picMain.Top = 0
If Bground = 1 Then
    frmMain.picMain.BackColor = vbWhite
ElseIf Bground = 2 Then
    frmMain.picMain.BackColor = Bcolor
Else
    frmMain.picMain.BackColor = Picture1.BackColor
End If
    MDIForm1.Label4.Caption = frmMain.picMain.ScaleHeight & " X " & frmMain.picMain.ScaleWidth

Unload Me
Call Reset
Call Resetprop
IsOpen = False



frmMain.VScroll1.Left = frmMain.ScaleWidth - VScroll1.Width
If frmMain.HScroll1.Visible = False Then
    frmMain.VScroll1.Height = frmMain.ScaleHeight
    frmMain.Picture2.Visible = False

Else
    frmMain.VScroll1.Height = frmMain.ScaleHeight - 300
    frmMain.Picture2.Visible = True
End If
If frmMain.VScroll1.Visible = False Then
    frmMain.HScroll1.Width = frmMain.ScaleWidth
    frmMain.Picture2.Visible = False
Else
    frmMain.HScroll1.Width = frmMain.ScaleWidth - 300
    frmMain.Picture2.Visible = True
End If
frmMain.HScroll1.Top = frmMain.ScaleHeight - frmMain.HScroll1.Height
frmMain.Picture2.Left = frmMain.ScaleWidth - 300
frmMain.Picture2.Top = frmMain.ScaleHeight - 300


If frmMain.VScroll1.Visible = False Then
    If frmMain.ScaleWidth > frmMain.picMain.Width Then
        frmMain.HScroll1.Visible = False
    Else
        frmMain.HScroll1.Visible = True
    End If
Else
    If (frmMain.ScaleWidth - 300) > frmMain.picMain.Width Then
        frmMain.HScroll1.Visible = False
    Else
        frmMain.HScroll1.Visible = True
    End If
End If

If frmMain.HScroll1.Visible = False Then
    If frmMain.ScaleHeight > frmMain.picMain.Height Then
        frmMain.VScroll1.Visible = False
    Else
        frmMain.VScroll1.Visible = True
    End If
Else
    If (frmMain.ScaleHeight - 300) > frmMain.picMain.Height Then
        frmMain.VScroll1.Visible = False
    Else
        frmMain.VScroll1.Visible = True
    End If
End If

If frmMain.ScaleWidth > frmMain.picMain.Width Then
    If frmMain.VScroll1.Visible = False Then
        frmMain.picMain.Left = (frmMain.ScaleWidth / 2) - (frmMain.picMain.Width / 2)
    Else
        frmMain.picMain.Left = ((frmMain.ScaleWidth - 300) / 2) - (frmMain.picMain.Width / 2)
    End If
End If
If frmMain.ScaleHeight > frmMain.picMain.Height Then
    If frmMain.HScroll1.Visible = False Then
        frmMain.picMain.Top = (frmMain.ScaleHeight / 2) - (frmMain.picMain.Height / 2)
    Else
        frmMain.picMain.Top = ((frmMain.ScaleHeight - 300) / 2) - (frmMain.picMain.Height / 2)
    End If
End If
If frmMain.VScroll1.Visible = False Or frmMain.HScroll1.Visible = False Then
frmMain.VScroll1.Max = (frmMain.picMain.Height - frmMain.ScaleHeight)
frmMain.HScroll1.Max = (frmMain.picMain.Width - frmMain.ScaleWidth)
Else
frmMain.VScroll1.Max = (frmMain.picMain.Height - frmMain.ScaleHeight) + frmMain.HScroll1.Height
frmMain.HScroll1.Max = (frmMain.picMain.Width - frmMain.ScaleWidth) + frmMain.VScroll1.Width
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Form_Load()

MDIForm1.Toolbar2.Buttons(2).Enabled = False
MDIForm1.mnuOpen.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Toolbar2.Buttons(2).Enabled = True
MDIForm1.mnuOpen.Enabled = True
End Sub

Private Sub Option1_Click()
Bground = 1
Picture1.Enabled = False
End Sub

Private Sub Option2_Click()
Bground = 2
Picture1.Enabled = False
End Sub

Private Sub Option3_Click()
Bground = 3
Picture1.Enabled = True
End Sub

Private Sub Picture1_Click()
On Error GoTo Traperr
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.Color
Traperr:
Exit Sub
End Sub
