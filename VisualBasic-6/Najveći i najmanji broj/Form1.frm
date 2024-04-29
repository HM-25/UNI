VERSION 5.00
Begin VB.Form Form1  
   BackColor       =   &H00FFFF00&
   Caption         =   "Najveæi i najmanji broj"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIzadi 
      Caption         =   "&Izaði"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdIzbrisi 
      Caption         =   "Izbriši"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdIzracunaj 
      Caption         =   "Izraèunaj"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Najmanji broj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Najveæi broj "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdIzadi_Click()
Unload Me
End Sub

Private Sub cmdIzbrisi_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub cmdIzracunaj_Click()
 Dim i, s, l, n(9) As Integer
    For i = 0 To 9
        n(i) = InputBox("Enter Number", "Integer Needed")
    Next i
    l = n(0)
    s = n(0)
    For i = 0 To 9
        If n(i) > l Then
            l = n(i)
        End If
        If n(i) < s Then
            s = n(i)
        End If
    Next i
    Text1.Text = l
    Text2.Text = s
End Sub

