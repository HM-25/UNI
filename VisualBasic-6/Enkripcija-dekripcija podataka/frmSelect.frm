VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption/Decryption"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select an option"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Decryption"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   8400
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    If Option1.Value = True Then
        frmEncrypt.a = True
        frmEncrypt.Show
    Else
        frmDecrypt.b = True
        frmDecrypt.Show
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
