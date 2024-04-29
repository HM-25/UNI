VERSION 5.00
Begin VB.Form frmEncrypt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrpyt"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   2295
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   -360
      X2              =   8040
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the word(s) to be encrypted."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public a As Boolean

Private Sub cmdClear_Click()
    txtEncrypt.Text = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEncrypt_Click()
    
    If txtEncrypt.Text = "" Then
        MsgBox "Please enter a string", 48
        Exit Sub
    End If
    
    a = False
    frmDecrypt.Show
    frmDecrypt.txtDecrypt.Text = ""
    
    Dim i As Integer
    Dim j As String
    
    For i = 1 To Len(txtEncrypt.Text) 'take each letter for the string to encrypt
        'changing each character to keyascii code
        j = (Asc(Mid(txtEncrypt.Text, i, 1)) + 2) 'moving 2 keyascii codes to the right (transposition)
        frmDecrypt.txtDecrypt.Text = frmDecrypt.txtDecrypt.Text + Chr(j) 'converting the keycodes to character
    Next i
End Sub

Public Sub VButton(bVal As Boolean)
    cmdEncrypt.Visible = bVal
    cmdClear.Visible = bVal
End Sub

Private Sub Form_Load()
    If a = True Then
        VButton True
        Label1.Caption = "Type the word(s) to be encrypted "
    Else
        VButton False
        Label1.Caption = "The decrypted word(s) are "
    End If
End Sub
