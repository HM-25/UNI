VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransaction 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATM 24 Hour Service [Transaction Form]"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3825
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2400
         Top             =   3000
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   5145
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton Command5 
            Caption         =   " PIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Enter"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   24
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Statement"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Deposit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Withdraw"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1215
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   0
            ScaleHeight     =   2175
            ScaleWidth      =   5175
            TabIndex        =   8
            Top             =   840
            Width           =   5175
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4935
               TabIndex        =   10
               Top             =   -120
               Width           =   4935
               Begin VB.ComboBox Combo1 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   1440
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   240
                  Width           =   3495
               End
               Begin VB.Label Label2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Fast Cash"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   12
                  Top             =   360
                  Width           =   1455
               End
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   9
               Top             =   1320
               Width           =   4935
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Enter Amount"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   960
               Width           =   4935
            End
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2025
            ScaleWidth      =   4905
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   4935
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   1080
               TabIndex        =   21
               Top             =   1560
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   20709377
               CurrentDate     =   38472
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   1080
               TabIndex        =   20
               Top             =   1080
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   20709377
               CurrentDate     =   38472
            End
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Range"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   2895
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ALL"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   18
               Top             =   120
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   360
               TabIndex        =   23
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "From"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   360
               TabIndex        =   22
               Top             =   1080
               Width           =   735
            End
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2025
            ScaleWidth      =   4905
            TabIndex        =   26
            Top             =   840
            Visible         =   0   'False
            Width           =   4935
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   1920
               PasswordChar    =   "#"
               TabIndex        =   29
               Top             =   1200
               Width           =   2775
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   1920
               PasswordChar    =   "#"
               TabIndex        =   28
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "New PIN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   600
               TabIndex        =   31
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Current PIN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Change PIN   "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   4935
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   4905
         TabIndex        =   1
         Top             =   1440
         Width           =   4935
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1800
            MaxLength       =   4
            PasswordChar    =   "#"
            TabIndex        =   4
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Account #"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   5175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Authentication Form"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Balance As Double           'Holds last balance extracted from database
Dim Trans2day As Integer       'Holds transaction of the current day for the account in case
Dim Withdraw2day As Double  'Holds the total amount withdrawn in the current day by the account in case
Dim Trans As Integer              'Specifies which transaction (Withdrawal,Deposit,Statement,PIN-[1,2,3,4])
Dim AccNo As Long                'Holds Account number
Dim rptTyp As Integer            'Specifies report type to be generated.[All or Range]-[1 or 2]
Const PREFIX = "ATM"            'Prefixes auto generated accound number

Private Sub Combo1_Click()
    Text3.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandler
Trans2day = TransToday
Set oCm = New Command
Set oRs = New Recordset
Connect
With oCm
    .ActiveConnection = oCn
    .CommandType = adCmdText
    
    If Text3.Text = Empty And (Trans <> 3 And Trans <> 4) Then
        MsgBox "You must enter an amount"
    Else
        If Val(Text3.Text) > 1000000 And (Trans <> 2 And Trans <> 3 And Trans <> 4) Then
            MsgBox "Sorry, You Cannot Withdraw more than 1,000,000.00 in a day" & vbCrLf & " Try again tomorrow", vbInformation
            GoTo DisCon
        End If
        If (Withdraw2day + Val(Text3.Text) > 1000000 Or Withdraw2day = 1000000) And (Trans <> 2 And Trans <> 3 And Trans <> 4) Then
            MsgBox "Sorry, You Cannot Withdraw more than 1,000,000.00 in a day" & vbCrLf & " Try again tomorrow", vbInformation
            GoTo DisCon
        End If
        If Trans2day = 3 And (Trans <> 3 And Trans <> 2 And Trans <> 4) Then
            MsgBox "Sorry, You have exhausted your daily transactions" & vbCrLf & "Try again tomorrow", vbInformation
        Else
                If Trans = 1 Then
                     .CommandText = "insert into [Transaction]  values (" & AccNo & "," & Date & ", " & Val(Text3.Text) & "," & "'Withdrawal')"
                     .Execute
                     .CommandText = "insert into [Statement] (AccountNo, TransDate, Comment, Withdrawal,Deposit,Balance ) values (" & AccNo & "," & Date & ",'Withdrawal + B/C'," & Val(Text3.Text) & ",0," & Balance - Val(Text3.Text) - 2000 & ")"
                     .Execute
                ElseIf Trans = 2 Then
                     .CommandText = "insert into [Transaction]  values (" & AccNo & "," & Date & ", " & Val(Text3.Text) & "," & "'Deposit')"
                     .Execute
                     .CommandText = "insert into [Statement] (AccountNo, TransDate, Comment, Withdrawal,Deposit,Balance ) values (" & AccNo & "," & Date & ",'Deposit',0," & Val(Text3.Text) & "," & Balance + Val(Text3.Text) & ")"
                     .Execute
                ElseIf Trans = 3 Then
                Disconnect
                      If Option1.Value = True Then
                            rptStatement.Generate_Report AccNo, 1
                      ElseIf Option2.Value = True Then
                            rptStatement.Generate_Report AccNo, 2, DTPicker1.Value, DTPicker2.Value
                      Else
                            MsgBox "Select an option to Generate Report"
                      End If
                      Command2_Click
                ElseIf Trans = 4 Then
                      If Text4.Text <> Text2.Text Then
                            MsgBox "Invalid PIN", vbInformation, "Invalid"
                            Text4.Text = Empty: Text5.Text = Empty
                      Else
                            .CommandText = "update [customers] set PIN='" & Text5.Text & "' where AccountNo=" & Rip_Number(Text1.Text)
                            .Execute
                            MsgBox "PIN Change Successful. Please keep your PIN secrete." & vbCrLf & "Incase forgotten, contact bankers", vbInformation, "PIN Changed"
                      End If
                      Command2_Click
                End If
            End If
        End If
    End With
DisCon:
    If Not (rptStatement.Visible = True) Then Disconnect
     Exit Sub
ErrHandler:
MsgBox Err.Description & " " & Err.Source
    If (Trans = 1 And Trans = 2) Then
        MsgBox "Could not Insert Record"
    ElseIf Trans = 4 Then
        MsgBox "Unable to Change PIN", vbInformation, "Error"
    Else
        MsgBox "Report Cannot be generated at this time"
    End If
    Disconnect
End Sub

Private Sub Command2_Click()
    Picture4.Visible = True
    Picture5.Visible = True
    Picture6.Visible = False
    Picture7.Visible = False
    Trans = 1
End Sub

Private Sub Command3_Click()
    Picture4.Visible = True
    Picture5.Visible = True
    Picture6.Visible = False
    Picture7.Visible = False
    Trans = 2
End Sub

Private Sub Command4_Click()
    Picture4.Visible = False
    Picture5.Visible = False
    Picture6.Visible = True
    Picture7.Visible = False
    Trans = 3
End Sub

Private Sub Command5_Click()
    Picture4.Visible = False
    Picture5.Visible = False
    Picture6.Visible = False
    Picture7.Visible = True
    Trans = 4
End Sub

Private Sub Form_Load()
    Center Me
    Label5.Caption = Format(Now, "Mmm dd,yyyy    hh:mm:ss ampm")
    Timer1.Enabled = True
End Sub

Private Sub Option1_Click()
rptTyp = 1
End Sub

Private Sub Option2_Click()
rptTyp = 2
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
'All these are absurd. Just making it work. No optmization now
    If KeyAscii = 13 Then
    If Text1.Text = "admin" Then
         Text2.SetFocus
        Text2.ZOrder (0)
        Label1.Caption = "PIN #"
        Exit Sub
    End If
        If UCase(Left(Text1.Text, 3)) <> PREFIX Then
            MsgBox "Sorry, Your Account Number is not right", vbInformation, "Invalid #"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Exit Sub
        End If
        If UCase(Left(Right(Text1.Text, 5), 1)) <> "M" And Val(Right(Text1.Text, 4)) < 1000 Then
            MsgBox "Sorry, Your Account Number is not right", vbInformation, "Invalid #"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Exit Sub
        End If
        Text2.SetFocus
        Text2.ZOrder (0)
        Label1.Caption = "PIN #"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'Use this for admin
    '********************ABSURD*******************
    If Text1.Text = "admin" And Text2.Text = "9999" Then
        frmCustomer.Show
        Unload Me
        Exit Sub
    End If
    '*********************************************
    
    Dim i As Integer
    Set oCm = New ADODB.Command
    Set oRs = New ADODB.Recordset
    Connect
    With oCm
        .ActiveConnection = oCn
        .CommandType = adCmdText
        .CommandText = "select * from customers"
        Set oRs = .Execute
    End With
    With oRs
        While Not .EOF
        AccNo = Rip_Number(Text1.Text)
            If AccNo = ![AccountNo] And Text2.Text = ![PIN] Then
                Picture2.Visible = False
                Picture3.Visible = True
                GoTo Free
            End If
            .MoveNext
        Wend
        MsgBox "Invalid Entries!" & vbCrLf & "Please Check your Acc# or Pin#", vbCritical, "Authentication Invalid"
        Text1.Text = Empty
        Text2.Text = Empty
        Text1.SetFocus
        Text1.ZOrder (0)
        Label1.Caption = "Account #"
    End With
Free:
    oRs.Close
    Set oCm = Nothing
    Set oRs = Nothing
    Disconnect
    Get_Balance
Else
    Validate_Numeric KeyAscii
End If
End Sub

Private Sub Fill_Fast(ByVal pMax As Double)
Dim ToVal As Long
    With Combo1
        If pMax > 100000 And pMax < 1000000 Then
            ToVal = pMax
        ElseIf pMax > 100000 And pMax > 1000000 Then
            ToVal = 1000000
        End If
        .Clear
        For i = 100000 To ToVal Step 100000
            .AddItem i
        Next i
        If (ToVal \ i) >= 1 Then Combo1.ListIndex = 0
    End With
End Sub

Private Sub Get_Balance()
'This part is also absurd
    Set oCm = New ADODB.Command
    Set oRs = New ADODB.Recordset
    Connect
    With oCm
        .ActiveConnection = oCn
        .CommandType = adCmdText
        .CommandText = "select balance from statement where ReferenceID=(select Max(ReferenceID) from statement where AccountNo=" & AccNo & ") and AccountNo=" & AccNo
        Set oRs = .Execute
    End With
    With oRs
        If Not .EOF Then
             Balance = ![Balance]
        End If
        Fill_Fast (Balance)
        .Close
    End With
    Set oCm = Nothing
    Set oRs = Nothing
    Disconnect
End Sub

Private Function TransToday() As Long
'This is the most absurd part
    Set oCm = New ADODB.Command
    Set oRs = New ADODB.Recordset
    Connect
    With oCm
        .ActiveConnection = oCn
        .CommandType = adCmdText
        '.CommandText = "select count(*) from [transaction] where AccountNo=" & Text1.Text & " and TransDate=(datevalue(" & Date & ")"
        .CommandText = "select * from [transaction] where AccountNo=" & AccNo & " and TransDate=" & Date
        Set oRs = .Execute
    End With
    With oRs
        Withdraw2day = 0
        While Not .EOF
            If ![TransactionType] = "Withdrawal" Then
                Withdraw2day = Withdraw2day + ![Amount]
            End If
            i = i + 1
            .MoveNext
        Wend
        TransToday = i
        .Close
    End With
    Set oCm = Nothing
    Set oRs = Nothing
    Disconnect
End Function



Private Sub Timer1_Timer()
    Label5.Caption = Format(Now, "Mmm dd,yyyy    hh:mm:ss ampm")
End Sub

Private Sub Greet(ByVal i As Integer)
'This is the most absurd
If Format(Now, "ampm") = "am" Then
    Label11.Caption = "Good Morning, " & ""
ElseIf Format(Now, "ampm") = "pm" And Val(Format(Now, "hh")) < 5 Then
    Label11.Caption = "Good Afternoon, " & ""
Else
    Label11.Caption = "Good Evening, " & ""
End If
End Sub

