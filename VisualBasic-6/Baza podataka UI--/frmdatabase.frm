VERSION 5.00
Begin VB.Form frmdatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Sample"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Caption         =   "?"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ">>"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<<"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AddNew"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim db As Database
Private Sub Command1_Click()
Call Clear_All_Text
rs.AddNew
Call Get_Text
End Sub

Private Sub Command2_Click()
rs.Delete
rs.MoveNext
Call Fill_Text
End Sub

Private Sub Command3_Click()
Call Get_Text
rs.Update
End Sub

Private Sub Command4_Click()
rs.Edit
End Sub

Private Sub Command5_Click()
On Error Resume Next
rs.MovePrevious
Call Fill_Text
End Sub

Private Sub Command6_Click()
On Error Resume Next
rs.MoveNext
Call Fill_Text
End Sub

Private Sub Command7_Click()
On Error Resume Next
rs.MoveFirst
Call Fill_Text
End Sub

Private Sub Command8_Click()
On Error Resume Next
rs.MoveLast
Call Fill_Text
End Sub

Private Sub Command9_Click()
frmabout.Show
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase("Contacts.mdb") 'The Database Path
    Set rs = db.OpenRecordset("Info", dbOpenTable)
    Call Fill_Text
End Sub

Public Sub Fill_Text()
On Error Resume Next
Text1.Text = rs.Fields("Name")
Text2.Text = rs.Fields("Address")
Text3.Text = rs.Fields("Phone Number")
End Sub

Public Sub Get_Text()
rs("Name") = Text1.Text
rs("Address") = Text2.Text
rs("Phone Number") = Text3.Text
End Sub

Public Sub Clear_All_Text()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub
