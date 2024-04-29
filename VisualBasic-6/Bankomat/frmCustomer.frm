VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomer 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Area - [Customer Information]"
   ClientHeight    =   5160
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8745
   Icon            =   "frmCustomer.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   19
      Top             =   4770
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7250
            MinWidth        =   7250
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2893
            MinWidth        =   2893
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:26 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/20/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      Width           =   1335
   End
   Begin MSComctlLib.ImageList img 
      Left            =   8040
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":1DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":3AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":4922
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":65FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":6686
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomer.frx":783A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstVw 
      Height          =   3015
      Left            =   3960
      TabIndex        =   12
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Acc #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PIN #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3345
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PIN #"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acc #"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "List Of Customers  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer Information"
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
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8775
   End
   Begin VB.Menu mnu 
      Caption         =   "Process"
      Visible         =   0   'False
      Begin VB.Menu mnuEncDyc 
         Caption         =   "Decrypt PIN"
      End
      Begin VB.Menu gold 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatement 
         Caption         =   "Get Statement"
      End
      Begin VB.Menu brk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out"
      End
      Begin VB.Menu brk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu brk5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pItem As ListItem
Dim AccNo As Long
Dim Mode As String
Const PREFIX = "ATM"

Private Sub Command1_Click()
On Error GoTo ErrHandler
If Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty Then
    MsgBox "You must enter values for all the text boxes"
Else
    Set oCm = New Command
    Set oRs = New Recordset
    Connect
    With oCm
        .ActiveConnection = oCn
        .CommandType = adCmdText
        .CommandText = "insert into [customers] (PIN,Name,Address) values ('" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
        .Execute
        .CommandText = "select Max(AccountNo) from customers "
        Set oRs = .Execute
        If Not oRs.EOF Then
            AccNo = oRs.Fields(0).Value
        End If
       .CommandText = "insert into [Transaction]  values (" & AccNo & "," & Date & ", " & Val(Text5.Text) & "," & "'Deposit')"
        .Execute
        .CommandText = "insert into [Statement] (AccountNo, TransDate, Comment, Withdrawal,Deposit,Balance ) values (" & AccNo & "," & Date & ",'Deposit',0," & Val(Text5.Text) & "," & Val(Text5.Text) & ")"
        .Execute
    End With
    MsgBox "Record Inserted"
    Disconnect
    If mnuEncDyc.Caption = "Encrypt PIN" Then Load_Records (1) Else Load_Records (2)
 Exit Sub
ErrHandler:
'MsgBox Err.Number & Err.Source & Err.Description
MsgBox "Could not Insert Record"
Disconnect
End If
End Sub

Private Sub Load_Records(ByVal pEncDyc As Integer)
    Dim i As Integer, Enc As String
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
        lstVw.ListItems.Clear
        While Not .EOF
            i = i + 1
            lstVw.ListItems.Add i, , Format_Number(![AccountNo], PREFIX), , 1
            If pEncDyc = 1 Then
            Enc = ""
                For j = 1 To Len(![PIN])
                    Enc = Enc + "#"
                Next j
                    lstVw.ListItems(i).SubItems(1) = Enc
            Else
                lstVw.ListItems(i).SubItems(1) = ![PIN]
            End If
           
            lstVw.ListItems(i).SubItems(2) = ![name]
            lstVw.ListItems(i).SubItems(3) = ![Address]
            .MoveNext
        Wend
        .Close
    End With
    Set oCm = Nothing
    Set oRs = Nothing
    Disconnect
End Sub

Private Sub Command2_Click()
Dim ctr As Control
For Each ctr In Me.Controls
    If TypeOf ctr Is TextBox Then
        ctr.Text = Empty
    End If
Next ctr
End Sub

Private Sub Command3_Click()
    Mode = "Edit"
    lstVw_DblClick
End Sub

Private Sub Form_Load()
    Center Me
    Image1.Picture = img.ListImages(1).Picture
    mnuEncDyc_Click
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnuMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnuMenu
End Sub

Private Sub lstVw_DblClick()
    If Not pItem.Text = "" Then
        With pItem
            Text1.Text = .Text
            Text2.Text = .SubItems(1)
            Text3.Text = .SubItems(2)
            Text4.Text = .SubItems(3)
        End With
    End If
End Sub

Private Sub lstVw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set pItem = Item
    Command3.Enabled = True
End Sub

Private Sub lstVw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnu
End Sub

Private Sub mnuEncDyc_Click()
If mnuEncDyc.Caption = "Decrypt PIN" Then
    mnuEncDyc.Caption = "Encrypt PIN"
    Load_Records (2)
Else
    mnuEncDyc.Caption = "Decrypt PIN"
    Load_Records (1)
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnuMenu
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Mode = "Edit" Then Command3.Caption = "Update"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Mode = "Edit" Then Command3.Caption = "Update"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Mode = "Edit" Then Command3.Caption = "Update"
End Sub

