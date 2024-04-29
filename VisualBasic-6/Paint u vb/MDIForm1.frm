VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   " "
   ClientHeight    =   8550
   ClientLeft      =   1950
   ClientTop       =   180
   ClientWidth     =   9975
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture7 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9975
      TabIndex        =   86
      Top             =   8070
      Width           =   9975
      Begin MSComctlLib.ProgressBar Bar 
         Height          =   300
         Left            =   120
         TabIndex        =   98
         Top             =   90
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   7200
         TabIndex        =   87
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open (Ctrl+O)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save as (Ctrl+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Step backward (Ctrl+Z)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Step forward (Ctrl+A)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cut (Ctrl+X)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Copy (Ctrl+C)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Paste (Ctrl+V)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete (Delete)"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      ForeColor       =   &H80000008&
      Height          =   7710
      Left            =   7065
      ScaleHeight     =   7680
      ScaleWidth      =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2910
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         Caption         =   "Brush options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   88
         Top             =   5640
         Visible         =   0   'False
         Width           =   2655
         Begin VB.PictureBox PicCopy 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   375
            Left            =   120
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   21
            TabIndex        =   99
            Top             =   2880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H8000000E&
            Caption         =   "Color mix"
            Height          =   255
            Left            =   720
            TabIndex        =   97
            Top             =   3000
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000009&
            Caption         =   "Blue channel"
            Height          =   195
            Left            =   720
            TabIndex        =   96
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txtBrushwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   94
            Text            =   "1"
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000009&
            Caption         =   "Green channel"
            Height          =   195
            Left            =   720
            TabIndex        =   92
            Top             =   2280
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000009&
            Caption         =   "Red channel"
            Height          =   195
            Left            =   720
            TabIndex        =   91
            Top             =   1920
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000009&
            Caption         =   "Grayscale"
            Height          =   195
            Left            =   720
            TabIndex        =   90
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000009&
            Caption         =   "Normal brush"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000009&
            Caption         =   "Artistic brushes"
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
            TabIndex        =   95
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000009&
            Caption         =   "Draw width"
            Height          =   255
            Left            =   720
            TabIndex        =   93
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.PictureBox picCrop 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   375
         Left            =   600
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   85
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picSave 
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   78
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   62
         Top             =   5640
         Width           =   2655
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Low quality"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   77
            Top             =   1920
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Normal"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   76
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Heigh quality"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   75
            Top             =   1440
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "MDIForm1.frx":0442
            Left            =   960
            List            =   "MDIForm1.frx":045E
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "MDIForm1.frx":04D1
            Left            =   960
            List            =   "MDIForm1.frx":04E4
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtEraserwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   70
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtFCirwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   69
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtFRectwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   68
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtCirwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   67
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtrectwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   66
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtLinewidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   65
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtPenwidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   63
            Text            =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fill style"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Draw style"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Draw width"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custom colors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   2655
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   111
            TabIndex        =   60
            Top             =   2085
            Width           =   1695
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2000
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   59
            ToolTipText     =   "Draw color"
            Top             =   1560
            Width           =   495
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2000
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   58
            ToolTipText     =   "Background color"
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2000
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   57
            ToolTipText     =   "Fill color"
            Top             =   360
            Width           =   495
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   120
            Picture         =   "MDIForm1.frx":0512
            ScaleHeight     =   111
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   111
            TabIndex        =   56
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "B:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   84
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "G:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   83
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "R:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "128"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   81
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "128"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   80
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "128"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   79
            Top             =   2520
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   525
            Left            =   1980
            Top             =   340
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Basic colors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   47
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   51
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00400040&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   50
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   46
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   49
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   45
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   44
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   47
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   43
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   46
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   42
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   45
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   41
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   44
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   39
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   42
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   38
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   41
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   37
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   40
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   36
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   39
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   35
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   38
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   34
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   37
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   33
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   36
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   32
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   35
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   34
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   33
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   29
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   32
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   28
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   31
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   27
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   30
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   29
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   28
            Top             =   900
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   26
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   25
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   24
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   23
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   22
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   21
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C000C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   20
            Top             =   1170
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   19
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   18
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   17
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   16
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   15
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   14
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   1860
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   13
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00800080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2130
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   12
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   11
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000040&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   10
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   780
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   9
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00004040&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1050
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   8
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1320
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   7
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1590
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   6
            Top             =   1710
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   510
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   5
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   40
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   43
            Top             =   630
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   240
            MousePointer    =   99  'Custom
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   27
            Top             =   1170
            Width           =   255
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FF0000&
            Height          =   278
            Left            =   230
            Top             =   340
            Width           =   278
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7710
      Left            =   0
      ScaleHeight     =   7680
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   360
      Width           =   765
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   120
         Top             =   7800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   23
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":35A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3C6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":432C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":49EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":50B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":5772
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":5E34
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":6512
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":6BF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":72CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":79AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":806E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8730
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   120
         Top             =   8640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Jpeg (*.jpg)|*.jpg|Bitmap (*.bmp)|*.bmp"
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   120
         Top             =   8640
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
               Picture         =   "MDIForm1.frx":8DF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8F04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9016
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9128
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":923A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":934C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":945E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9570
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9682
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   8640
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   23
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9794
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9E56
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":A518
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":ABDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":B29C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":B95E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":C020
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":C6E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":CDA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":D466
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":DB28
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":E1EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":E8C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":EFA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":F684
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   8640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "All files (*.*)|*.*"
      End
      Begin VB.PictureBox picFillcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   52
         ToolTipText     =   "Fill color"
         Top             =   6240
         Width           =   375
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   5460
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   9631
         ButtonWidth     =   794
         ButtonHeight    =   741
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pencil"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Line"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rectangle"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Square"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Filled Rectangle"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Filled Square"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eraser"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Brush"
               ImageIndex      =   11
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Crop"
               ImageIndex      =   12
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eye Dropper"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paint Bucket"
               ImageIndex      =   9
               Style           =   2
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Clone Stamp"
               ImageIndex      =   10
               Style           =   2
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Select"
               ImageIndex      =   13
               Style           =   2
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.PictureBox picDrawcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   53
         ToolTipText     =   "Draw color"
         Top             =   5880
         Width           =   495
      End
      Begin VB.PictureBox picBackcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   54
         ToolTipText     =   "Background color"
         Top             =   6360
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   720
         Y1              =   5760
         Y2              =   5760
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save as..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuBackward 
         Caption         =   "&Step backward"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuForward 
         Caption         =   "S&tep forward"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStandard 
         Caption         =   "&Standard tool bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Tool box"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOption 
         Caption         =   "&Option palette"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "St&atus bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "&Effects"
      Begin VB.Menu mnuGrayscale 
         Caption         =   "&Grayscale"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBrightness 
         Caption         =   "&Brightness"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDarkness 
         Caption         =   "&Darkness"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "&Emboss"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBnW 
         Caption         =   "&Black && White"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChannels 
         Caption         =   "&Channels"
         Begin VB.Menu mnuRed 
            Caption         =   "&Red channel"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "&Green channel"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue channel"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuQuality 
         Caption         =   "&Negative"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGlow 
         Caption         =   "&Glow"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
frmMain.picMain.DrawStyle = Combo1.ListIndex

End Sub

Private Sub Combo2_Click()
frmMain.picMain.FillStyle = Combo2.ListIndex
End Sub




Private Sub MDIForm_Load()
Curtool.Pencil = True
Dwidth = 1
Fcol = True
DColor = MDIForm1.picDrawcolor.BackColor
Fcolor = MDIForm1.picFillcolor.BackColor
Bcolor = MDIForm1.picBackcolor.BackColor
MDIForm1.Picture6.BackColor = MDIForm1.picDrawcolor.BackColor
MDIForm1.Picture4.BackColor = MDIForm1.picFillcolor.BackColor
MDIForm1.Picture5.BackColor = MDIForm1.picBackcolor.BackColor
Picture11.ScaleWidth = 256
For a = 0 To 255
    Picture11.Line (a, 0)-(a, Picture11.Height), RGB(a, a, a)

Next
Picture11.Picture = Picture11.Image
End Sub

Private Sub mnuBackward_Click()
    Undo = Undo - 1
    IsUndo = True
    Toolbar2.Buttons(5).Enabled = True
    mnuForward.Enabled = True
    If Undo = 0 Then
        Toolbar2.Buttons(4).Enabled = False
        mnuBackward.Enabled = False
        If IsOpen = True Then
            frmMain.picMain.Picture = frmMain.Picture1.Picture
            Exit Sub
        End If
    End If
    frmMain.picMain.Picture = MDIForm1.picSave(Undo).Picture

End Sub

Private Sub mnuBlue_Click()
On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            SetPixel frmMain.picMain.hdc, i, j, RGB(0, 0, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            SetPixel frmMain.picMain.hdc, i, j, RGB(0, 0, b)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub

Private Sub mnuBnW_Click()

On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            If r < 200 And g < 200 And b < 200 Then
                Color = vbBlack
            Else
                Color = vbWhite
            End If
            SetPixel frmMain.picMain.hdc, i, j, Color
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            If r < 200 And g < 200 And b < 200 Then
                Color = vbBlack
            Else
                Color = vbWhite
            End If
            SetPixel frmMain.picMain.hdc, i, j, Color
        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False


End Sub

Private Sub mnuBrightness_Click()

On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            c = Abs((r + g + b) \ 3)
            c = c \ 10
            r = Abs(r + c)
            g = Abs(g + c)
            b = Abs(b + c)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            c = Abs((r + g + b) \ 3)
            c = c \ 10
            r = Abs(r + c)
            g = Abs(g + c)
            b = Abs(b + c)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False


End Sub

Private Sub mnuColor_Click()
Form1.Show
End Sub

Private Sub mnuCopy_Click()
If Sel = True Then
    frmMain.PicCopy.Height = LY - FY
    frmMain.PicCopy.Width = LX - FX
    frmMain.PicCopy.PaintPicture frmMain.picMain.Picture, 0, 0, frmMain.PicCopy.ScaleWidth, frmMain.PicCopy.ScaleHeight, FX, FY, (LX - FX), (LY - FY)
    frmMain.PicCopy.Picture = frmMain.PicCopy.Image
    Clipboard.Clear
    Clipboard.SetData frmMain.PicCopy.Picture
End If
End Sub

Private Sub mnuCut_Click()
Dim Col As Long
If Sel = True Then
    Sel = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    Toolbar2.Buttons(6).Enabled = False
    Toolbar2.Buttons(7).Enabled = False
    mnuDelete.Enabled = False
    Toolbar2.Buttons(9).Enabled = False
    frmMain.PicCopy.Height = LY - FY
    frmMain.PicCopy.Width = LX - FX
    frmMain.PicCopy.PaintPicture frmMain.picMain.Picture, 0, 0, frmMain.PicCopy.ScaleWidth, frmMain.PicCopy.ScaleHeight, FX, FY, (LX - FX), (LY - FY)
    frmMain.PicCopy.Picture = frmMain.PicCopy.Image
    Clipboard.Clear
    Clipboard.SetData frmMain.PicCopy.Picture
    frmMain.picMain.Cls
    For i = FX + 1 To LX - 1
        For j = FY + 1 To LY - 1
            Col = Bcolor
            r = Col Mod 256
            g = (Col \ 256) Mod 256
            b = Col \ 256 \ 256
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
    Next
    frmMain.picMain.Refresh
    Call Setpicture
End If

End Sub

Private Sub mnuDarkness_Click()

On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            c = Abs((r + g + b) \ 3)
            c = c \ 10
            r = Abs(r - c)
            g = Abs(g - c)
            b = Abs(b - c)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            c = Abs((r + g + b) \ 3)
            c = c \ 10
            r = Abs(r - c)
            g = Abs(g - c)
            b = Abs(b - c)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False


End Sub

Private Sub mnuDelete_Click()
Dim Col As Long
If Sel = True Then
    Sel = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    Toolbar2.Buttons(6).Enabled = False
    Toolbar2.Buttons(7).Enabled = False
    mnuDelete.Enabled = False
    Toolbar2.Buttons(9).Enabled = False
    frmMain.picMain.Cls
    For i = FX + 1 To LX - 1
        For j = FY + 1 To LY - 1
            Col = Bcolor
            r = Col Mod 256
            g = (Col \ 256) Mod 256
            b = Col \ 256 \ 256
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
    Next
    frmMain.picMain.Refresh
    Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub

Private Sub mnuEmboss_Click()

On Error Resume Next
Dim Color As Long
Dim Color1 As Long
Dim r1 As Long
Dim g1 As Long
Dim b1 As Long
Dim i As Long
Dim j As Long
Dim c As Long
frmMain.picMain.Cls
If Sel = True Then
Sel = False

    For i = FX + 1 To LX - 1
        For j = FY + 1 To LY - 1
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            Color1 = GetPixel(frmMain.picMain.hdc, i + 2, j + 2)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            
            r1 = Color1 Mod 256
            g1 = (Color1 \ 256) Mod 256
            b1 = Color1 \ 256 \ 256
            
            r = Abs((r - r1) + 128)
            g = Abs((g - g1) + 128)
            b = Abs((b - b1) + 128)
            
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            Color1 = GetPixel(frmMain.picMain.hdc, i + 2, j + 2)
            
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            
            r1 = Color1 Mod 256
            g1 = (Color1 \ 256) Mod 256
            b1 = Color1 \ 256 \ 256
            
            r = Abs((r - r1) + 128)
            g = Abs((g - g1) + 128)
            b = Abs((b - b1) + 128)
            
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False


End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuForward_Click()
    Undo = Undo + 1
    Toolbar2.Buttons(4).Enabled = True
    MDIForm1.mnuBackward.Enabled = True
    frmMain.picMain.Picture = MDIForm1.picSave(Undo).Picture
    If Undo > Save Then
        Toolbar2.Buttons(5).Enabled = False
        mnuForward.Enabled = False
    End If

End Sub

Private Sub mnuGlow_Click()

On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            If r > 220 Or g > 220 Or b > 220 Then
                c = Abs((r + g + b) \ 3)
                c = c \ 10
                r = Abs(r + c)
                g = Abs(g + c)
                b = Abs(b + c)
            ElseIf r > 150 Or g > 150 Or b > 150 Then
                c = Abs((r + g + b) \ 3)
                c = c \ 20
                r = Abs(r + c)
                g = Abs(g + c)
                b = Abs(b + c)
            
            End If
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            If r > 180 Or g > 180 Or b > 180 Then
                c = Abs((r + g + b) \ 3)
                c = c \ 10
                r = Abs(r + c)
                g = Abs(g + c)
                b = Abs(b + c)
            ElseIf r > 150 Or g > 150 Or b > 150 Then
                c = Abs((r + g + b) \ 3)
                c = c \ 20
                r = Abs(r + c)
                g = Abs(g + c)
                b = Abs(b + c)
                
            End If
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False


End Sub

Private Sub mnuGrayscale_Click()
On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            c = (r * 0.33 + g * 0.33 + b * 0.33)
            SetPixel frmMain.picMain.hdc, i, j, RGB(c, c, c)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            c = r * 0.33 + g * 0.33 + b * 0.33
            SetPixel frmMain.picMain.hdc, i, j, RGB(c, c, c)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub





Private Sub mnuGreen_Click()
On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            SetPixel frmMain.picMain.hdc, i, j, RGB(0, g, 0)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            SetPixel frmMain.picMain.hdc, i, j, RGB(0, g, 0)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub

Private Sub mnuNew_Click()
frmNew.Show
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Traperr

    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    frmMain.picMain.Picture = LoadPicture(CommonDialog1.FileName)
    frmMain.Caption = CommonDialog1.FileName
    frmMain.WindowState = 2
    Label4.Caption = frmMain.picMain.ScaleHeight & " X " & frmMain.picMain.ScaleWidth
    frmMain.Picture1.Picture = frmMain.picMain.Picture
    
    Call Reset
    Call Resetprop

    IsOpen = True
    
    
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

    
    
    
Traperr:
    If Err.Number = 481 Then
        MsgBox "Invalid picture type", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub


End Sub

Private Sub mnuOption_Click()
If mnuOption.Checked = True Then
    mnuOption.Checked = False
    Picture2.Visible = False
Else
   mnuOption.Checked = True
   Picture2.Visible = True
End If

End Sub

Private Sub mnuPaste_Click()
frmMain.picMain.Cls
frmMain.PicCopy.Picture = Clipboard.GetData
frmMain.PicCopy.Visible = True
frmMain.PicCopy.Top = 0
frmMain.PicCopy.Left = 0


End Sub

Private Sub mnuQuality_Click()
On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            r = Abs(r - 200)
            g = Abs(g - 220)
            b = Abs(b - 220)
            
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            r = Abs(r - 200)
            g = Abs(g - 230)
            b = Abs(b - 250)
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, g, b)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub

Private Sub mnuRed_Click()
On Error Resume Next
Dim Color As Long
Dim i As Long
Dim j As Long
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
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, 0, 0)
        Next
        Bar.Max = (LX - 1) - (FX + 1)
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
Else
    For i = 0 To frmMain.picMain.ScaleWidth
        For j = 0 To frmMain.picMain.ScaleHeight
            Color = GetPixel(frmMain.picMain.hdc, i, j)
            r = Color Mod 256
            g = (Color \ 256) Mod 256
            b = Color \ 256 \ 256
            SetPixel frmMain.picMain.hdc, i, j, RGB(r, 0, 0)

        Next
        Bar.Max = frmMain.picMain.ScaleWidth
        Bar.value = Bar.value + 1
    Next
Bar.value = 0
frmMain.picMain.Refresh
Call Setpicture
End If
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
mnuBackward.Enabled = True
mnuForward.Enabled = False

End Sub

Private Sub mnuSave_Click()
On Error GoTo Traperr
    Call Reset
    CommonDialog2.CancelError = True
    CommonDialog2.ShowSave
    SavePicture frmMain.picMain.Picture, CommonDialog2.FileName
Traperr:
Exit Sub
End Sub

Private Sub mnuStandard_Click()
If mnuStandard.Checked = True Then
    mnuStandard.Checked = False
    Toolbar2.Visible = False
Else
   mnuStandard.Checked = True
   Toolbar2.Visible = True
End If
End Sub

Private Sub mnuStatus_Click()
If mnuStatus.Checked = True Then
    mnuStatus.Checked = False
    Picture7.Visible = False
Else
   mnuStatus.Checked = True
   Picture7.Visible = True
End If

End Sub

Private Sub mnuTool_Click()
If mnuTool.Checked = True Then
    mnuTool.Checked = False
    Picture1.Visible = False
Else
   mnuTool.Checked = True
   Picture1.Visible = True
End If

End Sub

Private Sub Option1_Click(Index As Integer)
Quality = Option1(Index).Index
End Sub

Private Sub picBackcolor_Click()
On Error GoTo Traperr
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
picBackcolor.BackColor = CommonDialog1.Color
Bcolor = CommonDialog1.Color
Picture5.BackColor = Bcolor

Label5.Caption = GetColorR(Picture5.hdc, 20, 20)
Label6.Caption = GetColorG(Picture5.hdc, 20, 20)
Label7.Caption = GetColorB(Picture5.hdc, 20, 20)


Traperr:
Exit Sub

End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    If Fcol = True Then
        Picture4.BackColor = picColor.Point(X, Y)
        picFillcolor.BackColor = picColor.Point(X, Y)
        Fcolor = picColor.Point(X, Y)
    ElseIf Bcol = True Then
        Picture5.BackColor = picColor.Point(X, Y)
        picBackcolor.BackColor = picColor.Point(X, Y)
        Bcolor = picColor.Point(X, Y)
    Else
        Picture6.BackColor = picColor.Point(X, Y)
        picDrawcolor.BackColor = picColor.Point(X, Y)
        DColor = picColor.Point(X, Y)
    End If
Label5.Caption = GetColorR(picColor.hdc, X, Y)
Label6.Caption = GetColorG(picColor.hdc, X, Y)
Label7.Caption = GetColorB(picColor.hdc, X, Y)

End If

End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim value As Long
Dim r As Long
Dim g As Long
Dim b As Long

On Error Resume Next
If Button = 1 Then
    If Fcol = True Then
        Picture4.BackColor = picColor.Point(X, Y)
        picFillcolor.BackColor = picColor.Point(X, Y)
        Fcolor = picColor.Point(X, Y)
    ElseIf Bcol = True Then
        Picture5.BackColor = picColor.Point(X, Y)
        picBackcolor.BackColor = picColor.Point(X, Y)
        Bcolor = picColor.Point(X, Y)
    Else
        Picture6.BackColor = picColor.Point(X, Y)
        picDrawcolor.BackColor = picColor.Point(X, Y)
        DColor = picColor.Point(X, Y)
    End If
   

Label5.Caption = GetColorR(picColor.hdc, X, Y)
Label6.Caption = GetColorG(picColor.hdc, X, Y)
Label7.Caption = GetColorB(picColor.hdc, X, Y)
End If
End Sub


Private Sub picDrawcolor_Click()
On Error GoTo Traperr
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
picDrawcolor.BackColor = CommonDialog1.Color
DColor = CommonDialog1.Color
Picture6.BackColor = DColor

Label5.Caption = GetColorR(Picture6.hdc, 20, 20)
Label6.Caption = GetColorG(Picture6.hdc, 20, 20)
Label7.Caption = GetColorB(Picture6.hdc, 20, 20)

Traperr:
Exit Sub

End Sub

Private Sub picFillcolor_Click()
On Error GoTo Traperr
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
picFillcolor.BackColor = CommonDialog1.Color
Fcolor = CommonDialog1.Color
Picture4.BackColor = Fcolor

Label5.Caption = GetColorR(Picture4.hdc, 20, 20)
Label6.Caption = GetColorG(Picture4.hdc, 20, 20)
Label7.Caption = GetColorB(Picture4.hdc, 20, 20)

Traperr:
Exit Sub

End Sub

Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    If Fcol = True Then
        Picture4.BackColor = Picture11.Point(X, Y)
        picFillcolor.BackColor = Picture11.Point(X, Y)
        Fcolor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture4.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture4.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture4.hdc, 20, 20)

    ElseIf Bcol = True Then
        Picture5.BackColor = Picture11.Point(X, Y)
        picBackcolor.BackColor = Picture11.Point(X, Y)
        Bcolor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture5.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture5.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture5.hdc, 20, 20)

    Else
        Picture6.BackColor = Picture11.Point(X, Y)
        picDrawcolor.BackColor = Picture11.Point(X, Y)
        DColor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture6.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture6.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture6.hdc, 20, 20)

    End If
    
End If

End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    If Fcol = True Then
        Picture4.BackColor = Picture11.Point(X, Y)
        picFillcolor.BackColor = Picture11.Point(X, Y)
        Fcolor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture4.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture4.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture4.hdc, 20, 20)
        
    ElseIf Bcol = True Then
        Picture5.BackColor = Picture11.Point(X, Y)
        picBackcolor.BackColor = Picture11.Point(X, Y)
        Bcolor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture5.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture5.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture5.hdc, 20, 20)

    Else
        Picture6.BackColor = Picture11.Point(X, Y)
        picDrawcolor.BackColor = Picture11.Point(X, Y)
        DColor = Picture11.Point(X, Y)
        Label5.Caption = GetColorR(Picture6.hdc, 20, 20)
        Label6.Caption = GetColorG(Picture6.hdc, 20, 20)
        Label7.Caption = GetColorB(Picture6.hdc, 20, 20)
        
    End If

End If

End Sub

Private Sub Picture3_Click(Index As Integer)
Fcolor = Picture3(Index).BackColor
picFillcolor.BackColor = Fcolor
Picture4.BackColor = Fcolor
Label5.Caption = GetColorR(Picture3(Index).hdc, X, Y)
Label6.Caption = GetColorG(Picture3(Index).hdc, X, Y)
Label7.Caption = GetColorB(Picture3(Index).hdc, X, Y)
Shape2.Left = Picture3(Index).Left - 10
Shape2.Top = Picture3(Index).Top - 20
End Sub

Private Sub Picture4_Click()
Shape1.Top = Picture4.Top - 10
Fcol = True
Bcol = False
Dcol = False
End Sub

Private Sub Picture5_Click()
Shape1.Top = Picture5.Top - 10
Fcol = False
Bcol = True
Dcol = False
End Sub

Private Sub Picture6_Click()
Shape1.Top = Picture6.Top - 10
Fcol = False
Bcol = False
Dcol = True
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call Reset
        Curtool.Pencil = True
        Frame3.Caption = "Pencil properties"
        Call Resetprop
        txtPenwidth.Visible = True
        Label1.Visible = True
        frmMain.Shape1.Visible = False
    Case 2
        Call Reset
        Curtool.line = True
        Frame3.Caption = "Line properties"
        Call Resetprop
        txtLinewidth.Visible = True
        Label1.Visible = True
        Combo1.Visible = True
        Label2.Visible = True
        Combo1.Text = "Solid"
        frmMain.Shape1.Visible = False
    Case 3
        Call Reset
        Curtool.Square = True
        Frame3.Caption = "Rectangle properties"
        Call Resetprop
        txtrectwidth.Visible = True
        Label1.Visible = True
        Combo1.Visible = True
        Label2.Visible = True
        Combo1.Text = "Solid"
        frmMain.Shape1.Visible = False
    Case 4
        Call Reset
        Curtool.Cir = True
        Frame3.Caption = "Circle properties"
        Call Resetprop
        txtCirwidth.Visible = True
        Label1.Visible = True
        Combo1.Visible = True
        Label2.Visible = True
        Combo1.Text = "Solid"
        frmMain.Shape1.Visible = False
    Case 5
        Call Reset
        Curtool.FSquare = True
        Frame3.Caption = "Filled rectangle properties"
        Call Resetprop
        txtFRectwidth.Visible = True
        Label1.Visible = True
        Combo1.Visible = True
        Label2.Visible = True
        Combo2.Visible = True
        Label3.Visible = True
        Combo1.Text = "Solid"
        Combo2.Text = "Solid"
        Combo2.Top = 1110
        Label3.Top = 1110
        frmMain.Shape1.Visible = False
    Case 6
        Call Reset
        Curtool.FCircle = True
        Frame3.Caption = "Filled circle properties"
        Call Resetprop
        txtFCirwidth.Visible = True
        Label1.Visible = True
        Combo1.Visible = True
        Label2.Visible = True
        Combo2.Visible = True
        Label3.Visible = True
        Combo2.Text = "Solid"
        Combo1.Text = "Solid"
        Combo2.Top = 1110
        Label3.Top = 1110
        frmMain.Shape1.Visible = False
    Case 7
        Call Reset
        Curtool.Eraser = True
        Frame3.Caption = "Eraser properties"
        Call Resetprop
        txtEraserwidth.Visible = True
        Label1.Visible = True
        frmMain.Shape1.Visible = False
    Case 8
        Call Reset
        Call Resetprop
        Frame4.Visible = True
        Curtool.Brush = True
        frmMain.Shape1.Visible = False
    Case 9
        Call Reset
        Call Resetprop
        Curtool.Crop = True
        Frame3.Caption = ""
        frmMain.Shape1.Visible = False
        
    Case 10
        Call Reset
        Call Resetprop
        Frame3.Caption = ""
        Curtool.Eyedroper = True
        frmMain.Shape1.Visible = False
    Case 11
        Call Reset
        Curtool.Paint = True
        Frame3.Caption = "Paint bucket properties"
        Call Resetprop
        Combo1.Visible = False
        Label2.Visible = False
        Combo2.Visible = True
        Combo2.Text = "Solid"
        Label3.Visible = True
        Combo2.Top = 360
        Label3.Top = 360
        frmMain.Shape1.Visible = False
    Case 12
        Call Reset
        Curtool.Clone = True
        Frame3.Caption = "Clone stamp options"
        Call Resetprop
        Option1(1).Visible = True
        Option1(2).Visible = True
        Option1(3).Visible = True
        Option1(1).Top = 360
        Option1(2).Top = 360 + Option1(1).Height + 10
        Option1(3).Top = 360 + Option1(1).Height + Option1(2).Height + 10
        
    Case 13
        Call Reset
        Call Resetprop
        Curtool.Select = True
        Frame3.Caption = ""
        frmMain.Shape1.Visible = False
End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Traperr
Select Case Button.Index
Case 1
    frmNew.Show
Case 2

    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    frmMain.picMain.Picture = LoadPicture(CommonDialog1.FileName)
    frmMain.Caption = CommonDialog1.FileName
    frmMain.WindowState = 2
    Label4.Caption = frmMain.picMain.ScaleHeight & " X " & frmMain.picMain.ScaleWidth
    frmMain.Picture1.Picture = frmMain.picMain.Picture
    
    Call Reset
    Call Resetprop

    IsOpen = True
    
    
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

    
    
    
Traperr:
    If Err.Number = 481 Then
        MsgBox "Invalid picture type", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
Case 3
    Call Reset
    CommonDialog2.CancelError = True
    CommonDialog2.ShowSave
    SavePicture frmMain.picMain.Picture, CommonDialog2.FileName

Case 4
    Undo = Undo - 1
    IsUndo = True
    Toolbar2.Buttons(5).Enabled = True
    mnuForward.Enabled = True
    If Undo = 0 Then
        Toolbar2.Buttons(4).Enabled = False
        mnuBackward.Enabled = False
        If IsOpen = True Then
            frmMain.picMain.Picture = frmMain.Picture1.Picture
            Exit Sub
        End If
    End If
    frmMain.picMain.Picture = MDIForm1.picSave(Undo).Picture
    
Case 5
    Undo = Undo + 1
    Toolbar2.Buttons(4).Enabled = True
    MDIForm1.mnuBackward.Enabled = True
    frmMain.picMain.Picture = MDIForm1.picSave(Undo).Picture
    If Undo > Save Then
        Toolbar2.Buttons(5).Enabled = False
        mnuForward.Enabled = False
    End If
    
Case 6
    mnuCut_Click
    
Case 7
    mnuCopy_Click
    
Case 8
    mnuPaste_Click
    
Case 9
    mnuDelete_Click
End Select
End Sub

Private Sub txtCirwidth_LostFocus()
If Val(txtCirwidth.Text) <= 1 Then
    txtCirwidth.Text = 1
End If
Dwidth = Trim(Val(txtCirwidth.Text))

End Sub

Private Sub txtEraserwidth_LostFocus()
If Val(txtEraserwidth.Text) <= 1 Then
    txtEraserwidth.Text = 1
End If
Dwidth = Trim(Val(txtEraserwidth.Text))

End Sub

Private Sub txtFCirwidth_LostFocus()
If Val(txtFCirwidth.Text) <= 1 Then
    txtFCirwidth.Text = 1
End If
Dwidth = Trim(Val(txtFCirwidth.Text))

End Sub

Private Sub txtFRectwidth_LostFocus()
If Val(txtFRectwidth.Text) <= 1 Then
    txtFRectwidth.Text = 1
End If
Dwidth = Trim(Val(txtFRectwidth.Text))

End Sub

Private Sub txtLinewidth_LostFocus()
If Val(txtLinewidth.Text) <= 1 Then
    txtLinewidth.Text = 1
End If
Dwidth = Trim(Val(txtLinewidth.Text))

End Sub

Private Sub txtPenwidth_LostFocus()
If Val(txtPenwidth.Text) <= 1 Then
    txtPenwidth.Text = 1
End If
Dwidth = Trim(Val(txtPenwidth.Text))
End Sub

Private Sub txtrectwidth_LostFocus()
If Val(txtrectwidth.Text) <= 1 Then
    txtrectwidth.Text = 1
End If
Dwidth = Trim(Val(txtrectwidth.Text))

End Sub
