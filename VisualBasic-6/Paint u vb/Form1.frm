VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   7815
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   6375
      LargeChange     =   10
      Left            =   7440
      TabIndex        =   3
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   7440
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      LargeChange     =   10
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      Width           =   7455
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000D&
      Height          =   6495
      Left            =   0
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      ScaleHeight     =   431
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.PictureBox picCopy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         MousePointer    =   15  'Size All
         ScaleHeight     =   119
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   167
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   2475
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         Height          =   135
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstX As Long
Dim firstY As Long
Dim LastX As Long
Dim LastY As Long

Private Sub Form_Activate()

On Error Resume Next
MDIForm1.Toolbar2.Buttons(1).Enabled = False
MDIForm1.Toolbar2.Buttons(2).Enabled = False
MDIForm1.mnuNew.Enabled = False
MDIForm1.mnuOpen.Enabled = False

VScroll1.Left = Me.ScaleWidth - VScroll1.Width
If HScroll1.Visible = False Then
    VScroll1.Height = Me.ScaleHeight
    Picture2.Visible = False

Else
    VScroll1.Height = Me.ScaleHeight - 300
    Picture2.Visible = True
End If
If VScroll1.Visible = False Then
    HScroll1.Width = Me.ScaleWidth
    Picture2.Visible = False
Else
    HScroll1.Width = Me.ScaleWidth - 300
    Picture2.Visible = True
End If
HScroll1.Top = Me.ScaleHeight - HScroll1.Height
Picture2.Left = Me.ScaleWidth - 300
Picture2.Top = Me.ScaleHeight - 300


If VScroll1.Visible = False Then
    If Me.ScaleWidth > picMain.Width Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If
Else
    If (Me.ScaleWidth - 300) > picMain.Width Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If
End If

If HScroll1.Visible = False Then
    If Me.ScaleHeight > picMain.Height Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
Else
    If (Me.ScaleHeight - 300) > picMain.Height Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
End If

If Me.ScaleWidth > picMain.Width Then
    If VScroll1.Visible = False Then
        picMain.Left = (Me.ScaleWidth / 2) - (picMain.Width / 2)
    Else
        picMain.Left = ((Me.ScaleWidth - 300) / 2) - (picMain.Width / 2)
    End If
End If
If Me.ScaleHeight > picMain.Height Then
    If HScroll1.Visible = False Then
        picMain.Top = (Me.ScaleHeight / 2) - (picMain.Height / 2)
    Else
        picMain.Top = ((Me.ScaleHeight - 300) / 2) - (picMain.Height / 2)
    End If
End If
If VScroll1.Visible = False Or HScroll1.Visible = False Then
VScroll1.Max = (picMain.Height - Me.ScaleHeight)
HScroll1.Max = (picMain.Width - Me.ScaleWidth)
Else
VScroll1.Max = (picMain.Height - Me.ScaleHeight) + HScroll1.Height
HScroll1.Max = (picMain.Width - Me.ScaleWidth) + VScroll1.Width
End If

End Sub

Private Sub Form_Load()
MDIForm1.Toolbar1.Enabled = True
MDIForm1.Toolbar2.Buttons(1).Enabled = False
MDIForm1.Toolbar2.Buttons(2).Enabled = False
MDIForm1.Toolbar2.Buttons(3).Enabled = True
MDIForm1.Toolbar2.Buttons(8).Enabled = True
MDIForm1.mnuNew.Enabled = False
MDIForm1.mnuOpen.Enabled = False
MDIForm1.mnuSave.Enabled = True
MDIForm1.mnuPaste.Enabled = True
MDIForm1.mnuGrayscale.Enabled = True
MDIForm1.mnuBrightness.Enabled = True
MDIForm1.mnuDarkness.Enabled = True
MDIForm1.mnuEmboss.Enabled = True
MDIForm1.mnuBnW.Enabled = True
MDIForm1.mnuRed.Enabled = True
MDIForm1.mnuGreen.Enabled = True
MDIForm1.mnuBlue.Enabled = True
MDIForm1.mnuQuality.Enabled = True
MDIForm1.mnuColor.Enabled = True
MDIForm1.mnuGlow.Enabled = True


IsUndo = False

Save = 0
Undo = 0
End Sub


Private Sub Form_Resize()

On Error Resume Next

VScroll1.Left = Me.ScaleWidth - VScroll1.Width
If HScroll1.Visible = False Then
    VScroll1.Height = Me.ScaleHeight
    Picture2.Visible = False

Else
    VScroll1.Height = Me.ScaleHeight - 300
    Picture2.Visible = True
End If
If VScroll1.Visible = False Then
    HScroll1.Width = Me.ScaleWidth
    Picture2.Visible = False
Else
    HScroll1.Width = Me.ScaleWidth - 300
    Picture2.Visible = True
End If
HScroll1.Top = Me.ScaleHeight - HScroll1.Height
Picture2.Left = Me.ScaleWidth - 300
Picture2.Top = Me.ScaleHeight - 300


If VScroll1.Visible = False Then
    If Me.ScaleWidth > picMain.Width Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If
Else
    If (Me.ScaleWidth - 300) > picMain.Width Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If
End If

If HScroll1.Visible = False Then
    If Me.ScaleHeight > picMain.Height Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
Else
    If (Me.ScaleHeight - 300) > picMain.Height Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
End If

If Me.ScaleWidth > picMain.Width Then
    If VScroll1.Visible = False Then
        picMain.Left = (Me.ScaleWidth / 2) - (picMain.Width / 2)
    Else
        picMain.Left = ((Me.ScaleWidth - 300) / 2) - (picMain.Width / 2)
    End If
End If
If Me.ScaleHeight > picMain.Height Then
    If HScroll1.Visible = False Then
        picMain.Top = (Me.ScaleHeight / 2) - (picMain.Height / 2)
    Else
        picMain.Top = ((Me.ScaleHeight - 300) / 2) - (picMain.Height / 2)
    End If
End If
If VScroll1.Visible = False Or HScroll1.Visible = False Then
VScroll1.Max = (picMain.Height - Me.ScaleHeight)
HScroll1.Max = (picMain.Width - Me.ScaleWidth)
Else
VScroll1.Max = (picMain.Height - Me.ScaleHeight) + HScroll1.Height
HScroll1.Max = (picMain.Width - Me.ScaleWidth) + VScroll1.Width
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Toolbar1.Enabled = False
MDIForm1.Toolbar2.Buttons(1).Enabled = True
MDIForm1.Toolbar2.Buttons(2).Enabled = True
MDIForm1.Toolbar2.Buttons(3).Enabled = False
MDIForm1.Toolbar2.Buttons(4).Enabled = False
MDIForm1.Toolbar2.Buttons(5).Enabled = False
MDIForm1.Toolbar2.Buttons(6).Enabled = False
MDIForm1.Toolbar2.Buttons(7).Enabled = False
MDIForm1.Toolbar2.Buttons(8).Enabled = False
MDIForm1.mnuDelete.Enabled = False
MDIForm1.Toolbar2.Buttons(9).Enabled = False

MDIForm1.mnuNew.Enabled = True
MDIForm1.mnuOpen.Enabled = True
MDIForm1.mnuSave.Enabled = False
MDIForm1.mnuBackward.Enabled = False
MDIForm1.mnuForward.Enabled = False

MDIForm1.mnuPaste.Enabled = False
MDIForm1.mnuCut.Enabled = False
MDIForm1.mnuCopy.Enabled = False
MDIForm1.mnuGrayscale.Enabled = False
MDIForm1.mnuBrightness.Enabled = False
MDIForm1.mnuDarkness.Enabled = False
MDIForm1.mnuEmboss.Enabled = False
MDIForm1.mnuBnW.Enabled = False
MDIForm1.mnuRed.Enabled = False
MDIForm1.mnuGreen.Enabled = False
MDIForm1.mnuBlue.Enabled = False
MDIForm1.mnuQuality.Enabled = False
MDIForm1.mnuColor.Enabled = False
MDIForm1.mnuGlow.Enabled = False

Call Button
Call Reset
Call Resetprop

If Save >= 1 Then
    For a = 1 To Save + 1
        Unload MDIForm1.picSave(a)
    Next
End If
Undo = 0
Save = 0
IsUndo = False
ISsave = False
MDIForm1.Frame3.Caption = ""
'Call Resetprop
'Call Reset
End Sub

Private Sub HScroll1_Change()
picMain.Left = -HScroll1.value
End Sub

Private Sub HScroll1_Scroll()
picMain.Left = -HScroll1.value
End Sub

Private Sub PicCopy_DblClick()

On Error Resume Next
frmMain.picMain.PaintPicture frmMain.PicCopy.Picture, PicCopy.Left, PicCopy.Top, PicCopy.ScaleWidth, PicCopy.ScaleHeight, 0, 0, PicCopy.ScaleWidth, PicCopy.ScaleHeight
PicCopy.Visible = False
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

Private Sub PicCopy_LostFocus()
On Error Resume Next
frmMain.picMain.PaintPicture frmMain.PicCopy.Picture, PicCopy.Left, PicCopy.Top, PicCopy.ScaleWidth, PicCopy.ScaleHeight, 0, 0, PicCopy.ScaleWidth, PicCopy.ScaleHeight
PicCopy.Visible = False
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

Private Sub picCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    firstX = X
    firstY = Y
End If
End Sub

Private Sub picCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    LastX = X - firstX
    LastY = Y - firstY
    PicCopy.Top = (PicCopy.Top + LastY)
    PicCopy.Left = (PicCopy.Left + LastX)
End If
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
If DoCrop = True Then
MDIForm1.picCrop.Picture = LoadPicture("")
MDIForm1.picCrop.Width = (LX + 1 - FX + 1) * 15.065913370998
MDIForm1.picCrop.Height = (LY + 1 - FY + 1) * 15.065913370998

    For i = FX + 1 To LX
        For j = FY + 1 To LY
            Crop = GetPixel(picMain.hdc, i, j)
            r = Crop Mod 256
            g = (Crop \ 256) Mod 256
            b = Crop \ 256 \ 256
            SetPixel MDIForm1.picCrop.hdc, i - 2 - FX + 1, j - 2 - FY + 1, RGB(r, g, b)

        Next
    Next
MDIForm1.picCrop.Refresh
MDIForm1.picCrop.Picture = MDIForm1.picCrop.Image
picMain.Picture = MDIForm1.picCrop.Image
picMain.Width = (LX - FX) * 15.065913370998
picMain.Height = (LY - FY) * 15.065913370998
MDIForm1.picCrop.Width = (LX + 1 - FX + 1) * 15.065913370998
MDIForm1.picCrop.Height = (LY + 1 - FY + 1) * 15.065913370998
DoCrop = False

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
        MDIForm1.mnuBackward.Enabled = True
        Load MDIForm1.picSave(Save + 1)
        MDIForm1.picSave(Save + 1).Picture = frmMain.picMain.Picture
        MDIForm1.Toolbar2.Buttons(5).Enabled = False
        MDIForm1.mnuForward.Enabled = False
Form_Resize

End If
End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.picMain.Cls
Sel = False
If Button = 1 Then
    FX = X
    FY = Y
            If MDIForm1.Option6.value = True Then
            Call Blue(picMain.hdc, X, Y)
            picMain.Refresh
        End If

End If
If Button = 1 Then
    If Curtool.Paint = True Then
        Call Filling(picMain.Point(X, Y), MDIForm1.Combo2.ListIndex, X, Y)
    End If
End If

If Curtool.Clone = True Then

    If Cloneset = True Then
    
        If Button = 2 Then
            Cloneset = True
            CX = X
            CY = Y
            Shape1.Visible = True
            Shape1.Left = CX - (Shape1.Width / 2)
            Shape1.Top = CY - (Shape1.Height / 2)
        End If
    Else
        If Button = 2 Then
            Cloneset = True
            CX = X
            CY = Y
            Shape1.Visible = True
            Shape1.Left = CX
            Shape1.Top = CY

        End If
        If Button = 1 Then
        MsgBox "Select clone position with right mouse button", vbInformation, App.Title
        End If
    End If
End If

If Curtool.Eyedroper = True Then
    If Button = 1 Then
        If Fcol = True Then
            MDIForm1.Picture4.BackColor = picMain.Point(X, Y)
            MDIForm1.picFillcolor.BackColor = picMain.Point(X, Y)
            Fcolor = picMain.Point(X, Y)
            
        End If
         
        If Bcol = True Then
            MDIForm1.Picture5.BackColor = picMain.Point(X, Y)
            MDIForm1.picBackcolor.BackColor = picMain.Point(X, Y)
            Bcolor = picMain.Point(X, Y)
        End If
         
        If Dcol = True Then
            MDIForm1.Picture6.BackColor = picMain.Point(X, Y)
            MDIForm1.picDrawcolor.BackColor = picMain.Point(X, Y)
            DColor = picMain.Point(X, Y)
        End If
        MDIForm1.Label5.Caption = GetColorR(picMain.hdc, X, Y)
        MDIForm1.Label6.Caption = GetColorG(picMain.hdc, X, Y)
        MDIForm1.Label7.Caption = GetColorB(picMain.hdc, X, Y)
     End If
End If

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then

    If Curtool.Pencil = True Then
    picMain.DrawStyle = 0
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtPenwidth.Text)
        picMain.Line (FX, FY)-(X, Y)
        FX = X
        FY = Y
    End If
    
    If Curtool.line = True Then
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtLinewidth.Text)
        picMain.Cls
        picMain.Line (FX, FY)-(X, Y)
    End If
    
    If Curtool.Square = True Then
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtrectwidth.Text)
        picMain.FillStyle = 1
        picMain.Cls
        picMain.Line (FX, FY)-(X, Y), , B
    End If
    
    If Curtool.Cir = True Then
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtCirwidth.Text)
        picMain.FillStyle = 1
        picMain.Cls
        If X > FX Or Y > FY Then
            picMain.Circle (FX, FY), (X - FX) + (Y - FY)
        Else
            picMain.Circle (FX, FY), (FX - X) + (FY - Y)
        End If
    End If
    
    If Curtool.FSquare = True Then
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtFRectwidth.Text)
        picMain.FillColor = Fcolor
        picMain.Cls
        picMain.Line (FX, FY)-(X, Y), , B
    End If
    
    If Curtool.FCircle = True Then
        picMain.ForeColor = DColor
        picMain.DrawWidth = Val(MDIForm1.txtFCirwidth.Text)
        picMain.FillColor = Fcolor
        picMain.Cls
        If X > FX Or Y > FY Then
            picMain.Circle (FX, FY), (X - FX) + (Y - FY)
        Else
            picMain.Circle (FX, FY), (FX - X) + (FY - Y)
        End If
    End If
    
    If Curtool.Eraser = True Then
        picMain.ForeColor = Bcolor
        picMain.DrawStyle = 0
        picMain.DrawWidth = Val(MDIForm1.txtEraserwidth.Text)
        picMain.Line (FX, FY)-(X, Y)
        FX = X
        FY = Y
    End If
    If Curtool.Brush = True Then
        If MDIForm1.Option2.value = True Then
            Call Normbrush(FX, FY, X, Y)
            FX = X
            FY = Y
        End If
        
        If MDIForm1.Option3.value = True Then
            Call Grayscale(picMain.hdc, X, Y)
            picMain.Refresh
        End If
        
        If MDIForm1.Option4.value = True Then
            Call Red(picMain.hdc, X, Y)
            picMain.Refresh
        End If
        
        If MDIForm1.Option5.value = True Then
            Call Green(picMain.hdc, X, Y)
            picMain.Refresh
        End If
        
        If MDIForm1.Option6.value = True Then
            Call Blue(picMain.hdc, X, Y)
            picMain.Refresh
        End If
        
        If MDIForm1.Option7.value = True Then
            picMain.DrawStyle = 0
            picMain.DrawWidth = Val(MDIForm1.txtBrushwidth.Text)
            picMain.ForeColor = picMain.Point(X, Y)
            picMain.PSet (X, Y)
            
        End If
        
    End If
    
    If Curtool.Crop = True Then
        If X > FX And Y > FY Then
            DoCrop = True
            picMain.ForeColor = vbBlack
            picMain.DrawWidth = 1
            picMain.DrawStyle = 0
            picMain.FillStyle = 1
            picMain.Cls
            picMain.Line (FX, FY)-(X, Y), , B
        End If
    End If
    
    If Curtool.Eyedroper = True Then
        If Fcol = True Then
            MDIForm1.Picture4.BackColor = picMain.Point(X, Y)
            MDIForm1.picFillcolor.BackColor = picMain.Point(X, Y)
            Fcolor = picMain.Point(X, Y)
        End If
         
        If Bcol = True Then
            MDIForm1.Picture5.BackColor = picMain.Point(X, Y)
            MDIForm1.picBackcolor.BackColor = picMain.Point(X, Y)
            Bcolor = picMain.Point(X, Y)
        End If
         
        If Dcol = True Then
            MDIForm1.Picture6.BackColor = picMain.Point(X, Y)
            MDIForm1.picDrawcolor.BackColor = picMain.Point(X, Y)
            DColor = picMain.Point(X, Y)
        End If
        MDIForm1.Label5.Caption = GetColorR(picMain.hdc, X, Y)
        MDIForm1.Label6.Caption = GetColorG(picMain.hdc, X, Y)
        MDIForm1.Label7.Caption = GetColorB(picMain.hdc, X, Y)
         
    End If
        
    
    If Curtool.Clone = True Then
        DX = X - FX
        DY = Y - FY
        picMain.DrawStyle = 0
        Quality = MDIForm1.Option1(Index).Index
        picMain.DrawWidth = Quality
        picMain.PSet (X, Y), picMain.Point(CX + DX, CY + DY)
        picMain.PSet (X + 1, Y), picMain.Point((CX + DX) + 1, CY + DY)
        picMain.PSet (X, Y + 1), picMain.Point(CX + DX, (CY + DY) + 1)
        picMain.PSet (X - 1, Y), picMain.Point((CX + DX) - 1, CY + DY)
        picMain.PSet (X, Y - 1), picMain.Point(CX + DX, (CY + DY) - 1)
        picMain.PSet (X + 1, Y - 1), picMain.Point((CX + DX) + 1, (CY + DY) - 1)
        picMain.PSet (X + 1, Y + 1), picMain.Point((CX + DX) + 1, (CY + DY) + 1)
        picMain.PSet (X - 1, Y + 1), picMain.Point((CX + DX) - 1, (CY + DY) + 1)
        picMain.PSet (X - 1, Y - 1), picMain.Point((CX + DX) - 1, (CY + DY) - 1)
        
        picMain.PSet (X + 2, Y), picMain.Point((CX + DX) + 2, CY + DY)
        picMain.PSet (X, Y + 2), picMain.Point(CX + DX, (CY + DY) + 2)
        picMain.PSet (X - 2, Y), picMain.Point((CX + DX) - 2, CY + DY)
        picMain.PSet (X, Y - 2), picMain.Point(CX + DX, (CY + DY) - 2)
        
        
        picMain.PSet (X + 2, Y - 2), picMain.Point((CX + DX) + 2, (CY + DY) - 2)
        picMain.PSet (X + 2, Y + 2), picMain.Point((CX + DX) + 2, (CY + DY) + 2)
        picMain.PSet (X - 2, Y + 2), picMain.Point((CX + DX) - 2, (CY + DY) + 2)
        picMain.PSet (X - 2, Y - 2), picMain.Point((CX + DX) - 2, (CY + DY) - 2)
        
        Shape1.Left = (CX + DX) - (Shape1.Width / 2)
        Shape1.Top = (CY + DY) - (Shape1.Height / 2)

    End If
    
    If Curtool.Select = True Then
        If X > FX And Y > FY Then
            Sel = True
            picMain.ForeColor = vbBlack
            picMain.DrawWidth = 1
            picMain.DrawStyle = 0
            picMain.FillStyle = 1
            picMain.Cls
            picMain.Line (FX, FY)-(X, Y), , B
        End If
    End If
    
    
End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sel = True Then
    MDIForm1.mnuCut.Enabled = True
    MDIForm1.mnuCopy.Enabled = True
    MDIForm1.Toolbar2.Buttons(6).Enabled = True
    MDIForm1.Toolbar2.Buttons(7).Enabled = True
    MDIForm1.mnuDelete.Enabled = True
    MDIForm1.Toolbar2.Buttons(9).Enabled = True
Else
    MDIForm1.mnuCut.Enabled = False
    MDIForm1.mnuCopy.Enabled = False
    MDIForm1.Toolbar2.Buttons(6).Enabled = False
    MDIForm1.Toolbar2.Buttons(7).Enabled = False
    MDIForm1.mnuDelete.Enabled = False
    MDIForm1.Toolbar2.Buttons(9).Enabled = False
End If
    
If DoCrop = True Then
    LX = X
    LY = Y
    If LX > picMain.ScaleWidth Then
        LX = picMain.ScaleWidth
    End If
    If LY > picMain.ScaleHeight Then
        LY = picMain.ScaleHeight
    End If
    
ElseIf Sel = True Then
    LX = X
    LY = Y
    If LX > picMain.ScaleWidth Then
        LX = picMain.ScaleWidth
    End If
    If LY > picMain.ScaleHeight Then
        LY = picMain.ScaleHeight
    End If
Else
Call Setpicture
If Button = 1 Then
    If Curtool.Pencil = True Or Curtool.line = True Or Curtool.Square = True Or Curtool.Cir = True Or Curtool.FSquare = True Or Curtool.FCircle = True Or Curtool.Brush = True Or Curtool.Crop = True Or Curtool.Eraser = True Or Curtool.Paint = True Or Curtool.Clone = True Then
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
        MDIForm1.mnuBackward.Enabled = True
        Load MDIForm1.picSave(Save + 1)
        MDIForm1.picSave(Save + 1).Picture = frmMain.picMain.Picture
        MDIForm1.Toolbar2.Buttons(5).Enabled = False
        MDIForm1.mnuForward.Enabled = False
    End If
End If
End If
End Sub

Private Sub VScroll1_Change()
picMain.Top = -VScroll1.value
End Sub

Private Sub VScroll1_Scroll()
picMain.Top = -VScroll1.value
End Sub
