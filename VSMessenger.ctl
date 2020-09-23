VERSION 5.00
Begin VB.UserControl VSMessenger 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   Begin VB.Timer Timer1 
      Left            =   15
      Top             =   1365
   End
   Begin VB.ListBox lstMessage 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "VSMessenger.ctx":0000
      Left            =   30
      List            =   "VSMessenger.ctx":0002
      TabIndex        =   7
      Top             =   360
      Width           =   3420
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   0
      Top             =   0
      Width           =   3525
      Begin VB.PictureBox sbDNc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3660
         ScaleHeight     =   150
         ScaleWidth      =   225
         TabIndex        =   6
         Top             =   165
         Visible         =   0   'False
         Width           =   225
         Begin VB.Line Line44 
            BorderColor     =   &H00FFFFFF&
            X1              =   195
            X2              =   195
            Y1              =   30
            Y2              =   120
         End
         Begin VB.Line Line43 
            BorderColor     =   &H00FFFFFF&
            X1              =   30
            X2              =   210
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line42 
            BorderColor     =   &H00808080&
            X1              =   15
            X2              =   15
            Y1              =   15
            Y2              =   135
         End
         Begin VB.Line Line41 
            BorderColor     =   &H00808080&
            X1              =   15
            X2              =   210
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line40 
            BorderColor     =   &H00000000&
            X1              =   45
            X2              =   45
            Y1              =   45
            Y2              =   60
         End
         Begin VB.Line Line39 
            BorderColor     =   &H00000000&
            X1              =   60
            X2              =   60
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Line Line38 
            BorderColor     =   &H00000000&
            X1              =   75
            X2              =   75
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line37 
            BorderColor     =   &H00000000&
            X1              =   90
            X2              =   90
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line36 
            BorderColor     =   &H00000000&
            X1              =   165
            X2              =   165
            Y1              =   45
            Y2              =   60
         End
         Begin VB.Line Line35 
            BorderColor     =   &H00000000&
            X1              =   150
            X2              =   150
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Line Line34 
            BorderColor     =   &H00000000&
            X1              =   135
            X2              =   135
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line33 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line32 
            BorderColor     =   &H00000000&
            X1              =   105
            X2              =   105
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   150
            Left            =   0
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox sbUP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3285
         ScaleHeight     =   150
         ScaleWidth      =   225
         TabIndex        =   5
         ToolTipText     =   "Previous Message"
         Top             =   15
         Width           =   225
         Begin VB.Line Line48 
            BorderColor     =   &H00808080&
            X1              =   195
            X2              =   0
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line47 
            BorderColor     =   &H00808080&
            X1              =   195
            X2              =   195
            Y1              =   30
            Y2              =   135
         End
         Begin VB.Line Line46 
            BorderColor     =   &H00FFFFFF&
            X1              =   15
            X2              =   15
            Y1              =   15
            Y2              =   135
         End
         Begin VB.Line Line45 
            BorderColor     =   &H00FFFFFF&
            X1              =   15
            X2              =   210
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line31 
            BorderColor     =   &H00000000&
            X1              =   165
            X2              =   165
            Y1              =   90
            Y2              =   105
         End
         Begin VB.Line Line30 
            BorderColor     =   &H00000000&
            X1              =   150
            X2              =   150
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Line Line29 
            BorderColor     =   &H00000000&
            X1              =   135
            X2              =   135
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line28 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   120
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line27 
            BorderColor     =   &H00000000&
            X1              =   45
            X2              =   45
            Y1              =   90
            Y2              =   105
         End
         Begin VB.Line Line26 
            BorderColor     =   &H00000000&
            X1              =   60
            X2              =   60
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Line Line25 
            BorderColor     =   &H00000000&
            X1              =   75
            X2              =   75
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line24 
            BorderColor     =   &H00000000&
            X1              =   90
            X2              =   90
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line23 
            BorderColor     =   &H00000000&
            X1              =   105
            X2              =   105
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   150
            Left            =   0
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox sbUPc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3660
         ScaleHeight     =   150
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   225
         Begin VB.Line Line22 
            BorderColor     =   &H00FFFFFF&
            X1              =   195
            X2              =   195
            Y1              =   30
            Y2              =   135
         End
         Begin VB.Line Line21 
            BorderColor     =   &H00FFFFFF&
            X1              =   30
            X2              =   210
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line20 
            BorderColor     =   &H00808080&
            X1              =   15
            X2              =   15
            Y1              =   15
            Y2              =   135
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00808080&
            X1              =   15
            X2              =   210
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00000000&
            X1              =   165
            X2              =   165
            Y1              =   90
            Y2              =   105
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00000000&
            X1              =   150
            X2              =   150
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00000000&
            X1              =   135
            X2              =   135
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   120
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00000000&
            X1              =   45
            X2              =   45
            Y1              =   90
            Y2              =   105
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000000&
            X1              =   60
            X2              =   60
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00000000&
            X1              =   75
            X2              =   75
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            X1              =   90
            X2              =   90
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            X1              =   105
            X2              =   105
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   150
            Left            =   0
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox sbDN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3285
         ScaleHeight     =   150
         ScaleWidth      =   225
         TabIndex        =   3
         ToolTipText     =   "Next Message"
         Top             =   150
         Width           =   225
         Begin VB.Line Line52 
            BorderColor     =   &H00808080&
            X1              =   195
            X2              =   195
            Y1              =   105
            Y2              =   15
         End
         Begin VB.Line Line51 
            BorderColor     =   &H00808080&
            X1              =   30
            X2              =   210
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line50 
            BorderColor     =   &H00FFFFFF&
            X1              =   15
            X2              =   15
            Y1              =   15
            Y2              =   135
         End
         Begin VB.Line Line49 
            BorderColor     =   &H00FFFFFF&
            X1              =   15
            X2              =   210
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00000000&
            X1              =   105
            X2              =   105
            Y1              =   75
            Y2              =   120
         End
         Begin VB.Line Line15 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line16 
            BorderColor     =   &H00000000&
            X1              =   135
            X2              =   135
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00000000&
            X1              =   150
            X2              =   150
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Line Line18 
            BorderColor     =   &H00000000&
            X1              =   165
            X2              =   165
            Y1              =   45
            Y2              =   60
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00000000&
            X1              =   90
            X2              =   90
            Y1              =   60
            Y2              =   105
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00000000&
            X1              =   75
            X2              =   75
            Y1              =   45
            Y2              =   90
         End
         Begin VB.Line Line12 
            BorderColor     =   &H00000000&
            X1              =   60
            X2              =   60
            Y1              =   30
            Y2              =   75
         End
         Begin VB.Line Line14 
            BorderColor     =   &H00000000&
            X1              =   45
            X2              =   45
            Y1              =   45
            Y2              =   60
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   150
            Left            =   0
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   15
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   218
         TabIndex        =   1
         Top             =   15
         Width           =   3270
         Begin VB.Label lblMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Message"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   60
            TabIndex        =   2
            Top             =   45
            Width           =   3150
         End
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00808080&
         X1              =   234
         X2              =   234
         Y1              =   0
         Y2              =   21
      End
      Begin VB.Line Line54 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   21
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00808080&
         X1              =   1
         X2              =   234
         Y1              =   20
         Y2              =   20
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   235
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop"
      Begin VB.Menu mnuNext 
         Caption         =   "Next Message"
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "Previous Message"
      End
   End
End
Attribute VB_Name = "VSMessenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created  by Stjepan Vardic
'Send comments and suggestions to topctc2001@yahoo.co.uk

Option Explicit

Dim lstInt As Integer
Public Event UpClick()
Public Event DownClick()

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMessage.ToolTipText = "Message " & lstInt + 1 & " of " & lstMessage.ListCount
If lstInt <> lstMessage.ListCount - 1 And lstMessage.ListCount > 1 Then mnuNext.Enabled = True
If lstInt <> 0 And lstMessage.ListCount > 1 Then mnuPrev.Enabled = True

End Sub

Private Sub lblMessage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstInt = lstMessage.ListCount - 1 Then mnuNext.Enabled = False
If lstInt = 0 Then mnuPrev.Enabled = False
If Button = 2 Then UserControl.PopupMenu mnuPopUp

End Sub

Private Sub lstMessage_Click()
lblMessage.Caption = lstMessage.Text

End Sub

Private Sub mnuNext_Click()
lstInt = lstInt + 1
lblMessage.Caption = lstMessage.List(lstInt)
If lstInt > lstMessage.ListCount - 1 Then Beep: _
   lstInt = lstInt - 1: lblMessage.Caption = lstMessage.List(lstInt): Exit Sub

End Sub

Private Sub mnuPrev_Click()
lstInt = lstInt - 1
lblMessage.Caption = lstMessage.List(lstInt)
If lstInt < 0 Then Beep: lstInt = 0: _
   lblMessage.Caption = lstMessage.List(lstInt): Exit Sub

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then UserControl.PopupMenu mnuPopUp

End Sub

Private Sub sbDN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbDN.Visible = False
sbDNc.Left = sbDN.Left
sbDNc.Top = sbDN.Top
sbDNc.Visible = True

End Sub

Private Sub sbDN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent DownClick
sbDN.Visible = True
sbDNc.Visible = False
lstInt = lstInt + 1
lblMessage.Caption = lstMessage.List(lstInt)
If lstInt > lstMessage.ListCount - 1 Then Beep: _
   lstInt = lstInt - 1: lblMessage.Caption = lstMessage.List(lstInt): Exit Sub
End Sub

Private Sub sbUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbUP.Visible = False
sbUPc.Left = sbUP.Left
sbUPc.Top = sbUP.Top
sbUPc.Visible = True

End Sub

Private Sub sbUP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent UpClick
sbUP.Visible = True
sbUPc.Visible = False
lstInt = lstInt - 1
lblMessage.Caption = lstMessage.List(lstInt)
If lstInt < 0 Then Beep: lstInt = 0: _
   lblMessage.Caption = lstMessage.List(lstInt): Exit Sub

End Sub

Private Sub Timer1_Timer()
'lstInt = lstInt + 1
'lblMessage.Caption = lstMessage.List(lstInt)
'If lstInt > lstMessage.ListCount - 1 Then _
   lstInt = 0: lblMessage.Caption = lstMessage.List(lstInt)

End Sub

Private Sub UserControl_Initialize()
sbUPc.Move sbUP.Left, sbUP.Top
sbDNc.Move sbDN.Left, sbDN.Top

End Sub

Public Sub AddMessage(ByVal Item As String, Optional ByVal Index As Variant)
lstMessage.Additem Item, Index

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
ForeColor = PropBag.ReadProperty("ForeColor", &H0&)

End Sub

Private Sub UserControl_Resize()
UserControl.Height = 315
picMain.Move 0, 0, ScaleWidth, 315
sbUP.Move ScaleWidth - 16, 1
sbDN.Move ScaleWidth - 16, 10
Line56.X1 = ScaleWidth - 1
Line56.X2 = ScaleWidth - 1
Line55.X2 = Line56.X2
Line53.X2 = Line56.X1
Picture3.Width = ScaleWidth
lblMessage.Width = Picture3.Width - 25


End Sub

Private Sub UserControl_Show()
lblMessage.Caption = lstMessage.List(0)
lstInt = 0

End Sub
Public Property Get MessageCount() As Integer
Attribute MessageCount.VB_MemberFlags = "400"
MessageCount = lstMessage.ListCount

End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = lblMessage.ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
lblMessage.ForeColor = NewForeColor
PropertyChanged "ForeColor"

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "ForeColor", ForeColor, &H0&

End Sub

Public Sub CurrentMessage(ByVal MessageNumber As Integer)
lblMessage.Caption = lstMessage.List(MessageNumber)
lstInt = MessageNumber
If lstInt > lstMessage.ListCount - 1 Then lstInt = lstMessage.ListCount - 1: lblMessage.Caption = lstMessage.List(lstInt): Exit Sub
If lstInt < 0 Then lstInt = 0: lblMessage.Caption = lstMessage.List(lstInt): Exit Sub

End Sub

Public Property Get Message() As String
Message = lblMessage.Caption

End Property
